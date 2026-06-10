"""Compound File Binary (CFB / OLE2) reader + repacker.

Purpose: rewrite streams inside xl/vbaProject.bin with *different sizes*,
which olefile cannot do (its write_stream requires identical size).

Strategy (lowest risk): keep the directory-entry array byte-identical in
order, names, types, tree pointers (left/right/child), CLSIDs and times —
only each entry's start-sector and size fields are updated. The FAT,
miniFAT, DIFAT and sector layout are rebuilt from scratch. Because no
entries are added, removed or reordered, the red-black tree Excel wrote
stays exactly as-is.

Limitations (fine for vbaProject.bin): version 3 files only (512-byte
sectors), DIFAT must fit the header (≤109 FAT sectors ≈ 6.9 MB file).
"""
import struct

ENDOFCHAIN = 0xFFFFFFFE
FATSECT = 0xFFFFFFFD
FREESECT = 0xFFFFFFFF
NOSTREAM = 0xFFFFFFFF

SECT = 512
MINISECT = 64
MINI_CUTOFF = 4096


class Entry:
    __slots__ = ("raw", "name", "obj_type", "left", "right", "child",
                 "start", "size", "eid")

    def __init__(self, raw, eid):
        self.raw = bytearray(raw)
        self.eid = eid
        name_len = struct.unpack_from("<H", raw, 64)[0]
        self.name = raw[: max(0, name_len - 2)].decode("utf-16-le") if name_len >= 2 else ""
        self.obj_type = raw[66]
        self.left, self.right, self.child = struct.unpack_from("<III", raw, 68)
        self.start = struct.unpack_from("<I", raw, 116)[0]
        self.size = struct.unpack_from("<Q", raw, 120)[0]

    def packed(self, new_start, new_size):
        out = bytearray(self.raw)
        struct.pack_into("<I", out, 116, new_start)
        struct.pack_into("<Q", out, 120, new_size)
        return bytes(out)


class CfbFile:
    def __init__(self, data: bytes):
        if data[:8] != b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
            raise ValueError("not a CFB file")
        (self.minor, self.major) = struct.unpack_from("<HH", data, 24)
        if self.major != 3:
            raise ValueError(f"only v3 (512-byte sectors) supported, got v{self.major}")
        self.data = data
        (self.n_fat,) = struct.unpack_from("<I", data, 44)
        (self.first_dir,) = struct.unpack_from("<I", data, 48)
        (self.first_minifat, self.n_minifat) = struct.unpack_from("<II", data, 60)
        (self.first_difat, self.n_difat) = struct.unpack_from("<II", data, 68)

        # DIFAT → list of FAT sector ids
        fat_sects = list(struct.unpack_from("<109I", data, 76))
        difat_sect = self.first_difat
        while difat_sect not in (ENDOFCHAIN, FREESECT):
            raw = self._sector(difat_sect)
            ids = struct.unpack("<128I", raw)
            fat_sects.extend(ids[:127])
            difat_sect = ids[127]
        self.fat = []
        for s in fat_sects:
            if s in (FREESECT, ENDOFCHAIN):
                continue
            self.fat.extend(struct.unpack("<128I", self._sector(s)))

        # directory entries
        dir_bytes = b"".join(self._sector(s) for s in self._chain(self.first_dir))
        self.entries = [Entry(dir_bytes[i: i + 128], i // 128)
                        for i in range(0, len(dir_bytes), 128)]
        self.root = next(e for e in self.entries if e.obj_type == 5)

        # miniFAT + ministream
        self.minifat = []
        for s in self._chain(self.first_minifat):
            self.minifat.extend(struct.unpack("<128I", self._sector(s)))
        self.ministream = b"".join(self._sector(s) for s in self._chain(self.root.start))

        # path map
        self.paths = {}
        self._walk(self.root, ())

    def _sector(self, sid):
        off = 512 + sid * SECT
        return self.data[off: off + SECT]

    def _chain(self, start):
        out, sid, seen = [], start, set()
        while sid != ENDOFCHAIN and sid != FREESECT:
            if sid in seen:
                raise ValueError("FAT chain loop")
            seen.add(sid)
            out.append(sid)
            sid = self.fat[sid]
        return out

    def _walk(self, entry, prefix):
        if entry.child != NOSTREAM:
            stack = [self.entries[entry.child]]
            while stack:
                e = stack.pop()
                p = prefix + (e.name,)
                self.paths[p] = e
                for sib in (e.left, e.right):
                    if sib != NOSTREAM:
                        stack.append(self.entries[sib])
                if e.obj_type == 1 and e.child != NOSTREAM:  # storage
                    self._walk(e, p)

    def read_stream(self, path):
        e = self.paths[tuple(path.split("/"))] if isinstance(path, str) else self.paths[tuple(path)]
        if e.size < MINI_CUTOFF and e is not self.root:
            out, sid = b"", e.start
            chunks = []
            while sid != ENDOFCHAIN and sid != FREESECT:
                chunks.append(self.ministream[sid * MINISECT:(sid + 1) * MINISECT])
                sid = self.minifat[sid]
            out = b"".join(chunks)
        else:
            out = b"".join(self._sector(s) for s in self._chain(e.start))
        return out[: e.size]

    def stream_paths(self):
        return ["/".join(p) for p, e in self.paths.items() if e.obj_type == 2]


def repack(original: bytes, replacements: dict) -> bytes:
    """Rebuild the CFB with some streams replaced (any sizes).

    replacements: {"VBA/Module1": b"...", ...}
    Directory topology and entry metadata are preserved verbatim;
    only start/size fields and the allocation structures are rebuilt.
    """
    src = CfbFile(original)
    repl = {tuple(k.split("/")): v for k, v in replacements.items()}
    unknown = set(repl) - set(src.paths)
    if unknown:
        raise KeyError(f"streams not in file: {unknown}")

    # final content for every stream entry (by entry id)
    content = {}
    for path, e in src.paths.items():
        if e.obj_type != 2:
            continue
        content[e.eid] = repl.get(path, src.read_stream(path))

    # --- split into mini / regular streams (root ministream excluded)
    mini_ids = [eid for eid, c in content.items() if len(c) < MINI_CUTOFF and len(c) > 0]
    reg_ids = [eid for eid, c in content.items() if len(c) >= MINI_CUTOFF]
    # keep original entry order for stable output
    mini_ids.sort()
    reg_ids.sort()

    # --- build ministream + miniFAT
    ministream = bytearray()
    minifat = []
    mini_start = {}
    for eid in mini_ids:
        c = content[eid]
        n = (len(c) + MINISECT - 1) // MINISECT
        first = len(minifat)
        mini_start[eid] = first
        for i in range(n):
            minifat.append(first + i + 1 if i < n - 1 else ENDOFCHAIN)
        ministream.extend(c)
        pad = (-len(c)) % MINISECT
        ministream.extend(b"\x00" * pad)
    while len(minifat) % 128:
        minifat.append(FREESECT)
    n_minifat_sects = len(minifat) // 128 if minifat else 0

    # --- regular sector plan
    def nsect(nbytes):
        return (nbytes + SECT - 1) // SECT

    n_dir_sects = nsect(len(src.entries) * 128)
    n_ministream_sects = nsect(len(ministream))
    n_reg = {eid: nsect(len(content[eid])) for eid in reg_ids}
    other_sects = (n_dir_sects + n_minifat_sects + n_ministream_sects
                   + sum(n_reg.values()))

    n_fat = 0
    for _ in range(8):  # fixpoint: FAT sectors are themselves FAT-tracked
        total = other_sects + n_fat
        need = (total + 127) // 128
        if need == n_fat:
            break
        n_fat = need
    if n_fat > 109:
        raise ValueError("file too large for header-only DIFAT")

    # layout: [FAT][DIR][miniFAT][ministream][regular...]
    fat = []
    layout = []  # (kind, eid_or_None, first_sector, count)

    def alloc(count, kind, eid=None, fat_value_chain=True):
        first = len(fat)
        for i in range(count):
            if fat_value_chain:
                fat.append(first + i + 1 if i < count - 1 else ENDOFCHAIN)
            else:
                fat.append(FATSECT)
        layout.append((kind, eid, first, count))
        return first

    fat_first = alloc(n_fat, "fat", fat_value_chain=False)
    dir_first = alloc(n_dir_sects, "dir")
    minifat_first = alloc(n_minifat_sects, "minifat") if n_minifat_sects else ENDOFCHAIN
    ministream_first = alloc(n_ministream_sects, "ministream") if n_ministream_sects else ENDOFCHAIN
    reg_start = {}
    for eid in reg_ids:
        reg_start[eid] = alloc(n_reg[eid], "stream", eid)

    total_sects = len(fat)
    while len(fat) % 128:
        fat.append(FREESECT)

    # --- directory bytes (preserve everything but start/size)
    dir_out = bytearray()
    for e in src.entries:
        if e.obj_type == 5:  # root: owns the ministream
            dir_out += e.packed(
                ministream_first if n_ministream_sects else ENDOFCHAIN,
                len(ministream))
        elif e.obj_type == 2:
            c = content[e.eid]
            if len(c) == 0:
                dir_out += e.packed(ENDOFCHAIN, 0)
            elif len(c) < MINI_CUTOFF:
                dir_out += e.packed(mini_start[e.eid], len(c))
            else:
                dir_out += e.packed(reg_start[e.eid], len(c))
        else:  # storage / unused: keep raw bytes
            dir_out += bytes(e.raw)
    dir_out += b"\x00" * ((-len(dir_out)) % SECT)

    # --- header
    header = bytearray(512)
    header[0:8] = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
    struct.pack_into("<HH", header, 24, 0x003E, 0x0003)   # minor, major
    struct.pack_into("<H", header, 28, 0xFFFE)            # little-endian
    struct.pack_into("<HH", header, 30, 9, 6)             # sector/minisector shift
    struct.pack_into("<I", header, 44, n_fat)
    struct.pack_into("<I", header, 48, dir_first)
    struct.pack_into("<I", header, 56, MINI_CUTOFF)
    struct.pack_into("<I", header, 60, minifat_first if n_minifat_sects else ENDOFCHAIN)
    struct.pack_into("<I", header, 64, n_minifat_sects)
    struct.pack_into("<I", header, 68, ENDOFCHAIN)        # no DIFAT sectors
    struct.pack_into("<I", header, 72, 0)
    for i in range(109):
        struct.pack_into("<I", header, 76 + 4 * i,
                         fat_first + i if i < n_fat else FREESECT)

    # --- body
    body = bytearray(total_sects * SECT)

    def put(first, blob):
        body[first * SECT: first * SECT + len(blob)] = blob

    fat_bytes = b"".join(struct.pack("<I", v) for v in fat)
    put(fat_first, fat_bytes)
    put(dir_first, bytes(dir_out))
    if n_minifat_sects:
        put(minifat_first, b"".join(struct.pack("<I", v) for v in minifat))
    if n_ministream_sects:
        put(ministream_first, bytes(ministream))
    for eid in reg_ids:
        put(reg_start[eid], content[eid])

    return bytes(header) + bytes(body)


if __name__ == "__main__":
    import sys
    # self test: repack with no changes and verify all streams match
    path = sys.argv[1]
    with open(path, "rb") as f:
        original = f.read()
    rebuilt = repack(original, {})
    a, b = CfbFile(original), CfbFile(rebuilt)
    assert sorted(a.stream_paths()) == sorted(b.stream_paths())
    for p in a.stream_paths():
        if a.read_stream(p) != b.read_stream(p):
            raise SystemExit(f"MISMATCH: {p}")
    print(f"self-test OK: {len(a.stream_paths())} streams, "
          f"{len(original)} → {len(rebuilt)} bytes")
