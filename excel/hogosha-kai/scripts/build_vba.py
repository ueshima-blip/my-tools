"""Build the .xlsm from edited vba/*.bas|.cls|.frm sources (any size).

Pipeline per module:
  1. read vba/<name>.<ext> (cp932, CRLF)
  2. MS-OVBA compress
  3. new module stream = zero-filled PerformanceCache prefix (original
     textoffset bytes) + compressed source

Then the whole vbaProject.bin OLE container is rebuilt (scripts/cfb.py), the
__SRP_* caches are zero-filled and the _VBA_PROJECT version field cleared so
Excel recompiles everything from source on next open, and xl/vbaProject.bin
is swapped into the xlsm zip.

Usage:
    python3 scripts/build_vba.py [path/to/file.xlsm]
"""
import io
import json
import shutil
import sys
import zipfile
from pathlib import Path

import olefile
from oletools.olevba import decompress_stream

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from ms_ovba_compress import compress_stream
from cfb import repack, CfbFile


def source_bytes(path: Path) -> bytes:
    """Repo sources are UTF-8/LF → VBA wants cp932/CRLF."""
    text = path.read_text(encoding="utf-8")
    text = text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
    return text.encode("cp932")


def main():
    root = HERE.parent
    meta = json.loads((root / "vba" / "_meta.json").read_text(encoding="utf-8"))
    xlsm = root / (sys.argv[1] if len(sys.argv) > 1 else meta["xlsm_path"])

    with zipfile.ZipFile(xlsm) as z:
        names = z.namelist()
        original_bin = z.read("xl/vbaProject.bin")

    src_cfb = CfbFile(original_bin)
    replacements = {}
    for m in meta["modules"]:
        src_path = root / m["file"]
        if not src_path.exists():
            print(f"  ! {m['name']}: {m['file']} missing, kept as-is")
            continue
        raw = source_bytes(src_path)
        compressed = compress_stream(raw)
        # verify round trip before committing
        assert bytes(decompress_stream(bytearray(compressed))) == raw, m["name"]
        prefix = b"\x00" * m["textoffset"]
        replacements[f"VBA/{m['name']}"] = prefix + compressed
        print(f"  ✓ {m['name']}: src {len(raw)} B → stream {len(prefix) + len(compressed)} B")

    # invalidate compiled caches → force recompile from source
    for p in src_cfb.stream_paths():
        if p.startswith("VBA/__SRP_"):
            replacements[p] = b"\x00" * len(src_cfb.read_stream(p))
    vp = bytearray(src_cfb.read_stream("VBA/_VBA_PROJECT"))
    vp[2] = vp[3] = 0
    replacements["VBA/_VBA_PROJECT"] = bytes(vp)

    new_bin = repack(original_bin, replacements)

    # paranoia: parse result and confirm every module decompresses to its source
    chk = olefile.OleFileIO(io.BytesIO(new_bin))
    for m in meta["modules"]:
        src_path = root / m["file"]
        if not src_path.exists():
            continue
        stream = chk.openstream(f"VBA/{m['name']}").read()
        got = bytes(decompress_stream(bytearray(stream[m["textoffset"]:])))
        assert got == source_bytes(src_path), f"verify failed: {m['name']}"
    chk.close()

    tmp = xlsm.with_suffix(".xlsm.tmp")
    with zipfile.ZipFile(xlsm) as zin, \
            zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as zout:
        for n in names:
            zout.writestr(n, new_bin if n == "xl/vbaProject.bin" else zin.read(n))
    shutil.move(tmp, xlsm)
    print(f"✓ built {xlsm.name} (vbaProject.bin {len(original_bin)} → {len(new_bin)} B)")


if __name__ == "__main__":
    main()
