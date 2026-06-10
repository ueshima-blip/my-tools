"""Extract VBA module sources from the .xlsm into vba/*.bas (.cls) text files.

Usage:
    python3 scripts/extract_vba.py [path/to/file.xlsm]

Writes vba/<ModuleName>.<ext> (cp932) and vba/_meta.json (textoffset = size of
the PerformanceCache prefix that build_vba.py must keep, zero-filled).
Unlike the shindo-yotei tooling there is NO size constraint: build_vba.py
rebuilds the OLE container, so sources may grow freely.
"""
import io
import json
import struct
import sys
from pathlib import Path
import zipfile

import olefile
from oletools.olevba import decompress_stream

DOC_MODULES = {"ThisWorkbook", "Sheet1", "Sheet2", "Sheet3", "Sheet5"}
FORM_MODULES = {"UserForm1", "UserForm2", "UserForm3", "UserForm4"}

DEFAULT_XLSM = "保護者会調整_ver200.xlsm"


def parse_dir(dir_data):
    modules = []
    cur = None
    i = 0
    while i < len(dir_data) - 6:
        if dir_data[i:i + 2] == b"\x19\x00":
            rec_size = struct.unpack("<I", dir_data[i + 2:i + 6])[0]
            if 0 < rec_size < 100:
                name = dir_data[i + 6:i + 6 + rec_size].decode("cp932", errors="replace")
                cur = {"name": name}
                modules.append(cur)
        elif dir_data[i:i + 2] == b"\x31\x00":
            rec_size = struct.unpack("<I", dir_data[i + 2:i + 6])[0]
            if rec_size == 4 and cur is not None:
                cur["offset"] = struct.unpack("<I", dir_data[i + 6:i + 10])[0]
        i += 1
    return modules


def main():
    here = Path(__file__).resolve().parent.parent
    xlsm = here / (sys.argv[1] if len(sys.argv) > 1 else DEFAULT_XLSM)
    vba_dir = here / "vba"
    vba_dir.mkdir(exist_ok=True)

    with zipfile.ZipFile(xlsm) as z:
        vba_bin = z.read("xl/vbaProject.bin")
    ole = olefile.OleFileIO(io.BytesIO(vba_bin))
    dir_data = decompress_stream(bytearray(ole.openstream("VBA/dir").read()))
    modules = parse_dir(dir_data)

    meta = {"xlsm_path": xlsm.name, "modules": []}
    for m in modules:
        name, off = m["name"], m.get("offset")
        if off is None:
            print(f"  ! {name}: no textoffset, skipped")
            continue
        stream = ole.openstream(f"VBA/{name}").read()
        src = decompress_stream(bytearray(stream[off:])).decode("cp932", errors="replace")
        ext = "cls" if name in DOC_MODULES else ("frm" if name in FORM_MODULES else "bas")
        out = vba_dir / f"{name}.{ext}"
        # repo files are UTF-8 (LF); build_vba.py converts to cp932/CRLF
        with open(out, "w", encoding="utf-8", newline="\n") as f:
            f.write(src.replace("\r\n", "\n"))
        meta["modules"].append({"name": name, "file": f"vba/{name}.{ext}",
                                "textoffset": off})
        print(f"  ✓ {name}: {len(src)} chars → {out.name}")
    ole.close()

    with open(vba_dir / "_meta.json", "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)
    print("✓ wrote vba/_meta.json")


if __name__ == "__main__":
    main()
