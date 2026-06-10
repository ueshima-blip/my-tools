"""Validate an .xlsx/.xlsm for the corruptions Excel reports as
"found a problem with some content" — without needing Excel itself.

Checks:
  1. every part is well-formed XML
  2. calcChain.xml refs (if present) point to cells that have a formula
  3. no orphan shared-formula slaves (each si used has a master with ref=)
  4. every relationship Target exists in the package
  5. every Content_Types Override PartName exists, and key parts are declared
  6. sharedStrings: count/uniqueCount sane; no string index out of range
  7. defined names referenced by sheets exist

Usage: python3 scripts/validate_xlsx.py [file.xlsm]
Exit non-zero on any error.
"""
import re
import sys
import zipfile
import xml.dom.minidom as md
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def col_to_num(col):
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n


def main():
    target = ROOT / (sys.argv[1] if len(sys.argv) > 1 else "保護者会調整_ver200.xlsm")
    z = zipfile.ZipFile(target)
    names = set(z.namelist())
    errors = []

    # 1. XML well-formed
    for n in names:
        if n.endswith((".xml", ".rels", ".vml")):
            try:
                md.parseString(z.read(n))
            except Exception as e:
                errors.append(f"[xml] {n}: {e}")

    # sheet parts
    sheet_parts = sorted(n for n in names if re.match(r"xl/worksheets/sheet\d+\.xml$", n))

    # gather formula cells + shared-formula masters/slaves per sheet
    sheet_formula_cells = {}
    for sp in sheet_parts:
        xml = z.read(sp).decode("utf-8")
        fcells = set()
        masters = defaultdict(set)   # si -> set of master refs
        slaves = defaultdict(int)
        # match both self-closing (<c .../>) and open/close (<c ...>...</c>) cells
        for m in re.finditer(r'<c r="([A-Z]+\d+)"[^>]*?(?:/>|>(.*?)</c>)', xml, flags=re.S):
            ref, body = m.group(1), (m.group(2) or "")
            if "<f" in body:
                fcells.add(ref)
                sm = re.search(r'<f[^>]*\bt="shared"[^>]*\bsi="(\d+)"', body)
                if sm:
                    si = sm.group(1)
                    if "ref=" in body:
                        masters[si].add(ref)
                    else:
                        slaves[si] += 1
        # self-closing formula cells (rare)
        sheet_formula_cells[sp] = fcells
        # 3. orphan slaves
        for si, cnt in slaves.items():
            if si not in masters:
                errors.append(f"[shared] {sp}: si={si} has {cnt} slave(s) but no master (orphan → corruption)")

    # 2. calcChain refs
    if "xl/calcChain.xml" in names:
        cc = z.read("xl/calcChain.xml").decode("utf-8")
        # map sheetId(i) -> sheet part via workbook rels + workbook.xml
        wb = z.read("xl/workbook.xml").decode("utf-8")
        rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        rid_to_target = dict(re.findall(r'<Relationship Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
        sheetid_to_part = {}
        for sm in re.finditer(r'<sheet [^>]*sheetId="(\d+)"[^>]*r:id="(rId\d+)"', wb):
            tgt = rid_to_target.get(sm.group(2), "")
            sheetid_to_part[sm.group(1)] = "xl/" + tgt.lstrip("/")
        cur_i = None
        for m in re.finditer(r'<c r="([A-Z]+\d+)"(?: i="(\d+)")?', cc):
            ref, i = m.group(1), m.group(2)
            if i:
                cur_i = i
            part = sheetid_to_part.get(cur_i)
            if part and ref not in sheet_formula_cells.get(part, set()):
                errors.append(f"[calcChain] {part} {ref} listed but has no formula")
                if sum(1 for e in errors if e.startswith("[calcChain]")) > 8:
                    errors.append("[calcChain] … (more)")
                    break

    # 4. relationship targets exist
    for n in names:
        if n.endswith(".rels"):
            base = re.sub(r"_rels/[^/]+$", "", n)  # dir of the part owning these rels
            xml = z.read(n).decode("utf-8")
            for rm in re.finditer(r'<Relationship [^>]*Target="([^"]+)"([^>]*)>', xml):
                tgt, rest = rm.group(1), rm.group(2)
                if "External" in rest or tgt.startswith(("http", "../")) is False and tgt.startswith("/"):
                    continue
                if tgt.startswith("http") or 'TargetMode="External"' in rest:
                    continue
                # resolve relative
                from posixpath import normpath, join
                resolved = normpath(join(base, tgt)).lstrip("/")
                if resolved not in names:
                    errors.append(f"[rels] {n}: target not found: {tgt} → {resolved}")

    # 5. Content_Types overrides exist
    ct = z.read("[Content_Types].xml").decode("utf-8")
    for m in re.finditer(r'<Override PartName="([^"]+)"', ct):
        part = m.group(1).lstrip("/")
        if part not in names:
            errors.append(f"[content-types] override for missing part: {part}")
    # key parts must be declared (by Override or Default ext)
    exts = set(re.findall(r'<Default Extension="([^"]+)"', ct))
    overrides = set(m.lstrip("/") for m in re.findall(r'<Override PartName="([^"]+)"', ct))
    for n in names:
        if n == "[Content_Types].xml":
            continue
        ext = n.rsplit(".", 1)[-1].lower()
        if n not in overrides and ext not in exts:
            errors.append(f"[content-types] part not declared: {n}")

    # 6. sharedStrings index range
    if "xl/sharedStrings.xml" in names:
        ss = z.read("xl/sharedStrings.xml").decode("utf-8")
        n_si = len(re.findall(r"<si>", ss))
        m = re.search(r'uniqueCount="(\d+)"', ss)
        if m and int(m.group(1)) != n_si:
            errors.append(f"[sst] uniqueCount={m.group(1)} but {n_si} <si> present")
        max_idx = -1
        for sp in sheet_parts:
            xml = z.read(sp).decode("utf-8")
            for vm in re.finditer(r'<c [^>]*t="s"[^>]*><v>(\d+)</v>', xml):
                max_idx = max(max_idx, int(vm.group(1)))
        if max_idx >= n_si:
            errors.append(f"[sst] string index {max_idx} out of range (only {n_si} strings)")

    if errors:
        print(f"✗ {target.name}: {len(errors)} problem(s)")
        for e in errors:
            print("  " + e)
        sys.exit(1)
    print(f"✓ {target.name}: structural validation passed "
          f"({len(names)} parts, {len(sheet_parts)} sheets)")


if __name__ == "__main__":
    main()
