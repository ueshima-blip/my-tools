"""One-time surgery on 保護者会調整_ver200.xlsm (idempotent):

1. Remove the two ActiveX CommandButtons (modern-Office compatibility):
   - <controls> blocks in sheet1.xml (Start!) / sheet3.xml (希望入力)
   - control/image relationships, activeX parts, EMF fallback images
   - the t201 button shapes inside vmlDrawing1/2.vml (comments stay)
   - activeX/emf entries in [Content_Types].xml
2. Replace drawing1.xml / drawing2.xml with shape buttons wired to macros:
   - Start!   : 初期設定 / 希望を取り込む / クイック入力 / 調査票を印刷
   - 希望入力 : メニュー (copied to every 調整 sheet via SheetCopy)
3. Update on-sheet help text (sharedStrings):
   - △ is now used by the engine (2nd choice)
   - failure now reports the cause instead of asking to retry

Usage: python3 scripts/patch_workbook_xml.py
"""
import re
import shutil
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
XLSM = ROOT / "保護者会調整_ver200.xlsm"

BTN_TMPL = (
    '<xdr:twoCellAnchor editAs="oneCell">'
    '<xdr:from><xdr:col>{c1}</xdr:col><xdr:colOff>{c1o}</xdr:colOff>'
    '<xdr:row>{r1}</xdr:row><xdr:rowOff>{r1o}</xdr:rowOff></xdr:from>'
    '<xdr:to><xdr:col>{c2}</xdr:col><xdr:colOff>{c2o}</xdr:colOff>'
    '<xdr:row>{r2}</xdr:row><xdr:rowOff>{r2o}</xdr:rowOff></xdr:to>'
    '<xdr:sp macro="[0]!{macro}" textlink="">'
    '<xdr:nvSpPr><xdr:cNvPr id="{id}" name="{name}"/><xdr:cNvSpPr/></xdr:nvSpPr>'
    '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>'
    '<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>'
    '<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
    '<a:ln><a:solidFill><a:srgbClr val="{line}"/></a:solidFill></a:ln></xdr:spPr>'
    '<xdr:txBody>'
    '<a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="none" '
    'lIns="18000" tIns="9000" rIns="18000" bIns="9000" anchor="ctr"/>'
    '<a:lstStyle/><a:p><a:pPr algn="ctr"/>'
    '<a:r><a:rPr lang="ja-JP" sz="{sz}" b="1">'
    '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr>'
    '<a:t>{text}</a:t></a:r></a:p></xdr:txBody></xdr:sp>'
    '<xdr:clientData fPrintsWithSheet="0"/></xdr:twoCellAnchor>'
)

WSDR_OPEN = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
)


def start_buttons():
    """Start! sheet: E列の説明文が無い行帯に配置（セル文字との重なり実測済み）。

    シート行ベースの空き帯: 3-5 / 7-9 / 22-24 / 29-31
    （0-based drawing rows: 2-4 / 6-8 / 21-23 / 28-30）
    横位置は E列内 x≈300px〜443px（旧ボタンより左へ広げて約150px幅を確保）。
    """
    btns = []
    rows = [(2, 60000, 4, 120000),       # シート3-5行: 取込（回収後の入口・目立つ位置）
            (6, 180975, 8, 104775),      # シート7-9行: 初期設定（旧ボタンと同じ・ラベル隣）
            (21, 60000, 23, 120000),     # シート22-24行: クイック入力
            (28, 60000, 30, 120000)]     # シート29-31行: 調査票を印刷
    spec = [("BtnImport", "希望を取り込む", "548235", "375623"),
            ("BtnSettings", "初期設定", "4472C4", "2F528F"),
            ("BtnQuickInput", "クイック入力", "548235", "375623"),
            ("BtnSurveyPrint", "調査票を印刷", "BF8F00", "806000")]
    for i, ((r1, r1o, r2, r2o), (macro, text, fill, line)) in enumerate(zip(rows, spec)):
        btns.append(BTN_TMPL.format(
            c1=4, c1o=2855000, r1=r1, r1o=r1o,
            c2=4, c2o=4219575, r2=r2, r2o=r2o,
            macro=macro, id=i + 2, name=macro, fill=fill, line=line,
            sz=1000, text=text))
    return WSDR_OPEN + "".join(btns) + "</xdr:wsDr>"


def kibou_buttons():
    """希望入力 sheet: メニュー button at the old A1 position."""
    btn = BTN_TMPL.format(
        c1=0, c1o=76200, r1=0, r1o=19050,
        c2=2, c2o=447675, r2=0, r2o=323850,
        macro="BtnMenu", id=2, name="BtnMenu", fill="4472C4", line="2F528F",
        sz=1000, text="メニュー")
    return WSDR_OPEN + btn + "</xdr:wsDr>"


def remove_controls_block(xml):
    out = re.sub(r"<controls>.*?</controls>", "", xml, flags=re.S)
    assert "<controls>" not in out
    return out


def remove_rels(xml, targets):
    for t in targets:
        xml = re.sub(r'<Relationship [^>]*Target="' + re.escape(t) + r'"[^>]*/>', "", xml)
        assert t not in xml
    return xml


def remove_vml_button(xml):
    """Drop the t201 shape (the ActiveX button) and its shapetype."""
    out = re.sub(r"<v:shape id=\"CommandButton1\".*?</v:shape>", "", xml, flags=re.S)
    out = re.sub(r"<v:shapetype id=\"_x0000_t201\".*?</v:shapetype>", "", out, flags=re.S)
    assert "_x0000_t201" not in out
    return out


def patch_shared_strings(xml):
    sis = re.findall(r"<si>.*?</si>", xml, flags=re.S)
    new_sankaku = ("<si><t>第2希望の\"△\"は、○だけでは調整できない場合に限り自動で使われます。"
                   "（△で確定した枠も青色になります）</t></si>")
    new_fail = ("<si><t>調整できない場合は、原因となる生徒が赤色表示され、理由（希望の重なり）"
                "が表示されます。希望(○/△)を増やすか、休憩・◎を見直してください。</t></si>")
    old_sankaku = [s for s in sis if "単なる覚え" in s]
    old_fail = [s for s in sis if "確定順の設定" in s]
    assert len(old_sankaku) == 1 and len(old_fail) == 1, "si lookup failed"
    xml = xml.replace(old_sankaku[0], new_sankaku, 1)
    xml = xml.replace(old_fail[0], new_fail, 1)
    return xml


def patch_content_types(xml):
    xml = re.sub(r'<Override PartName="/xl/activeX/[^"]*"[^>]*/>', "", xml)
    xml = re.sub(r'<Default Extension="emf"[^>]*/>', "", xml)
    assert "activeX" not in xml
    return xml


def main():
    with zipfile.ZipFile(XLSM) as z:
        files = {n: z.read(n) for n in z.namelist()}

    dec = lambda n: files[n].decode("utf-8")
    first_time = "xl/activeX/activeX1.xml" in files

    if first_time:
        # 1. worksheets
        files["xl/worksheets/sheet1.xml"] = remove_controls_block(dec("xl/worksheets/sheet1.xml")).encode("utf-8")
        files["xl/worksheets/sheet3.xml"] = remove_controls_block(dec("xl/worksheets/sheet3.xml")).encode("utf-8")

        # 2. rels
        files["xl/worksheets/_rels/sheet1.xml.rels"] = remove_rels(
            dec("xl/worksheets/_rels/sheet1.xml.rels"),
            ["../activeX/activeX1.xml", "../media/image1.emf"]).encode("utf-8")
        files["xl/worksheets/_rels/sheet3.xml.rels"] = remove_rels(
            dec("xl/worksheets/_rels/sheet3.xml.rels"),
            ["../activeX/activeX2.xml", "../media/image2.emf"]).encode("utf-8")

        # 3. VML (keep comment shapes)
        files["xl/drawings/vmlDrawing1.vml"] = remove_vml_button(dec("xl/drawings/vmlDrawing1.vml")).encode("utf-8")
        files["xl/drawings/vmlDrawing2.vml"] = remove_vml_button(dec("xl/drawings/vmlDrawing2.vml")).encode("utf-8")

        # 5. delete activeX parts & images & their rels
        for n in list(files):
            if n.startswith("xl/activeX/") or n in ("xl/media/image1.emf", "xl/media/image2.emf"):
                del files[n]

        # 6. content types + shared strings
        files["[Content_Types].xml"] = patch_content_types(dec("[Content_Types].xml")).encode("utf-8")
        files["xl/sharedStrings.xml"] = patch_shared_strings(dec("xl/sharedStrings.xml")).encode("utf-8")

    # 4. drawings → shape buttons（レイアウト変更時は再実行すればここだけ更新される）
    files["xl/drawings/drawing1.xml"] = start_buttons().encode("utf-8")
    files["xl/drawings/drawing2.xml"] = kibou_buttons().encode("utf-8")

    tmp = XLSM.with_suffix(".xlsm.tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
        for n, data in files.items():
            z.writestr(n, data)
    shutil.move(tmp, XLSM)
    if first_time:
        print(f"✓ patched {XLSM.name}: ActiveX removed, shape buttons added, help text updated")
    else:
        print(f"✓ {XLSM.name}: shape buttons regenerated (drawing1/drawing2)")


if __name__ == "__main__":
    main()
