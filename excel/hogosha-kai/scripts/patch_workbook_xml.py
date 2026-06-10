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


# ------------------------------------------------------------------
# 氏名1セル化 + 時刻ヘッダの「縮小して全体を表示」（常時実行・冪等）
# ------------------------------------------------------------------

# 希望入力シート 2行目（時刻ヘッダ）が使う書式番号。numFmtId=20(h:mm)・中央揃え。
TIME_HEADER_XFS = (86, 87, 96, 102, 107, 112, 117, 125)


def ensure_shared_string(xml, text):
    """sharedStrings に text の <si> が無ければ追加し、(xml, index) を返す。"""
    sis = re.findall(r"<si>.*?</si>", xml, flags=re.S)
    needle = f"<si><t>{text}</t></si>"
    for i, s in enumerate(sis):
        if s == needle:
            return xml, i
    idx = len(sis)
    xml = xml.replace("</sst>", needle + "</sst>")
    def bump(m):
        return f'count="{int(m.group(1)) + 2}" uniqueCount="{int(m.group(2)) + 1}"'
    xml = re.sub(r'count="(\d+)" uniqueCount="(\d+)"', bump, xml, count=1)
    return xml, idx


def patch_shrink_time_xfs(xml):
    """時刻ヘッダの xf に shrinkToFit を追加（#### 対策）。"""
    m = re.search(r"(<cellXfs count=\"\d+\">)(.*?)(</cellXfs>)", xml, flags=re.S)
    body = m.group(2)
    xfs = re.findall(r"<xf [^>]*?/>|<xf [^>]*?>.*?</xf>", body, flags=re.S)
    for i in TIME_HEADER_XFS:
        xf = xfs[i]
        if "shrinkToFit" in xf:
            continue
        assert 'numFmtId="20"' in xf, f"xf[{i}] is not a time format"
        xfs[i] = xf.replace("<alignment ", '<alignment shrinkToFit="1" ', 1)
    return xml[: m.start(2)] + "".join(xfs) + xml[m.end(2):]


def patch_meibo_sheet(xml, name_idx):
    """名簿: B1=氏名(1セル), C1は空に。B列を広げる。"""
    xml = re.sub(r'<c r="B1"( s="\d+")? t="s"><v>\d+</v></c>',
                 f'<c r="B1"\\1 t="s"><v>{name_idx}</v></c>', xml)
    xml = re.sub(r'<c r="C1"( s="\d+")? t="s"><v>\d+</v></c>', r'<c r="C1"\1/>', xml)
    if '<col min="2"' not in xml:
        xml = xml.replace('<col min="4"',
                          '<col min="2" max="2" width="14.5" customWidth="1"/><col min="4"', 1)
    return xml


def patch_kibou_sheet(xml, name_idx):
    """希望入力: B2=氏名, C2空, C3:C52 の VLOOKUP(…,3,0) を撤去, B列を広げる。"""
    xml = re.sub(r'<c r="B2"( s="\d+")? t="s"><v>\d+</v></c>',
                 f'<c r="B2"\\1 t="s"><v>{name_idx}</v></c>', xml)
    xml = re.sub(r'<c r="C2"( s="\d+")? t="s"><v>\d+</v></c>', r'<c r="C2"\1/>', xml)
    for row in range(3, 53):
        xml = re.sub(
            r'<c r="C%d"( s="\d+")?[^>]*>.*?</c>' % row,
            r'<c r="C%d"\1/>' % row, xml, flags=re.S)
    xml = xml.replace(
        '<col min="2" max="3" width="6.25" customWidth="1"/>',
        '<col min="2" max="2" width="12.5" customWidth="1"/>'
        '<col min="3" max="3" width="2.25" customWidth="1"/>')
    return xml


def patch_yotei_sheet(xml):
    """元_予定表: 「名」列(F/I/L/O/R)の数式を完全撤去して空セルに。

    1セル氏名では「氏」列(E/H/K/N/Q)の VLOOKUP(…,2,0) がフルネームを返すので、
    「名」列は不要。共有数式のマスターだけ消すとスレーブが孤児化して破損するため、
    マスター・スレーブとも丸ごと空セル化する。氏列を広げ名列を狭くする。
    """
    def strip(m):
        return f'<c r="{m.group(1)}"{m.group(2) or ""}/>'
    xml = re.sub(
        r'<c r="([FILOR]\d+)"( s="\d+")?(?: t="[^"]*")?>\s*'
        r'<f[^>]*?(?:/>|>.*?</f>)\s*(?:<v\s*/>|<v>.*?</v>)?\s*</c>',
        strip, xml, flags=re.S)
    assert ",氏名,3,0" not in xml
    for first, second in ((5, 6), (8, 9), (11, 12), (14, 15), (17, 18)):
        xml = xml.replace(
            f'<col min="{first}" max="{second}" width="8" style="4" customWidth="1"/>',
            f'<col min="{first}" max="{first}" width="12.5" style="4" customWidth="1"/>'
            f'<col min="{second}" max="{second}" width="3.5" style="4" customWidth="1"/>')
    return xml


def prune_dangling_rels(files):
    """すべての .rels から「存在しないパートを指す関係」を除去する。

    ActiveX/EMF を消したのに vmlDrawing*.vml.rels が画像を指したまま、等の
    取りこぼしを最終段でまとめて掃除する（Excel の破損検出を防ぐ）。
    外部参照(TargetMode=External)・http は対象外。
    """
    from posixpath import normpath, join
    names = set(files)
    for rel_name in [n for n in files if n.endswith(".rels")]:
        base = re.sub(r"_rels/[^/]+$", "", rel_name)
        xml = files[rel_name].decode("utf-8")
        changed = False

        def keep(m):
            nonlocal changed
            whole, tgt, rest = m.group(0), m.group(1), m.group(2)
            if "External" in rest or tgt.startswith("http"):
                return whole
            resolved = normpath(join(base, tgt)).lstrip("/")
            if resolved not in names:
                changed = True
                return ""
            return whole

        new = re.sub(r'<Relationship [^>]*Target="([^"]+)"([^>]*)/>', keep, xml)
        if changed:
            files[rel_name] = new.encode("utf-8")


def drop_calcchain(files):
    """calcChain.xml を削除（Excel が開いたとき再生成）。

    数式を削った（C列の VLOOKUP など）あとは calcChain に古い参照が残り、
    Excel が「内容に問題」エラーを出すため、丸ごと取り除くのが最も安全。
    Content_Types の Override と workbook.xml.rels の関係も外す。
    """
    if "xl/calcChain.xml" not in files:
        return
    del files["xl/calcChain.xml"]
    ct = files["[Content_Types].xml"].decode("utf-8")
    ct = re.sub(r'<Override PartName="/xl/calcChain\.xml"[^>]*/>', "", ct)
    files["[Content_Types].xml"] = ct.encode("utf-8")
    rels = files["xl/_rels/workbook.xml.rels"].decode("utf-8")
    rels = re.sub(r'<Relationship[^>]*Target="calcChain\.xml"[^>]*/>', "", rels)
    files["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")


def patch_workbook_calc(xml):
    """開いたときに全再計算（キャッシュ済みの古い値を一掃）。"""
    if "fullCalcOnLoad" in xml:
        return xml
    if "<calcPr" in xml:
        return xml.replace("<calcPr ", '<calcPr fullCalcOnLoad="1" ', 1)
    return xml.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>')


def patch_meibo_instruction(xml):
    """Start! の説明文「姓・名に分けて」→ 1セル氏名に合わせて更新。"""
    sis = re.findall(r"<si>.*?</si>", xml, flags=re.S)
    old = [s for s in sis if "姓・名に分けて" in s]
    if not old:
        return xml
    new = ("<si><t>クラス生徒の名簿を「出席番号（2桁）」と「氏名（1セル）」の2列で"
           "貼り付け（最大50名）</t></si>")
    return xml.replace(old[0], new, 1)


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

    # 7. 氏名1セル化 + 時刻ヘッダ表示 + 全再計算（冪等）
    ss_xml, name_idx = ensure_shared_string(dec("xl/sharedStrings.xml"), "氏名")
    ss_xml = patch_meibo_instruction(ss_xml)
    files["xl/sharedStrings.xml"] = ss_xml.encode("utf-8")
    files["xl/worksheets/sheet2.xml"] = patch_meibo_sheet(dec("xl/worksheets/sheet2.xml"), name_idx).encode("utf-8")
    files["xl/worksheets/sheet3.xml"] = patch_kibou_sheet(dec("xl/worksheets/sheet3.xml"), name_idx).encode("utf-8")
    files["xl/worksheets/sheet4.xml"] = patch_yotei_sheet(dec("xl/worksheets/sheet4.xml")).encode("utf-8")
    files["xl/styles.xml"] = patch_shrink_time_xfs(dec("xl/styles.xml")).encode("utf-8")
    files["xl/workbook.xml"] = patch_workbook_calc(dec("xl/workbook.xml")).encode("utf-8")

    # 8. 数式を削った影響で calcChain が陳腐化 → 削除（Excel が再生成）
    drop_calcchain(files)

    # 9. 最終掃除: 存在しないパートを指す関係を全 .rels から除去
    prune_dangling_rels(files)

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
