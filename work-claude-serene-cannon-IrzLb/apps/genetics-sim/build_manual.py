"""遺伝のモデル実験 シミュレーター 教員向けマニュアル PDF生成スクリプト"""

from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    BaseDocTemplate, Frame, PageTemplate,
    Paragraph, Spacer, Table, TableStyle,
    PageBreak, KeepTogether,
)

FONT_REGULAR = "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf"
FONT_PROP = "/usr/share/fonts/opentype/ipafont-gothic/ipagp.ttf"

pdfmetrics.registerFont(TTFont("IPAGothic", FONT_REGULAR))
pdfmetrics.registerFont(TTFont("IPAPGothic", FONT_PROP))

COLOR_PRIMARY = colors.HexColor("#246b46")
COLOR_LIGHT = colors.HexColor("#eaf4ee")
COLOR_BORDER = colors.HexColor("#cfdcd2")
COLOR_TEXT = colors.HexColor("#1f2d28")
COLOR_MUTED = colors.HexColor("#4f6056")
COLOR_WARN = colors.HexColor("#c0392b")

styles = {
    "title": ParagraphStyle(
        "title", fontName="IPAGothic", fontSize=22, leading=30,
        textColor=COLOR_PRIMARY, alignment=TA_CENTER, spaceAfter=8,
    ),
    "subtitle": ParagraphStyle(
        "subtitle", fontName="IPAPGothic", fontSize=13, leading=20,
        textColor=COLOR_MUTED, alignment=TA_CENTER, spaceAfter=24,
    ),
    "h1": ParagraphStyle(
        "h1", fontName="IPAGothic", fontSize=16, leading=24,
        textColor=colors.white, backColor=COLOR_PRIMARY,
        borderPadding=(6, 10, 6, 10), spaceBefore=16, spaceAfter=10,
    ),
    "h2": ParagraphStyle(
        "h2", fontName="IPAGothic", fontSize=13, leading=20,
        textColor=COLOR_PRIMARY, spaceBefore=12, spaceAfter=6,
        borderPadding=(0, 0, 4, 0),
    ),
    "h3": ParagraphStyle(
        "h3", fontName="IPAGothic", fontSize=11.5, leading=18,
        textColor=COLOR_TEXT, spaceBefore=8, spaceAfter=4,
    ),
    "body": ParagraphStyle(
        "body", fontName="IPAPGothic", fontSize=10.5, leading=17,
        textColor=COLOR_TEXT, spaceAfter=6, alignment=TA_LEFT,
    ),
    "list": ParagraphStyle(
        "list", fontName="IPAPGothic", fontSize=10.5, leading=17,
        textColor=COLOR_TEXT, leftIndent=14, bulletIndent=2, spaceAfter=2,
    ),
    "code": ParagraphStyle(
        "code", fontName="IPAGothic", fontSize=9.5, leading=14,
        textColor=COLOR_TEXT, backColor=COLOR_LIGHT,
        borderPadding=(6, 8, 6, 8), spaceAfter=6, leftIndent=4, rightIndent=4,
    ),
    "note": ParagraphStyle(
        "note", fontName="IPAPGothic", fontSize=9.5, leading=15,
        textColor=COLOR_MUTED, leftIndent=10, spaceAfter=6,
    ),
    "warn": ParagraphStyle(
        "warn", fontName="IPAPGothic", fontSize=10, leading=16,
        textColor=COLOR_WARN, leftIndent=10, spaceAfter=6,
    ),
}


def p(text, style="body"):
    return Paragraph(text, styles[style])


def li(text):
    return Paragraph(f"・{text}", styles["list"])


def step(n, text):
    return Paragraph(f"<font color='#246b46'><b>{n}.</b></font> {text}", styles["list"])


def code(text):
    return Paragraph(text.replace("\n", "<br/>"), styles["code"])


def hr():
    t = Table([[""]], colWidths=[170 * mm], rowHeights=[1])
    t.setStyle(TableStyle([("LINEBELOW", (0, 0), (-1, -1), 0.5, COLOR_BORDER)]))
    return t


def make_table(rows, col_widths=None, header=True):
    style = [
        ("FONTNAME", (0, 0), (-1, -1), "IPAPGothic"),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("GRID", (0, 0), (-1, -1), 0.4, COLOR_BORDER),
    ]
    if header:
        style.append(("BACKGROUND", (0, 0), (-1, 0), COLOR_PRIMARY))
        style.append(("TEXTCOLOR", (0, 0), (-1, 0), colors.white))
        style.append(("FONTNAME", (0, 0), (-1, 0), "IPAGothic"))

    cells = []
    for r_i, row in enumerate(rows):
        cells.append([
            Paragraph(c, ParagraphStyle(
                f"cell-{r_i}-{c_i}",
                fontName="IPAGothic" if (header and r_i == 0) else "IPAPGothic",
                fontSize=10, leading=14,
                textColor=colors.white if (header and r_i == 0) else COLOR_TEXT,
            ))
            for c_i, c in enumerate(row)
        ])
    t = Table(cells, colWidths=col_widths)
    t.setStyle(TableStyle(style))
    return t


def header_footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("IPAPGothic", 8)
    canvas.setFillColor(COLOR_MUTED)
    canvas.drawRightString(
        200 * mm, 10 * mm,
        f"− {doc.page} −"
    )
    canvas.drawString(
        15 * mm, 10 * mm,
        "遺伝のモデル実験 シミュレーター　教員マニュアル"
    )
    canvas.restoreState()


def build_story():
    s = []

    # ----- 表紙 -----
    s.append(Spacer(1, 60 * mm))
    s.append(p("遺伝のモデル実験", "title"))
    s.append(p("シミュレーター", "title"))
    s.append(Spacer(1, 8 * mm))
    s.append(p("教員向け 操作マニュアル", "subtitle"))
    s.append(Spacer(1, 30 * mm))
    s.append(p("中学校3年 理科<br/>「遺伝の規則性と遺伝子」探Q実習1 対応", "subtitle"))
    s.append(Spacer(1, 60 * mm))
    s.append(p("Google Chrome 専用<br/>サーバー・追加インストール不要 ／ 校内共有フォルダで動作", "note"))
    s.append(PageBreak())

    # ----- 1. このアプリについて -----
    s.append(p("1. このアプリについて", "h1"))
    s.append(p(
        "本アプリは、中学3年理科「遺伝の規則性と遺伝子」の探Q実習1（教科書 p.10／p.23）"
        "で行うトランプを使った遺伝モデル実験を、生徒が自分のタブレット上で再現できるようにしたものです。"
        "クラスの全員がリアルタイムに同期され、誰がどんな遺伝子型なのか、ペア交配の結果はどうなったかが、"
        "先生・生徒の両方の画面で同時に見られます。"
    ))
    s.append(p("特徴", "h2"))
    s.append(li("<b>サーバー不要・インストール不要</b>。校内共有フォルダにファイルを置いて、ブラウザで開くだけ"))
    s.append(li("Node.js や追加のソフトウェアは一切不要"))
    s.append(li("生徒どうしのペア交配・遺伝子交換が、画面操作だけで完結"))
    s.append(li("クラス全体の AA : Aa : aa の比、丸 : しわ の比がリアルタイムに集計される"))
    s.append(li("教員用パスワードでクラスの初期化や個別設定が可能"))

    s.append(p("仕組み（参考）", "h2"))
    s.append(p(
        "ブラウザの File System Access API を使って、共有フォルダ内の <b>data/</b> フォルダを"
        "データベースのように使います。生徒1人につき1ファイルを書き込み、約1.5秒間隔で全員が"
        "互いのファイルを読みに行くことで同期します。"
    ))

    # ----- 2. 動作環境 -----
    s.append(p("2. 動作環境", "h1"))
    s.append(make_table([
        ["項目", "条件"],
        ["OS", "Windows 10 / 11（Mac でも動作可）"],
        ["ブラウザ", "Google Chrome バージョン 86 以降（各端末にインストール済みであること）"],
        ["ネットワーク", "校内のファイル共有（SMB／共有フォルダ）にアクセスできること"],
        ["端末", "1人1台 or 1班1台のいずれでも可（同名ログインは不可）"],
        ["想定人数", "1クラス 40人程度まで"],
    ], col_widths=[35 * mm, 135 * mm]))
    s.append(Spacer(1, 4 * mm))
    s.append(p("⚠ 本ランチャーは Google Chrome を起動します（Edge/Firefox/Safari は使いません）。各端末に Chrome が入っていることを確認してください。", "warn"))
    s.append(p("⚠ 学校のセキュリティ設定により、ブラウザから共有フォルダへの書き込みが制限されている場合は動作しません。本格運用の前に、必ず1台で動作確認をしてください。", "warn"))

    # ----- 3. セットアップ -----
    s.append(p("3. セットアップ（先生の事前準備）", "h1"))

    s.append(p("⚠ 重要：生徒は必ず .bat ランチャー経由で起動", "h2"))
    s.append(p(
        "ブラウザの仕様により、<b>index.html を共有フォルダから直接ダブルクリックして開く</b>と、"
        "ファイルアクセス API がエラーも出さずに止まることがあります。"
        "（「共有フォルダを選ぶ」を押しても先に進まない症状）", "warn"
    ))
    s.append(p(
        "本アプリには専用の起動バッチファイル <b>「GeneticsApp起動.bat」</b> が同梱されています。"
        "これをダブルクリックすると、次の処理が自動で行われます：", "body"
    ))
    s.append(li("3ファイル（index.html／app.js／style.css）を <b>%TEMP%\\GeneticsApp\\</b> にコピー"))
    s.append(li("<b>専用プロファイルの Google Chrome</b> でローカルコピーを開く（普段使いの Chrome に影響しない）"))
    s.append(li("Chromeのウィンドウが閉じられたら、<b>コピーした一式と専用プロファイルを自動削除</b>（端末に何も残らない）"))
    s.append(p("<b>生徒は .bat をダブルクリック → 授業終了時は Chrome を × で閉じる、これだけです。</b>", "body"))

    s.append(p("3.1 共有フォルダにファイル一式を配置", "h2"))
    s.append(p("校内共有フォルダに、以下の <b>4ファイル</b> をまとめて置きます。"))
    s.append(code(
        "GeneticsApp/<br/>"
        "├── index.html<br/>"
        "├── app.js<br/>"
        "├── style.css<br/>"
        "└── GeneticsApp起動.bat   ← 生徒はこれをダブルクリック"
    ))
    s.append(p("例：<b>\\\\school-server\\shared\\3年理科\\GeneticsApp\\</b>", "note"))
    s.append(p("※ アプリが起動時に自動で data/ サブフォルダを作成します。生徒全員に「書き込み可」の権限を付けてください。", "note"))

    s.append(p("3.2 動作確認（重要）", "h2"))
    s.append(step(1, "先生の端末から共有フォルダを開く"))
    s.append(step(2, "<b>GeneticsApp起動.bat</b> をダブルクリック（一瞬黒い画面が出てから Chrome が開く）"))
    s.append(step(3, "「共有フォルダを選ぶ」ボタンをクリック"))
    s.append(step(4, "手順 3.1 で用意した <b>共有フォルダ</b> を選択"))
    s.append(step(5, "「このサイトに〜への変更を許可しますか？」→ <b>「編集を許可」</b>を選択"))
    s.append(step(6, "名前を入れて参加 → 自分の遺伝子カードが表示されれば成功"))
    s.append(p("⚠ もし反応しない・許可ダイアログが出ない場合は、ブラウザのアドレスバーを確認してください。<br/>"
               "正常なとき: <b>file:///C:/Users/.../AppData/Local/Temp/GeneticsApp/index.html</b><br/>"
               "問題があるとき: <b>file:////server-name/.../index.html</b>（共有フォルダ直接、.bat 経由になっていない）", "warn"))

    s.append(p("3.3 授業終了時（自動でかたづく）", "h2"))
    s.append(p("Chrome のウィンドウを <b>×</b> で閉じるだけで OK です。"
               ".bat が裏で待機しており、Chromeが閉じられたタイミングで"
               "<b>コピーした3ファイルと専用プロファイルを自動削除</b>します。"
               "（Chromeを閉じたあと、一瞬黒い画面が見えてから消えれば後片付け完了）", "body"))
    s.append(p("※ 共有フォルダの data/ 配下にある授業データは残ります。古いデータを消したい場合は、"
               "授業の最後に先生メニューの「古いデータを掃除」をお使いください。", "note"))

    s.append(p("3.4 アプリの更新方法", "h2"))
    s.append(p("先生がアプリを差し替えたいときは、共有フォルダ上の "
               "<b>index.html / app.js / style.css</b> を上書きするだけで OK です。"
               "次に生徒が .bat を起動した瞬間に、最新版が各PCにコピーされます。", "body"))

    s.append(p("3.5 先生用パスワードの確認・変更", "h2"))
    s.append(p("先生メニューを使うには、参加時に <b>「先生として参加する」</b> にチェックを入れ、パスワードを入力します。初期パスワードは次のとおりです。"))
    s.append(code("初期パスワード： sensei"))
    s.append(p("パスワードを変えたい場合は、メモ帳で <b>app.js</b> を開き、ファイル先頭付近の次の行を編集して保存してください。"))
    s.append(code("const TEACHER_PASSWORD = 'sensei';   ← この 'sensei' の部分を書き換える"))
    s.append(p("※ パスワードを変更したら、共有フォルダの app.js も差し替えてください。生徒が編集できないように、共有フォルダの権限設定で <b>書き込み禁止＋読み取り専用</b> にすることを推奨します（生徒の data/ への書き込みは別途必要です）。", "note"))

    s.append(PageBreak())

    # ----- 4. 授業の流れ -----
    s.append(p("4. 授業での使い方", "h1"))

    s.append(p("4.1 全体の流れ", "h2"))
    s.append(make_table([
        ["段階", "教科書の対応", "アプリでの操作"],
        ["導入", "純系の親（AA）と純系の親（aa）を用意", "先生メニュー：「AA と aa を半々に」で初期化"],
        ["ステップ1", "親（AA × aa）から子をつくる", "生徒どうしで「ペアで子をつくる」を実行 → 子は必ず Aa"],
        ["途中の整理", "全員の遺伝子型を Aa（雑種）にして子の世代に揃える", "先生メニュー：「全員 Aa」で初期化"],
        ["ステップ2", "子（Aa × Aa）から孫をつくる", "生徒どうしで「ペアで子をつくる」を40回程度繰り返す"],
        ["まとめ", "孫の世代の AA : Aa : aa の比を表に書く", "右パネルの「クラス集計」を見ながら全体で確認"],
        ["後片付け", "—", "先生メニュー：「古いデータを掃除」"],
    ], col_widths=[22 * mm, 65 * mm, 83 * mm]))

    s.append(p("4.2 各画面の見方", "h2"))
    s.append(p("<b>左パネル「あなたの遺伝子」</b>", "h3"))
    s.append(li("自分の遺伝子型（AA／Aa／aa）と形質（丸／しわ）が表示されます。"))
    s.append(li("「ペアで子をつくる」「遺伝子を交換する」の2つのモードを切り替えて操作します。"))
    s.append(p("<b>中央パネル「クラスのみんな」</b>", "h3"))
    s.append(li("参加中の生徒一覧が表示されます。各生徒の現在の遺伝子型・形質が見えます。"))
    s.append(li("先生でログインしている場合は、各生徒の右に「設定」ボタンが表示され、個別に遺伝子型を変更できます。"))
    s.append(p("<b>右パネル「クラス集計」</b>", "h3"))
    s.append(li("AA／Aa／aa の人数と、丸：しわの比、AA:Aa:aa の比がリアルタイムに集計されます。"))
    s.append(li("「活動の記録」には、誰が誰と交配したか、誰が初期化したか、などの履歴が時刻順に並びます。"))

    # ----- 5. 先生メニュー -----
    s.append(p("5. 先生メニュー（初期化）", "h1"))
    s.append(p(
        "先生としてログインすると、右下に「先生メニュー」が表示されます。クラス全体の遺伝子型を一度に揃えたいときに使います。"
    ))

    s.append(p("5.1 初期化モードの種類", "h2"))
    s.append(make_table([
        ["モード", "結果", "おもな用途"],
        ["全員 AA", "参加中の全員を AA（純系・丸）に揃えます。", "純系AAクラス全体での確認"],
        ["全員 Aa", "参加中の全員を Aa（雑種）に揃えます。", "ステップ2（孫の世代）の出発点"],
        ["全員 aa", "参加中の全員を aa（純系・しわ）に揃えます。", "純系aaクラス全体での確認"],
        ["AA と aa を半々に", "ランダム順で半分を AA、残り半分を aa にします（奇数人なら AA を1人多く）。", "ステップ1（親の世代）の出発点"],
        ["配分を指定…", "AA：Aa：aa の比率を指定し、クラス人数で按分してランダムに配ります。", "1:2:1（メンデルの予想比）の確認、特殊な分布での実験"],
    ], col_widths=[40 * mm, 80 * mm, 50 * mm]))

    s.append(p("5.2 「配分を指定」モードの使い方", "h2"))
    s.append(step(1, "「初期化モード」のドロップダウンで <b>「配分を指定…」</b> を選択"))
    s.append(step(2, "下に表示される 3つの数値入力欄に、希望する比率を入力（例：AA=1、Aa=2、aa=1）"))
    s.append(step(3, "「初期化すると：AA X人／Aa Y人／aa Z人」のプレビューが自動更新される"))
    s.append(step(4, "「全員を初期化」をクリック → 確認ダイアログで OK"))
    s.append(Spacer(1, 4 * mm))
    s.append(p(
        "<b>例：</b>30人クラスで <b>1 : 2 : 1</b> を指定した場合 → AA 8人／Aa 15人／aa 8人（端数は Aa で吸収）。"
        "毎回ランダム順にシャッフルするので、出席番号順や入室順に偏りません。", "note"
    ))

    s.append(p("5.3 個別の生徒を変更したい場合", "h2"))
    s.append(p(
        "「クラスのみんな」パネルで、生徒名の右にある <b>「設定」</b> ボタンをクリックすると、その生徒だけ"
        "AA／Aa／aa を指定して変更できます。実験中に「この生徒は最初 AA、もう一人は最初 aa」のような"
        "細かい配置をしたいときに使います。"
    ))

    s.append(p("5.4 古いデータを掃除", "h2"))
    s.append(p(
        "授業終了後や別クラスで使う前に押してください。data/events、data/requests、退出した生徒のファイルを"
        "削除し、活動の記録（履歴）もクリアします。"
    ))

    # ----- 6. 生徒の操作 -----
    s.append(p("6. 生徒の操作（指導の参考に）", "h1"))

    s.append(p("6.1 参加するまで", "h2"))
    s.append(step(1, "共有フォルダを開き、<b>GeneticsApp起動.bat</b> をダブルクリック（Chromeが自動で開く）"))
    s.append(step(2, "「共有フォルダを選ぶ」をクリックし、<b>同じ共有フォルダ</b> を選択"))
    s.append(step(3, "「編集を許可」をクリック（1回だけ表示されます）"))
    s.append(step(4, "名前を入れて「参加する」"))
    s.append(p("授業の終わりは Chrome ウィンドウを <b>×</b> で閉じるだけ。.bat が自動で後片付けをします。", "note"))

    s.append(p("6.2 ペアで子をつくる（メイン操作）", "h2"))
    s.append(step(1, "「ペアで子をつくる」タブが選ばれていることを確認"))
    s.append(step(2, "ドロップダウンから相手を選ぶ（相手の遺伝子型が表示される）"))
    s.append(step(3, "「交配を申し込む」をクリック"))
    s.append(step(4, "相手の画面に申し込みダイアログが出る → 相手が「受ける」をクリック"))
    s.append(step(5, "両方の画面に <b>「子の遺伝子ができました！」</b> ダイアログが表示される"))
    s.append(p("親役の遺伝子は変化しません。1回の交配ごとに「子1人分」のデータが出るので、生徒は記録用紙にメモしてから次の相手と交配します。", "note"))

    s.append(p("6.3 遺伝子を交換する（発展操作）", "h2"))
    s.append(p(
        "「自分のカード1枚」と「相手のカード1枚」を交換します。教科書のメイン手順では使いませんが、"
        "「もし途中で遺伝子が混ざったらどうなるか？」を試したいときに便利です。"
    ))
    s.append(li("申込側：渡したいカード（1枚目／2枚目）を選んで「交換を申し込む」"))
    s.append(li("受け側：自分のどちらと入れ替えるかを選んで「受ける」"))
    s.append(li("両方の遺伝子型が変化します"))

    # ----- 7. トラブルシューティング -----
    s.append(p("7. トラブルシューティング", "h1"))
    tbl = make_table([
        ["症状", "対処"],
        [
            "フォルダを選んでも何も起きない／許可ダイアログも出ない（最頻出）",
            "共有フォルダの index.html を直接開いている可能性が高いです。いったんChromeを閉じて、共有フォルダの <b>「GeneticsApp起動.bat」</b> をダブルクリックして開き直してください。"
        ],
        [
            ".batを起動したら「Google Chromeが見つかりません」と出る",
            "その端末に Chrome がインストールされていません。https://www.google.com/chrome/ からインストールしてください（Edgeでは動作させない設計です）。"
        ],
        [
            "Chromeを閉じても黒い画面が残る／一瞬で消えない",
            "後片付けの処理中です。数秒以内に自動で消えます。途中で閉じてもPCに残るのは %TEMP%\\GeneticsApp\\ 配下のキャッシュファイルのみで、Windowsの自動クリーンアップで消えます。"
        ],
        [
            "「共有フォルダを選ぶ」ボタンが押せない／対応していないと表示される",
            "Google Chrome がインストールされていません。.bat は Chrome を探して起動する仕組みのため、Chromeを各端末にインストールしてください（https://www.google.com/chrome/）。"
        ],
        [
            "「閲覧のみ」を選んでしまった／書き込みできないと出る",
            "ページを再読込（F5）し、もう一度「共有フォルダを選ぶ」を試してください。許可ダイアログで <b>「編集を許可」</b> を選びます。"
        ],
        [
            "フォルダ選択時に「アクセスを拒否しました」と出る",
            "共有フォルダへの書き込み権限がない可能性があります。まず先生の端末でテストし、ダメなら学校のシステム管理者に相談してください。"
        ],
        [
            "他の生徒が表示されない／同期が遅い",
            "ネットワークが遅い可能性があります。1〜2秒待ってください。それでもダメならページ再読込（F5）→ もう一度同じフォルダを選択。"
        ],
        [
            "「その名前はすでに使われています」と出る",
            "1分ほど待つ（30秒で自動的に「退出」扱い）か、先生メニューの「古いデータを掃除」を実行してください。"
        ],
        [
            "ボタンを押しても反応がない／画面が真っ白",
            "F12キーを押して開発者ツールを開き、Console タブのエラー内容を確認。多くの場合、ブラウザの再起動で解決します。"
        ],
        [
            "data/ フォルダにファイルが大量にたまっている",
            "授業の終わりに先生メニューの「古いデータを掃除」を実行。それでも残るファイルは手動で data/ フォルダごと削除して構いません（次回起動時に自動で再作成されます）。"
        ],
        [
            "選択肢が3つしか表示されない（半々／配分を指定が出ない）",
            "古い index.html が残っています。最新版で上書きし、ブラウザで Ctrl+Shift+R（強制リロード）を実行してください。"
        ],
        [
            "新しい選択肢が古い設定を上書きする",
            "先生メニューでの初期化は <b>参加中の生徒のみ</b> に適用されます。授業開始前に全員に参加してもらってから初期化してください。"
        ],
    ], col_widths=[55 * mm, 115 * mm])
    s.append(tbl)

    # ----- 8. 制約事項 -----
    s.append(p("8. 制約事項・注意点", "h1"))
    s.append(li("ブラウザを閉じてから1分ほどは、その生徒がまだ参加中扱いになります（同名再参加ができないことがあります）。"))
    s.append(li("ファイル共有のディスク I/O 速度に依存します。古いネットワーク機器の場合、同期に2〜3秒かかることがあります。"))
    s.append(li("生徒どうしが同時に同じ動作（同じ相手と交配など）を行うと、まれに同期が遅れる場合があります。"))
    s.append(li("先生メニューでの初期化は、その時点で参加中の生徒のみに適用されます。後から参加した生徒は <b>config.json の initialGenotype</b> の値で参加します。"))
    s.append(li("生徒の操作内容は data/ フォルダにJSON形式で保存されます。個人情報を入力させないようご注意ください（名前のみ入力します）。"))

    # ----- 9. 参考 -----
    s.append(p("9. 参考：ファイル構成", "h1"))
    s.append(code(
        "GeneticsApp/                  ← 共有フォルダに置くフォルダ<br/>"
        "├── index.html                ← 開くファイル（生徒・先生共通）<br/>"
        "├── app.js                    ← プログラム本体<br/>"
        "├── style.css                 ← デザイン<br/>"
        "├── README.md                 ← 簡易マニュアル<br/>"
        "└── data/                     ← 起動時に自動作成<br/>"
        "    ├── config.json           ← クラスの設定（先生のみ更新）<br/>"
        "    ├── students/             ← 生徒1人＝1ファイル<br/>"
        "    ├── events/               ← 活動の履歴<br/>"
        "    └── requests/             ← 交配・交換の申し込み（処理後に自動削除）"
    ))

    s.append(Spacer(1, 12 * mm))
    s.append(hr())
    s.append(Spacer(1, 4 * mm))
    s.append(p(
        "本マニュアルおよびアプリは教育用途で自由にお使いいただけます。<br/>"
        "改善要望・不具合報告は、お手元の管理者または開発者までお知らせください。",
        "note"
    ))

    return s


def main():
    out = "遺伝のモデル実験_教員マニュアル.pdf"
    doc = BaseDocTemplate(
        out, pagesize=A4,
        leftMargin=18 * mm, rightMargin=18 * mm,
        topMargin=18 * mm, bottomMargin=18 * mm,
        title="遺伝のモデル実験 シミュレーター 教員マニュアル",
        author="探Q実習1 教材",
    )
    frame = Frame(
        doc.leftMargin, doc.bottomMargin,
        doc.width, doc.height, id="main",
    )
    doc.addPageTemplates([
        PageTemplate(id="default", frames=[frame], onPage=header_footer),
    ])
    doc.build(build_story())
    print(f"Generated: {out}")


if __name__ == "__main__":
    main()
