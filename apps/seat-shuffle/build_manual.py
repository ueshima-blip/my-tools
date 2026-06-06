"""席替えアプリ 教員向け操作マニュアル PDF生成スクリプト

ページ送りで見出しが本文と分離しないよう、次の対策を入れている：
- 章見出し(h1)は必ず改ページしてページ先頭から始める（ページ末での孤立を防止）
- 小見出し(h2/h3)は keepWithNext=1 で直後の本文と必ず同じページに保持する
"""

from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
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

COLOR_PRIMARY = colors.HexColor("#2d6a4f")
COLOR_LIGHT = colors.HexColor("#e8f4ee")
COLOR_BORDER = colors.HexColor("#d4ddd6")
COLOR_TEXT = colors.HexColor("#1a2620")
COLOR_MUTED = colors.HexColor("#566159")
COLOR_WARN = colors.HexColor("#c0392b")

styles = {
    "title": ParagraphStyle(
        "title", fontName="IPAGothic", fontSize=26, leading=34,
        textColor=COLOR_PRIMARY, alignment=TA_CENTER, spaceAfter=8,
    ),
    "subtitle": ParagraphStyle(
        "subtitle", fontName="IPAPGothic", fontSize=14, leading=22,
        textColor=COLOR_MUTED, alignment=TA_CENTER, spaceAfter=24,
    ),
    "h1": ParagraphStyle(
        "h1", fontName="IPAGothic", fontSize=16, leading=24,
        textColor=colors.white, backColor=COLOR_PRIMARY,
        borderPadding=(7, 10, 7, 10), spaceBefore=2, spaceAfter=12,
    ),
    "h2": ParagraphStyle(
        "h2", fontName="IPAGothic", fontSize=13, leading=20,
        textColor=COLOR_PRIMARY, spaceBefore=12, spaceAfter=6,
        keepWithNext=1,
    ),
    "h3": ParagraphStyle(
        "h3", fontName="IPAGothic", fontSize=11.5, leading=18,
        textColor=COLOR_TEXT, spaceBefore=8, spaceAfter=4,
        keepWithNext=1,
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
        "code", fontName="IPAGothic", fontSize=9.5, leading=15,
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
    return Paragraph(f"<font color='#2d6a4f'><b>{n}.</b></font> {text}", styles["list"])


def code(text):
    return Paragraph(text.replace("\n", "<br/>"), styles["code"])


def hr():
    t = Table([[""]], colWidths=[174 * mm], rowHeights=[1])
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
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, COLOR_LIGHT]),
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
    t = Table(cells, colWidths=col_widths, repeatRows=1 if header else 0)
    t.setStyle(TableStyle(style))
    return t


def chapter(s, title):
    """章見出し(h1)。直前が改ページでなければ改ページを入れ、ページ先頭から始める。"""
    if s and not isinstance(s[-1], PageBreak):
        s.append(PageBreak())
    s.append(p(title, "h1"))


def header_footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("IPAPGothic", 8)
    canvas.setFillColor(COLOR_MUTED)
    canvas.drawRightString(195 * mm, 10 * mm, f"− {doc.page} −")
    if doc.page > 1:
        canvas.drawString(18 * mm, 10 * mm, "席替えアプリ　操作マニュアル")
    canvas.setStrokeColor(COLOR_BORDER)
    canvas.setLineWidth(0.4)
    canvas.line(18 * mm, 13 * mm, 195 * mm, 13 * mm)
    canvas.restoreState()


def build_story():
    s = []

    # ===== 表紙 =====
    s.append(Spacer(1, 58 * mm))
    s.append(p("席替えアプリ", "title"))
    s.append(Spacer(1, 6 * mm))
    s.append(p("操作マニュアル（教員向け）", "subtitle"))
    s.append(Spacer(1, 40 * mm))
    s.append(p(
        "クラス名簿（Excel）を取り込み、男女配置・班分け・座席固定・距離条件などを"
        "考慮して、瞬時／ルーレットで席替えができる単一HTMLファイルのアプリです。",
        "subtitle",
    ))
    s.append(Spacer(1, 40 * mm))
    s.append(p(
        "単一HTMLファイル ／ インストール・サーバー不要 ／ パソコン・iPad 対応<br/>"
        "データはお使いの端末のブラウザ内にのみ保存されます（外部送信なし）",
        "note",
    ))
    s.append(PageBreak())

    # ===== 1. このアプリについて =====
    chapter(s, "1. このアプリについて")
    s.append(p(
        "本アプリは、先生がクラスの席替えを手早く・公平に行うための支援ツールです。"
        "名簿を読み込み、教室のレイアウトや男女の配置、班分け、「この2人は離す」といった"
        "条件を設定したうえで、ボタンひとつで席替えを実行できます。結果はそのまま印刷・"
        "画像保存でき、教室に配付・掲示できます。"
    ))
    s.append(p("おもな特徴", "h2"))
    s.append(li("<b>クラス×教室ごとに独立した座席設定</b>。同じ「理科室」でも1組と2組で別のレイアウトを保持"))
    s.append(li("<b>男女の配置・班分け・距離条件・座席固定</b>を考慮した席替え"))
    s.append(li("<b>3つの席替え方式</b>（瞬時／一人ずつルーレット／一括ルーレット）"))
    s.append(li("<b>ルーレットは自動で全画面表示</b>。教室のモニター・プロジェクターに大きく映せる"))
    s.append(li("<b>印刷・JPG保存</b>。A4横1ページに自動で収まり、デザインも選べる"))
    s.append(li("<b>履歴</b>。過去の席替えをいつでも復元できる"))
    s.append(li("<b>バックアップ</b>。全データを1ファイルにまとめて外部に保存・復元できる"))
    s.append(p(
        "すべての処理はお使いの端末（ブラウザ）の中だけで完結します。インターネットに"
        "生徒の情報を送信することは一切ありません。", "note"
    ))

    # ===== 2. 準備と基本の流れ =====
    chapter(s, "2. 準備と基本の流れ")
    s.append(p("起動のしかた", "h2"))
    s.append(li("パソコン：<b>index.html</b> を Google Chrome／Microsoft Edge などで開きます"))
    s.append(li("iPad：<b>index.html</b> を Safari で開きます"))
    s.append(p(
        "インストールやサーバーは不要です。ファイルをダブルクリックして開くだけで使えます。", "note"
    ))

    s.append(p("基本の流れ（クイックスタート）", "h2"))
    s.append(step(1, "<b>「名簿」</b>タブで、クラス名簿の Excel ファイルを読み込む"))
    s.append(step(2, "<b>「配置」</b>タブで、教室の列数・行数を決め、座席を編集する"))
    s.append(step(3, "必要に応じて<b>「条件」</b>タブで「離す」「席を固定」などを設定する"))
    s.append(step(4, "画面上部の<b>「瞬時に決定」</b>または<b>ルーレット</b>で席替えを実行する"))
    s.append(step(5, "<b>「デザイン」</b>を選び、<b>「印刷」</b>または<b>「JPG保存」</b>で出力する"))
    s.append(step(6, "ときどき<b>「設定」</b>タブからバックアップを外部に保存する"))

    # ===== 3. 画面の構成 =====
    chapter(s, "3. 画面の構成")
    s.append(p("画面は大きく「上部のヘッダー」「左side（タブ）」「中央の座席表」に分かれます。", "body"))

    s.append(p("3.1 ヘッダー（最上部）", "h2"))
    s.append(li("<b>クラスの選択・追加</b>：今あつかうクラスを切り替えます"))
    s.append(li("<b>教室の選択・追加</b>：普通教室・理科室など、教室ごとに座席配置を保存できます"))
    s.append(li("<b>「データ消去」</b>：履歴のみ／名簿のみ／教室のみ／全データ の消去メニューを開きます"))

    s.append(p("3.2 左側のタブ", "h2"))
    s.append(make_table([
        ["タブ", "できること"],
        ["名簿", "Excel名簿の読み込み、生徒の追加・編集・削除"],
        ["配置", "列・行の設定、座席の有効／無効／使わない席、男子席・女子席の指定、班色"],
        ["班色", "座席をグループに分けて色をつける（班ごとの色分け）"],
        ["条件", "「2人を離す」などの距離条件、特定の生徒を特定の席に固定"],
        ["履歴", "過去の席替え結果の確認・復元・削除"],
        ["設定", "効果音、クラス・教室の管理、バックアップ、データ消去"],
    ], col_widths=[28 * mm, 146 * mm]))

    s.append(p("3.3 中央上部のツールバー", "h2"))
    s.append(li("<b>席替え</b>：「新しい席替え」「瞬時に決定」「一人ずつルーレット」「一括ルーレット」"))
    s.append(li("<b>入替</b>：2つの席を選んで座席を入れ替える"))
    s.append(li("<b>視点</b>：「生徒視点」と「教師視点」（180度回転）を切り替える"))
    s.append(li("<b>表示</b>：番号・ふりがな・男女色・班色 の表示ON/OFF"))
    s.append(li("<b>デザイン</b>：印刷・JPGの見た目を選ぶ（第8章）"))
    s.append(li("<b>JPG保存・印刷</b>：座席表を出力する"))

    # ===== 4. 名簿を読み込む =====
    chapter(s, "4. 名簿を読み込む（名簿タブ）")
    s.append(p(
        "「名簿」タブで、クラス名簿の Excel ファイル（.xlsx）を読み込みます。"
        "次の列を用意してください。", "body"
    ))
    s.append(make_table([
        ["列", "内容", "必須"],
        ["出席番号", "1, 2, 3 … の番号", "推奨"],
        ["氏名", "生徒の名前（座席に表示されます）", "必須"],
        ["ふりがな", "氏名の読み（座席や印刷に表示できます）", "任意"],
        ["性別", "男子／女子（男女の配置やルーレット表示に使います）", "任意"],
    ], col_widths=[28 * mm, 116 * mm, 30 * mm]))
    s.append(p("テンプレートを使うと簡単です", "h2"))
    s.append(p(
        "名簿タブの<b>「名簿テンプレート（Excel）をダウンロード」</b>を押すと、上の列があらかじめ"
        "用意された空の Excel ファイルが手に入ります。記入例の2行を上書きして名簿を作り、"
        "<b>「Excelファイルを読み込む」</b>で取り込んでください。", "body"
    ))
    s.append(p(
        "性別の欄は「男／女」（または male／female）と入力します。氏名が空の行は読み込み時に"
        "とばされるので、人数より多めの行が残っていても問題ありません。", "note"
    ))
    s.append(p("生徒の追加・修正", "h2"))
    s.append(li("読み込み後、一覧から生徒をタップすると、名前・ふりがな・性別・出席番号を修正できます"))
    s.append(li("座席表のうえで生徒をタップしても、その生徒の編集ができます"))
    s.append(p(
        "ふりがな・性別は任意ですが、男女別の配置や、ルーレット中の名前表示に使われます。"
        "入力しておくとより便利です。", "note"
    ))

    # ===== 5. 教室レイアウトを作る =====
    chapter(s, "5. 教室レイアウトを作る（配置タブ）")
    s.append(p("5.1 列・行の設定", "h2"))
    s.append(p(
        "「配置」タブで教室の<b>列数</b>と<b>行数</b>を指定すると、その大きさの座席が並びます。"
        "教卓は前（上）に表示されます。", "body"
    ))

    s.append(p("5.2 座席の編集（使う席・使わない席・無効）", "h2"))
    s.append(p("レイアウト編集では、座席をタップするたびに次のように状態が変わります。", "body"))
    s.append(make_table([
        ["状態", "意味"],
        ["通常の席", "生徒が座る席"],
        ["使わない席", "席はあるが生徒は座らせない（空席として表示）"],
        ["無効", "席そのものを置かない（机がない場所）"],
    ], col_widths=[34 * mm, 140 * mm]))

    s.append(p("5.3 男子席・女子席の指定", "h2"))
    s.append(p("座席に「男子席」「女子席」を割り当てると、その配置に合わせて席替えされます。指定方法は3通りです。", "body"))
    s.append(li("<b>クリック設定</b>：座席を1つずつタップして男子席／女子席を指定"))
    s.append(li("<b>列ごと</b>：列単位でまとめて男女を割り当て"))
    s.append(li("<b>交互</b>：男女が交互になるように自動で割り当て"))
    s.append(p(
        "注意：男子席・女子席の数と、名簿の男子・女子の人数が合っていないと、席替え実行時に"
        "警告が出ます。数をそろえてください。", "warn"
    ))

    s.append(p("5.4 他の教室からコピー", "h2"))
    s.append(p(
        "すでに作った別の教室のレイアウトを、いまの教室にコピーできます。"
        "似た配置を一から作り直す手間が省けます。", "body"
    ))
    s.append(p("座席配置は<b>クラス×教室ごとに別々に保存</b>されます。1組と2組で同じ教室でも別の配置を持てます。", "note"))

    # ===== 6. 班分け・条件・固定 =====
    chapter(s, "6. 班分け・離す条件・座席固定")
    s.append(p("6.1 班色（班色タブ）", "h2"))
    s.append(p(
        "座席を選んでグループ（班）の色をつけられます。班ごとに色分けして印刷したいときに使います。"
        "色は座席表・印刷・JPGに反映されます。", "body"
    ))

    s.append(p("6.2 距離条件（条件タブ）", "h2"))
    s.append(p("「この組み合わせは避けたい」といった条件を設定できます。", "body"))
    s.append(li("<b>2人を離す</b>：指定した2人が近くにならないようにする"))
    s.append(li("<b>前の方にする</b>：特定の生徒を前から数行以内に配置する"))
    s.append(p(
        "条件が多すぎたり、互いに矛盾していると席替えできないことがあります。"
        "その場合は警告が出るので、条件を見直して減らしてください。", "warn"
    ))

    s.append(p("6.3 座席を固定する", "h2"))
    s.append(p(
        "特定の生徒を特定の席に固定できます。配慮が必要な生徒を決まった席にしたいときに便利です。"
        "固定した生徒は、席替えを実行してもその席のまま動きません。", "body"
    ))
    s.append(p("固定や条件も、クラス×教室ごとに保存されます。", "note"))

    # ===== 7. 席替えを実行する =====
    chapter(s, "7. 席替えを実行する")
    s.append(p("7.1 3つの席替え方式", "h2"))
    s.append(make_table([
        ["方式", "動き", "向いている場面"],
        ["瞬時に決定", "ボタンを押すと即座に席が決まる", "すばやく決めたいとき"],
        ["一人ずつルーレット", "1人ずつ順番に、回って止めて席を決める", "1人ずつ盛り上げたいとき"],
        ["一括ルーレット", "全員分が同時に回り、止めると全員確定", "全体を一気に決めたいとき"],
    ], col_widths=[38 * mm, 86 * mm, 50 * mm]))

    s.append(p("7.2 ルーレットの操作", "h2"))
    s.append(li("<b>一人ずつ</b>：「START」で回り、「STOP」でその生徒の席が決定。これを人数分くり返します"))
    s.append(li("<b>一括</b>：「START」で全席が回り、「STOP」で全員の席が一度に決まります"))

    s.append(p("7.3 自動で全画面表示（教室モニター向け）", "h2"))
    s.append(p(
        "ルーレットを始めると、自動的にブラウザが全画面表示になります。教室のモニターや"
        "プロジェクターに大きく映して、生徒に分かりやすく見せられます。", "body"
    ))
    s.append(li("終わったら、緑色の<b>「全画面を解除して終了」</b>ボタンを押すと、全画面が解除され座席表（結果）に戻ります"))
    s.append(li("上部の<b>「全画面」</b>ボタンでも、いつでも全画面の入／切を切り替えられます"))
    s.append(li("キーボードの<b>Esc キー</b>でも全画面を解除できます"))
    s.append(p(
        "全画面はブラウザの仕様上、ボタン操作をきっかけにのみ有効になります（勝手に全画面に"
        "なることはありません）。ルーレット開始ボタンがそのきっかけになります。", "note"
    ))

    s.append(p("7.4 途中で中断しても再開できる", "h2"))
    s.append(p(
        "ルーレットの途中でブラウザが閉じたり固まっても、次に開いたときに「続きから再開」"
        "できます。確定済みの席はそのまま、残りの生徒から続けられます。", "body"
    ))

    s.append(p("7.5 席を手動で入れ替える", "h2"))
    s.append(p(
        "「入替」ボタンを押し、入れ替えたい2つの席を順にタップすると、その2人の席が入れ替わります。"
        "席替えの結果を少しだけ手直ししたいときに使います。", "body"
    ))

    # ===== 8. 印刷・JPG保存 =====
    chapter(s, "8. 座席表の印刷・JPG保存")
    s.append(p(
        "「印刷」または「JPG保存」を押すと、<b>印字する月日</b>と<b>デザイン</b>を選ぶ画面が出ます。"
        "選んでから出力してください。座席表は<b>A4横1ページ</b>に自動で収まります。", "body"
    ))

    s.append(p("8.1 視点（生徒視点・教師視点）", "h2"))
    s.append(li("<b>生徒視点</b>：教卓が上。生徒に配る・掲示する向き"))
    s.append(li("<b>教師視点</b>：全体を180度回転し、教卓が下（先生の立ち位置）。先生が教室で見やすい向き"))
    s.append(p("どちらの視点でも、名前や見出しは正しく読める向きで印刷されます。", "note"))

    s.append(p("8.2 表示の切り替え", "h2"))
    s.append(p(
        "「番号」「ふりがな」「男女色」「班色」の表示を切り替えられます。印刷・JPGのときは、"
        "見やすいように番号とふりがなを自動で表示します。", "body"
    ))

    s.append(p("8.3 デザインを選ぶ", "h2"))
    s.append(p("中学・高校でも使える落ち着いたデザインを4種類から選べます。", "body"))
    s.append(make_table([
        ["デザイン", "内容"],
        ["シンプル（標準）", "装飾なしの標準デザイン"],
        ["机（モダン）", "普通教室向け。座席ひとつひとつが学校机の形。前面にホワイトボードと教卓"],
        ["机（木目）", "同じく机の形で、木のあたたかい色合い。前面に黒板"],
        ["シック", "無地で上品なミニマルデザイン"],
    ], col_widths=[38 * mm, 136 * mm]))
    s.append(li("「机」デザインでは、<b>生徒のいる席だけ</b>が机の形になり、空席は平らな枠で表示されます"))
    s.append(li("男女色・班色は、机の天板の色としてそのまま残ります"))
    s.append(li("選んだデザインは画面にもすぐ反映され、次に開いたときも記憶されています"))
    s.append(p(
        "印刷はブラウザの印刷機能を使います。用紙は<b>A4・横向き</b>を選んでください。"
        "背景の色がうすく出る場合は、印刷設定の「背景のグラフィック」をオンにします。", "note"
    ))

    # ===== 9. 履歴 =====
    chapter(s, "9. 履歴（過去の席替え）")
    s.append(p(
        "席替えを実行するたびに、結果が自動で保存されます（クラス×教室ごとに最大20件）。"
        "「履歴」タブから、過去の座席表を確認・復元できます。", "body"
    ))
    s.append(li("<b>大きく見る</b>：過去の座席表をプレビュー表示"))
    s.append(li("<b>復元</b>：その席替えを現在の座席表として呼び戻す"))
    s.append(li("<b>削除</b>：不要な履歴を消す"))

    # ===== 10. バックアップ =====
    chapter(s, "10. バックアップ（外部保存）と復元")
    s.append(p(
        "このアプリのデータは、お使いの端末のブラウザの中だけに保存されます。"
        "そのため、ブラウザのデータ消去・端末の故障や買い替えなどで<b>消えてしまう可能性</b>があります。"
        "万が一に備えて、定期的に外部へバックアップしてください。", "body"
    ))
    s.append(p("10.1 バックアップのしかた（設定タブ）", "h2"))
    s.append(li("<b>「バックアップを共有／送信」</b>（iPad など対応端末）：iCloud Drive・Google Drive・メールなどへワンタップで送れます"))
    s.append(li("<b>「ファイルに書き出す」</b>（パソコンなど）：バックアップ用の JSON ファイルをダウンロードし、クラウドや USB に保存します"))
    s.append(p(
        "しばらくバックアップしていないと、「設定」タブに注意（●印）が表示されます。"
        "学期はじめ・席替えの節目などに保存する習慣をおすすめします。", "note"
    ))
    s.append(p("10.2 復元のしかた", "h2"))
    s.append(step(1, "「設定」タブの<b>「バックアップから復元」</b>を押す"))
    s.append(step(2, "以前に書き出した（または共有した）バックアップファイルを選ぶ"))
    s.append(step(3, "読み込むと、その内容で上書きされ、自動で読み込み直されます"))
    s.append(p(
        "バックアップには、名簿・教室・座席配置・班・条件・座席固定・履歴・デザイン設定など、"
        "<b>このアプリのすべての情報</b>が含まれます。端末の引っ越しにもそのまま使えます。", "note"
    ))
    s.append(p("注意：復元すると、いまのデータはすべて上書きされます。必要なら先に現在のデータをバックアップしてください。", "warn"))

    # ===== 11. データの消去 =====
    chapter(s, "11. データの消去")
    s.append(p(
        "年度末・引き継ぎ・端末の譲渡などのときに、データを消去できます。ヘッダーの「データ消去」"
        "ボタン、または「設定」タブから、消す範囲を選べます。", "body"
    ))
    s.append(make_table([
        ["メニュー", "消える範囲"],
        ["履歴のみ", "過去の席替え履歴だけを消す"],
        ["名簿のみ", "生徒名簿だけを消す"],
        ["教室のみ", "教室・座席配置だけを消す"],
        ["全データ完全消去", "このアプリの保存内容をすべて消し、初期状態に戻す"],
    ], col_widths=[40 * mm, 134 * mm]))
    s.append(p("注意：消去したデータは元に戻せません。「全データ完全消去」の前には、必ずバックアップを取ってください。", "warn"))

    # ===== 12. 困ったときに =====
    chapter(s, "12. 困ったときに（よくある質問）")
    s.append(make_table([
        ["症状・質問", "対処"],
        ["Excel名簿が読み込めない", "ファイル形式が .xlsx か、列（出席番号・氏名・ふりがな・性別）がそろっているか確認してください。"],
        ["「男女がそろわない」と警告が出る", "男子席・女子席の数と、名簿の男子・女子の人数が一致しているか確認してください（第5章）。"],
        ["条件を付けたら席替えできない", "条件が多すぎる・矛盾している可能性があります。条件タブで設定を減らすか見直してください。"],
        ["ルーレットが全画面にならない", "ボタン操作をきっかけにのみ全画面になります。対応ブラウザ（Chrome／Edge／Safari）でお試しください。解除はEscキーでも可能です。"],
        ["印刷で座席表がはみ出す", "用紙をA4・横向きにし、余白を「標準」にしてください。色を出したいときは「背景のグラフィック」をオンにします。"],
        ["データが消えてしまった", "バックアップファイルがあれば「復元」で戻せます。日ごろから外部へのバックアップをおすすめします。"],
        ["iPadで使いたい", "Safari で index.html を開いてください。動作確認済みです。"],
    ], col_widths=[52 * mm, 122 * mm]))

    # ===== 13. 動作環境・データの保管 =====
    chapter(s, "13. 動作環境・データの保管について")
    s.append(make_table([
        ["項目", "内容"],
        ["形式", "単一HTMLファイル（インストール・サーバー不要）"],
        ["ブラウザ", "Google Chrome／Microsoft Edge／Safari"],
        ["端末", "パソコン・iPad（Safari）で動作確認済み"],
        ["データ保存先", "お使いの端末のブラウザ内（localStorage）。外部への送信はありません"],
        ["インターネット", "Excel読み込み・JPG保存の部品を配信サイトから読み込むため、初回などは接続が必要です"],
    ], col_widths=[34 * mm, 140 * mm]))
    s.append(p("データの保管についての注意", "h2"))
    s.append(li("データは<b>端末ごと・ブラウザごと</b>に保存されます。別の端末には自動では引き継がれません（バックアップで移行します）"))
    s.append(li("ブラウザの「閲覧データの消去」を行うと、アプリのデータも消えます"))
    s.append(li("プライベートブラウズ（シークレットモード）では、閉じるとデータが残りません"))
    s.append(li("大切なデータは、第10章のバックアップで定期的に外部へ保存してください"))

    s.append(Spacer(1, 10 * mm))
    s.append(hr())
    s.append(Spacer(1, 4 * mm))
    s.append(p(
        "本マニュアルおよびアプリは教育用途で自由にお使いいただけます。"
        "改善のご要望・不具合のご報告は、お手元の管理者または開発者までお知らせください。",
        "note",
    ))

    return s


def main():
    out = "席替えアプリ_操作マニュアル.pdf"
    doc = BaseDocTemplate(
        out, pagesize=A4,
        leftMargin=18 * mm, rightMargin=18 * mm,
        topMargin=18 * mm, bottomMargin=20 * mm,
        title="席替えアプリ 操作マニュアル",
        author="席替えアプリ 教材",
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
