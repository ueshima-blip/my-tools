"""完全オフライン版（席替えアプリ.html）生成スクリプト

index.html が参照している外部部品（SheetJS / html2canvas）をHTMLに直接埋め込み、
Googleフォントの外部読み込みを外して、インターネット不要の単一ファイル
「席替えアプリ.html」を作る。人に配付する用。

部品は、同じフォルダにローカルファイルがあればそれを使い、無ければCDNから取得する：
  - xlsx.full.min.js    （SheetJS：Excel読み込み・名簿テンプレート生成）
  - html2canvas.min.js  （JPG保存）
ローカルに置く場合は、index.html が参照しているのと同じバージョンを置くこと。
"""
import os
import re
import sys
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "index.html")
OUT = os.path.join(HERE, "席替えアプリ.html")

LIBS = [
    ("https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js", "xlsx.full.min.js"),
    ("https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js", "html2canvas.min.js"),
]
FONT_LINK = ('<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+JP'
             ':wght@400;500;700&display=swap" rel="stylesheet">')


def fetch(url, fname):
    local = os.path.join(HERE, fname)
    if os.path.exists(local):
        print(f"  ローカルを使用: {fname}")
        with open(local, encoding="utf-8") as f:
            return f.read()
    print(f"  ダウンロード: {url}")
    with urllib.request.urlopen(url, timeout=60) as r:
        return r.read().decode("utf-8")


def main():
    with open(SRC, encoding="utf-8") as f:
        html = f.read()

    # Googleフォントの外部読み込みを削除（端末の標準の日本語フォントで表示）
    html = html.replace(FONT_LINK + "\n", "").replace(FONT_LINK, "")

    # 各ライブラリの <script src=...> をインライン化
    for url, fname in LIBS:
        code = fetch(url, fname)
        if "</script" in code.lower():
            sys.exit(f"ERROR: {fname} に </script> が含まれ、インライン化できません")
        tag = f'<script src="{url}"></script>'
        if tag not in html:
            sys.exit(f"ERROR: 置換対象のタグが見つかりません: {tag}")
        html = html.replace(tag, "<script>\n" + code + "\n</script>")

    with open(OUT, "w", encoding="utf-8") as f:
        f.write(html)

    leftover = re.findall(r'(?:src|href)="https?://[^"]+"', html)
    print(f"Generated: {OUT}  ({os.path.getsize(OUT) // 1024} KB)")
    if leftover:
        print("WARNING: 外部参照が残っています:", leftover)
    else:
        print("OK: 外部参照なし（完全オフライン）")


if __name__ == "__main__":
    main()
