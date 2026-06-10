# 保護者会調整ツール — Claude 向け作業ガイド

中学校の保護者会（個人面談）日程を自動調整する Excel マクロブック。
フリーソフト ver1.11 を全面改修した ver2.0。**Excel の無い環境で開発する**ための
ツール一式がそろっている（shindo-yotei と同系だが、こちらは**サイズ制約なし**）。

## ファイル構成

```
保護者会調整_ver200.xlsm          # 本体（成果物）
保護者会調整_ver111_オリジナル.xlsm # 改造前（絶対に触らない）
vba/                              # VBA ソース（UTF-8/LF・Source of truth）
  Module1.bas   # エンジン・通知票・予定表・PDF・ボタン入口
  Module2.bas   # 取込(CSV/xlsx/OMR)・調査票印刷・クイック入力
  Sheet1.cls    # Start!（ほぼ空）
  Sheet2.cls    # 希望入力（行ハイライト・入れ替えモード）※調整シートにコピーされる
  Sheet3/5.cls  # 名簿・元_予定表（空）
  UserForm1-4.frm, ThisWorkbook.cls
scripts/
  cfb.py                 # OLE(CFB)リパッカ — ストリームサイズ変更可・要の技術
  ms_ovba_compress.py    # MS-OVBA 圧縮
  extract_vba.py         # xlsm → vba/（cp932→UTF-8 変換）
  build_vba.py           # vba/ → xlsm（UTF-8→cp932、コンテナ再構築）
  lint_vba.py            # 簡易リント（未宣言変数・ブロック対応）
  test_engine_logic.py   # 調整エンジンの Python ミラー + 総当たり検証
  patch_workbook_xml.py  # ActiveX除去等の一回きり手術（適用済み・冪等）
```

## 標準ワークフロー（VBA を直す）

```bash
pip install oletools olefile          # 初回のみ（Setup script 推奨）
# 1. vba/*.bas|cls|frm を編集（UTF-8 のまま）
python3 scripts/lint_vba.py           # 2. リント（必ず通す）
python3 scripts/build_vba.py          # 3. ビルド（往復検証つき）
python3 scripts/test_engine_logic.py  # 4. エンジンを触ったら必ず
```

`extract_vba.py` は xlsm 側が正のとき（手動で Excel から変更が入ったとき）だけ使う。

## 重要な設計知識

- **モジュール構成（名前と個数）は変えない**。dir ストリームと PROJECT ストリームに
  触れずに済む唯一の方法。新コードは Module1/Module2 に追記する。
  UserForm の**デザイン（コントロール）も変更不可**（designer ストリームのため）。
  新しい UI はシート上の図形ボタン + InputBox/MsgBox で作る。
- ビルドは PerformanceCache をゼロ埋め（textoffset 維持）、__SRP_* ゼロ埋め、
  _VBA_PROJECT のバージョン欄クリア → Excel が次回起動時にソースから再コンパイル。
- **調整エンジン**は scripts/test_engine_logic.py の `solve()` と 1:1 対応。
  VBA を変えるときは Python ミラーも同じ変更をして 4000 ケース検証を通すこと。
  フェーズ: ○のみ → 未確定者の△ → 全△。失敗時は Hall violator 集合を診断表示。
- **セルの意味**: 希望入力/調整シート。FIRST=D2（時刻行）、日付=1行目、生徒=3..52行。
  値: ○◎△。色: kColor24=休憩, mColor33=確定, pColor47=◎確定。
  補助列: BL=RowNow, BM=RowFail(1=未入力,2=問題), BN-BQ=集計, **BS(71)=確定コマ裏データ**
  （行ごとのコマ番号 j=(d-1)*kk+k、ヘッダ行に "STATE2"）。
- **名前解決のトリック**: 調整シートは 希望入力 のコピーなので、FIRST 等の名前は
  シートローカル名としてコピーされ、`Range("FIRST")` はアクティブシートで解決される。
  コードはこの性質に依存している（v1.11 から）。
- 生徒数 = `MAX(名簿!A2:A51)`（人数でなく最大番号）。コマ数セルは浮動小数
  （10.999…）なので読み込みは丸める（EnsureDims 参照）。
- 図形ボタン: drawing1.xml(Start!)/drawing2.xml(希望入力) に `macro="[0]!BtnXXX"`。
  希望入力の「メニュー」ボタンは SheetCopy で調整シートに複製される。

## OMR 連携（apps/hogosha-omr/）

- Excel 側 `MakeSurveySheets` が調査票 + `調査票定義.csv`（実測ジオメトリ・pt単位・
  左上マーカー原点）を出力。Web 側 `omr.js` がマーカー検出→双一次写像→読み取り。
- ID は B列の縦8ビット帯（b6..b0+奇数パリティ）。下マーカーは上2つから**予測**して
  探す（フォーム高さがコマ数依存のため画像下端にあるとは限らない）。
- 出力 CSV はマトリクス形式 `番号,氏名,<6/20 15:30>...` で、ImportWishes が
  Googleフォーム等と同じ経路で取り込む。
- テスト: `cd apps/hogosha-omr/test && python3 make_synthetic.py && node test_omr.js`
  （回転±3°・100/150/200dpi・薄い鉛筆の合成スキャンで検証）

## 検証（Excel が無い環境での確認手段）

```bash
python3 scripts/lint_vba.py                       # コンパイルエラー予防
python3 scripts/build_vba.py                      # 内蔵の往復検証
olevba 保護者会調整_ver200.xlsm | head            # パース可能か
soffice --headless --convert-to xlsx --outdir /tmp 保護者会調整_ver200.xlsm  # 構造検証
```

実機 Excel での最終確認はユーザーに依頼する（マクロ有効化が必要）。
