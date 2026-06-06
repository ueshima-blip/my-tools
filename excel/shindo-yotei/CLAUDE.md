# 進度予定表システム — Claude 向け作業ガイド

このリポジトリは中学校教諭用の進度予定表（Excel マクロ有効ブック）です。VBA を含むバイナリ `.xlsm` を Python スクリプト経由で安全に編集します。

## ファイル構成

```
進度予定表システム_鳴海.xlsm            # 本体（編集対象）
進度予定表システム_鳴海.xlsm.backup     # オリジナル（触らない）
進度予定表システム_鳴海.xlsm.before_*   # 過去スナップショット（触らない）
vba/                                # VBA ソース（Source of truth）
  ThisWorkbook.cls
  Sheet1.cls
  Module1.bas
  _meta.json                        # textoffset / target_compressed_size など
scripts/
  ms_ovba_compress.py               # MS-OVBA 圧縮ライブラリ
  extract_vba.py                    # xlsm → vba/*.bas (.cls)
  build_vba.py                      # vba/ → xlsm（VBA だけを書き戻し）
```

## 標準ワークフロー（VBA を直す場合）

1. **最新版を取り込む**
   ```bash
   python3 scripts/extract_vba.py
   ```
   `vba/*.bas` / `vba/*.cls` が現在の VBA で上書きされる。

2. **`vba/Module1.bas` や `vba/ThisWorkbook.cls` を編集**
   - ふつうのテキスト編集でよい（cp932 / Shift-JIS で書く）
   - VBA の構文・属性（`Attribute VB_Name =` 等）は触らない

3. **ビルドして `.xlsm` を更新**
   ```bash
   python3 scripts/build_vba.py
   ```

4. **コミット & プッシュ**
   ```bash
   git add 進度予定表システム_鳴海.xlsm vba/
   git commit -m "fix: ..."
   git push
   ```

## VBA を編集する際の注意

- **サイズ上限**: `build_vba.py` は元の OLE ストリームと同じバイト数で書き戻す制約があるので、`_meta.json` の `target_compressed_size` を超えるとビルドが失敗する
  - その場合は別の場所のコードを削るか、`.xlsm.backup` から作り直す
- **VBA キャッシュ無効化は自動**: build 時に `_VBA_PROJECT` のバージョン欄と `__SRP_*`、各モジュールの PerformanceCache をゼロクリアする。Excel が次回開いたときソースから再コンパイルする
- **AUTO_PADDING**: ビルド時に末尾にランダムなコメントが付くが、extract で自動的に剥がされる。気にしなくてよい

## ワークシートの内容（時間割や日付）を編集する場合

`xl/worksheets/sheetN.xml` を直接編集する。
- `sheet1.xml` = 基本設定
- `sheet2.xml` = 時間割マスター
- `sheet3.xml` = 週案_テンプレート
- `sheet4.xml`〜 = 第01週、第02週、…

`.xlsm` は単なる ZIP なので、`unzip` → 編集 → `zip` で再パッケージ。
`scripts/build_vba.py` の `os.walk(work_dir)` を参考にできる。

## このシステムで実装済みの機能

1. **総合のカウント** — 集計欄が col=15 までループ
2. **累計を更新の範囲制限** — 現在シートの週番号以下のシートだけ合算
3. **テスト機能**
   - 基本設定 B14 に「テスト」が登録されており、時間割セルのドロップダウンに表示される
   - 「テスト」を選ぶと `Workbook_SheetChange` が `HandleTestEntry` を呼び出し、対象クラスを聞く InputBox を表示
   - セル値を `テスト[,2-1,3-5,]` 形式で記録（前後のカンマで一意マッチを担保）
   - 集計式（PLAN 行）が `COUNTIF(範囲, "*,"&クラス名&",*")` で対象クラスにのみ +1

## Web 版 Claude Code（claude.ai/code）での作業時

リポジトリをクローンしたら最初に下記が必要：

```bash
pip install oletools olefile
```

Setup script（環境設定で自動実行）に書いておくと毎回不要。

## 困ったとき

- `build_vba.py` が `target_compressed_size` 超過で失敗 → 編集量を減らす or 別のモジュールから削る
- `extract_vba.py` で AUTO_PADDING が混じる → `extract_vba.py` の `strip_padding()` のヒューリスティクを調整
- Excel が「VBA プロジェクトを読み込めません」 → `.xlsm.backup` から作り直す
