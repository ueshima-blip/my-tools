# 進度予定表システム（鳴海中用）

中学校教諭用の週単位の進度予定表（Excel マクロ有効ブック）。

## ファイル

| パス | 説明 |
|---|---|
| `進度予定表システム_鳴海.xlsm` | 最新版（テスト機能対応済み） |
| `進度予定表システム_鳴海.xlsm.backup` | 一番最初のオリジナル |
| `進度予定表システム_鳴海.xlsm.before_test_feature` | テスト機能追加前のバージョン |
| `vba/` | VBA ソース（テキスト・コミットされる Source of truth） |
| `scripts/` | xlsm の VBA を抜き出す / 書き戻す Python ツール |
| `CLAUDE.md` | Claude (CLI / Web 版) 向け作業ガイド |
| `SETUP_WEB.md` | Web 版 Claude Code セットアップ手順 |

## 実装済みの機能

1. **総合の集計** — 集計欄ループを col=15 まで拡張
2. **累計の範囲制限** — 現在シートの週番号以下のみ合算
3. **テスト機能** — 時間割セルで「テスト」を選ぶと対象クラスを聞くダイアログが出る。選んだクラスのみ +1 カウント

## 編集ワークフロー

```bash
# 1. 最新の VBA を取り出す
python3 scripts/extract_vba.py

# 2. vba/Module1.bas などをエディタで編集

# 3. .xlsm に反映
python3 scripts/build_vba.py

# 4. コミット
git add 進度予定表システム_鳴海.xlsm vba/
git commit -m "..."
git push
```

詳細は `CLAUDE.md` と `scripts/README.md` を参照。

## 別 PC で作業する

```bash
git clone https://github.com/ueshima-blip/shindo-yotei-narumi.git
cd shindo-yotei-narumi
pip install oletools olefile
```

職場 PC のように何もインストールできない場合は、ブラウザで https://claude.ai/code を使う方法があります。`SETUP_WEB.md` を参照してください。
