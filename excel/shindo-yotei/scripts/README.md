# scripts/

VBA を含む `.xlsm` を安全に編集するための Python ツール集。

## 必要なライブラリ

```bash
pip install oletools olefile
```

Web 版 Claude Code で使う場合は Setup script に書いておくと毎セッション自動で入ります。

## ファイル

| ファイル | 役割 |
|---|---|
| `ms_ovba_compress.py` | MS-OVBA 圧縮アルゴリズム（VBA ソース埋め込み用） |
| `extract_vba.py` | `.xlsm` → `vba/*.bas` (.cls) にソースを取り出す |
| `build_vba.py` | `vba/*.bas` (.cls) → `.xlsm` に書き戻す |

## 典型的なワークフロー

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

## 仕組み

### extract_vba.py
1. `.xlsm` (実体は ZIP) を読み、`xl/vbaProject.bin` を取り出す
2. vbaProject.bin は OLE 複合ドキュメント。`olefile` でストリームを読む
3. `VBA/dir` を MS-OVBA 解凍し、各モジュール（Module1, ThisWorkbook, Sheet1）の `textoffset`（圧縮ソース開始位置）を取得
4. 各モジュールストリームの `[textoffset:]` を解凍 → cp932 でテキスト化
5. 末尾のオートパディング（旧版で挿入したランダムコメント、現版の MS-OVBA 空チャンクで挿入される末尾空白）をヒューリスティクで除去
6. `vba/<Name>.<ext>` に保存。`vba/_meta.json` に各モジュールの `textoffset` / `target_compressed_size` / `stream_size` を記録

### build_vba.py
1. `vba/_meta.json` から各モジュールのメタ情報を読む
2. `vba/<Name>.<ext>` を読んで MS-OVBA で圧縮
3. **サイズが target に満たない場合のパディング**
   - 元の OLE ストリームと同じバイト数で書き戻す必要があるため、`compressed_size` を `target_compressed_size` に **完全一致** させる必要がある
   - 戦略 (2段階):
     - **ソース側**: ランダムなコメント行を末尾に追加。圧縮後サイズが target にできるだけ近く、かつ gap が 0 または 3 以上になるよう探索
     - **圧縮側**: 残った 0/3+バイトの隙間を MS-OVBA の「空チャンク」（3 バイトで 0 バイト解凍）と微小チャンク（4-5 バイトで 1-2 文字解凍）で埋める
4. `olefile.write_stream()` で vbaProject.bin に書き戻し
5. **キャッシュ無効化**: 各モジュールの PerformanceCache、`_VBA_PROJECT` のバージョン欄、`__SRP_*` ストリームを全部ゼロクリア。Excel が次に開いたときソースから再コンパイルする
6. xl/* 全体を再 ZIP して `.xlsm` に上書き

## サイズ上限に当たったら

`build_vba.py` は `compressed_size > target_compressed_size` でエラーになります。
このとき選択肢は:

1. **編集量を減らす**: 不要なコメント・空白を削る、ロジックを圧縮する
2. **別のモジュールから領域を分けてもらう**: 大きい修正は Module1 以外に分散
3. **ベースラインから作り直す**: `.xlsm.backup` をコピーしてから再構築（履歴がリセットされる）

target サイズは元の `.xlsm` を Excel が保存した際に決まったストリーム長です。Excel で `.xlsm` を一度開いて保存し直すと、現状のソースサイズに合わせて target が広げ直されます（手元 PC で Excel が使える場合の回避策）。
