# Web 版 Claude Code セットアップガイド

職場 PC など、何もインストールできない環境からブラウザで Claude Code に作業させる手順。

## 1. プラン確認

Claude Code on the web は **Pro / Max / Team** プラン（リサーチプレビュー）または **Enterprise**（premium / Chat+Claude Code seat）で使えます。
Free プランでは使えません。

## 2. GitHub と連携

ブラウザで https://claude.ai/code を開き、GitHub アカウント (`ueshima-blip`) を連携。

- 連携方法: GitHub App をインストール（推奨）
- 連携後、`ueshima-blip/shindo-yotei-narumi` リポジトリへのアクセスを許可

## 3. 環境設定（Environment）の作成

claude.ai/code 上で新しい Environment を作成し、**Setup script** に以下を貼り付ける:

```bash
#!/bin/bash
pip install --quiet oletools olefile
```

これでセッション起動時に自動で必要ライブラリが入ります。
（一度実行されると filesystem snapshot にキャッシュされ、以降のセッションは即起動）

## 4. タスク開始

claude.ai/code でリポジトリを選び、自然言語でタスクを指示します。例:

> 累計を更新ボタンの挙動を変更してください。今は実施数の合計を出していますが、予定数の合計に切り替えてください。

Claude が自動で:
1. リポジトリをクローン
2. `CLAUDE.md` を読んでプロジェクト構造を把握
3. `python3 scripts/extract_vba.py` で `vba/Module1.bas` を取り出す
4. 該当箇所を編集
5. `python3 scripts/build_vba.py` で `.xlsm` を再構築
6. ブランチを切ってコミット
7. PR を作成

## 5. PR レビューと反映

PR のリンクが返ってくるので、ブラウザで内容を確認してマージ。

マージ後、自宅 PC では次回作業前に:

```bash
git pull
```

## 注意点

- **同時編集の競合**: 自宅と職場で同時に編集すると競合します。作業開始前に必ず最新を取り込む
- **`.xlsm` のサイズ制約**: VBA を大幅追加する場合、`build_vba.py` が `target_compressed_size` 超過でエラーになる可能性あり。その場合は別の場所のコード削減が必要
- **Excel での確認**: Web 版では Excel を実行できません。最終確認はローカル PC で `.xlsm` を開いて行う
- **バイナリ差分**: `.xlsm` は ZIP バイナリなので Git の diff では中身が見えません。PR では `vba/*.bas` の diff を確認することで内容を把握できます
