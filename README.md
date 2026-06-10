# my-tools — 授業で使う自作ツール置き場

これまで作ってきた **ブラウザ／iPad で動く HTML アプリ** や **Excel ツール** を、1 か所にまとめて保存・整理するリポジトリです。作品が増えても探しやすいよう、種類ごとにフォルダを分けています。

## 📁 名前のしくみ（GitHub に慣れていない方へ）

GitHub では、持ち物は必ず `アカウント名 / フォルダ名` の形で表示されます。

**`ueshima-blip` / `my-tools`**

- **`ueshima-blip`**（前）… あなたの **アカウント名（ログイン ID）**。すべての持ち物の前に付く「表札」で、フォルダではありません。
- **`my-tools`**（後ろ）… この **フォルダ（リポジトリ）の名前**。中に下記のツールが入っています。

## 🗂 フォルダ構成

```
my-tools/
├── apps/    ブラウザ／iPad で動く HTML アプリ
│   ├── kanji-note/          かんじノート（漢字まちがい判定つき）
│   ├── genetics-sim/        遺伝のモデル実験シミュレーター
│   ├── ethanol-sim/         エタノール沸騰の実験シミュレーション
│   ├── seat-shuffle/        席替えアプリ（Excel名簿から自動席替え）
│   ├── submission-tracker/  提出物チェックシステム
│   └── hogosha-omr/         保護者会調査票の読み取り（スキャンPDF→CSV）
└── excel/   Excel で作った教材・ツール
    ├── shindo-yotei/        進度予定表システム（鳴海中用・学校共通用）
    └── hogosha-kai/         保護者会調整ツール（希望日程の自動調整）
```

- **1 作品 = 1 フォルダ**。フォルダ名は半角英数（リンクや共有でトラブルが起きにくい）。日本語の正式名称はこの目次で対応づけます。

## 🧰 作品一覧

| 作品名 | 種類 | 内容 | 場所 |
|--------|------|------|------|
| かんじノート | HTMLアプリ（iPad） | Apple Pencil で書ける学習用ノート。書いた漢字の正誤を自動判定。 | [`apps/kanji-note/`](apps/kanji-note/) |
| 遺伝のモデル実験シミュレーター | HTMLアプリ（PC・タブレット） | 中3理科「遺伝の規則性」をブラウザで再現。教員マニュアル(PDF)・記録用紙(Word)つき。 | [`apps/genetics-sim/`](apps/genetics-sim/) |
| エタノール沸騰シミュレーション | HTMLアプリ | エタノールが沸騰する温度を調べる理科の実験シミュレーション。タッチ操作対応。 | [`apps/ethanol-sim/`](apps/ethanol-sim/) |
| 席替えアプリ | HTMLアプリ | クラス名簿(Excel)を取り込み、男女配置・班・距離条件などを考慮して瞬時／ルーレットで席替え。印刷・JPG保存・履歴つき。 | [`apps/seat-shuffle/`](apps/seat-shuffle/) |
| 提出物チェックシステム | HTMLアプリ | クラスの提出物を管理。Excel から生徒を一括登録でき、複数クラス対応。 | [`apps/submission-tracker/`](apps/submission-tracker/) |
| 進度予定表システム（鳴海中用・学校共通用） | Excel（マクロ） | 中学校教諭用の週単位の進度予定表。鳴海中用と学校共通用の 2 種類。VBA を編集するツール・手順つき。 | [`excel/shindo-yotei/`](excel/shindo-yotei/) |
| 保護者会調整ツール | Excel（マクロ） | 保護者会（個人面談）の希望を集めて、ボタン一つで全員の日程を自動調整。個別通知票の自動作成・Googleフォーム/Forms/紙スキャンからの取込・PDF出力つき。 | [`excel/hogosha-kai/`](excel/hogosha-kai/) |
| 調査票読み取り（OMR） | HTMLアプリ | 紙の保護者会調査票をスキャン（PDF/JPG）して自動読み取り、上記Excelツール用のCSVに変換。 | [`apps/hogosha-omr/`](apps/hogosha-omr/) |

## ▶️ ツールの開き方

### HTML アプリ（`apps/` の中）

GitHub の画面でファイルをクリックしても「中身のプログラム（コード）」が出るだけで、教材としては動きません。実際に使うには次のどちらかです。

**方法1：GitHub Pages をオンにする（おすすめ・設定は一度だけ）**

1. 上部メニュー **Settings** → 左側の **Pages**
2. **Branch** を `main` にして **Save**

数分後、次のアドレスでブラウザから直接開けます（リンクを保存しておけば毎回すぐ使えます）。

- かんじノート：`https://ueshima-blip.github.io/my-tools/apps/kanji-note/`
- 遺伝のモデル実験：`https://ueshima-blip.github.io/my-tools/apps/genetics-sim/`
- エタノール沸騰：`https://ueshima-blip.github.io/my-tools/apps/ethanol-sim/`
- 席替えアプリ：`https://ueshima-blip.github.io/my-tools/apps/seat-shuffle/席替えアプリ_オンライン版.html`
- 提出物チェック：`https://ueshima-blip.github.io/my-tools/apps/submission-tracker/`

**方法2：ダウンロードして開く**

フォルダごとダウンロードし、中の `index.html` をダブルクリック（席替えアプリは `席替えアプリ_オフライン版.html`）。

### Excel ツール（`excel/` の中）

`excel/shindo-yotei/` の `.xlsm` をダウンロードして Excel で開きます（マクロを有効化）。編集方法は同フォルダの `README.md` / `CLAUDE.md` を参照してください。

## ➕ 新しい作品を追加するには

### A. Claude Code に頼む（おすすめ）

このリポジトリで Claude Code を開いて、例えばこう伝えてください。

> 「新しいアプリ一式を `apps/作品名/` に入れて、目次にも追加して」

→ フォルダへの保存・この目次への追記・保存（コミット）まで自動でやります。

### B. 自分で置く

1. 入れる場所を決める（HTMLアプリ → `apps/作品名/`、Excel → `excel/作品名/`）
2. ファイルを置く
3. この `README.md` の「作品一覧」に 1 行足す

## ✅ 自動チェック（CI）

変更を提案（プルリクエスト）するたびに、GitHub が `apps/` の各アプリを自動でブラウザに読み込み、**プログラムのエラーで動かなくなっていないか**を確認します。結果はプルリクエスト画面に ✓（成功）／✗（失敗）で表示されます。配付や共有の前の「うっかり壊し」を防ぐための仕組みです。

- 仕組み：[`.github/workflows/ci.yml`](.github/workflows/ci.yml) と [`scripts/smoke-test-apps.mjs`](scripts/smoke-test-apps.mjs)
- 手元で同じ確認をする場合（任意）：このフォルダで `npm install` のあと `npm test`
- チェック対象は自動で増えます（`apps/◯◯/` に `.html` を置けば次回から自動でチェックされます）

## 🔒 メモ（GitHub に不慣れな方へ）

- **公開設定**：教材を人に見せたくない場合は、このリポジトリを **Private（非公開）** にできます（Settings）。
- **名前の変更**：フォルダ名（リポジトリ名）`my-tools` は Settings の一番上で、アカウント名 `ueshima-blip` はアカウントの Settings → Account で変更できます（影響が大きいので、変更前に一度ご相談ください）。
