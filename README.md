# Color Picker Add-in for PowerPoint

## 概要

PowerPoint 用のカラーピッカーアドインです。使いたい色のスウォッチをクリックすると、HEXコードがクリップボードにコピーされます。

## 動作環境

- PowerPoint for Microsoft 365（Windows）

## GitHub Pages を使ったセットアップ

この方法は、一度設定すれば、どのPCからでも使えます。

### ステップ1：GitHub にリポジトリを作成してファイルをアップロード

1. [GitHub](https://github.com) にログイン（アカウントがなければ作成）
2. 新しいリポジトリを作成（例：`color-picker-addin`）
3. このフォルダのファイルをすべてリポジトリにプッシュ

```bash
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/YOUR-USERNAME/color-picker-addin.git
git push -u origin main
```

### ステップ2：GitHub Pages を有効化

1. リポジトリの **Settings** → **Pages** を開く
2. **Source** を `Deploy from a branch` に設定
3. Branch を `main` / `/(root)` に設定して **Save**
4. しばらく待つと以下の URL が発行される

```
https://YOUR-USERNAME.github.io/color-picker-addin/
```

### ステップ3：manifest.xml の URL を書き換える

`manifest.xml` 内の `https://localhost:3000` をすべて GitHub Pages の URL に置換します。

**例：** `YOUR-USERNAME` が `myname` の場合、

```
https://localhost:3000  →  https://myname.github.io/color-picker-addin
```

書き換えたら、変更を GitHub にプッシュします。

```bash
git add manifest.xml
git commit -m "Update URLs to GitHub Pages"
git push
```

### ステップ4：PowerPoint にアドインを登録（初回のみ）

`manifest.xml` ファイル自体はローカルに置いたままにしておきます。PowerPoint はこのファイルを経由してアドインを読み込み、ファイルの内部に書かれた GitHub Pages の URL からコンテンツを取得します。

PowerPoint に manifest.xml の場所（フォルダ）を登録する手順：

1. **フォルダを共有する**
   - このフォルダ（`color-picker-addin`）を右クリック → **プロパティ** → **共有** タブ
   - **共有** をクリック → 自分のアカウントを選択して **追加** → **共有** → OK
   - 「ネットワーク パスとして共有されています」欄に表示された UNCパスをコピーする
   - 例：`\\PCNAME\Users\username\...\color-picker-addin`

2. **PowerPoint のトラスト センターに登録する**
   - **[ファイル] → [オプション] → [トラスト センター] → [トラスト センターの設定]**
   - **[信頼できるアドイン カタログ]** を選択
   - 「カタログのURL」にコピーした UNCパスを貼り付けて **[カタログの追加]**
   - **「メニューに表示する」にチェック** → OK

3. PowerPoint を**再起動**

### ステップ5：アドインを起動する

**[挿入] → [アドイン] → [個人用アドイン] → [共有フォルダ]** → **Color Picker** を追加

> **[アドイン] が見つからない場合：** 画面上部の検索ボックスに「アドイン」と入力してください。


## 使い方

- サイドパネルに 14 色 × 6 段階（明→暗）のカラーパレットが表示されます
- スウォッチにカーソルを合わせると、色名・HEXコード・RGB値がポップアップ表示されます
- スウォッチをクリックすると HEXコード（例：`#F5C000`）がクリップボードにコピーされます


## 色のカスタマイズ

`taskpane.js` の `BASE_COLORS` 配列を編集すると基準色を変更できます。

```javascript
const BASE_COLORS = [
    { hex: '#5C7A8C', name: 'Steel Blue Gray' },
    { hex: '#7A7A7A', name: 'Gray' },
    { hex: '#F5C000', name: 'Golden Yellow' },
    // ...
];
```

基準色を変更すると、明るい3段階・暗い2段階も自動で計算されます。
変更後は GitHub にプッシュするだけで反映されます（サーバー再起動不要）。


## ローカルサーバーでの開発（オプション）

GitHub Pages を使わずにローカルで動かす方法です。Node.js とローカルサーバーを使用します。
HTTPS 証明書が必要なため、初回はいくつかのセットアップ手順があります。

### 必要なもの

- [Node.js](https://nodejs.org)（インストール後、ターミナルで `node -v` が表示されれば OK）

### 初回セットアップ

#### 1. SSL 証明書のインストール（管理者権限が必要）

PowerPoint の Office アドインは HTTPS での配信が必須です。ローカルで HTTPS を使うために自己署名証明書をインストールします。

**管理者としてターミナルを起動**（スタートメニューで「PowerShell」または「cmd」を右クリック → 管理者として実行）し、このフォルダに移動してから実行：

```bash
npm run setup
```

これにより `%USERPROFILE%\.office-addin-dev-certs\` に証明書ファイルが生成され、Windows の信頼済み証明書ストアに登録されます。

> **証明書の有効期限は 365 日です。** 期限が切れたら再度 `npm run setup`（管理者）を実行してください。

#### 2. manifest.xml の URL がローカル用になっていることを確認

`manifest.xml` 内の URL がすべて `https://localhost:3000` になっていることを確認します。
GitHub Pages 用に書き換えた場合は `localhost:3000` に戻してください。

#### 3. PowerPoint にアドインを登録（初回のみ）

GitHub Pages と同様に、UNCパス経由で `manifest.xml` のフォルダを PowerPoint に登録します。

1. このフォルダ（`color-picker-addin`）を右クリック → **プロパティ** → **共有** タブ → **共有**
2. 表示された UNCパスをコピー（例：`\\PCNAME\Users\username\...\color-picker-addin`）
3. **[ファイル] → [オプション] → [トラスト センター] → [トラスト センターの設定]**
4. **[信頼できるアドイン カタログ]** → UNCパスを入力 → **[カタログの追加]**
5. **「メニューに表示する」にチェック** → OK → PowerPoint を**再起動**

### サーバーの起動と使用

```bash
# サーバー起動（毎回）
npm start
```

サーバーが起動したら PowerPoint を開き、**[挿入] → [アドイン] → [個人用アドイン] → [共有フォルダ]** から **Color Picker** を追加します。

使用終了後は `Ctrl + C` でサーバーを停止します。

> ローカルサーバーを使う場合は、アドインを使用するたびにサーバーを起動しておく必要があります。

### アイコンの再生成

```bash
npm run icons
```


## ファイル構成

| ファイル | 説明 |
|---|---|
| `manifest.xml` | アドイン定義ファイル（PowerPoint に登録するファイル） |
| `taskpane.html` | アドインのUI |
| `taskpane.css` | スタイル |
| `taskpane.js` | ロジック（色データ・コピー処理） |
| `commands.html` | Office コマンド用プレースホルダー |
| `favicon.svg` | ブラウザタブ用アイコン |
| `icon-16/32/80.png` | PowerPoint リボン用アイコン |
| `generate-icons.js` | アイコン PNG 生成スクリプト |

## 開発者について
- [Hayate.H](https://github.com/cloud8high/profile) with Claude Code