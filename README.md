<!--
Copyright 2023 tsukuboshi

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

      http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
-->
# gas-gmail-to-line

## 概要

Gmailの特定のラベルのメールを確認し、未読であればLINE Notifyに通知した上で既読をつけるGoogle Apps ScriptをTypeScriptで実装しています。

## 前提条件

以下のソフトウェアの使用を前提とします。

- Node.js：v20以降
- TypeScript：v5以降

また、GASを実行するためのGoogleアカウントが必要です。  

## デプロイ手順

1. リポジトリをクローンし、ディレクトリを移動

```bash
git clone https://github.com/tsukuboshi/gas-gmail-to-line.git
cd gas-gmail-to-line
```

2. npmパッケージをインストール

```bash
npm init -y
npm install
```

3. asideを初期化しGASプロジェクトを作成

- Project Title：(任意のプロジェクト名)
- `license-header.txt`の上書き：No
- `rollup.config.mjs`の上書き：No
- Script ID：空文字
- Script ID for production environment：空文字

```
npx @google/aside init
✔ Project Title: … gas-gmail-to-line
✔ Adding scripts...
✔ Saving package.json...
✔ Installing dependencies...
license-header.txt already exists
✔ Overwrite … No
rollup.config.mjs already exists
✔ Overwrite … No
✔ Script ID (optional): … 
✔ Script ID for production environment (optional): … 
✔ Creating gas-gmail-to-line...

-> Google Sheets Link: https://drive.google.com/open?id=xxx
-> Apps Script Link: https://script.google.com/d/xxx/edit
```

4. Googleアカウントの認証を実施

```bash
npx clasp login
```

5. 環境変数として、Google Calender IDを配列で設定(複数指定できるよう配列にしているだけなので、配列の中身は単一でも問題ありません)

```bash
export HOME_LABEL_ARRAY="['sampleA','sampleB','sampleC']"
export LINE_TOKEN="xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
```

1. 5/6で指定した環境変数を使用し、GAS実行用ファイルである`src/index.ts`を作成

```bash
cat <<EOF > src/index.ts
import { main } from './example-module';

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function handler(): void {
  main(
    ${HOME_LABEL_ARRAY},
    "${LINE_TOKEN}"
}
EOF
```

8. GASアプリをデプロイ

```bash
npm run deploy
```
