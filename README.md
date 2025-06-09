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

Gmailの特定のラベルのメールを確認し、未読であればLINE Botに通知した上で既読をつけるGoogle Apps ScriptをTypeScriptで実装しています。

### 主な機能

- 📧 Gmail の特定ラベルの未読メール自動取得
- 🤖 LINE Bot への見やすいフォーマットでの通知
- ⏰ 1時間ごとの自動実行（トリガー設定）
- 🛠️ スプレッドシートでの簡単設定管理
- ✅ 包括的なテストスイート
- 🔧 統合テスト機能

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

5. GASアプリをデプロイ

```bash
npm run deploy
```

6. 以下コマンドでスプレッドシートを開く

```bash
npx clasp open --addon
```

7. スプレッドシートでLINE Bot設定を行う

- A列：LINE Channel Access Token を入力
- B-G列：通知したいGmailラベル名を入力（最大6つ）
- 複数行設定可能（最大4つのトークン）

8. カスタムメニューから初期設定・テストを実行

スプレッドシートで「Gmail to LINE」メニューから以下を実行：

- **設定を初期化**：スプレッドシートのフォーマットを初期化
- **接続・動作確認テスト**：LINE Bot接続とメール取得のテスト
- **メール通知実行**：手動でメール通知を実行
- **通知トリガーを1時間で設定**：自動実行トリガーを設定

## メッセージフォーマット

LINE に送信されるメッセージは以下の見やすい形式で配信されます：

```
📧 送信者：

example@gmail.com

📋 件名：

【重要】会議の議題について

📄 内容：

お疲れ様です。
明日の会議の議題を送付いたします...
```

### フォーマット設定

- **本文制限**：500文字（超過時は「...」で切り詰め）
- **メッセージ全体制限**：5000文字
- **区切り文字**：各セクション間に2行改行

## 開発・テスト

### テスト実行

```bash
# 全テストの実行
npm run test

# Lintチェック
npm run lint

# ビルド
npm run build
```

### 開発時のトラブルシューティング

#### TypeScript バージョン警告

ESLint で TypeScript バージョンの警告が表示される場合がありますが、機能には影響ありません。

#### LINE Bot 接続エラー

1. LINE Channel Access Token が正しく設定されているか確認
2. スプレッドシートの「接続・動作確認テスト」でトークン検証を実行
3. LINE Bot の友達追加が完了しているか確認

#### メール取得エラー

1. Gmail ラベル名が正確に設定されているか確認
2. Gmail API の権限が適切に設定されているか確認

## 設定詳細

### 環境設定値

主要な設定値は `src/env.ts` で管理されています：

- `NUMBER_OF_TOKENS`: 4（設定可能なトークン数）
- `NUMBER_OF_LABELS`: 6（設定可能なラベル数）
- `BODY_MAX_LENGTH`: 500（メール本文の最大文字数）
- `TRIGGER_SETTINGS.INTERVAL_HOURS`: 1（実行間隔）

### LINE Bot API 設定

- **API URL**: `https://api.line.me/v2/bot/message/broadcast`（ブロードキャスト配信）
- **送信形式**: JSON形式のメッセージオブジェクト
- **認証**: Bearer Token方式

## ライセンス

Apache License 2.0
