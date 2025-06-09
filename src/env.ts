/**
 * Copyright 2023 tsukuboshi
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *       http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

// LINE Messaging API設定
export const LINE_API_BASE_URL = 'https://api.line.me/v2/bot/message/broadcast';

// LINEチャネルアクセストークンの個数
export const NUMBER_OF_TOKENS = 4;

// Gmailラベルの個数
export const NUMBER_OF_LABELS = 6;

// メール本文の最大文字数
export const BODY_MAX_LENGTH = 500;

// スプレッドシートの列定義
export const SPREADSHEET_COLUMNS = {
  LINE_TOKEN: 1, // A列: LINEチャネルアクセストークン
  LABEL_START: 2, // B列以降: Gmailラベル名
} as const;

// スプレッドシートの行定義
export const SPREADSHEET_ROWS = {
  HEADER_ROW: 1, // 1行目: ヘッダー
  DATA_START_ROW: 2, // 2行目以降: データ
} as const;

// Gmail検索設定
export const GMAIL_SEARCH = {
  UNREAD_CONDITION: 'is:unread', // 未読メール条件
  LABEL_FORMAT: 'label:%s', // ラベル形式
} as const;

// トリガー設定
export const TRIGGER_SETTINGS = {
  INTERVAL_HOURS: 1, // トリガー実行間隔（時間）
  FUNCTION_NAME: 'main', // トリガー対象関数名
} as const;

// シート設定
export const SHEET_SETTINGS = {
  DEFAULT_NAME: 'プロパティ', // デフォルトシート名
  TOKEN_HEADER_COLOR: 'lightgreen', // トークンヘッダー背景色
  LABEL_HEADER_COLOR: 'lightyellow', // ラベルヘッダー背景色
  BORDER_COLOR: 'black', // 枠線色
} as const;

// メッセージフォーマット設定
export const MESSAGE_FORMAT = {
  SEPARATOR: '\n\n', // セクション区切り文字
  SUBJECT_PREFIX: '件名：', // 件名プレフィックス
  CONTENT_PREFIX: '内容：', // 内容プレフィックス
} as const;

// API呼び出し設定
export const API_SETTINGS = {
  REQUEST_DELAY: 100, // API呼び出し間の待機時間（ミリ秒）
  MAX_RETRIES: 3, // 最大リトライ回数
  TIMEOUT: 30000, // タイムアウト時間（ミリ秒）
} as const;

// バリデーション設定
export const VALIDATION = {
  MIN_TOKEN_LENGTH: 40, // LINEトークンの最小文字数
  VALID_TOKEN_PATTERN: /^[A-Za-z0-9+/=]+$/, // 有効なトークン文字パターン
  MAX_MESSAGE_LENGTH: 5000, // LINEメッセージの最大文字数
} as const;
