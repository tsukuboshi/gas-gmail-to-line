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

import {
  API_SETTINGS,
  BODY_MAX_LENGTH,
  GMAIL_SEARCH,
  LINE_API_BASE_URL,
  MESSAGE_FORMAT,
  NUMBER_OF_LABELS,
  NUMBER_OF_TOKENS,
  SHEET_SETTINGS,
  SPREADSHEET_COLUMNS,
  SPREADSHEET_ROWS,
  TRIGGER_SETTINGS,
  VALIDATION,
} from './env';

// --- 型定義 ---
interface GmailLabelConfig {
  lineToken: string;
  labelNames: string[];
}

interface FormattedMessage {
  from: string;
  subject: string;
  body: string;
  formattedText: string;
}

interface ValidationResult {
  isValid: boolean;
  message: string;
}

// --- スプレッドシートから設定情報を取得 ---
function getConfigFromSheet(): GmailLabelConfig[] {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const configs: GmailLabelConfig[] = [];

    console.log('スプレッドシートから設定情報を取得中...');

    // ヘッダー行の次の行からLINEチャネルアクセストークン毎に設定を取得
    for (
      let i = SPREADSHEET_ROWS.DATA_START_ROW;
      i <= NUMBER_OF_TOKENS + 1;
      i++
    ) {
      // A列のセルよりLINEチャネルアクセストークンを取得
      const lineTokenRaw = sheet
        .getRange(i, SPREADSHEET_COLUMNS.LINE_TOKEN)
        .getValue();

      if (lineTokenRaw) {
        const lineToken = String(lineTokenRaw)
          .trim()
          .replace(/[\r\n]/g, '');

        // トークンの検証
        const tokenValidation = validateLineToken(lineToken);
        if (!tokenValidation.isValid) {
          console.warn(
            `行${i}のLINEトークンが無効です: ${tokenValidation.message}`
          );
          continue;
        }

        // 該当行のB列以降のラベル名を取得
        const labelNamesRaw = sheet
          .getRange(i, SPREADSHEET_COLUMNS.LABEL_START, 1, NUMBER_OF_LABELS)
          .getValues()
          .flat()
          .filter(label => label && String(label).trim() !== '');

        const labelNames = labelNamesRaw.map(label => String(label).trim());

        if (labelNames.length > 0) {
          configs.push({
            lineToken,
            labelNames,
          });

          console.log(
            `行${i}: トークン設定済み, ラベル数: ${labelNames.length}`
          );
          console.log(`  ラベル: ${labelNames.join(', ')}`);
        } else {
          console.warn(
            `行${i}: トークンは設定されているが、ラベルが設定されていません`
          );
        }
      }
    }

    if (configs.length === 0) {
      throw new Error(
        '有効な設定が見つかりません。LINEトークンとGmailラベルを確認してください。'
      );
    }

    console.log(`有効な設定数: ${configs.length}`);
    return configs;
  } catch (error) {
    console.error('設定情報の取得中にエラーが発生しました:', error);
    throw error;
  }
}

// --- LINE Token形式チェック関数 ---
function validateLineToken(token: string): ValidationResult {
  if (!token || token.trim() === '') {
    return { isValid: false, message: 'トークンが空です' };
  }

  if (token.length < VALIDATION.MIN_TOKEN_LENGTH) {
    return {
      isValid: false,
      message: `トークンが短すぎます (${token.length}文字)`,
    };
  }

  if (!VALIDATION.VALID_TOKEN_PATTERN.test(token)) {
    return { isValid: false, message: '無効な文字が含まれています' };
  }

  return { isValid: true, message: 'トークン形式は正常です' };
}

// --- メールを取得する関数 ---
function fetchHomeMail(homeLabel: string): FormattedMessage[] {
  try {
    console.log(`ラベル "${homeLabel}" の未読メールを検索中...`);

    // 未読メールの検索条件
    const searchTerms = Utilities.formatString(
      `${GMAIL_SEARCH.LABEL_FORMAT} ${GMAIL_SEARCH.UNREAD_CONDITION}`,
      homeLabel
    );

    console.log(`検索条件: ${searchTerms}`);

    // 未読メールのスレッドを取得
    const fetchedThreads = GmailApp.search(searchTerms);
    console.log(`見つかったスレッド数: ${fetchedThreads.length}`);

    if (fetchedThreads.length === 0) {
      console.log(`ラベル "${homeLabel}" に未読メールはありません`);
      return [];
    }

    // 未読メールの取得
    const fetchedMessages = GmailApp.getMessagesForThreads(fetchedThreads);

    // 送信メッセージ配列
    const formattedMessages: FormattedMessage[] = [];

    // 送信メッセージ配列に詰める
    fetchedMessages.forEach((messageArray, threadIndex) => {
      messageArray.forEach((message, messageIndex) => {
        // メッセージが未読の場合
        if (message.isUnread()) {
          console.log(
            `処理中: スレッド${threadIndex + 1}のメッセージ${messageIndex + 1}`
          );

          // メッセージを整形
          const formattedMessage = gmailToFormattedMessage(message);
          formattedMessages.push(formattedMessage);

          // メッセージを既読にマーク
          message.markRead();
          console.log(
            `メッセージを既読にマークしました: "${formattedMessage.subject}"`
          );
        }
      });
    });

    // LINE用に配列の順番を反転させる（新しいメールが上に来るように）
    const reversedMessages = formattedMessages.reverse();
    console.log(
      `ラベル "${homeLabel}" で ${reversedMessages.length}件の未読メールを処理しました`
    );

    return reversedMessages;
  } catch (error) {
    console.error(
      `ラベル "${homeLabel}" のメール取得中にエラーが発生しました:`,
      error
    );
    throw error;
  }
}

// --- メールを整形する関数 ---
function gmailToFormattedMessage(
  mail: GoogleAppsScript.Gmail.GmailMessage
): FormattedMessage {
  try {
    // メールの情報を取得
    const from = mail.getFrom();
    const subject = mail.getSubject();
    const bodyRaw = mail.getPlainBody();

    // 本文を制限文字数でカット
    const body =
      bodyRaw.length > BODY_MAX_LENGTH
        ? bodyRaw.slice(0, BODY_MAX_LENGTH) + '...'
        : bodyRaw;

    // フォーマット済みテキストを作成
    const formattedText = Utilities.formatString(
      '\n%s%s%s%s%s%s%s%s%s%s%s',
      MESSAGE_FORMAT.SENDER_PREFIX,
      MESSAGE_FORMAT.SEPARATOR,
      from,
      MESSAGE_FORMAT.SEPARATOR,
      MESSAGE_FORMAT.SUBJECT_PREFIX,
      MESSAGE_FORMAT.SEPARATOR,
      subject,
      MESSAGE_FORMAT.SEPARATOR,
      MESSAGE_FORMAT.CONTENT_PREFIX,
      MESSAGE_FORMAT.SEPARATOR,
      body
    );

    // メッセージ長を検証
    if (formattedText.length > VALIDATION.MAX_MESSAGE_LENGTH) {
      console.warn(
        `メッセージが長すぎます (${formattedText.length}文字), 切り詰めます`
      );
      const truncatedText =
        formattedText.slice(0, VALIDATION.MAX_MESSAGE_LENGTH - 3) + '...';
      return {
        from,
        subject,
        body,
        formattedText: truncatedText,
      };
    }

    return {
      from,
      subject,
      body,
      formattedText,
    };
  } catch (error) {
    console.error('メールの整形中にエラーが発生しました:', error);
    throw error;
  }
}

// --- LINE Botヘ送信する関数 ---
function sendLine(message: string, channelAccessToken: string): void {
  try {
    console.log('LINEメッセージ送信中...');
    console.log(`メッセージ長: ${message.length}文字`);

    if (!message || message.trim() === '') {
      throw new Error('送信するメッセージが空です');
    }

    if (!channelAccessToken || channelAccessToken.trim() === '') {
      throw new Error('LINE Channel Access Tokenが設定されていません');
    }

    // 送信内容
    const payload = {
      messages: [
        {
          type: 'text',
          text: message,
        },
      ],
    };

    // 送信オプション
    const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + channelAccessToken,
        'Content-Type': 'application/json',
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true, // HTTPエラーでも例外を投げない
    };

    // 送信
    const response = UrlFetchApp.fetch(LINE_API_BASE_URL, options);

    if (response.getResponseCode() !== 200) {
      throw new Error(
        `LINE API エラー: ${response.getResponseCode()} - ${response.getContentText()}`
      );
    }

    console.log('LINEメッセージ送信成功');

    // API制限を考慮して少し待機
    if (API_SETTINGS.REQUEST_DELAY > 0) {
      Utilities.sleep(API_SETTINGS.REQUEST_DELAY);
    }
  } catch (error) {
    console.error('LINE メッセージ送信エラー:', error);
    throw error;
  }
}

// --- スプレッドシート初期化関数 ---
function initializeSpreadsheet(): void {
  try {
    console.log('スプレッドシートの初期化を開始します...');

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getActiveSheet();

    // シート名を「プロパティ」に変更（Sheet1の場合のみ）
    if (sheet.getName() === 'Sheet1' || sheet.getName().startsWith('シート')) {
      sheet.setName(SHEET_SETTINGS.DEFAULT_NAME);
      console.log(`シート名を「${SHEET_SETTINGS.DEFAULT_NAME}」に変更しました`);
    }

    // シートをクリア
    sheet.clear();

    // ヘッダー行の値を動的に生成
    const tokenHeader = ['LINE Channel Access Token'];
    const labelHeaders = Array.from(
      { length: NUMBER_OF_LABELS },
      (_, i) => `Gmail Label Name ${i + 1}`
    );
    const headers = tokenHeader.concat(labelHeaders);

    // ヘッダー行を設定
    const headerRange = sheet.getRange(
      SPREADSHEET_ROWS.HEADER_ROW,
      1,
      1,
      NUMBER_OF_LABELS + 1
    );
    headerRange.setValues([headers]);
    headerRange.setFontWeight('bold');

    // ヘッダー行トークン列の色を設定
    const tokenHeaderRange = sheet.getRange(
      SPREADSHEET_ROWS.HEADER_ROW,
      SPREADSHEET_COLUMNS.LINE_TOKEN,
      1,
      1
    );
    tokenHeaderRange.setBackground(SHEET_SETTINGS.TOKEN_HEADER_COLOR);

    // ヘッダー行ラベル列の色を設定
    const labelHeadersRange = sheet.getRange(
      SPREADSHEET_ROWS.HEADER_ROW,
      SPREADSHEET_COLUMNS.LABEL_START,
      1,
      NUMBER_OF_LABELS
    );
    labelHeadersRange.setBackground(SHEET_SETTINGS.LABEL_HEADER_COLOR);

    // データ行のサンプルを追加
    const sampleRow = SPREADSHEET_ROWS.DATA_START_ROW;
    sheet
      .getRange(sampleRow, SPREADSHEET_COLUMNS.LINE_TOKEN)
      .setNote('LINEのChannel Access Tokenを入力してください');
    sheet
      .getRange(sampleRow, SPREADSHEET_COLUMNS.LABEL_START)
      .setNote('Gmailのラベル名を入力してください（例: 重要なメール）');

    // すべてのセルの範囲を取得（ヘッダー + データ行分）
    const allRange = sheet.getRange(
      1,
      1,
      NUMBER_OF_TOKENS + 1,
      NUMBER_OF_LABELS + 1
    );
    allRange.setBorder(
      true,
      true,
      true,
      true,
      true,
      true,
      SHEET_SETTINGS.BORDER_COLOR,
      SpreadsheetApp.BorderStyle.SOLID
    );

    // 列の幅を自動調整
    sheet.autoResizeColumns(1, NUMBER_OF_LABELS + 1);

    // 説明の追加
    const instructionStartRow = NUMBER_OF_TOKENS + 3;
    sheet.getRange(instructionStartRow, 1).setValue('設定方法:');
    sheet
      .getRange(instructionStartRow + 1, 1)
      .setValue('• A列: LINE Channel Access Token');
    sheet
      .getRange(instructionStartRow + 2, 1)
      .setValue('• B〜G列: Gmail ラベル名');
    sheet
      .getRange(instructionStartRow + 3, 1)
      .setValue('• 複数行設定可能（最大4行）');

    const instructionRange = sheet.getRange(instructionStartRow, 1, 4, 1);
    instructionRange.setFontStyle('italic');
    instructionRange.setFontColor('#666666');

    console.log('スプレッドシートの初期化が完了しました');
  } catch (error) {
    console.error('スプレッドシート初期化エラー:', error);
    throw error;
  }
}

// --- カスタムメニューをUIに追加する関数 ---
function addCustomMenuToUi(): void {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Gmail to LINE')
      .addItem('設定を初期化', 'initializeSpreadsheet')
      .addItem('接続・動作確認テスト', 'testGmailToLineIntegration')
      .addItem('メール通知実行', 'main')
      .addItem('通知トリガーを1時間で設定', 'setupOrUpdateTrigger')
      .addToUi();
    console.log('カスタムメニューが追加されました。');
  } catch (error) {
    console.error('カスタムメニュー追加エラー:', error);
  }
}

// --- プロパティをデフォルトシートに書き込む関数 ---
function writePropertyToSheet1(): void {
  try {
    console.log('プロパティシートの作成を開始します...');

    // スプレッドシートをIDで読み込む
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getActiveSheet();

    // シート名が未設定の場合のみ初期化
    if (sheet.getName() === 'Sheet1' || sheet.getName().startsWith('シート')) {
      initializeSpreadsheet();
    } else {
      console.log('既存のプロパティシートを使用します');
    }
  } catch (error) {
    console.error('プロパティシート作成エラー:', error);
  }
}

// --- トリガー設定関数 ---
function setupOrUpdateTrigger(): void {
  try {
    console.log('トリガー設定を開始します...');

    // 1時間ごとのトリガーを設定
    const triggers = ScriptApp.getProjectTriggers();
    let mainTriggerExists = false;

    // 既存のmain関数のトリガーを確認
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === TRIGGER_SETTINGS.FUNCTION_NAME) {
        mainTriggerExists = true;
        console.log('既存のトリガーが見つかりました');
        // トリガーの設定が1時間ごとでない場合は更新
        if (
          trigger.getEventType() !== ScriptApp.EventType.CLOCK ||
          trigger.getTriggerSource() !== ScriptApp.TriggerSource.CLOCK
        ) {
          console.log('トリガーを更新します');
          ScriptApp.deleteTrigger(trigger);
          mainTriggerExists = false;
        }
      }
    });

    // main関数のトリガーが存在しない場合は新規作成
    if (!mainTriggerExists) {
      ScriptApp.newTrigger(TRIGGER_SETTINGS.FUNCTION_NAME)
        .timeBased()
        .everyHours(TRIGGER_SETTINGS.INTERVAL_HOURS)
        .create();
      console.log(
        `${TRIGGER_SETTINGS.INTERVAL_HOURS}時間ごとのトリガーが設定されました。`
      );

      // UIアラート表示
      try {
        const ui = SpreadsheetApp.getUi();
        ui.alert(
          'トリガー設定完了',
          `${TRIGGER_SETTINGS.INTERVAL_HOURS}時間ごとにメール通知が実行されます。`,
          ui.ButtonSet.OK
        );
      } catch (uiError) {
        console.log('UI表示エラー:', uiError);
      }
    } else {
      console.log('既存のトリガーを使用します');
    }
  } catch (error) {
    console.error('トリガー設定エラー:', error);
    throw error;
  }
}

// --- 統合テスト関数 ---
function testGmailToLineIntegration(sendTestMessage: boolean = true): void {
  try {
    console.log('Gmail to LINE 統合テストを開始します...');

    // 1. 設定情報の検証
    const configs = getConfigFromSheet();
    console.log('設定情報の取得に成功しました');
    console.log(`設定数: ${configs.length}`);

    // 各設定をテスト
    configs.forEach((config, index) => {
      console.log(`設定${index + 1}のテスト開始`);
      console.log(`  ラベル数: ${config.labelNames.length}`);
      console.log(`  ラベル: ${config.labelNames.join(', ')}`);

      // トークンの詳細検証
      const tokenValidation = validateLineToken(config.lineToken);
      console.log(`  トークン検証: ${tokenValidation.message}`);

      if (!tokenValidation.isValid) {
        throw new Error(
          `設定${index + 1}のトークンが無効です: ${tokenValidation.message}`
        );
      }

      // テストメッセージ送信（オプション）
      if (sendTestMessage) {
        const testMessage = `🧪 Gmail to LINE アプリのテストメッセージです。\n設定${index + 1}の動作確認中...`;
        sendLine(testMessage, config.lineToken);
        console.log(`設定${index + 1}のテストメッセージ送信完了`);
      }

      // 各ラベルのメール取得テスト
      config.labelNames.forEach((labelName, labelIndex) => {
        console.log(`  ラベル${labelIndex + 1} "${labelName}" のテスト中...`);
        try {
          const messages = fetchHomeMail(labelName);
          console.log(`    未読メール数: ${messages.length}件`);

          if (messages.length > 0) {
            // 最初のメッセージの詳細をログ出力
            const firstMessage = messages[0];
            console.log(`    最初のメッセージ: ${firstMessage.subject}`);
            console.log(`    送信者: ${firstMessage.from}`);

            // テスト実行時は実際にLINE送信
            if (sendTestMessage && messages.length > 0) {
              const testResultMessage = `📧 ラベル "${labelName}" のテスト結果\n未読メール: ${messages.length}件\n\n最新メッセージ:\n${firstMessage.formattedText}`;
              sendLine(testResultMessage, config.lineToken);
              console.log(`ラベル "${labelName}" のテスト結果を送信しました`);
            }
          }
        } catch (labelError) {
          console.warn(`ラベル "${labelName}" のテスト中にエラー:`, labelError);
        }
      });
    });

    // 結果表示
    const resultMessage = `統合テスト完了\n設定数: ${configs.length}\nテストメッセージ送信: ${sendTestMessage ? 'あり' : 'なし'}`;
    console.log(resultMessage);

    // UIアラート表示
    try {
      const ui = SpreadsheetApp.getUi();
      ui.alert('統合テスト結果', resultMessage, ui.ButtonSet.OK);
    } catch (uiError) {
      console.log('UI表示エラー:', uiError);
    }
  } catch (error) {
    console.error('統合テストでエラーが発生しました:', error);
    throw error;
  }
}

// --- メインの実行関数 ---
function main(): void {
  try {
    console.log('Gmail to LINE メール通知を開始します...');

    // スプレッドシートから設定を取得
    const configs = getConfigFromSheet();
    console.log(`設定数: ${configs.length}`);

    let totalProcessedMessages = 0;
    let totalErrors = 0;

    // 各設定で処理
    configs.forEach((config, configIndex) => {
      try {
        console.log(`設定${configIndex + 1}/${configs.length}の処理を開始...`);

        config.labelNames.forEach((labelName, labelIndex) => {
          try {
            console.log(
              `  ラベル${labelIndex + 1}/${config.labelNames.length}: "${labelName}"`
            );

            // メールを取得
            const newMessages = fetchHomeMail(labelName);

            if (newMessages.length > 0) {
              console.log(`    ${newMessages.length}件の未読メールを処理中...`);

              // LINEに送信
              newMessages.forEach((message, messageIndex) => {
                try {
                  console.log(
                    `      メッセージ${messageIndex + 1}/${newMessages.length}を送信中...`
                  );
                  sendLine(message.formattedText, config.lineToken);
                  totalProcessedMessages++;
                } catch (messageError) {
                  console.error(`メッセージ送信エラー:`, messageError);
                  totalErrors++;
                }
              });
            } else {
              console.log(`    未読メールはありません`);
            }
          } catch (labelError) {
            console.error(
              `ラベル "${labelName}" の処理中にエラー:`,
              labelError
            );
            totalErrors++;
          }
        });
      } catch (configError) {
        console.error(`設定${configIndex + 1}の処理中にエラー:`, configError);
        totalErrors++;
      }
    });

    // 結果のメッセージを作成
    const resultMessage = `処理完了: ${totalProcessedMessages}件のメッセージを送信, エラー: ${totalErrors}件`;
    console.log(resultMessage);

    // スプレッドシートのUIにも結果を表示
    try {
      const ui = SpreadsheetApp.getUi();
      const statusMessage =
        totalErrors > 0 ? `⚠️ ${resultMessage}` : `✅ ${resultMessage}`;
      ui.alert('実行結果', statusMessage, ui.ButtonSet.OK);
    } catch (uiError) {
      console.log('UI表示エラー:', uiError);
    }
  } catch (error) {
    console.error('メイン処理でエラーが発生しました:', error);

    // エラーが発生した場合はUIにのみ表示
    try {
      const ui = SpreadsheetApp.getUi();
      const errorMessage = `❌ Gmail to LINE でエラーが発生しました:\n${error instanceof Error ? error.message : String(error)}`;
      ui.alert('エラー', errorMessage, ui.ButtonSet.OK);
    } catch (uiError) {
      console.log('UI表示エラー:', uiError);
    }

    throw error;
  }
}

// --- スプレッドシート自動初期化（onOpen時に実行） ---
function onOpen(): void {
  try {
    // カスタムメニューの追加
    addCustomMenuToUi();

    // シート1にプロパティを書き込む
    writePropertyToSheet1();
  } catch (error) {
    console.error('onOpen実行エラー:', error);
  }
}
