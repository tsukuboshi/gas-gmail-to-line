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

// LINEアクセストークンの個数(必要に応じて調整してください)
const numLineTokens = 3;

// ヘッダー行の値
const headers = [
  'LINE Access Token',
  'Gmail Label Name 1',
  'Gmail Label Name 2',
  'Gmail Label Name 3',
  'Gmail Label Name 4',
  'Gmail Label Name 5',
];

// ヘッダー行の値の個数
const numHeaders = headers.length;

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function onOpen() {
  // UIにカスタムメニューを追加
  addCustomMenuToUi();

  // シート1にプロパティを書き込む
  writePropertyToSheet1();
}

// カスタムメニューをUIに追加する関数
function addCustomMenuToUi() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('カスタムメニュー')
    .addItem('通知テストの実施', 'main')
    .addItem('通知トリガーを1時間で設定', 'setupOrUpdateTrigger')
    .addToUi();
  console.log('カスタムメニューが追加されました。');
}

// プロパティをデフォルトシートに書き込む関数
function writePropertyToSheet1(): void {
  // スプレッドシートをIDで読み込む
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();

  // シート1名を変更
  const sheetName = 'プロパティ';
  sheet.setName(sheetName);

  // ヘッダー行を設定
  const headerRange = sheet.getRange(1, 1, 1, numHeaders);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('lightblue');

  // すべてのセルの範囲を取得
  const allRange = sheet.getRange(1, 1, numLineTokens + 1, numHeaders);

  // すべてのセルに枠線を設定
  allRange.setBorder(
    true,
    true,
    true,
    true,
    true,
    true,
    'black',
    SpreadsheetApp.BorderStyle.SOLID
  );

  // 列の幅を自動調整
  sheet.autoResizeColumns(1, numHeaders);

  console.log('プロパティシートが作成されました。');
}

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function setupOrUpdateTrigger() {
  // 1時間ごとのトリガーを設定
  const triggers = ScriptApp.getProjectTriggers();
  let mainTriggerExists = false;

  // 既存のmain関数のトリガーを確認
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'main') {
      mainTriggerExists = true;
      // トリガーの設定が1時間ごとでない場合は更新
      if (
        trigger.getEventType() !== ScriptApp.EventType.CLOCK ||
        trigger.getTriggerSource() !== ScriptApp.TriggerSource.CLOCK
      ) {
        ScriptApp.deleteTrigger(trigger);
        mainTriggerExists = false;
      }
    }
  });

  // main関数のトリガーが存在しない場合は新規作成
  if (!mainTriggerExists) {
    ScriptApp.newTrigger('main').timeBased().everyHours(1).create();
    console.log('1時間ごとのトリガーが設定されました。');
  }
}

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function main(): void {
  // スプレッドシートをIDで読み込む
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();

  // ヘッダー行の次の行からLINEアクセストークン毎に繰り返し処理を実施
  for (let i = 2; i <= numLineTokens + 1; i++) {
    // A列のセルよりLINE Notifyトークンを取得
    const lineToken = sheet.getRange(i, 1).getValue();
    // const lineToken = sheet.getRange(`A${i}`).getValue();

    // 該当行のA列にLINE Notifyトークンがある場合、同じ行に存在するB-F列のラベル名を取得して処理
    if (lineToken) {
      const homeLabelArray = sheet
        .getRange(i, 2, 1, numHeaders - 1)
        .getValues()
        .flat()
        .filter(label => label.trim() !== '');
      // const homeLabelArray = sheet.getRange(`B${i}:F${i}`).getValues().flat();

      homeLabelArray.forEach(homeLabel => {
        // メールを取得
        const newMessages: string[] = fetchHomeMail(homeLabel);
        // LINEに送信
        newMessages.forEach(message => {
          sendLine(message, lineToken);
        });
      });
    }
  }
}

// メールを取得する関数
function fetchHomeMail(homeLabel: string): string[] {
  // 未読メールの検索条件
  const searchTerms: string = Utilities.formatString(
    'label:%s is:unread',
    homeLabel
  );

  // 未読メールのスレッドを取得
  const fetchedThreads: GoogleAppsScript.Gmail.GmailThread[] =
    GmailApp.search(searchTerms);

  // 未読メールの取得
  const fetchedMessages: GoogleAppsScript.Gmail.GmailMessage[][] =
    GmailApp.getMessagesForThreads(fetchedThreads);

  // 送信メッセージ配列
  const sendMessages: string[] = [];

  // 送信メッセージ配列に詰める
  fetchedMessages.forEach(messageArray => {
    messageArray.forEach(message => {
      // メッセージが未読の場合
      if (message.isUnread()) {
        // メッセージを整形
        const formattedMessage = gmailToString(message);
        // 送信メッセージ配列に追加
        sendMessages.push(formattedMessage);
        // メッセージを既読にマーク
        message.markRead();
      }
    });
  });

  // LINE用に配列の順番を反転させる
  return sendMessages.reverse();
}

// メールを整形する関数
function gmailToString(mail: GoogleAppsScript.Gmail.GmailMessage): string {
  // メールの情報を取得
  const mailFrom: string = mail.getFrom();
  const subject: string = mail.getSubject();
  const body: string = mail.getPlainBody().slice(0, 300);

  return Utilities.formatString(
    // 送信者、件名、内容を整形
    '\n%s\n\n件名：\n%s\n\n内容：\n%s',
    mailFrom,
    subject,
    body
  );
}

// LINEヘ送信する関数
function sendLine(message: string, lineToken: string): void {
  // 送信内容
  const payload = { message: message };

  // 送信オプション
  const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
    method: 'post',
    headers: { Authorization: 'Bearer ' + lineToken },
    payload: payload,
  };

  // 送信
  UrlFetchApp.fetch('https://notify-api.line.me/api/notify', options);
}
