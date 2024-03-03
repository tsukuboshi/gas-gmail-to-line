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
// メールを取得してLINEに送信する関数
export function main(homeLabelArray: string[], lineToken: string): void {
  homeLabelArray.forEach(homeLabel => {
    // メールを取得
    const newMessages: string[] = fetchHomeMail(homeLabel);
    // LINEに送信
    newMessages.forEach(message => {
      sendLine(message, lineToken);
    });
  });
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
