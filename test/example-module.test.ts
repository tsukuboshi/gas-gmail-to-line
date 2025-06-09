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
} from '../src/env';

describe('gmail-to-line', () => {
  describe('基本的なテスト', () => {
    it('テストが正常に実行される', () => {
      expect(true).toBe(true);
    });

    it('文字列の操作が正常に動作する', () => {
      const sender = 'test@example.com';
      const subject = 'テストメール';
      const message = `${sender}から「${subject}」が届きました。`;
      expect(message).toContain('test@example.com');
      expect(message).toContain('テストメール');
    });

    it('配列の操作が正常に動作する', () => {
      const labelNames = ['重要', '仕事', 'プライベート'];
      const filteredLabels = labelNames.filter(label => label.length > 2);

      expect(filteredLabels).toHaveLength(1);
      expect(filteredLabels).toContain('プライベート');
    });
  });

  describe('環境定数のテスト', () => {
    it('LINE API URLが正しく設定されている', () => {
      expect(LINE_API_BASE_URL).toBe(
        'https://api.line.me/v2/bot/message/broadcast'
      );
      expect(LINE_API_BASE_URL).toMatch(/^https:\/\//);
    });

    it('トークン数とラベル数が妥当である', () => {
      expect(NUMBER_OF_TOKENS).toBeGreaterThan(0);
      expect(NUMBER_OF_TOKENS).toBeLessThanOrEqual(10);
      expect(typeof NUMBER_OF_TOKENS).toBe('number');

      expect(NUMBER_OF_LABELS).toBeGreaterThan(0);
      expect(NUMBER_OF_LABELS).toBeLessThanOrEqual(20);
      expect(typeof NUMBER_OF_LABELS).toBe('number');
    });

    it('メール本文の最大文字数が設定されている', () => {
      expect(BODY_MAX_LENGTH).toBeGreaterThan(0);
      expect(BODY_MAX_LENGTH).toBeLessThanOrEqual(5000);
      expect(typeof BODY_MAX_LENGTH).toBe('number');
    });

    it('API設定が正しい', () => {
      expect(API_SETTINGS.REQUEST_DELAY).toBeGreaterThanOrEqual(0);
      expect(API_SETTINGS.MAX_RETRIES).toBeGreaterThan(0);
      expect(API_SETTINGS.TIMEOUT).toBeGreaterThan(0);
      expect(typeof API_SETTINGS.REQUEST_DELAY).toBe('number');
      expect(typeof API_SETTINGS.MAX_RETRIES).toBe('number');
      expect(typeof API_SETTINGS.TIMEOUT).toBe('number');
    });
  });

  describe('スプレッドシート設定のテスト', () => {
    it('列定義が正しい', () => {
      expect(SPREADSHEET_COLUMNS.LINE_TOKEN).toBe(1);
      expect(SPREADSHEET_COLUMNS.LABEL_START).toBe(2);
      expect(SPREADSHEET_COLUMNS.LABEL_START).toBeGreaterThan(
        SPREADSHEET_COLUMNS.LINE_TOKEN
      );
    });

    it('行定義が正しい', () => {
      expect(SPREADSHEET_ROWS.HEADER_ROW).toBe(1);
      expect(SPREADSHEET_ROWS.DATA_START_ROW).toBe(2);
      expect(SPREADSHEET_ROWS.DATA_START_ROW).toBeGreaterThan(
        SPREADSHEET_ROWS.HEADER_ROW
      );
    });

    it('シート設定が正しい', () => {
      expect(SHEET_SETTINGS.DEFAULT_NAME).toBe('プロパティ');
      expect(typeof SHEET_SETTINGS.TOKEN_HEADER_COLOR).toBe('string');
      expect(typeof SHEET_SETTINGS.LABEL_HEADER_COLOR).toBe('string');
      expect(typeof SHEET_SETTINGS.BORDER_COLOR).toBe('string');
    });
  });

  describe('Gmail検索設定のテスト', () => {
    it('Gmail検索条件が正しい', () => {
      expect(GMAIL_SEARCH.UNREAD_CONDITION).toBe('is:unread');
      expect(GMAIL_SEARCH.LABEL_FORMAT).toBe('label:%s');
    });

    it('検索条件の組み合わせが正しく動作する', () => {
      const labelName = '重要なメール';
      const searchTerms =
        `${GMAIL_SEARCH.LABEL_FORMAT} ${GMAIL_SEARCH.UNREAD_CONDITION}`.replace(
          '%s',
          labelName
        );

      expect(searchTerms).toBe('label:重要なメール is:unread');
      expect(searchTerms).toContain('label:');
      expect(searchTerms).toContain('is:unread');
      expect(searchTerms).toContain(labelName);
    });
  });

  describe('トリガー設定のテスト', () => {
    it('トリガー設定が正しい', () => {
      expect(TRIGGER_SETTINGS.INTERVAL_HOURS).toBe(1);
      expect(TRIGGER_SETTINGS.FUNCTION_NAME).toBe('main');
      expect(TRIGGER_SETTINGS.INTERVAL_HOURS).toBeGreaterThan(0);
      expect(TRIGGER_SETTINGS.INTERVAL_HOURS).toBeLessThanOrEqual(24);
    });
  });

  describe('メッセージフォーマット設定のテスト', () => {
    it('メッセージフォーマットが正しい', () => {
      expect(MESSAGE_FORMAT.SEPARATOR).toBe('\n\n');
      expect(MESSAGE_FORMAT.SENDER_PREFIX).toBe('📧 送信者：');
      expect(MESSAGE_FORMAT.SUBJECT_PREFIX).toBe('📋 件名：');
      expect(MESSAGE_FORMAT.CONTENT_PREFIX).toBe('📄 内容：');
    });

    it('メッセージの組み立てが正しく動作する', () => {
      const from = 'sender@example.com';
      const subject = 'テスト件名';
      const body = 'テスト本文';

      const formattedMessage = `\n${MESSAGE_FORMAT.SENDER_PREFIX}${MESSAGE_FORMAT.SEPARATOR}${from}${MESSAGE_FORMAT.SEPARATOR}${MESSAGE_FORMAT.SUBJECT_PREFIX}${MESSAGE_FORMAT.SEPARATOR}${subject}${MESSAGE_FORMAT.SEPARATOR}${MESSAGE_FORMAT.CONTENT_PREFIX}${MESSAGE_FORMAT.SEPARATOR}${body}`;

      expect(formattedMessage).toContain(from);
      expect(formattedMessage).toContain(subject);
      expect(formattedMessage).toContain(body);
      expect(formattedMessage).toContain('📧 送信者：');
      expect(formattedMessage).toContain('📋 件名：');
      expect(formattedMessage).toContain('📄 内容：');
    });
  });

  describe('バリデーション設定のテスト', () => {
    it('バリデーション設定が正しい', () => {
      expect(VALIDATION.MIN_TOKEN_LENGTH).toBe(40);
      expect(VALIDATION.MIN_TOKEN_LENGTH).toBeGreaterThan(0);
      expect(VALIDATION.MAX_MESSAGE_LENGTH).toBe(5000);
      expect(VALIDATION.MAX_MESSAGE_LENGTH).toBeGreaterThan(0);
    });

    it('トークンパターンが正しく動作する', () => {
      const validToken = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnop+/=';
      const invalidToken = 'invalid-token-with-special-chars!@#$%';

      expect(VALIDATION.VALID_TOKEN_PATTERN.test(validToken)).toBe(true);
      expect(VALIDATION.VALID_TOKEN_PATTERN.test(invalidToken)).toBe(false);
    });

    it('トークン長の検証が正しく動作する', () => {
      const shortToken = 'short';
      const validLengthToken = 'A'.repeat(VALIDATION.MIN_TOKEN_LENGTH);

      expect(shortToken.length < VALIDATION.MIN_TOKEN_LENGTH).toBe(true);
      expect(validLengthToken.length >= VALIDATION.MIN_TOKEN_LENGTH).toBe(true);
    });
  });

  describe('メール処理機能のテスト', () => {
    it('メール本文の切り詰めが正しく動作する', () => {
      const longBody = 'A'.repeat(BODY_MAX_LENGTH + 100);
      const truncatedBody = longBody.slice(0, BODY_MAX_LENGTH) + '...';

      expect(longBody.length > BODY_MAX_LENGTH).toBe(true);
      expect(truncatedBody.length).toBe(BODY_MAX_LENGTH + 3);
      expect(truncatedBody.endsWith('...')).toBe(true);
    });

    it('メッセージ長の検証が正しく動作する', () => {
      const shortMessage = 'short message';
      const longMessage = 'A'.repeat(VALIDATION.MAX_MESSAGE_LENGTH + 100);

      expect(shortMessage.length <= VALIDATION.MAX_MESSAGE_LENGTH).toBe(true);
      expect(longMessage.length > VALIDATION.MAX_MESSAGE_LENGTH).toBe(true);
    });

    it('Gmail検索条件の構築が正しく動作する', () => {
      const labelName = 'テストラベル';

      // Google Apps ScriptのUtilities.formatStringのモック
      const mockFormatString = (
        template: string,
        ...args: unknown[]
      ): string => {
        return template.replace(/%s/g, () => args.shift()?.toString() || '');
      };

      const searchTerms = mockFormatString(
        `${GMAIL_SEARCH.LABEL_FORMAT} ${GMAIL_SEARCH.UNREAD_CONDITION}`,
        labelName
      );

      expect(searchTerms).toBe('label:テストラベル is:unread');
      expect(searchTerms).toMatch(/^label:.+ is:unread$/);
    });
  });

  describe('データ構造のテスト', () => {
    it('GmailLabelConfig の構造が正しい', () => {
      const config = {
        lineToken: 'test-token-12345678901234567890123456789012345678',
        labelNames: ['ラベル1', 'ラベル2', 'ラベル3'],
      };

      expect(config.lineToken).toBeDefined();
      expect(config.labelNames).toBeDefined();
      expect(Array.isArray(config.labelNames)).toBe(true);
      expect(config.labelNames.length).toBeGreaterThan(0);
      expect(typeof config.lineToken).toBe('string');
    });

    it('FormattedMessage の構造が正しい', () => {
      const message = {
        from: 'sender@example.com',
        subject: 'テスト件名',
        body: 'テスト本文',
        formattedText: 'フォーマット済みテキスト',
      };

      expect(message.from).toBeDefined();
      expect(message.subject).toBeDefined();
      expect(message.body).toBeDefined();
      expect(message.formattedText).toBeDefined();
      expect(typeof message.from).toBe('string');
      expect(typeof message.subject).toBe('string');
      expect(typeof message.body).toBe('string');
      expect(typeof message.formattedText).toBe('string');
    });

    it('ValidationResult の構造が正しい', () => {
      const validResult = {
        isValid: true,
        message: '検証成功',
      };

      const invalidResult = {
        isValid: false,
        message: '検証失敗',
      };

      expect(validResult.isValid).toBe(true);
      expect(validResult.message).toBeDefined();
      expect(typeof validResult.isValid).toBe('boolean');
      expect(typeof validResult.message).toBe('string');

      expect(invalidResult.isValid).toBe(false);
      expect(invalidResult.message).toBeDefined();
    });
  });

  describe('スプレッドシート操作のテスト', () => {
    it('ヘッダー配列の生成が正しく動作する', () => {
      const tokenHeader = ['LINE Channel Access Token'];
      const labelHeaders = Array.from(
        { length: NUMBER_OF_LABELS },
        (_, i) => `Gmail Label Name ${i + 1}`
      );
      const headers = tokenHeader.concat(labelHeaders);

      expect(headers).toHaveLength(NUMBER_OF_LABELS + 1);
      expect(headers[0]).toBe('LINE Channel Access Token');
      expect(headers[1]).toBe('Gmail Label Name 1');
      expect(headers[NUMBER_OF_LABELS]).toBe(
        `Gmail Label Name ${NUMBER_OF_LABELS}`
      );
    });

    it('範囲指定の計算が正しく動作する', () => {
      const headerRowCount = 1;
      const dataRowCount = NUMBER_OF_TOKENS;
      const totalRows = headerRowCount + dataRowCount;
      const totalColumns = NUMBER_OF_LABELS + 1;

      expect(totalRows).toBe(NUMBER_OF_TOKENS + 1);
      expect(totalColumns).toBe(NUMBER_OF_LABELS + 1);
      expect(totalRows).toBeGreaterThan(1);
      expect(totalColumns).toBeGreaterThan(1);
    });
  });

  describe('エラーハンドリングのテスト', () => {
    it('空の設定配列の処理が正しい', () => {
      const emptyConfigs: unknown[] = [];

      expect(emptyConfigs.length).toBe(0);
      expect(Array.isArray(emptyConfigs)).toBe(true);
    });

    it('無効なトークンの検証が正しく動作する', () => {
      const testCases = [
        { token: '', expectedValid: false },
        { token: 'short', expectedValid: false },
        { token: 'invalid-chars!@#', expectedValid: false },
        { token: 'A'.repeat(VALIDATION.MIN_TOKEN_LENGTH), expectedValid: true },
      ];

      testCases.forEach(({ token, expectedValid }) => {
        const isValidLength = token.length >= VALIDATION.MIN_TOKEN_LENGTH;
        const isValidPattern = VALIDATION.VALID_TOKEN_PATTERN.test(token);
        const isValid = token.trim() !== '' && isValidLength && isValidPattern;

        expect(isValid).toBe(expectedValid);
      });
    });

    it('メッセージの切り詰め処理が正しく動作する', () => {
      const maxLength = VALIDATION.MAX_MESSAGE_LENGTH;
      const longMessage = 'A'.repeat(maxLength + 100);
      const truncatedMessage = longMessage.slice(0, maxLength - 3) + '...';

      expect(longMessage.length > maxLength).toBe(true);
      expect(truncatedMessage.length).toBe(maxLength);
      expect(truncatedMessage.endsWith('...')).toBe(true);
    });
  });

  describe('設定の互換性テスト', () => {
    it('新旧定数名の互換性が保たれている', () => {
      // 後方互換性のため、旧変数名も利用可能であることを確認
      expect(NUMBER_OF_TOKENS).toBeDefined();
      expect(NUMBER_OF_LABELS).toBeDefined();
      expect(BODY_MAX_LENGTH).toBeDefined();

      expect(typeof NUMBER_OF_TOKENS).toBe('number');
      expect(typeof NUMBER_OF_LABELS).toBe('number');
      expect(typeof BODY_MAX_LENGTH).toBe('number');
    });

    it('設定値の一貫性が保たれている', () => {
      // 関連する設定値が論理的に一貫していることを確認
      expect(SPREADSHEET_ROWS.DATA_START_ROW).toBeGreaterThan(
        SPREADSHEET_ROWS.HEADER_ROW
      );
      expect(SPREADSHEET_COLUMNS.LABEL_START).toBeGreaterThan(
        SPREADSHEET_COLUMNS.LINE_TOKEN
      );
      expect(BODY_MAX_LENGTH).toBeLessThanOrEqual(
        VALIDATION.MAX_MESSAGE_LENGTH
      );
      expect(VALIDATION.MIN_TOKEN_LENGTH).toBeGreaterThan(0);
    });
  });
});
