// Copyright 2021 Taro TSUKAGOSHI
// 
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
// 
//     http://www.apache.org/licenses/LICENSE-2.0
// 
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.
//
// See https://github.com/ttsukagoshi/mail-merge-for-gmail for latest information.

/* exported LocalizedMessage  */

const MESSAGES = {
  'en_US': {
    'menuName': 'Mail Merge',
    'menuCreateDraft': 'Create Draft',
    'menuSendEmails': 'Send Emails',
    'alertConfirmAccountCreateDraft': 'Are you sure you want to create draft email(s) as {{myEmail}}?',
    'alertConfirmAccountSendEmail': 'Are you sure you want to send email(s) as {{myEmail}}?',
    'errorCanceled': 'Mail merge canceled.',
    'promptEnterSubjectOfTemplateDraft': 'Enter the subject of Gmail template draft.',
    'errorNoTextEntered': 'No text entered.',
    'errorTwoOrMoreDraftsWithSameSubject': 'There are 2 or more Gmail drafts with the subject you entered. Enter a unique subject text.',
    'errorNoMatchingTemplateDraft': 'No template Gmail draft with matching subject found.',
    'alertGroupMergeFieldMarkerDetected': 'Group merge field marker detected. Do you want to enable group merge function?',
    'alertTitleConfirmation': 'Confirmation',
    'errorInvalidRECIPIENT_COL_NAME': 'Invalid RECIPIENT_COL_NAME. Check sheet "Config" to make sure it refers to an existing column name.',
    'alertCompleteAllDraftsCreated': 'Complete: All draft(s) created.',
    'alertCompleteAllMailsSent': 'Complete: All mail(s) sent.',
    'alertSkippedLabeling': '\nLabeling was skipped for {{skipLabelingCount}} messages. See log for details at https://script.google.com/home/projects/{{scriptId}}/executions?run_as=1'
  },
  'ja': {
    'menuName': 'Mail Merge（メール差込）',
    'menuCreateDraft': 'テスト差込（メール下書き作成）',
    'menuSendEmails': '差込み送信（メール送信）',
    'alertConfirmAccountCreateDraft': '{{myEmail}}として差し込みメールを作成し、下書きとして保存します。',
    'alertConfirmAccountSendEmail': '{{myEmail}}として差し込みメールを作成し、送信します。',
    'errorCanceled': 'メール差込が中断されました。',
    'promptEnterSubjectOfTemplateDraft': 'テンプレートとなるGmail下書きメールの件名を入力してください。',
    'errorNoTextEntered': '件名テキストが入力されていません。',
    'errorTwoOrMoreDraftsWithSameSubject': '入力された件名を持つ下書きメールが2件以上あります。固有の件名を指定してください。',
    'errorNoMatchingTemplateDraft': '入力された件名を持つ下書きメールが見つかりません。',
    'alertGroupMergeFieldMarkerDetected': '「まとめ差し込み（Group Merge）」のマーカーが検出されました。「まとめ差し込み」を有効にしますか？',
    'alertTitleConfirmation': '確認',
    'errorInvalidRECIPIENT_COL_NAME': '無効な RECIPIENT_COL_NAME 。設定（Config）シートで、実際に存在するフィールド（列）が指定されていることを確認してください。',
    'alertCompleteAllDraftsCreated': '完了：すべての下書きが作成されました。',
    'alertCompleteAllMailsSent': '完了：すべての差し込みメールの送信が完了しました。',
    'alertSkippedLabeling': '\n{{skipLabelingCount}}通のメールについて、ラベル付け処理がスキップされました。詳細は次のURLからログを確認してください： https://script.google.com/home/projects/{{scriptId}}/executions?run_as=1'
  }
};

class LocalizedMessage {
  constructor(userLocale) {
    this.DEFAULT_LOCALE = 'en_US';
    this.locale = (MESSAGES[userLocale] ? userLocale : this.DEFAULT_LOCALE);
    this.messageList = MESSAGES[this.locale];
    Object.keys(MESSAGES[this.DEFAULT_LOCALE]).forEach(key => {
      if (!this.messageList[key]) {
        this.messageList[key] = MESSAGES[this.DEFAULT_LOCALE][key];
      }
    });
  }
  /**
   * Replace placeholder values in the designated text. String.prototype.replace() is executed using regular expressions with the 'global' flag on.
   * @param {string} text 
   * @param {array} placeholderValues Array of objects containing a placeholder string expressed in regular expression and its corresponding value.
   * @returns {string} The replaced text.
   */
  replacePlaceholders_(text, placeholderValues = []) {
    let replacedText = placeholderValues.reduce((acc, cur) => acc.replace(new RegExp(cur.regexp, 'g'), cur.value), text);
    return replacedText;
  }
  /**
   * Replace placeholder string in this.messageList.alertConfirmAccountCreateDraft or .alertConfirmAccountSendEmail
   * @param {boolean} draftMode
   * @param {string} myEmail
   * @returns {string} The replaced text.
   */
  replaceConfirmAccount(draftMode, myEmail) {
    let text = (draftMode ? this.messageList.alertConfirmAccountCreateDraft : this.messageList.alertConfirmAccountSendEmail);
    let placeholderValues = [
      {
        'regexp': '\{\{myEmail\}\}',
        'value': myEmail
      }
    ];
    text = this.replacePlaceholders_(text, placeholderValues);
    return text;
  }
  /**
   * Replace placeholder string in this.messageList.alertSkippedLabeling
   * @param {string} skipLabelingCount
   * @param {string} scriptId
   * @returns {string} The replaced text.
   */
  replaceAlertSkippedLabeling(skipLabelingCount, scriptId) {
    let text = this.messageList.alertSkippedLabeling;
    let placeholderValues = [
      {
        'regexp': '\{\{skipLabelingCount\}\}',
        'value': skipLabelingCount
      },
      {
        'regexp': '\{\{scriptId\}\}',
        'value': scriptId
      }
    ];
    text = this.replacePlaceholders_(text, placeholderValues);
    return text;
  }
}