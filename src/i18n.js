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
  'en': {
    'cardMessage': 'Send personalized emails based on Gmail template to multiple recipients. Together with an optional <b>Group Merge</b> feature.\n\nEnter the following items and select <b><i>Create Drafts</i></b> or <b><i>Send Emails</i></b>. <b>All items are required fields</b> unless otherwise noted.',
    'cardRecipientListSettings': 'Recipient List',
    'cardEnterSpreadsheetUrl': 'Spreadsheet URL',
    'cardHintSpreadsheetUrl': 'Enter the full URL of the Google Sheets spreadsheet used as the list of recipients.',
    'cardEnterSheetName': 'Sheet Name',
    'cardHintSheetName': 'Name of the worksheet of the recipient list. Make sure that the first row of this sheet is used for merge field names.',
    'cardEnterRecipientColName': 'Recipient Field (column) Name',
    'cardHintRecipientColName': 'Field (column) name that lists the recipient\'s email address.',
    'cardTemplateDraftSettings': 'Template Draft',
    'cardEnterTemplateSubject': 'Subject of Template Draft Mail',
    'cardHintTemplateSubject': 'Be sure to make the template subject unique; an error will be returned if there are two or more draft messages with the same subject.',
    'cardSwitchEnableGroupMerge': 'Enable Group Merge',
    'cardAdvancedSettings': 'Advanced Settings',
    'cardSwitchEnableReplyTo': 'Enable Reply-To',
    'cardEnterReplyTo': 'Reply-To',
    'cardHintReplyTo': '[Required only if Reply-To is enabled] The email address to set as reply-to. Placeholders can be used to set the value depending on the individual data, i.e., {{replyTo}}@mydomain.com or simply, {{replyToAddress}}',
    'cardEnterReplaceValue': 'Replace Value',
    'cardHintReplaceValue': 'Text that will replace markers with empty data.',
    'cardEnterMergeFieldMarker': 'Merge Field Marker',
    'cardHintMergeFieldMarker': 'Text to be processed in RegExp() constructor to define merge field(s). Note that the backslash itself does not need to be escaped, i.e., does not need to be repeated.',
    'cardEnterGroupFieldMarker': 'Group Field Marker',
    'cardHintGroupFieldMarker': 'Text to be processed in RegExp() constructor to define group merge field(s). Note that the backslash itself does not need to be escaped, i.e., does not need to be repeated.',
    'cardEnterRowIndexMarker': 'Row Index Marker',
    'cardHintRowIndexMarker': 'Marker for merging row index number in a group merge.',
    //////////
    // 'menuName': 'Mail Merge',
    'buttonCreateDrafts': 'Create Drafts',
    'buttonSendDrafts': 'Send the Created Drafts',
    'buttonSendEmails': 'Send Emails',
    'alertConfirmSendingOfDraft': 'Are you sure you want to send the draft email(s) as {{myEmail}}?\nOnly the drafts created by "Create Drafts" will be sent.',
    'errorSendDraftsCanceled': 'Sending drafts is canceled.',
    'errorNoDraftToSend': 'No draft found.',
    'alertConfirmAccountCreateDraft': 'Are you sure you want to create draft email(s) as {{myEmail}}?',
    'alertConfirmAccountSendEmail': 'Are you sure you want to send email(s) as {{myEmail}}?',
    'errorMailMergeCanceled': 'Mail merge canceled.',
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
    'cardMessage': '複数の宛先に差し込みメールを送信。同じ宛先への複数行にわたる情報を1通のメールにまとめられる<b>まとめ差し込み（Group Merge）</b>機能付き。\n\n以下の項目を入力して「<b>テスト差込（下書きメール作成）</b>」または「<b>差込み送信（メール送信）</b>」を洗濯してください。\n[Optional]とある項目以外は<b>すべて必須の入力項目</b>です。',
    'cardRecipientListSettings': '宛先リストの設定',
    'cardEnterSpreadsheetUrl': 'スプレッドシートURL',
    'cardHintSpreadsheetUrl': '宛先リストのスプレッドシートURLを入力',
    'cardEnterSheetName': 'シート名',
    'cardHintSheetName': '宛先リストが記載されているシートの名前。1行目が差込フィールド名となっていることをご確認ください。',
    'cardEnterRecipientColName': '宛先メールアドレスの列名',
    'cardHintRecipientColName': '宛先となるメールアドレスが記載されているフィールド（列）の名前',
    'cardTemplateDraftSettings': 'テンプレートの設定',
    'cardEnterTemplateSubject': 'テンプレートとなる下書きメールの件名',
    'cardHintTemplateSubject': 'テンプレートとなる下書きメールの件名に重複がないようご注意ください。入力した件名を持つ下書きメールが2通以上ある場合、エラーとなります。',
    'cardSwitchEnableGroupMerge': 'まとめ差し込み（Group Merge）を有効にする',
    'cardAdvancedSettings': '高度な設定',
    'cardSwitchEnableReplyTo': 'Reply-To設定を有効にする',
    'cardEnterReplyTo': 'Reply-Toメールアドレス',
    'cardHintReplyTo': '[Reply-To設定が有効の場合にのみ入力必須] Reply-Toとして設定するメールアドレス。ここでも差し込みフィールドを使うことができる（例：{{replyTo}}@mydomain.com や {{replyToAddress}}',
    'cardEnterReplaceValue': 'データ不在時の差し込みテキスト',
    'cardHintReplaceValue': '該当するデータがない場合に差し込まれる文字列',
    'cardEnterMergeFieldMarker': '差し込みフィールドのマーカー',
    'cardHintMergeFieldMarker': '差し込みフィールドを定義する正規表現。バックスラッシュ（\\）自体はエスケープ不要。',
    'cardEnterGroupFieldMarker': 'まとめ差し込みフィールドのマーカー',
    'cardHintGroupFieldMarker': 'まとめ差し込みフィールドを定義する正規表現。バックスラッシュ（\\）自体はエスケープ不要。',
    'cardEnterRowIndexMarker': 'まとめ差し込み番号マーカー',
    'cardHintRowIndexMarker': 'まとめ差し込みで、同一宛先内の差し込み順序を番号付けするためのマーカー',
    //////////
    // 'menuName': 'Mail Merge（メール差込）',
    'buttonCreateDrafts': '差し込みテスト（下書きメール作成）',
    'buttonSendDrafts': '作成済みの下書きメールを送信',
    'buttonSendEmails': '差し込み送信（メール送信）',
    'alertConfirmSendingOfDraft': 'Are you sure you want to send the draft email(s) as {{myEmail}}?\nOnly the drafts created by "Create Drafts" will be sent.',
    'errorSendDraftsCanceled': 'Sending drafts is canceled.',
    'errorNoDraftToSend': '送信する下書きメールが見つかりません。',
    'alertConfirmAccountCreateDraft': '{{myEmail}}として差し込みメールを作成し、下書きとして保存します。',
    'alertConfirmAccountSendEmail': '{{myEmail}}として差し込みメールを作成し、送信します。',
    'errorMailMergeCanceled': 'メール差込が中断されました。',
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
    this.DEFAULT_LOCALE = 'en';
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
   * Replace placeholder string in this.messageList.alertConfirmSendingOfDraft
   * @param {string} myEmail
   * @returns {string} The replaced text.
   */
  replaceConfirmSendingOfDraft(myEmail) {
    let text = this.messageList.alertConfirmSendingOfDraft;
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