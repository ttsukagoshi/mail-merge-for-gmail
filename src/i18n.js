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

// Message Naming Rules:
// Name of the message (key) must start with 'error' for error messages
// to distinguish between expected and unexpected errors

const MESSAGES = {
  en: {
    cardHomepageMessage:
      '<b>Group Merge: Mail Merge for Gmail</b>\nOpen-sourced add-on to send personalized emails based on Gmail template to multiple recipients. The unique <b>Group Merge</b> feature allows sender to group multiple contents for the same recipient in a single email.\n\nEnter the following items and select <b><i>Create Drafts</i></b> or <b><i>Send Emails</i></b>. <b>All items are required fields</b> unless otherwise noted.',
    cardRecipientListSettings: '1. Recipient List',
    cardEnterSpreadsheetUrl: 'Spreadsheet URL',
    cardHintSpreadsheetUrl:
      'Enter the full URL of the Google Sheets spreadsheet used as the list of recipients.',
    cardEnterSheetName: 'Sheet Name',
    cardHintSheetName:
      'Name of the worksheet of the recipient list. Make sure that the first row of this sheet is used for merge field names.',
    cardEnterTo: 'To',
    cardHintTo:
      "Recipient's email address. Use field markers for only one address, e.g. {{Email}}. If you want to send the respective personalized mails to more than one account, make use of the CC and BCC fields.",
    cardEnterCc: 'CC',
    cardHintCc:
      "[Optional] CC recipients' email address. Use commas to seperate multiple addresses.",
    cardEnterBcc: 'BCC',
    cardHintBcc:
      "[Optional] BCC recipients' email address. Use commas to seperate multiple addresses.",
    cardTemplateDraftSettings: '2. Template Draft',
    cardEnterTemplateSubject: 'Subject of Template Draft Mail',
    cardHintTemplateSubject:
      'Be sure to make the template subject unique; an error will be returned if there are two or more draft messages with the same subject.',
    cardSwitchEnableGroupMerge: 'Enable Group Merge',
    cardAdvancedSettings: '3. Advanced Settings',
    cardSwitchEnableReplyTo: 'Enable Reply-To',
    cardEnterReplyTo: 'Reply-To',
    cardHintReplyTo:
      '[Required only if Reply-To is enabled] The email address to set as reply-to. Placeholders can be used to set the value depending on the individual data, i.e., {{replyTo}}@mydomain.com or simply, {{replyToAddress}}',
    cardEnterReplaceValue: 'Replace Value',
    cardHintReplaceValue: 'Text that will replace markers with empty data.',
    cardEnterMergeFieldMarker: 'Merge Field Marker',
    cardHintMergeFieldMarker:
      'Text to be processed in RegExp() constructor to define merge field(s). Be sure to use parentheses to denote the capturing group to mark the content of the field, i.e., the word "field" in {{field}}. Note that the backslash itself does not need to be escaped.',
    cardEnterGroupFieldMarker: 'Group Field Marker',
    cardHintGroupFieldMarker:
      'Text to be processed in RegExp() constructor to define group merge field(s). Be sure to use parentheses to denote the capturing group to mark the content of the field. Note that the backslash itself does not need to be escaped.',
    cardEnterRowIndexMarker: 'Row Index Marker',
    cardHintRowIndexMarker:
      'Marker for merging row index number in a group merge.',
    cardSwitchEnableDebugMode: 'Enable debug mode',
    cardMessageUnexpectedError: 'Unexpected Error:\n',
    cardMessageDebugInfo:
      "Debug Info\nBe careful to exclude any sensitive information, like your clients' email addresses, if you are going to copy & paste the below debug information elsewhere:\n\n",
    cardMessageSentDebugInfo:
      '<b>Email of the debug info is sent to your Gmail account.</b>',
    buttonRestoreUserConfig: 'Restore User Settings',
    buttonRestoreDefault: 'Restore Default Settings',
    buttonReadDocument: 'Website',
    buttonReadDocumentUrl:
      'https://www.scriptable-assets.page/add-ons/group-merge/',
    buttonSaveUserConfig: 'Save User Settings',
    buttonCreateDrafts: 'Create Drafts',
    buttonSendDrafts: 'Send Created Drafts',
    buttonSendEmails: 'Send Emails',
    buttonReturnHome: 'Return',
    errorSpreadsheetNotFound:
      '[Mail Merge Error]\nThe Google Sheets URL that you entered is not valid. Check if you have entered the correct URL, and that you have access to that spreadsheet.',
    errorSheetNotFound:
      '[Mail Merge Error]\nThe worksheet of the list of recipients in the spreadsheet cannot found. Check if you have entered the correct sheet name.',
    alertConfirmAccountCreateDraft:
      'Are you sure you want to create draft email(s) as {{myEmail}}?',
    alertConfirmAccountSendEmail:
      'Are you sure you want to send email(s) as {{myEmail}}?',
    errorMailMergeCanceled: 'Mail merge canceled.',
    errorNoSubjectTextEntered:
      '[Mail Merge Error]\nEnter the subject of the Gmail draft message to use as template for mail merge.',
    errorTwoOrMoreDraftsWithSameSubject:
      '[Mail Merge Error]\nThere are 2 or more Gmail drafts with the subject you entered. Enter a unique subject text.',
    errorNoMatchingTemplateDraft:
      '[Mail Merge Error]\nNo template Gmail draft message with matching subject found. Make sure the subject that you entered is correct.',
    errorNoToEntered:
      '[Mail Merge Error]\nNo To recipients entered. Make sure you have entered the placeholder value of the email address to use as To.',
    errorMultipleToEntered:
      "[Mail Merge Error]\nThere are multiple To's entered as merge fields. Use only one.",
    errorInvalidTo:
      '[Mail Merge Error]\nInvalid Recipient Field (column) Name. Make sure the placeholder value in "To" refers to an existing column name.',
    alertCompleteAllDraftsCreated:
      "Complete: {{messageCount}} draft(s) created. Reload Gmail's Drafts page if you are not seeing the full drafted list of merged mails.",
    alertCompleteAllMailsSent: 'Complete: {{messageCount}} mail(s) sent.',
    alertConfirmSendingOfDraft:
      'Are you sure you want to send the draft email(s) as {{myEmail}}?\nOnly the drafts created by "Create Drafts" will be sent.',
    errorSendDraftsCanceled: 'Sending drafts is canceled.',
    errorNoDraftToSend:
      '[Mail Merge Error]\nNo draft found. Execute "Create Drafts" to create merged drafts.',
    alertCompleteSavedUserConfig: 'Complete: Saved user settings.\n\n',
    appsScriptMessageErrorOnOpenByUrlStartsWith:
      'Unexpected error while getting the method or property openByUrl',
    appsScriptMessageNoPermissionErrorStartsWith:
      'You do not have permission to access the requested document.',
    proceedingToPostProcess:
      'Since the current mail merge process is expected to exceed the {{actionLimitTimeInSec}}-second execution time limit for Google Workspace add-ons, the remaining process will be conducted on a time-based trigger. A mail notification will be sent to {{myEmail}} once the whole process is complete.\nDraft Mode: {{draftMode}}\nMessage Count: {{messageCount}}',
    continuingPostProcess:
      'CONTINUING MAIL MERGE\nSince the current mail merge process is expected to exceed the {{appsScriptLimitTimeInSec}}-second execution time limit for Google Apps Script, the remaining process will be continued on a time-based trigger. A mail notification will be sent to {{myEmail}} once the whole process is complete.\nDraft Mode: {{draftMode}}\nMessage Count: {{messageCount}}',
    subjectPostProcessUpdate: '[GROUP MERGE] Mail Merge Notice',
  },
  ja: {
    cardHomepageMessage:
      '<b>Group Merge: Mail Merge for Gmail</b>\nオープンソースで利用可能なGmailのための差し込みメール作成用アドオン。同じ宛先への複数行にわたる情報を1通のメールにまとめられる<b>まとめ差し込み（Group Merge）</b>機能付き。\n\n以下の項目を入力して「<b>テスト差込（下書きメール作成）</b>」または「<b>差込み送信（メール送信）</b>」を選択してください。\n別途説明がある項目を除けば<b>すべて必須の入力項目</b>です。',
    cardRecipientListSettings: '1. 宛先リストの設定',
    cardEnterSpreadsheetUrl: 'スプレッドシートURL',
    cardHintSpreadsheetUrl: '宛先リストのスプレッドシートURLを入力',
    cardEnterSheetName: 'シート名',
    cardHintSheetName:
      '宛先リストが記載されているシートの名前。1行目が差込フィールド名となっていることをご確認ください。',
    cardEnterTo: '宛先メールアドレス',
    cardHintTo:
      'Toとして登録する宛先メールアドレス。必ず差し込みフィールドを1つだけ使ってください（例：{{Email}}）。複数の宛先に同時に送りたい場合は、CCやBCC欄をご活用ください。',
    cardEnterCc: 'CC',
    cardHintCc:
      '[オプション] CCとして登録するメールアドレス。差し込みフィールド使用可。カンマ区切りで複数宛先を登録可。（例：{{cc}}},{{anotherCc}},fixedEmail@sample.com）',
    cardEnterBcc: 'BCC',
    cardHintBcc:
      '[オプション] BCCとして登録するメールアドレス。差し込みフィールド使用可。カンマ区切りで複数宛先を登録可。（例：{{bcc}},{{anotherBcc}},fixedEmail@sample.com）',
    cardTemplateDraftSettings: '2. テンプレートの設定',
    cardEnterTemplateSubject: 'テンプレートとなる下書きメールの件名',
    cardHintTemplateSubject:
      'テンプレートとなる下書きメールの件名に重複がないようご注意ください。入力した件名を持つ下書きメールが2通以上ある場合、エラーとなります。',
    cardSwitchEnableGroupMerge: 'まとめ差し込み（Group Merge）を有効にする',
    cardAdvancedSettings: '3. 高度な設定',
    cardSwitchEnableReplyTo: 'Reply-To設定を有効にする',
    cardEnterReplyTo: 'Reply-Toメールアドレス',
    cardHintReplyTo:
      '[Reply-To設定が有効の場合にのみ入力必須] Reply-Toとして設定するメールアドレス。ここでも差し込みフィールドを使うことができる（例：{{replyTo}}@mydomain.com や {{replyToAddress}}）',
    cardEnterReplaceValue: 'データ不在時の差し込みテキスト',
    cardHintReplaceValue: '該当するデータがない場合に差し込まれる文字列',
    cardEnterMergeFieldMarker: '差し込みフィールドのマーカー',
    cardHintMergeFieldMarker:
      '差し込みフィールドを定義する正規表現。フィールド名そのもの（例：{{名前}}の中の「名前」という文字列）を検出するためのキャプチャグループ（"( )"）を使用すること。バックスラッシュ（\\）自体はエスケープ不要。',
    cardEnterGroupFieldMarker: 'まとめ差し込みフィールドのマーカー',
    cardHintGroupFieldMarker:
      'まとめ差し込みフィールドを定義する正規表現。フィールド名そのもの（例：[[ミーティング{{i}}]]の中の「ミーティング{{i}}」という文字列）を検出するためのキャプチャグループ（"( )"）を使用すること。バックスラッシュ（\\）自体はエスケープ不要。',
    cardEnterRowIndexMarker: 'まとめ差し込み番号マーカー',
    cardHintRowIndexMarker:
      'まとめ差し込みで、同一宛先内の差し込み順序を番号付けするためのマーカー',
    cardSwitchEnableDebugMode: 'デバッグモードを有効にする',
    cardMessageUnexpectedError: '予期しないエラー：\n',
    cardMessageDebugInfo:
      'デバッグ情報\n以下のデバッグ情報をどこかにコピー＆ペーストする際は、顧客のメールアドレスなどの機密情報を取り除くよう注意してください：\n\n',
    cardMessageSentDebugInfo:
      '<b>デバッグ情報がメールにてお使いのGoogleアカウントに送信されました。</b>',
    buttonRestoreUserConfig: '保存した設定を使用',
    buttonRestoreDefault: '初期設定に戻す',
    buttonReadDocument: 'ウェブサイトへ',
    buttonReadDocumentUrl:
      'https://www.scriptable-assets.page/ja/add-ons/group-merge/',
    buttonSaveUserConfig: 'ユーザ設定として保存',
    buttonCreateDrafts: '下書き作成（差込テスト）',
    buttonSendDrafts: '作成済みの下書きを送信',
    buttonSendEmails: '送信（差込メール送信）',
    buttonReturnHome: '戻る',
    errorSpreadsheetNotFound:
      '[Mail Merge Error]\n無効なGoogleスプレッドシートURLです。URLに誤りがないことや、当該スプレッドシートにアクセス権があることを確認してください。',
    errorSheetNotFound:
      '[Mail Merge Error]\n宛先リストが記載されているシートが見つかりません。シート名が正しく入力されていることを確認してください。',
    alertConfirmAccountCreateDraft:
      '{{myEmail}}として差し込みメールを作成し、下書きとして保存します。',
    alertConfirmAccountSendEmail:
      '{{myEmail}}として差し込みメールを作成し、送信します。',
    errorMailMergeCanceled: 'メールの差し込みが中断されました。',
    errorNoSubjectTextEntered:
      '[Mail Merge Error]\nテンプレートとなるGmail下書きメッセージの件名が入力されていません。',
    errorTwoOrMoreDraftsWithSameSubject:
      '[Mail Merge Error]\n入力された件名を持つ下書きメッセージが2件以上あります。固有の件名で指定してください。',
    errorNoMatchingTemplateDraft:
      '[Mail Merge Error]\n入力された件名を持つ下書きメッセージが見つかりません。件名が正しく入力されていることを確認してください。',
    errorNoToEntered:
      '[Mail Merge Error]\nToとして使用する宛先メールアドレスが空白です。メールアドレスとなる文字列や差し込みフィールドが入力されていることを確認してください。',
    errorMultipleToEntered:
      '[Mail Merge Error]\nToとして複数の宛先メールアドレスの差し込みフィールドが入力されています。差し込みフィールドとして使える宛先メールアドレスは1つのみです。',
    errorInvalidTo:
      '[Mail Merge Error]\n入力された宛先メールアドレスの列（フィールド）名が無効です。実際に存在するフィールド（列）が指定されていることを確認してください。',
    alertCompleteAllDraftsCreated:
      '完了：{{messageCount}}通の下書きが作成されました。差し込み済みのメールが表示されない場合、Gmailの下書き画面を更新してください。',
    alertCompleteAllMailsSent:
      '完了：{{messageCount}}通の差し込みメールの送信が完了しました。',
    alertConfirmSendingOfDraft:
      '{{myEmail}}として、作成された下書きメールを送信してもよろしいですか？\n「下書き作成（差込テスト）」によって作成された下書きのみが送信されます。',
    errorSendDraftsCanceled: '下書きメールの送信がキャンセルされました。',
    errorNoDraftToSend:
      '[Mail Merge Error]\n送信する下書きメールが見つかりません。改めて「下書き作成（差込テスト）」を行ってください。',
    alertCompleteSavedUserConfig:
      '完了：ユーザ設定として現在の設定が保存されました。\n\n',
    appsScriptMessageErrorOnOpenByUrl:
      'SpreadsheetApp オブジェクトでの openByUrl メソッドまたはプロパティの取得中に予期しないエラーが発生',
    appsScriptMessageNoPermissionErrorStartsWith:
      'リクエストされたドキュメントにアクセスする権限がありません。',
    proceedingToPostProcess:
      '現在の処理はGoogle Workspaceアドオンに設けられた{{actionLimitTimeInSec}}秒の実行時間制限を超過する見込みです。残りの差し込み処理はトリガーによって自動的に実行され、完了は{{myEmail}}宛のメールで通知されます。\n下書きモード：{{draftMode}}\n現在までに完了したメッセージ数：{{messageCount}}',
    continuingPostProcess:
      '差し込み処理が継続中です\n現在の処理はGoogle Apps Scriptに設けられた{{appsScriptLimitTimeInSec}}秒の実行時間制限を超過する見込みです。残りの差し込み処理はトリガーによって自動的に継続実行され、完了は{{myEmail}}宛のメールで通知されます。\n下書きモード：{{draftMode}}\n現在までに完了したメッセージ数：{{messageCount}}',
    subjectCompletePostProcess: '[GROUP MERGE] 差し込み処理のお知らせ',
  },
};

class LocalizedMessage {
  constructor(userLocale) {
    this.DEFAULT_LOCALE = 'en';
    this.locale = MESSAGES[userLocale] ? userLocale : this.DEFAULT_LOCALE;
    this.messageList = MESSAGES[this.locale];
    Object.keys(MESSAGES[this.DEFAULT_LOCALE]).forEach((key) => {
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
    let replacedText = placeholderValues.reduce(
      (acc, cur) => acc.replace(new RegExp(cur.regexp, 'g'), cur.value),
      text
    );
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
        regexp: '{{myEmail}}',
        value: myEmail,
      },
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
    let text = draftMode
      ? this.messageList.alertConfirmAccountCreateDraft
      : this.messageList.alertConfirmAccountSendEmail;
    let placeholderValues = [
      {
        regexp: '{{myEmail}}',
        value: myEmail,
      },
    ];
    text = this.replacePlaceholders_(text, placeholderValues);
    return text;
  }
  /**
   * Replace placeholder string in this.messageList.alertCompleteAllDraftsCreated or .alertCompleteAllMailsSent
   * @param {boolean} draftMode
   * @param {number} messageCount Number of emails drafted or sent
   * @returns {string} The replaced text.
   */
  replaceCompleteMessage(draftMode, messageCount) {
    let text = draftMode
      ? this.messageList.alertCompleteAllDraftsCreated
      : this.messageList.alertCompleteAllMailsSent;
    let placeholderValues = [
      {
        regexp: '{{messageCount}}',
        value: messageCount,
      },
    ];
    text = this.replacePlaceholders_(text, placeholderValues);
    return text;
  }
  /**
   * Replace placeholder string in this.messageList.proceedingToPostProcess
   * @param {number} actionLimitTimeInSec
   * @param {string} myEmail
   * @param {boolean} draftMode
   * @param {number} messageCount Number of emails drafted or sent
   * @returns {string} The replaced text.
   */
  replaceProceedingToPostProcess(
    actionLimitTimeInSec,
    myEmail,
    draftMode,
    messageCount
  ) {
    let text = this.messageList.proceedingToPostProcess;
    let placeholderValues = [
      {
        regexp: '{{actionLimitTimeInSec}}',
        value: actionLimitTimeInSec,
      },
      {
        regexp: '{{myEmail}}',
        value: myEmail,
      },
      {
        regexp: '{{draftMode}}',
        value: draftMode,
      },
      {
        regexp: '{{messageCount}}',
        value: messageCount,
      },
    ];
    text = this.replacePlaceholders_(text, placeholderValues);
    return text;
  }
  /**
   * Replace placeholder string in this.messageList.continuingPostProcess
   * @param {number} appsScriptLimitTimeInSec
   * @param {string} myEmail
   * @param {boolean} draftMode
   * @param {number} messageCount Number of emails drafted or sent
   * @returns {string} The replaced text.
   */
  replaceContinuingPostProcess(
    appsScriptLimitTimeInSec,
    myEmail,
    draftMode,
    messageCount
  ) {
    let text = this.messageList.continuingPostProcess;
    let placeholderValues = [
      {
        regexp: '{{appsScriptLimitTimeInSec}}',
        value: appsScriptLimitTimeInSec,
      },
      {
        regexp: '{{myEmail}}',
        value: myEmail,
      },
      {
        regexp: '{{draftMode}}',
        value: draftMode,
      },
      {
        regexp: '{{messageCount}}',
        value: messageCount,
      },
    ];
    text = this.replacePlaceholders_(text, placeholderValues);
    return text;
  }
}
