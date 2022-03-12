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

///////////////////////////
// Language Localization //
///////////////////////////

const MESSAGES = {
  // Message Naming Rules:
  // Name of error messages (keys) must start with 'error'
  // to distinguish between expected and unexpected errors
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
    proceedingToPostProcessMailMerge:
      'Since the current mail merge process is expected to exceed the {{actionLimitTimeInSec}}-second execution time limit for Google Workspace add-ons, the remaining process will be conducted on a time-based trigger. An email notification will be sent to {{myEmail}} once the whole process is complete.\nDraft Mode: {{draftMode}}\nMessage Count: {{messageCount}}',
    continuingPostProcessMailMerge:
      'CONTINUING MAIL MERGE\nSince the current mail merge process is expected to exceed the {{appsScriptLimitTimeInSec}}-second execution time limit for Google Apps Script, the remaining process will be continued on a time-based trigger. An email notification will be sent to {{myEmail}} once the whole process is complete.\nDraft Mode: {{draftMode}}\nMessage Count: {{messageCount}}',
    proceedingToPostProcessSendDrafts:
      'Since the current send-drafts process is expected to exceed the {{actionLimitTimeInSec}}-second execution time limit for Google Workspace add-ons, the remaining process will be conducted on a time-based trigger. An email notification will be sent to {{myEmail}} once the whole process is complete.\nMessage Count: {{messageCount}}',
    continuingPostProcessSendDrafts:
      'CONTINUING "SEND DRAFTS"\nSince the current send-drafts process is expected to exceed the {{appsScriptLimitTimeInSec}}-second execution time limit for Google Apps Script, the remaining process will be continued on a time-based trigger. An email notification will be sent to {{myEmail}} once the whole process is complete.\nMessage Count: {{messageCount}}',
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
    proceedingToPostProcessMailMerge:
      '現在の差し込み処理はGoogle Workspaceアドオンに設けられた{{actionLimitTimeInSec}}秒の実行時間制限を超過する見込みです。残りの処理はトリガーによって自動的に実行され、完了は{{myEmail}}宛のメールで通知されます。\n\n下書きモード：{{draftMode}}\n現在までに完了したメッセージ数：{{messageCount}}',
    continuingPostProcessMailMerge:
      '差し込み処理が継続中です\n現在の処理はGoogle Apps Scriptに設けられた{{appsScriptLimitTimeInSec}}秒の実行時間制限を超過する見込みです。残りの処理はトリガーによって自動的に継続実行され、完了は{{myEmail}}宛のメールで通知されます。\n\n下書きモード：{{draftMode}}\n現在までに完了したメッセージ数：{{messageCount}}',
    proceedingToPostProcessSendDrafts:
      '現在の「作成済みの下書きを送信」処理はGoogle Workspaceアドオンに設けられた{{actionLimitTimeInSec}}秒の実行時間制限を超過する見込みです。残りの処理はトリガーによって自動的に実行され、完了は{{myEmail}}宛のメールで通知されます。\n現在までに完了したメッセージ数：{{messageCount}}',
    continuingPostProcessSendDrafts:
      '「作成済みの下書きを送信」処理が継続中です\n現在の処理はGoogle Apps Scriptに設けられた{{appsScriptLimitTimeInSec}}秒の実行時間制限を超過する見込みです。残りの処理はトリガーによって自動的に継続実行され、完了は{{myEmail}}宛のメールで通知されます。\n\n現在までに完了したメッセージ数：{{messageCount}}',
    subjectPostProcessUpdate: '[GROUP MERGE] 差し込みメール処理',
  },
};

class LocalizedMessage {
  /**
   * @param {String} userLocale
   */
  constructor(userLocale) {
    this.DEFAULT_LOCALE = 'en';
    if (userLocale.match(/^[a-z]{2}(-[A-Z]{2})?$/g)) {
      // Event objects in Google Workspace Add-ons contain user locale info
      // in the format two-letter ISO 639 language code (e.g., "en" and "ja")
      // or with an additiona hyphen followed by ISO 3166 country/region code (e.g., "en-GB")
      // https://developers.google.com/apps-script/add-ons/concepts/event-objects#common_event_object
      let availableLocales = Object.keys(MESSAGES);
      let twoLetterUserLocale = userLocale.slice(0, 2);
      if (availableLocales.includes(userLocale)) {
        this.locale = userLocale;
      } else if (availableLocales.includes(twoLetterUserLocale)) {
        this.locale = twoLetterUserLocale;
      } else {
        this.locale = this.DEFAULT_LOCALE;
      }
    } else {
      this.locale = this.DEFAULT_LOCALE;
    }
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
   * Replace placeholder string in this.messageList.proceedingToPostProcessMailMerge
   * @param {number} actionLimitTimeInSec
   * @param {string} myEmail
   * @param {boolean} draftMode
   * @param {number} messageCount Number of emails drafted or sent
   * @returns {string} The replaced text.
   */
  replaceProceedingToPostProcessMailMerge(
    actionLimitTimeInSec,
    myEmail,
    draftMode,
    messageCount
  ) {
    let text = this.messageList.proceedingToPostProcessMailMerge;
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
   * Replace placeholder string in this.messageList.continuingPostProcessMailMerge
   * @param {number} appsScriptLimitTimeInSec
   * @param {string} myEmail
   * @param {boolean} draftMode
   * @param {number} messageCount Number of emails drafted or sent
   * @returns {string} The replaced text.
   */
  replaceContinuingPostProcessMailMerge(
    appsScriptLimitTimeInSec,
    myEmail,
    draftMode,
    messageCount
  ) {
    let text = this.messageList.continuingPostProcessMailMerge;
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
  /**
   * Replace placeholder string in this.messageList.proceedingToPostProcessSendDrafts
   * @param {number} actionLimitTimeInSec
   * @param {string} myEmail
   * @param {number} messageCount Number of emails drafted or sent
   * @returns {string} The replaced text.
   */
  replaceProceedingToPostProcessSendDrafts(
    actionLimitTimeInSec,
    myEmail,
    messageCount
  ) {
    let text = this.messageList.proceedingToPostProcessSendDrafts;
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
        regexp: '{{messageCount}}',
        value: messageCount,
      },
    ];
    text = this.replacePlaceholders_(text, placeholderValues);
    return text;
  }
  /**
   * Replace placeholder string in this.messageList.continuingPostProcessSendDrafts
   * @param {number} appsScriptLimitTimeInSec
   * @param {string} myEmail
   * @param {number} messageCount Number of emails drafted or sent
   * @returns {string} The replaced text.
   */
  replaceContinuingPostProcessSendDrafts(
    appsScriptLimitTimeInSec,
    myEmail,
    messageCount
  ) {
    let text = this.messageList.continuingPostProcessSendDrafts;
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
        regexp: '{{messageCount}}',
        value: messageCount,
      },
    ];
    text = this.replacePlaceholders_(text, placeholderValues);
    return text;
  }
}

////////////////////////////
// Other Global Variables //
////////////////////////////

// Default configurations
const DEFAULT_CONFIG = {
  SPREADSHEET_URL: '', // 'https://docs.google.com/spreadsheets/d/*****/edit'
  DATA_SHEET_NAME: '', // 'sheetName'
  TO: '{{Email}}',
  CC: '',
  BCC: '',
  TEMPLATE_SUBJECT: '', // 'Enter subject here'
  REPLACE_VALUE: 'NA',
  MERGE_FIELD_MARKER_TEXT: '[{]{2}([^{^}]+)[}]{2}',
  MERGE_FIELD_MARKER: /[{]{2}([^{^}]+)[}]{2}/g, // deprecated
  ENABLE_GROUP_MERGE: true,
  GROUP_FIELD_MARKER_TEXT: '[[]{2}([^[^\\]]+)[\\]]{2}',
  GROUP_FIELD_MARKER: /[[]{2}([^[^\]]+)[\]]{2}/g, // deprecated
  ROW_INDEX_MARKER: '{{i}}',
  ENABLE_REPLY_TO: false,
  REPLY_TO: 'replyTo@email.com',
  ENABLE_DEBUG_MODE: false,
};
// Key(s) for storing and calling settings stored as user property
const UP_KEY_CREATED_DRAFT_IDS = 'createdDraftIds';
const UP_KEY_PREV_MAIL_MERGE_CONFIG = 'prevMailMergeConfig';
const UP_KEY_PREV_SEND_DRAFTS_CONFIG = 'prevSendDraftsConfig';
const UP_KEY_USER_CONFIG = 'userConfig';
// Key lengths in milliseconds regarding time limit of add-on card actions.
const ADDON_TIME_LIMIT = 30 * 1000; // Card actions have a limited execution time of maximum 30 seconds, in milliseconds. See https://developers.google.com/workspace/add-ons/concepts/actions#callback_functions
const ADDON_TIME_LIMIT_OFFSET = 5 * 1000; // Milliseconds prior to the actual ADDON_TIME_LIMIT at which the script will break the current mail merge process for a scheduled trigger.
const APPS_SCRIPT_TIME_LIMIT = 6 * 60 * 1000; // Apps Script execution time limit. 6 minutes, in milliseconds, for all users. See https://developers.google.com/apps-script/guides/services/quotas#current_limitations
const APPS_SCRIPT_TIME_LIMIT_OFFSET = 30 * 1000; // Milliseconds prior to the actual APPS_SCRIPT_TIME_LIMIT at which the script will break the current mail merge process for a scheduled trigger.
const EXECUTE_TRIGGER_AFTER = 5 * 1000; // Milliseconds after which the mail merge process will resume by a scheduled trigger.

//////////////////////////
// Add-on Card Builders //
//////////////////////////

/**
 * Function to create the homepage card for the add-on.
 * @param {Object} event Google Workspace Add-on Event object.
 * @see https://developers.google.com/workspace/add-ons/concepts/event-objects
 */
function buildHomepage(event) {
  var up = PropertiesService.getUserProperties();
  var config =
    JSON.parse(up.getProperty(UP_KEY_USER_CONFIG)) ||
    JSON.parse(up.getProperty(UP_KEY_PREV_MAIL_MERGE_CONFIG)) ||
    DEFAULT_CONFIG;
  return createMailMergeCard(
    event.commonEventObject.userLocale,
    event.commonEventObject.hostApp,
    config
  );
}

/**
 * Function to reset (re-create) the homepage card with user config values.
 * @param {Object} event Google Workspace Add-on Event object.
 * @see https://developers.google.com/workspace/add-ons/concepts/event-objects
 */
function buildHomepageRestoreUserConfig(event) {
  var config =
    JSON.parse(
      PropertiesService.getUserProperties().getProperty(UP_KEY_USER_CONFIG)
    ) || DEFAULT_CONFIG;
  return createMailMergeCard(
    event.commonEventObject.userLocale,
    event.commonEventObject.hostApp,
    config
  );
}

/**
 * Function to reset (re-create) the homepage card with default config values.
 * @param {Object} event Google Workspace Add-on Event object.
 * @see https://developers.google.com/workspace/add-ons/concepts/event-objects
 */
function buildHomepageRestoreDefault(event) {
  PropertiesService.getUserProperties()
    .setProperty(UP_KEY_USER_CONFIG, '[{}]')
    .setProperty(UP_KEY_PREV_MAIL_MERGE_CONFIG, '[{}]');
  return createMailMergeCard(
    event.commonEventObject.userLocale,
    event.commonEventObject.hostApp,
    DEFAULT_CONFIG
  );
}

/**
 * Homepage card builder
 * @param {string} userLocale User locale obtained from the add-on event object
 * https://developers.google.com/workspace/add-ons/concepts/event-objects#common_event_object
 * @param {string} hostApp Name of host that the add-on was call on; obtained from the add-on event object
 * @param {Object} config A set of pre-defined configurations
 */
function createMailMergeCard(userLocale, hostApp, config) {
  // Load localized messages
  const localizedMessage = new LocalizedMessage(userLocale);
  // Load user properties
  const userProperties = PropertiesService.getUserProperties();
  const createdDraftIds = JSON.parse(
    userProperties.getProperty(UP_KEY_CREATED_DRAFT_IDS)
  );
  const disableSendDrafts = !createdDraftIds || createdDraftIds.length == 0;
  // Get URL of currently open spreadsheet if host is Google Sheets
  let ssUrl = null;
  if (hostApp == 'SHEETS') {
    ssUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  }
  // Build card
  let builder = CardService.newCardBuilder();
  // Message Section
  builder.addSection(
    CardService.newCardSection()
      .addWidget(
        CardService.newTextParagraph().setText(
          localizedMessage.messageList.cardHomepageMessage
        )
      )
      .addWidget(
        CardService.newButtonSet().addButton(
          CardService.newTextButton()
            .setText(localizedMessage.messageList.buttonSendDrafts)
            .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
            .setOnClickAction(
              CardService.newAction().setFunctionName('sendCreatedDrafts')
            )
            .setDisabled(disableSendDrafts)
        )
      )
  );
  // Recipient List Section
  builder.addSection(
    CardService.newCardSection()
      .setHeader(localizedMessage.messageList.cardRecipientListSettings)
      .addWidget(
        CardService.newTextInput()
          .setFieldName('SPREADSHEET_URL')
          .setTitle(localizedMessage.messageList.cardEnterSpreadsheetUrl)
          .setHint(localizedMessage.messageList.cardHintSpreadsheetUrl)
          .setValue(
            ssUrl || config.SPREADSHEET_URL || DEFAULT_CONFIG.SPREADSHEET_URL
          )
      )
      .addWidget(
        CardService.newTextInput()
          .setFieldName('DATA_SHEET_NAME')
          .setTitle(localizedMessage.messageList.cardEnterSheetName)
          .setHint(localizedMessage.messageList.cardHintSheetName)
          .setValue(config.DATA_SHEET_NAME || DEFAULT_CONFIG.DATA_SHEET_NAME)
      )
      .addWidget(
        CardService.newTextInput()
          .setFieldName('TO')
          .setTitle(localizedMessage.messageList.cardEnterTo)
          .setHint(localizedMessage.messageList.cardHintTo)
          .setValue(config.TO || DEFAULT_CONFIG.TO)
      )
      .addWidget(
        CardService.newTextInput()
          .setFieldName('CC')
          .setTitle(localizedMessage.messageList.cardEnterCc)
          .setHint(localizedMessage.messageList.cardHintCc)
          .setValue(config.CC || DEFAULT_CONFIG.CC)
      )
      .addWidget(
        CardService.newTextInput()
          .setFieldName('BCC')
          .setTitle(localizedMessage.messageList.cardEnterBcc)
          .setHint(localizedMessage.messageList.cardHintBcc)
          .setValue(config.BCC || DEFAULT_CONFIG.BCC)
      )
  );
  // Template Draft Section
  builder.addSection(
    CardService.newCardSection()
      .setHeader(localizedMessage.messageList.cardTemplateDraftSettings)
      .addWidget(
        CardService.newTextInput()
          .setFieldName('TEMPLATE_SUBJECT')
          .setTitle(localizedMessage.messageList.cardEnterTemplateSubject)
          .setHint(localizedMessage.messageList.cardHintTemplateSubject)
          .setValue(config.TEMPLATE_SUBJECT || DEFAULT_CONFIG.TEMPLATE_SUBJECT)
      )
      .addWidget(
        CardService.newDecoratedText()
          .setText(localizedMessage.messageList.cardSwitchEnableGroupMerge)
          .setSwitchControl(
            CardService.newSwitch()
              .setSelected(
                typeof config.ENABLE_GROUP_MERGE == 'boolean'
                  ? config.ENABLE_GROUP_MERGE
                  : DEFAULT_CONFIG.ENABLE_GROUP_MERGE
              )
              .setFieldName('ENABLE_GROUP_MERGE')
              .setValue('enabled')
          )
      )
  );
  // Advanced Settings Section
  builder.addSection(
    CardService.newCardSection()
      .setHeader(localizedMessage.messageList.cardAdvancedSettings)
      .setCollapsible(true)
      .addWidget(
        CardService.newDecoratedText()
          .setText(localizedMessage.messageList.cardSwitchEnableReplyTo)
          .setSwitchControl(
            CardService.newSwitch()
              .setSelected(
                typeof config.ENABLE_REPLY_TO == 'boolean'
                  ? config.ENABLE_REPLY_TO
                  : DEFAULT_CONFIG.ENABLE_REPLY_TO
              )
              .setFieldName('ENABLE_REPLY_TO')
              .setValue('enabled')
          )
      )
      .addWidget(
        CardService.newTextInput()
          .setFieldName('REPLY_TO')
          .setTitle(localizedMessage.messageList.cardEnterReplyTo)
          .setHint(localizedMessage.messageList.cardHintReplyTo)
          .setValue(config.REPLY_TO || DEFAULT_CONFIG.REPLY_TO)
      )
      .addWidget(
        CardService.newTextInput()
          .setFieldName('REPLACE_VALUE')
          .setTitle(localizedMessage.messageList.cardEnterReplaceValue)
          .setHint(localizedMessage.messageList.cardHintReplaceValue)
          .setValue(config.REPLACE_VALUE || DEFAULT_CONFIG.REPLACE_VALUE)
      )
      .addWidget(
        CardService.newTextInput()
          .setFieldName('MERGE_FIELD_MARKER_TEXT')
          .setTitle(localizedMessage.messageList.cardEnterMergeFieldMarker)
          .setHint(localizedMessage.messageList.cardHintMergeFieldMarker)
          .setValue(
            config.MERGE_FIELD_MARKER_TEXT ||
              DEFAULT_CONFIG.MERGE_FIELD_MARKER_TEXT
          )
      )
      .addWidget(
        CardService.newTextInput()
          .setFieldName('GROUP_FIELD_MARKER_TEXT')
          .setTitle(localizedMessage.messageList.cardEnterGroupFieldMarker)
          .setHint(localizedMessage.messageList.cardHintGroupFieldMarker)
          .setValue(
            config.GROUP_FIELD_MARKER_TEXT ||
              DEFAULT_CONFIG.GROUP_FIELD_MARKER_TEXT
          )
      )
      .addWidget(
        CardService.newTextInput()
          .setFieldName('ROW_INDEX_MARKER')
          .setTitle(localizedMessage.messageList.cardEnterRowIndexMarker)
          .setHint(localizedMessage.messageList.cardHintRowIndexMarker)
          .setValue(config.ROW_INDEX_MARKER || DEFAULT_CONFIG.ROW_INDEX_MARKER)
      )
      .addWidget(
        CardService.newDecoratedText()
          .setText(localizedMessage.messageList.cardSwitchEnableDebugMode)
          .setSwitchControl(
            CardService.newSwitch()
              .setSelected(
                typeof config.ENABLE_DEBUG_MODE == 'boolean'
                  ? config.ENABLE_DEBUG_MODE
                  : DEFAULT_CONFIG.ENABLE_DEBUG_MODE
              )
              .setFieldName('ENABLE_DEBUG_MODE')
              .setValue('enabled')
          )
      )
  );
  // Buttons Section
  builder.addSection(
    CardService.newCardSection().addWidget(
      CardService.newButtonSet()
        .addButton(
          CardService.newTextButton()
            .setText(localizedMessage.messageList.buttonRestoreUserConfig)
            .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
            .setOnClickAction(
              CardService.newAction().setFunctionName(
                'buildHomepageRestoreUserConfig'
              )
            )
        )
        .addButton(
          CardService.newTextButton()
            .setText(localizedMessage.messageList.buttonSaveUserConfig)
            .setOnClickAction(
              CardService.newAction().setFunctionName('saveUserConfig')
            )
        )
        .addButton(
          CardService.newTextButton()
            .setText(localizedMessage.messageList.buttonRestoreDefault)
            .setOnClickAction(
              CardService.newAction().setFunctionName(
                'buildHomepageRestoreDefault'
              )
            )
        )
        .addButton(
          CardService.newTextButton()
            .setText(localizedMessage.messageList.buttonReadDocument)
            .setOpenLink(
              CardService.newOpenLink().setUrl(
                localizedMessage.messageList.buttonReadDocumentUrl
              )
            )
        )
      /*
      .addButton(CardService.newTextButton()
        .setText('test')
        .setOnClickAction(CardService.newAction().setFunctionName('test')))
      */
    )
  );
  // Fixed Footer
  builder.setFixedFooter(
    CardService.newFixedFooter()
      .setPrimaryButton(
        CardService.newTextButton()
          .setText(localizedMessage.messageList.buttonCreateDrafts)
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setOnClickAction(
            CardService.newAction().setFunctionName('createDraftEmails')
          )
          .setDisabled(false)
      )
      .setSecondaryButton(
        CardService.newTextButton()
          .setText(localizedMessage.messageList.buttonSendEmails)
          .setOnClickAction(
            CardService.newAction().setFunctionName('sendEmails')
          )
          .setDisabled(false)
      )
  );
  return builder.build();
}

/*
function test(event) {
  console.log(JSON.stringify(event));
}
*/

/**
 * Builder for message cards to present error and other messages to the add-on user.
 * @param {String} message Message string that can accept some basic HTML formatting
 * described in https://developers.google.com/workspace/add-ons/concepts/widgets?hl=en#text_formatting
 * @param {String} userLocale Two-letter user locale like 'ja', included in
 * the Google Workspace Add-on event object at event.commonEventObject.userLocale
 */
function createMessageCard(message, userLocale) {
  var localizedMessage = new LocalizedMessage(userLocale);
  var builder = CardService.newCardBuilder()
    .addSection(
      CardService.newCardSection().addWidget(
        CardService.newTextParagraph().setText(message)
      )
    )
    .addSection(
      CardService.newCardSection().addWidget(
        CardService.newButtonSet().addButton(
          CardService.newTextButton()
            .setText(localizedMessage.messageList.buttonReturnHome)
            .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
            .setOnClickAction(
              CardService.newAction().setFunctionName('buildHomepage')
            )
        )
      )
    );
  return builder.build();
}

////////////////////////
// Mail Merge Actions //
////////////////////////

/**
 * Save the form values as config to user property.
 * @param {Object} event Google Workspace Add-on Event object.
 * @see https://developers.google.com/workspace/add-ons/concepts/event-objects
 */
function saveUserConfig(event) {
  const config = parseConfig_(event);
  // Save on user property
  PropertiesService.getUserProperties().setProperty(
    UP_KEY_USER_CONFIG,
    JSON.stringify(config)
  );
  // Construct complete message
  const localizedMessage = new LocalizedMessage(config.userLocale);
  let cardMessage = localizedMessage.messageList.alertCompleteSavedUserConfig;
  if (config.ENABLE_DEBUG_MODE) {
    for (let k in config) {
      cardMessage += `<b>${k}</b>: ${config[k]}\n`;
    }
  }
  return createMessageCard(cardMessage, config.userLocale);
}

/**
 * Create draft of personalized email(s)
 * @param {Object} event Google Workspace Add-on Event object.
 * @see https://developers.google.com/workspace/add-ons/concepts/event-objects
 */
function createDraftEmails(event) {
  const draftMode = true;
  const config = parseConfig_(event);
  return createMessageCard(
    mailMerge(draftMode, config),
    event.commonEventObject.userLocale
  );
}

/**
 * Send personalized email(s)
 * @param {Object} event Google Workspace Add-on Event object.
 * @see https://developers.google.com/workspace/add-ons/concepts/event-objects
 */
function sendEmails(event) {
  const draftMode = false;
  const config = parseConfig_(event);
  return createMessageCard(
    mailMerge(draftMode, config),
    event.commonEventObject.userLocale
  );
}

/**
 * Core function of mail merge
 * @param {Boolean} draftMode Creates Gmail draft(s) instead of sending email. Defaults to true.
 * @param {Object} config Object returned by parseConfig_(eventObj)
 * @param {Object} prevProperties Properties handed over from the original mail merge process. Required on triggered executions.
 * @return {String} Message to display as the add-on card, be it an error message or a notification that the merge process is complete.
 */
function mailMerge(
  draftMode = true,
  config = DEFAULT_CONFIG,
  prevProperties = {}
) {
  // Determine whether this script is executed manually or by a time-trigger
  const isTimeTriggered = Object.keys(prevProperties).length > 0;
  const executionTimeThreshold = isTimeTriggered
    ? APPS_SCRIPT_TIME_LIMIT - APPS_SCRIPT_TIME_LIMIT_OFFSET
    : ADDON_TIME_LIMIT - ADDON_TIME_LIMIT_OFFSET;
  const myEmail = Session.getActiveUser().getEmail();
  const localizedMessage = new LocalizedMessage(config.userLocale);
  // Designate name of fields without placeholders, i.e. values that can be skipped for the merge process later on
  const noPlaceholder = ['from', 'attachments', 'inLineImages', 'labels'];
  const userProperties = PropertiesService.getUserProperties();
  // Save current settings in user property
  userProperties.setProperty(
    UP_KEY_PREV_MAIL_MERGE_CONFIG,
    JSON.stringify(config)
  );
  let debugInfo = {
    completedRecipients: [],
    config: config,
    draftMode: draftMode,
    exeRoundsMailMerge: isTimeTriggered ? prevProperties.exeRoundsMailMerge : 0,
    prevCompletedRecipients: isTimeTriggered
      ? prevProperties.completedRecipients
      : [],
    processTime: [],
    start: new Date().getTime(),
  };
  let cardMessage = '';
  let templateDraftIds = isTimeTriggered ? prevProperties.templateDraftIds : [];
  // Reset list of created drafts
  let createdDraftIds = isTimeTriggered ? prevProperties.createdDraftIds : [];
  if (config.hostApp == 'SHEETS' && !isTimeTriggered) {
    var ui = SpreadsheetApp.getUi();
  }
  try {
    let ss = SpreadsheetApp.openByUrl(config.SPREADSHEET_URL);
    // Get data of field(s) to merge in form of 2d array
    let dataSheet = ss.getSheetByName(config.DATA_SHEET_NAME);
    if (dataSheet == null) {
      throw new Error(localizedMessage.messageList.errorSheetNotFound);
    }
    let mergeData = dataSheet.getDataRange().getDisplayValues();
    // Convert line breaks in the spreadsheet (in LF format, i.e., '\n')
    // to CRLF format ('\r\n') for merging into Gmail plain text
    let mergeDataEolReplaced = mergeData.map((element) =>
      element.map((value) => value.replace(/\n|\r|\r\n/g, '\r\n'))
    );
    debugInfo.processTime.push(
      `Retrieved recipient list data at ${
        new Date().getTime() - debugInfo.start
      } (millisec from start)`
    );
    if (config.hostApp == 'SHEETS' && !isTimeTriggered) {
      // Confirmation before sending email
      let confirmAccount = localizedMessage.replaceConfirmAccount(
        draftMode,
        myEmail
      );
      let answer = ui.alert(confirmAccount, ui.ButtonSet.OK_CANCEL);
      if (answer !== ui.Button.OK) {
        throw new Error(localizedMessage.messageList.errorMailMergeCanceled);
      }
      debugInfo.processTime.push(
        `Passed UI confirmation on SHEETS at ${
          new Date().getTime() - debugInfo.start
        } (millisec from start)`
      );
    }
    // Verify value of TO
    var Tos = [...config.TO.matchAll(config.MERGE_FIELD_MARKER)];
    if (config.TO == null || config.TO == '') {
      throw new Error(localizedMessage.messageList.errorNoToEntered);
    } else if (Tos.length > 1) {
      throw new Error(localizedMessage.messageList.errorMultipleToEntered);
    }
    // Verify value of TEMPLATE_SUBJECT
    if (config.TEMPLATE_SUBJECT == null || config.TEMPLATE_SUBJECT == '') {
      throw new Error(localizedMessage.messageList.errorNoSubjectTextEntered);
    }
    // Get an array of GmailMessage class objects whose subject matches config.TEMPLATE_SUBJECT
    // or if templateDraftIds is not empty, directly retrieve the GmailMessage object by its draft ID.
    let draftMessages = [];
    if (templateDraftIds.length > 0) {
      draftMessages = templateDraftIds.map((draftId) =>
        GmailApp.getDraft(draftId).getMessage()
      );
    } else {
      // draftMessages = GmailApp.getDraftMessages().filter(element => element.getSubject() == config.TEMPLATE_SUBJECT);
      draftMessages = GmailApp.getDrafts().reduce((messages, draft) => {
        let message = draft.getMessage();
        if (message.getSubject() === config.TEMPLATE_SUBJECT) {
          messages.push(message);
          templateDraftIds.push(draft.getId());
        }
        return messages;
      }, []);
    }
    // Check for duplicates
    if (draftMessages.length > 1) {
      throw new Error(
        localizedMessage.messageList.errorTwoOrMoreDraftsWithSameSubject
      );
    }
    // Throw error if no draft template is found
    if (draftMessages.length == 0) {
      throw new Error(
        localizedMessage.messageList.errorNoMatchingTemplateDraft
      );
    }
    // Store template into an object
    let [messageTo, messageCc, messageBcc] = [
      draftMessages[0].getTo(),
      draftMessages[0].getCc(),
      draftMessages[0].getBcc(),
    ];
    let template = {
      subject: config.TEMPLATE_SUBJECT,
      plainBody: draftMessages[0].getPlainBody(),
      htmlBody: draftMessages[0].getBody(),
      from: draftMessages[0].getFrom(),
      to: config.TO + (messageTo ? `,${messageTo}` : ''),
      cc: config.CC + (messageCc ? `,${messageCc}` : ''),
      bcc: config.BCC + (messageBcc ? `,${messageBcc}` : ''),
      attachments: draftMessages[0].getAttachments({
        includeInlineImages: false,
        includeAttachments: true,
      }),
      inLineImages: draftMessages[0].getAttachments({
        includeInlineImages: true,
        includeAttachments: false,
      }),
      labels: draftMessages[0].getThread().getLabels(),
      replyTo: config.ENABLE_REPLY_TO ? config.REPLY_TO : '',
    };
    template.to =
      template.to.slice(0, 1) === ',' ? template.to.slice(1) : template.to;
    template.cc =
      template.cc.slice(0, 1) === ',' ? template.cc.slice(1) : template.cc;
    template.bcc =
      template.bcc.slice(0, 1) === ',' ? template.bcc.slice(1) : template.bcc;
    debugInfo.processTime.push(
      `Retrieved template draft data at ${
        new Date().getTime() - debugInfo.start
      } (millisec from start)`
    );
    // Check template format; plain or HTML text.
    let isPlainText = template.plainBody === template.htmlBody;
    let inLineImageBlobs = {};
    if (!isPlainText) {
      // If the template is composed in HTML, check for in-line images
      // and create a mapping object of cid (content ID) and its corresponding in-line image blob.
      let regExpImgTag =
        /<img data-surl="cid:(?<cidDataSurl>[^"]+)"[^>]+src="cid:(?<cidSrc>[^"]+)"[^>]+alt="(?<blobName>[^"]+)"[^>]+>/g;
      let inLineImageTags = [...template.htmlBody.matchAll(regExpImgTag)];
      inLineImageBlobs = template.inLineImages.reduce((obj, blob) => {
        let cid = inLineImageTags.find(
          (element) => element.groups.blobName == blob.getName()
        ).groups.cidSrc;
        obj[cid] = blob;
        return obj;
      }, {});
      debugInfo.processTime.push(
        `Template is composed in HTML. Retrieved ${
          template.inLineImages.length
        } in-line image data at ${
          new Date().getTime() - debugInfo.start
        } (millisec from start)`
      );
    }
    // Convert the 2d-array merge data into object, grouped by recipient(s) if group merge is enabled
    let mergeDataHeader = mergeDataEolReplaced.shift();
    let mergeDataKeyIndex = mergeDataHeader.indexOf(Tos[0][1]);
    if (mergeDataKeyIndex < 0) {
      // Throw error if there were no column header matching the placeholder value of To
      throw new Error(localizedMessage.messageList.errorInvalidTo);
    }
    let groupedMergeData = mergeDataEolReplaced.reduce(
      (obj, dataRow, rowIndex) => {
        let key = dataRow[mergeDataKeyIndex];
        key += config.ENABLE_GROUP_MERGE ? '' : `_${rowIndex}`;
        if (!obj[key]) {
          obj[key] = [];
        }
        obj[key].push(
          mergeDataHeader.reduce((o, k, i) => {
            o[k] = dataRow[i];
            return o;
          }, {})
        );
        return obj;
      },
      {}
    );
    debugInfo.processTime.push(
      `Group Merge is ${
        config.ENABLE_GROUP_MERGE
          ? 'enabled. Grouped recipient list by recipient email address'
          : 'disabled. Data stored in groupedMergeData'
      } at ${new Date().getTime() - debugInfo.start} (millisec from start)`
    ); // Create draft or send email based on the template.
    let fillInTemplate_options = {
      excludeFromTemplate: noPlaceholder,
      asHtml: [isPlainText ? '' : 'htmlBody'],
      replaceValue: config.REPLACE_VALUE,
      mergeFieldMarker: config.MERGE_FIELD_MARKER,
      enableGroupMerge: config.ENABLE_GROUP_MERGE,
      groupFieldMarker: config.GROUP_FIELD_MARKER,
      rowIndexMarker: config.ROW_INDEX_MARKER,
    };
    // Create draft for each recipient and, depending on the value of draftMode, send it.
    for (let k in groupedMergeData) {
      if (debugInfo.prevCompletedRecipients.includes(k)) {
        continue;
      }
      let messageData = fillInTemplate_(
        template,
        groupedMergeData[k],
        fillInTemplate_options
      );
      let options = {
        htmlBody: isPlainText ? null : messageData.htmlBody,
        from: messageData.from,
        cc: messageData.cc,
        bcc: messageData.bcc,
        attachments: messageData.attachments ? messageData.attachments : null,
        inlineImages: isPlainText ? null : inLineImageBlobs,
        replyTo: config.ENABLE_REPLY_TO ? messageData.replyTo : null,
      };
      let draft = GmailApp.createDraft(
        messageData.to,
        messageData.subject,
        messageData.plainBody,
        options
      );
      // Add the same Gmail labels as those on the template draft message.
      let draftThread = draft.getMessage().getThread();
      messageData.labels.forEach((label) => draftThread.addLabel(label));
      if (!draftMode) {
        draft.send();
      } else {
        // List the created draft ID
        createdDraftIds.push(draft.getId());
      }
      debugInfo.completedRecipients.push(k);
      // Determine current process time lapse.
      let timeElapsed = new Date().getTime() - debugInfo.start;
      debugInfo.processTime.push(
        `Message for ${k} ${
          draftMode ? 'drafted' : 'sent'
        } at ${timeElapsed} (millisec from start)`
      );
      if (timeElapsed >= executionTimeThreshold) {
        // If the script execution time has passed more than executionTimeThreshold,
        // break the process to complete by execution on time-based triggers.
        debugInfo.completedRecipients =
          debugInfo.prevCompletedRecipients.concat(
            debugInfo.completedRecipients
          );
        userProperties.setProperties(
          {
            completedRecipients: JSON.stringify(debugInfo.completedRecipients),
            createdDraftIds: JSON.stringify(createdDraftIds),
            draftMode: draftMode,
            exeRoundsMailMerge: debugInfo.exeRoundsMailMerge,
            templateDraftIds: JSON.stringify(templateDraftIds),
          },
          false
        );
        if (!isTimeTriggered) {
          // If this is the first execution of mailMerge,
          // i.e, it this is manually executed, and not triggered,
          // create a new time-based trigger
          ScriptApp.newTrigger('postProcessMailMerge')
            .timeBased()
            .after(EXECUTE_TRIGGER_AFTER)
            .create();
        } else if (debugInfo.exeRoundsMailMerge == 1) {
          // If this is executed on a time-based trigger for the first time,
          // create a new hourly trigger for further post-process.
          // Hourly, because Google places a limitation on
          // the recurrence interval for an Add-on trigger;
          // it must be at least one hour.
          ScriptApp.newTrigger('postProcessMailMerge')
            .timeBased()
            .everyHours(1)
            .create();
        }
        debugInfo.processTime.push(
          `Saved property and trigger set for post-process at ${
            new Date().getTime() - debugInfo.start
          } (millisec from start)`
        );
        let notificationMessage = isTimeTriggered
          ? localizedMessage.replaceContinuingPostProcessMailMerge(
              APPS_SCRIPT_TIME_LIMIT / 1000,
              myEmail,
              draftMode,
              debugInfo.completedRecipients.length
            )
          : localizedMessage.replaceProceedingToPostProcessMailMerge(
              ADDON_TIME_LIMIT / 1000,
              myEmail,
              draftMode,
              debugInfo.completedRecipients.length
            );
        throw new Error(notificationMessage);
      }
    }
    // Notification
    debugInfo.completedRecipients = debugInfo.prevCompletedRecipients.concat(
      debugInfo.completedRecipients
    );
    cardMessage = localizedMessage.replaceCompleteMessage(
      draftMode,
      debugInfo.completedRecipients.length
    );
    // Delete time-based triggers
    if (isTimeTriggered) {
      ScriptApp.getProjectTriggers().forEach((trigger) =>
        ScriptApp.deleteTrigger(trigger)
      );
    }
  } catch (error) {
    let knownErrorMessages = getKnownErrorMessages_(config.userLocale);
    if (
      knownErrorMessages.includes(error.message) ||
      error.message.startsWith(
        localizedMessage.messageList.proceedingToPostProcessMailMerge.slice(
          0,
          10
        )
      ) ||
      error.message.startsWith(
        localizedMessage.messageList.continuingPostProcessMailMerge.slice(0, 10)
      )
    ) {
      cardMessage = error.message;
    } else if (
      error.message.startsWith(
        localizedMessage.messageList.appsScriptMessageErrorOnOpenByUrl
      ) ||
      error.message.startsWith(
        localizedMessage.messageList
          .appsScriptMessageNoPermissionErrorStartsWith
      )
    ) {
      cardMessage = localizedMessage.messageList.errorSpreadsheetNotFound;
    } else {
      cardMessage =
        localizedMessage.messageList.cardMessageUnexpectedError + error.stack;
    }
  }
  userProperties.setProperty(
    UP_KEY_CREATED_DRAFT_IDS,
    JSON.stringify(createdDraftIds)
  );
  if (config.ENABLE_DEBUG_MODE) {
    let debugInfoText = localizedMessage.messageList.cardMessageDebugInfo;
    for (let k in debugInfo) {
      debugInfoText += `${k}: ${JSON.stringify(debugInfo[k])}\n`;
    }
    MailApp.sendEmail(
      myEmail,
      `[GROUP MERGE] Debug Info`,
      `${cardMessage}\n\n${debugInfoText}`
    );
    cardMessage += `\n\n${localizedMessage.messageList.cardMessageSentDebugInfo}`;
  }
  if (isTimeTriggered) {
    MailApp.sendEmail(
      myEmail,
      localizedMessage.messageList.subjectPostProcessUpdate,
      cardMessage
    );
  }
  return cardMessage;
}

/**
 * Post-process for mailMerge(); continues the process of mail merge
 * for executions that are expected to exceed the Google Workspace Add-ons' 30-second time limit.
 */
function postProcessMailMerge() {
  const userPropertiesValues =
    PropertiesService.getUserProperties().getProperties();
  const draftMode = userPropertiesValues['draftMode'] === 'true';
  let config = JSON.parse(userPropertiesValues[UP_KEY_PREV_MAIL_MERGE_CONFIG]);
  config.MERGE_FIELD_MARKER = new RegExp(config.MERGE_FIELD_MARKER_TEXT, 'g');
  config.GROUP_FIELD_MARKER = new RegExp(config.GROUP_FIELD_MARKER_TEXT, 'g');
  const prevProperties = {
    exeRoundsMailMerge: parseInt(userPropertiesValues.exeRoundsMailMerge) + 1,
    completedRecipients: JSON.parse(userPropertiesValues.completedRecipients),
    createdDraftIds: JSON.parse(userPropertiesValues.createdDraftIds),
    templateDraftIds: JSON.parse(userPropertiesValues.templateDraftIds),
  };
  mailMerge(draftMode, config, prevProperties);
}

/**
 * Send the drafts created by createDraftEmails()
 * @param {Object} event Google Workspace Add-on Event object.
 * @see https://developers.google.com/workspace/add-ons/concepts/event-objects
 */
function sendCreatedDrafts(event) {
  const config = parseConfig_(event);
  return createMessageCard(sendDrafts(config), config.userLocale);
}

/**
 * The core function for sendCreatedDrafts
 * @param {Object} config Object returned by parseConfig_(eventObj)
 * @param {Object} prevProperties Properties handed over from the original mail merge process. Required on triggered executions.
 * @returns {String}  Message to display as the add-on card, be it an error message or a notification that the merge process is complete.
 */
function sendDrafts(config, prevProperties = {}) {
  // Determine whether this script is executed manually or by a time-trigger
  const isTimeTriggered = Object.keys(prevProperties).length > 0;
  const executionTimeThreshold = isTimeTriggered
    ? APPS_SCRIPT_TIME_LIMIT - APPS_SCRIPT_TIME_LIMIT_OFFSET
    : ADDON_TIME_LIMIT - ADDON_TIME_LIMIT_OFFSET;
  const myEmail = Session.getActiveUser().getEmail();
  const localizedMessage = new LocalizedMessage(config.userLocale);
  const userProperties = PropertiesService.getUserProperties();
  // Save current settings in user property
  userProperties.setProperty(
    UP_KEY_PREV_SEND_DRAFTS_CONFIG,
    JSON.stringify(config)
  );
  let debugInfo = {
    config: config,
    exeRoundsSendDrafts: isTimeTriggered
      ? prevProperties.exeRoundsSendDrafts
      : 0,
    processTime: [],
    start: new Date().getTime(),
  };
  let cardMessage = '';
  try {
    let messageCount = 0;
    if (config.hostApp == 'SHEETS' && !isTimeTriggered) {
      // Confirmation before sending email
      var ui = SpreadsheetApp.getUi();
      let answer = ui.alert(
        localizedMessage.replaceConfirmSendingOfDraft(myEmail),
        ui.ButtonSet.OK_CANCEL
      );
      if (answer !== ui.Button.OK) {
        throw new Error(localizedMessage.messageList.errorSendDraftsCanceled);
      }
    }
    // Get the values of createdDraftIds, the string of draft IDs to send, stored in the user property
    let createdDraftIds = isTimeTriggered
      ? prevProperties.createdDraftIds
      : JSON.parse(userProperties.getProperty(UP_KEY_CREATED_DRAFT_IDS));
    if (!createdDraftIds || createdDraftIds.length == 0) {
      // Throw error if no draft ID is stored.
      throw new Error(localizedMessage.messageList.errorNoDraftToSend);
    }
    // Send emails
    let drafts = GmailApp.getDrafts();
    debugInfo.processTime.push(
      `Retrieved Gmail drafts at ${
        new Date().getTime() - debugInfo.start
      } (millisec from start)`
    );
    for (let i = 0; i < drafts.length; i++) {
      let draft = drafts[i];
      let draftId = draft.getId();
      let isSent = false;
      if (createdDraftIds.includes(draftId)) {
        draft.send();
        messageCount += 1;
        isSent = true;
        createdDraftIds = createdDraftIds.filter((id) => id !== draftId);
      }
      // Determine current process time lapse.
      let timeElapsed = new Date().getTime() - debugInfo.start;
      debugInfo.processTime.push(
        `Draft (ID: ${draftId}) ${
          isSent ? 'sent' : 'checked'
        } at ${timeElapsed} (millisec from start)`
      );
      if (timeElapsed >= executionTimeThreshold) {
        // If the script execution time has passed more than executionTimeThreshold,
        // break the process to complete by execution on time-based triggers.
        userProperties.setProperties(
          {
            createdDraftIds: JSON.stringify(createdDraftIds),
            exeRoundsSendDrafts: debugInfo.exeRoundsSendDrafts,
          },
          false
        );
        if (!isTimeTriggered) {
          // If this is the first execution of sendDrafts,
          // i.e, it this is manually executed, and not triggered,
          // create a new time-based trigger
          ScriptApp.newTrigger('postProcessSendDrafts')
            .timeBased()
            .after(EXECUTE_TRIGGER_AFTER)
            .create();
        } else if (debugInfo.exeRoundsSendDrafts == 1) {
          // If this is executed on a time-based trigger for the first time,
          // create a new hourly trigger for further post-process.
          // Hourly, because Google places a limitation on
          // the recurrence interval for an Add-on trigger;
          // it must be at least one hour.
          ScriptApp.newTrigger('postProcessSendDrafts')
            .timeBased()
            .everyHours(1)
            .create();
        }
        debugInfo.processTime.push(
          `Saved property and trigger set for post-process at ${
            new Date().getTime() - debugInfo.start
          } (millisec from start)`
        );
        let notificationMessage = isTimeTriggered
          ? localizedMessage.replaceContinuingPostProcessSendDrafts(
              APPS_SCRIPT_TIME_LIMIT / 1000,
              myEmail,
              messageCount
            )
          : localizedMessage.replaceProceedingToPostProcessSendDrafts(
              ADDON_TIME_LIMIT / 1000,
              myEmail,
              messageCount
            );
        throw new Error(notificationMessage);
      }
    }
    // Empty createdDraftIds
    createdDraftIds = [];
    userProperties.setProperty(
      UP_KEY_CREATED_DRAFT_IDS,
      JSON.stringify(createdDraftIds)
    );
    cardMessage = localizedMessage.replaceCompleteMessage(false, messageCount);
    // Delete time-based triggers
    if (isTimeTriggered) {
      ScriptApp.getProjectTriggers().forEach((trigger) =>
        ScriptApp.deleteTrigger(trigger)
      );
    }
  } catch (error) {
    let knownErrorMessages = getKnownErrorMessages_(config.userLocale);
    if (
      knownErrorMessages.includes(error.message) ||
      error.message.startsWith(
        localizedMessage.messageList.proceedingToPostProcessSendDrafts.slice(
          0,
          10
        )
      ) ||
      error.message.startsWith(
        localizedMessage.messageList.continuingPostProcessSendDrafts.slice(
          0,
          10
        )
      )
    ) {
      cardMessage = error.message;
    } else {
      cardMessage =
        localizedMessage.messageList.cardMessageUnexpectedError + error.stack;
    }
  }
  if (config.ENABLE_DEBUG_MODE) {
    let debugInfoText = localizedMessage.messageList.cardMessageDebugInfo;
    for (let k in debugInfo) {
      debugInfoText += `${k}: ${JSON.stringify(debugInfo[k])}\n`;
    }
    MailApp.sendEmail(
      myEmail,
      `[GROUP MERGE] Debug Info`,
      `${cardMessage}\n\n${debugInfoText}`
    );
    cardMessage += `\n\n${localizedMessage.messageList.cardMessageSentDebugInfo}`;
  }
  if (isTimeTriggered) {
    MailApp.sendEmail(
      myEmail,
      localizedMessage.messageList.subjectPostProcessUpdate,
      cardMessage
    );
  }
  return cardMessage;
}

/**
 * Post-process for sendDrafts(); continues the process of sending created drafts
 * for executions that are expected to exceed the Google Workspace Add-ons' 30-second time limit.
 */
function postProcessSendDrafts() {
  const userPropertiesValues =
    PropertiesService.getUserProperties().getProperties();
  let config = JSON.parse(userPropertiesValues[UP_KEY_PREV_SEND_DRAFTS_CONFIG]);
  config.MERGE_FIELD_MARKER = new RegExp(config.MERGE_FIELD_MARKER_TEXT, 'g');
  config.GROUP_FIELD_MARKER = new RegExp(config.GROUP_FIELD_MARKER_TEXT, 'g');
  const prevProperties = {
    exeRoundsSendDrafts: parseInt(userPropertiesValues.exeRoundsSendDrafts) + 1,
    // completedRecipients: JSON.parse(userPropertiesValues.completedRecipients),
    createdDraftIds: JSON.parse(userPropertiesValues.createdDraftIds),
    // templateDraftIds: JSON.parse(userPropertiesValues.templateDraftIds),
  };
  sendDrafts(config, prevProperties);
}

/**
 * Retrieve configuration values for mail merge from the input values of the Card widget interactions.
 * @param {Object} eventObj Add-on event object. https://developers.google.com/workspace/add-ons/concepts/event-objects
 * @return {Object}
 */
function parseConfig_(eventObj) {
  var targetSettings = [
    // A complete list of the form input items in createMailMergeCard
    'SPREADSHEET_URL',
    'DATA_SHEET_NAME',
    'TO',
    'CC',
    'BCC',
    'TEMPLATE_SUBJECT',
    'ENABLE_GROUP_MERGE',
    'ENABLE_REPLY_TO',
    'REPLY_TO',
    'REPLACE_VALUE',
    'MERGE_FIELD_MARKER_TEXT',
    'GROUP_FIELD_MARKER_TEXT',
    'ROW_INDEX_MARKER',
    'ENABLE_DEBUG_MODE',
  ];
  let configObj = targetSettings.reduce((config, item) => {
    let input = eventObj.commonEventObject.formInputs[item] || {
      stringInputs: { value: [''] },
    };
    let value = input.stringInputs.value[0];
    if (
      item == 'ENABLE_GROUP_MERGE' ||
      item == 'ENABLE_REPLY_TO' ||
      item == 'ENABLE_DEBUG_MODE'
    ) {
      value = value == 'enabled' || value == 'true';
    } else if (
      item == 'MERGE_FIELD_MARKER_TEXT' ||
      item == 'GROUP_FIELD_MARKER_TEXT'
    ) {
      let itemRegexp = item.replace('_TEXT', '');
      let valueRegexp = new RegExp(value, 'g');
      config[itemRegexp] = valueRegexp;
    }
    config[item] = value;
    return config;
  }, {});
  // Add host app and user locale info
  configObj['hostApp'] = eventObj.commonEventObject.hostApp;
  configObj['userLocale'] = eventObj.commonEventObject.userLocale;
  return configObj;
}

/**
 * Replaces markers in a template object with values defined in a JavaScript data object.
 * @param {Object} template Template object containing markers, as designated in regular expression in mergeFieldMarker
 * @param {array} data Array of object(s) with values to replace markers.
 * @param {Object} options [Optional] Options, as described below. Defaults to the respective values in DEFAULT_CONFIG
 * @param {array} options.excludeFromTemplate [Optional] Array of the keys in template to exclude from the replacement process, in strings.
 * @param {array} option.asHtml [Optional] Array of the keys in template to regard as HTML formatted.
 * @param {string} options.replaceValue [Optional] String to replace empty data of a marker.
 * @param {RegExp} options.mergeFieldMarker [Optional] Regular expression for the merge field marker.
 * @param {boolean} options.enableGroupMerge [Optional] Merged texts are returned in concatenated form when true.
 * @param {RegExp} options.groupFieldMarker [Optional] Regular expression for the group merge field marker.
 * @param {string} options.rowIndexMarker [Optional] Marker for merging row index number in a group merge.
 * @return {Object} Returns the template object with markers replaced for personalized text.
 */
function fillInTemplate_(template, data, options) {
  if (!('excludeFromTemplate' in options)) {
    options['excludeFromTemplate'] = [];
  }
  if (!('asHtml' in options)) {
    options['asHtml'] = [];
  }
  if (!('replaceValue' in options)) {
    options['replaceValue'] = DEFAULT_CONFIG.REPLACE_VALUE;
  }
  if (!('mergeFieldMarker' in options)) {
    options['mergeFieldMarker'] = new RegExp(
      DEFAULT_CONFIG.MERGE_FIELD_MARKER_TEXT,
      'g'
    );
  }
  if (!('enableGroupMerge' in options)) {
    options['enableGroupMerge'] = DEFAULT_CONFIG.ENABLE_GROUP_MERGE;
  }
  if (!('groupFieldMarker' in options)) {
    options['groupFieldMarker'] = new RegExp(
      DEFAULT_CONFIG.GROUP_FIELD_MARKER_TEXT,
      'g'
    );
  }
  if (!('rowIndexMarker' in options)) {
    options['rowIndexMarker'] = DEFAULT_CONFIG.ROW_INDEX_MARKER;
  }
  let messageData = {};
  for (let k in template) {
    if (options.excludeFromTemplate.includes(k)) {
      messageData[k] = template[k];
      continue;
    }
    let text = '';
    text = template[k];
    // Group merge
    if (options.enableGroupMerge) {
      // Create an array of group field marker(s) in the text
      let groupTexts = [...text.matchAll(options.groupFieldMarker)];
      // If the number of group field marker is not 0...
      if (groupTexts.length) {
        let groupTextsMerged = groupTexts.map((groupMergeField) => {
          // Create an array of merge field markers within a group merge marker
          let fieldVars = [
            ...groupMergeField[1].matchAll(options.mergeFieldMarker),
          ];
          if (!fieldVars.length) {
            return groupMergeField; // return the text of group merge field itself if no merge field marker is found within the group merge marker.
          } else {
            let fieldMerged = [];
            data.forEach((datum, i) => {
              let rowIndex = i + 1;
              let fieldRowIndexed = groupMergeField[1].replace(
                options.rowIndexMarker,
                rowIndex
              );
              fieldVars.forEach((variable, ind) => {
                let replaceValue =
                  datum[fieldVars[ind][1]] || options.replaceValue;
                replaceValue = options.asHtml.includes(k)
                  ? replaceValue.replace(/\r\n/g, '<br>')
                  : replaceValue;
                fieldRowIndexed = fieldRowIndexed.replace(
                  variable[0],
                  replaceValue
                );
              });
              fieldMerged.push(fieldRowIndexed);
            });
            let fieldMergedText = fieldMerged.join('');
            return fieldMergedText;
          }
        });
        groupTexts.forEach(
          (groupField, index) =>
            (text = text.replace(
              groupField[0],
              groupTextsMerged[index] || options.replaceValue
            ))
        );
      }
    }
    // merging the rest of the field based on the first row of data object array.
    let mergeData = data[0];
    // Create an array of all merge fields to be replaced
    let textVars = [...text.matchAll(options.mergeFieldMarker)];
    if (!textVars.length) {
      messageData[k] = text; // return text itself if no marker is found
    } else {
      // Replace variables in textVars with the actual values from the data object.
      // If no value is available, replace with replaceValue.
      textVars.forEach((variable, ind) => {
        let replaceValue = mergeData[textVars[ind][1]] || options.replaceValue;
        replaceValue = options.asHtml.includes(k)
          ? replaceValue.replace(/\r\n/g, '<br>')
          : replaceValue;
        text = text.replace(variable[0], replaceValue);
      });
      messageData[k] = text;
    }
  }
  return messageData;
}

/**
 * Get an array of all error messages
 * @param {String} userLocale
 * @returns {Array}
 */
function getKnownErrorMessages_(userLocale) {
  const localizedMessage = new LocalizedMessage(userLocale);
  let knownErrorMessages = [];
  for (let k in localizedMessage.messageList) {
    if (!k.startsWith('error')) {
      continue;
    }
    knownErrorMessages.push(localizedMessage.messageList[k]);
  }
  return knownErrorMessages;
}

if (typeof module === 'object') {
  module.exports = {
    MESSAGES,
    DEFAULT_CONFIG,
    buildHomepage,
    buildHomepageRestoreDefault,
    buildHomepageRestoreUserConfig,
    createMailMergeCard,
    createMessageCard,
    saveUserConfig,
    createDraftEmails,
    sendCreatedDrafts,
    sendDrafts,
    postProcessSendDrafts,
    sendEmails,
    mailMerge,
    postProcessMailMerge,
    parseConfig_,
    fillInTemplate_,
    getKnownErrorMessages_,
  };
}
