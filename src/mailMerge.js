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

/* global LocalizedMessage */
/* exported buildHomepage, buildHomepageRestoreDefault, buildHomepageRestoreUserConfig, createDraftEmails, postProcessMailMerge, saveUserConfig, sendDrafts, sendEmails */

//////////////////////
// Global Variables //
//////////////////////

// Default configurations
const DEFAULT_CONFIG = {
  SPREADSHEET_URL: 'https://docs.google.com/spreadsheets/d/*****/edit',
  DATA_SHEET_NAME: 'sheetName',
  TO: '{{Email}}',
  CC: '',
  BCC: '',
  TEMPLATE_SUBJECT: 'Enter subject here...',
  REPLACE_VALUE: 'NA',
  MERGE_FIELD_MARKER_TEXT: '\\{\\{([^\\}]+)\\}\\}',
  MERGE_FIELD_MARKER: /\{\{([^\}]+)\}\}/g, // deprecated
  ENABLE_GROUP_MERGE: true,
  GROUP_FIELD_MARKER_TEXT: '\\[\\[([^\\]]+)\\]\\]',
  GROUP_FIELD_MARKER: /\[\[([^\]]+)\]\]/g, // deprecated
  ROW_INDEX_MARKER: '{{i}}',
  ENABLE_REPLY_TO: false,
  REPLY_TO: 'replyTo@email.com',
  ENABLE_DEBUG_MODE: false
};
// Key(s) for storing and calling settings stored as user property
const UP_KEY_CREATED_DRAFT_IDS = 'createdDraftIds';
const UP_KEY_PREV_CONFIG = 'prevConfig';
const UP_KEY_USER_CONFIG = 'userConfig';
// Key lengths in milliseconds regarding time limit of add-on card actions.
const ACTION_LIMIT_TIME = 30 * 1000; // Card actions have a limited execution time of maximum 30 seconds https://developers.google.com/workspace/add-ons/concepts/actions#callback_functions
const ACTION_LIMIT_TIME_OFFSET = 5 * 1000; // Milliseconds prior to the actual ACTION_LIMIT_TIME at which the script will break the current mail merge process for a scheduled trigger.
const EXECUTE_TRIGGER_AFTER = 5 * 1000; // Milliseconds after which the mail merge process will resume by a scheduled trigger.

//////////////////////////
// Add-on Card Builders //
//////////////////////////

/**
 * Function to create the homepage card for the add-on.
 * @param {Object} event Google Workspace Add-on Event object. https://developers.google.com/workspace/add-ons/concepts/event-objects
 */
function buildHomepage(event) {
  var up = PropertiesService.getUserProperties();
  var config = JSON.parse(up.getProperty(UP_KEY_USER_CONFIG)) || JSON.parse(up.getProperty(UP_KEY_PREV_CONFIG)) || DEFAULT_CONFIG;
  return createMailMergeCard(event.commonEventObject.userLocale, event.commonEventObject.hostApp, config);
}

/**
 * Function to reset (re-create) the homepage card with user config values.
 * @param {Object} event Google Workspace Add-on Event object. https://developers.google.com/workspace/add-ons/concepts/event-objects
 */
function buildHomepageRestoreUserConfig(event) {
  var config = JSON.parse(PropertiesService.getUserProperties().getProperty(UP_KEY_USER_CONFIG)) || DEFAULT_CONFIG;
  return createMailMergeCard(event.commonEventObject.userLocale, event.commonEventObject.hostApp, config);
}

/**
 * Function to reset (re-create) the homepage card with default config values.
 * @param {Object} event Google Workspace Add-on Event object. https://developers.google.com/workspace/add-ons/concepts/event-objects
 */
function buildHomepageRestoreDefault(event) {
  PropertiesService.getUserProperties().setProperty(UP_KEY_USER_CONFIG, '[{}]').setProperty(UP_KEY_PREV_CONFIG, '[{}]');
  return createMailMergeCard(event.commonEventObject.userLocale, event.commonEventObject.hostApp, DEFAULT_CONFIG);
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
  var localizedMessage = new LocalizedMessage(userLocale);
  // Load user properties
  var userProperties = PropertiesService.getUserProperties();
  var createdDraftIds = JSON.parse(userProperties.getProperty(UP_KEY_CREATED_DRAFT_IDS));
  var disableSendDrafts = (!createdDraftIds || createdDraftIds.length == 0);
  // Get URL of currently open spreadsheet if host is Google Sheets
  var ssUrl = null;
  if (hostApp == 'SHEETS') {
    ssUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  }
  // Build card
  var builder = CardService.newCardBuilder();
  // Message Section
  builder.addSection(CardService.newCardSection()
    .addWidget(CardService.newTextParagraph()
      .setText(localizedMessage.messageList.cardHomepageMessage))
    .addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
        .setText(localizedMessage.messageList.buttonSendDrafts)
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOnClickAction(CardService.newAction().setFunctionName('sendDrafts'))
        .setDisabled(disableSendDrafts))));
  // Recipient List Section
  builder.addSection(CardService.newCardSection()
    .setHeader(localizedMessage.messageList.cardRecipientListSettings)
    .addWidget(CardService.newTextInput()
      .setFieldName('SPREADSHEET_URL')
      .setTitle(localizedMessage.messageList.cardEnterSpreadsheetUrl)
      .setHint(localizedMessage.messageList.cardHintSpreadsheetUrl)
      .setValue(ssUrl || config.SPREADSHEET_URL || DEFAULT_CONFIG.SPREADSHEET_URL))
    .addWidget(CardService.newTextInput()
      .setFieldName('DATA_SHEET_NAME')
      .setTitle(localizedMessage.messageList.cardEnterSheetName)
      .setHint(localizedMessage.messageList.cardHintSheetName)
      .setValue(config.DATA_SHEET_NAME || DEFAULT_CONFIG.DATA_SHEET_NAME))
    .addWidget(CardService.newTextInput()
      .setFieldName('TO')
      .setTitle(localizedMessage.messageList.cardEnterTo)
      .setHint(localizedMessage.messageList.cardHintTo)
      .setValue(config.TO || DEFAULT_CONFIG.TO))
    .addWidget(CardService.newTextInput()
      .setFieldName('CC')
      .setTitle(localizedMessage.messageList.cardEnterCc)
      .setHint(localizedMessage.messageList.cardHintCc)
      .setValue(config.CC || DEFAULT_CONFIG.CC))
    .addWidget(CardService.newTextInput()
      .setFieldName('BCC')
      .setTitle(localizedMessage.messageList.cardEnterBcc)
      .setHint(localizedMessage.messageList.cardHintBcc)
      .setValue(config.BCC || DEFAULT_CONFIG.BCC)));
  // Template Draft Section
  builder.addSection(CardService.newCardSection()
    .setHeader(localizedMessage.messageList.cardTemplateDraftSettings)
    .addWidget(CardService.newTextInput()
      .setFieldName('TEMPLATE_SUBJECT')
      .setTitle(localizedMessage.messageList.cardEnterTemplateSubject)
      .setHint(localizedMessage.messageList.cardHintTemplateSubject)
      .setValue(config.TEMPLATE_SUBJECT || DEFAULT_CONFIG.TEMPLATE_SUBJECT))
    .addWidget(CardService.newDecoratedText()
      .setText(localizedMessage.messageList.cardSwitchEnableGroupMerge)
      .setSwitchControl(CardService.newSwitch()
        .setSelected(typeof config.ENABLE_GROUP_MERGE == 'boolean' ? config.ENABLE_GROUP_MERGE : DEFAULT_CONFIG.ENABLE_GROUP_MERGE)
        .setFieldName('ENABLE_GROUP_MERGE')
        .setValue('enabled'))));
  // Advanced Settings Section
  builder.addSection(CardService.newCardSection()
    .setHeader(localizedMessage.messageList.cardAdvancedSettings)
    .setCollapsible(true)
    .addWidget(CardService.newDecoratedText()
      .setText(localizedMessage.messageList.cardSwitchEnableReplyTo)
      .setSwitchControl(CardService.newSwitch()
        .setSelected(typeof config.ENABLE_REPLY_TO == 'boolean' ? config.ENABLE_REPLY_TO : DEFAULT_CONFIG.ENABLE_REPLY_TO)
        .setFieldName('ENABLE_REPLY_TO')
        .setValue('enabled')))
    .addWidget(CardService.newTextInput()
      .setFieldName('REPLY_TO')
      .setTitle(localizedMessage.messageList.cardEnterReplyTo)
      .setHint(localizedMessage.messageList.cardHintReplyTo)
      .setValue(config.REPLY_TO || DEFAULT_CONFIG.REPLY_TO))
    .addWidget(CardService.newTextInput()
      .setFieldName('REPLACE_VALUE')
      .setTitle(localizedMessage.messageList.cardEnterReplaceValue)
      .setHint(localizedMessage.messageList.cardHintReplaceValue)
      .setValue(config.REPLACE_VALUE || DEFAULT_CONFIG.REPLACE_VALUE))
    .addWidget(CardService.newTextInput()
      .setFieldName('MERGE_FIELD_MARKER_TEXT')
      .setTitle(localizedMessage.messageList.cardEnterMergeFieldMarker)
      .setHint(localizedMessage.messageList.cardHintMergeFieldMarker)
      .setValue(config.MERGE_FIELD_MARKER_TEXT || DEFAULT_CONFIG.MERGE_FIELD_MARKER_TEXT))
    .addWidget(CardService.newTextInput()
      .setFieldName('GROUP_FIELD_MARKER_TEXT')
      .setTitle(localizedMessage.messageList.cardEnterGroupFieldMarker)
      .setHint(localizedMessage.messageList.cardHintGroupFieldMarker)
      .setValue(config.GROUP_FIELD_MARKER_TEXT || DEFAULT_CONFIG.GROUP_FIELD_MARKER_TEXT))
    .addWidget(CardService.newTextInput()
      .setFieldName('ROW_INDEX_MARKER')
      .setTitle(localizedMessage.messageList.cardEnterRowIndexMarker)
      .setHint(localizedMessage.messageList.cardHintRowIndexMarker)
      .setValue(config.ROW_INDEX_MARKER || DEFAULT_CONFIG.ROW_INDEX_MARKER))
    .addWidget(CardService.newDecoratedText()
      .setText(localizedMessage.messageList.cardSwitchEnableDebugMode)
      .setSwitchControl(CardService.newSwitch()
        .setSelected(typeof config.ENABLE_DEBUG_MODE == 'boolean' ? config.ENABLE_DEBUG_MODE : DEFAULT_CONFIG.ENABLE_DEBUG_MODE)
        .setFieldName('ENABLE_DEBUG_MODE')
        .setValue('enabled'))));
  // Buttons Section
  builder.addSection(CardService.newCardSection()
    .addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
        .setText(localizedMessage.messageList.buttonRestoreUserConfig)
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOnClickAction(CardService.newAction().setFunctionName('buildHomepageRestoreUserConfig')))
      .addButton(CardService.newTextButton()
        .setText(localizedMessage.messageList.buttonSaveUserConfig)
        .setOnClickAction(CardService.newAction().setFunctionName('saveUserConfig')))
      .addButton(CardService.newTextButton()
        .setText(localizedMessage.messageList.buttonRestoreDefault)
        .setOnClickAction(CardService.newAction().setFunctionName('buildHomepageRestoreDefault')))
      .addButton(CardService.newTextButton()
        .setText(localizedMessage.messageList.buttonReadDocument)
        .setOpenLink(CardService.newOpenLink()
          .setUrl(localizedMessage.messageList.buttonReadDocumentUrl)))
      /*
      .addButton(CardService.newTextButton()
        .setText('test')
        .setOnClickAction(CardService.newAction().setFunctionName('test')))
      */
    ));
  // Fixed Footer
  builder.setFixedFooter(CardService.newFixedFooter()
    .setPrimaryButton(CardService.newTextButton()
      .setText(localizedMessage.messageList.buttonCreateDrafts)
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction().setFunctionName('createDraftEmails'))
      .setDisabled(false))
    .setSecondaryButton(CardService.newTextButton()
      .setText(localizedMessage.messageList.buttonSendEmails)
      .setOnClickAction(CardService.newAction().setFunctionName('sendEmails'))
      .setDisabled(false)));
  return builder.build();
}

/*
function test(event) {
  console.log(JSON.stringify(event));
}
*/

/**
 * Builder for message cards to present error and other messages to the add-on user.
 * @param {String} message Message string that can accept some basic HTML formatting described in https://developers.google.com/workspace/add-ons/concepts/widgets?hl=en#text_formatting
 */
function createMessageCard(message, userLocale) {
  var localizedMessage = new LocalizedMessage(userLocale);
  var builder = CardService.newCardBuilder()
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText(message)))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText(localizedMessage.messageList.buttonReturnHome)
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setOnClickAction(CardService.newAction().setFunctionName('buildHomepage')))));
  return builder.build();
}

//////////////////////////
// Mail Merge Actions //
//////////////////////////

function saveUserConfig(event) {
  var config = parseConfig_(event);
  // Save on user property
  PropertiesService.getUserProperties().setProperty(UP_KEY_USER_CONFIG, JSON.stringify(config));
  // Construct complete message
  var localizedMessage = new LocalizedMessage(config.userLocale);
  var cardMessage = localizedMessage.messageList.alertCompleteSavedUserConfig;
  if (config.ENABLE_DEBUG_MODE) {
    for (let k in config) {
      cardMessage += `<b>${k}: ${config[k]}\n`;
    }
  }
  return createMessageCard(cardMessage, config.userLocale);
}

/**
 * Create draft of personalized email(s)
 */
function createDraftEmails(event) {
  const draftMode = true;
  const config = parseConfig_(event);
  return createMessageCard(mailMerge(draftMode, config), event.commonEventObject.userLocale);
}

/**
 * Send the drafts created by createDraftEmails()
 */
function sendDrafts(event) {
  const config = parseConfig_(event);
  var myEmail = Session.getActiveUser().getEmail();
  var localizedMessage = new LocalizedMessage(config.userLocale);
  var userProperties = PropertiesService.getUserProperties();
  var cardMessage = '';
  try {
    let messageCount = 0;
    if (config.hostApp == 'SHEETS') {
      // Confirmation before sending email
      var ui = SpreadsheetApp.getUi();
      let confirmSend = localizedMessage.replaceConfirmSendingOfDraft(myEmail);
      let answer = ui.alert(confirmSend, ui.ButtonSet.OK_CANCEL);
      if (answer !== ui.Button.OK) {
        throw new Error(localizedMessage.messageList.errorSendDraftsCanceled);
      }
    }
    // Get the values of createdDraftIds, the string of draft IDs to send, stored in the user property
    let createdDraftIds = JSON.parse(userProperties.getProperty(UP_KEY_CREATED_DRAFT_IDS));
    if (!createdDraftIds || createdDraftIds.length == 0) {
      // Throw error if no draft ID is stored.
      throw new Error(localizedMessage.messageList.errorNoDraftToSend);
    }
    // Send emails
    createdDraftIds.forEach(draftId => {
      GmailApp.getDraft(draftId).send();
      messageCount += 1;
    });
    // Empty createdDraftIds
    createdDraftIds = [];
    userProperties.setProperty(UP_KEY_CREATED_DRAFT_IDS, JSON.stringify(createdDraftIds));
    cardMessage = localizedMessage.replaceCompleteMessage(false, messageCount);
  } catch (error) {
    let knownErrorMessages = [];
    for (let k in localizedMessage.messageList) {
      if (!(k.startsWith('error'))) {
        continue;
      }
      knownErrorMessages.push(localizedMessage.messageList[k]);
    }
    if (knownErrorMessages.includes(error.message)) {
      cardMessage = error.message;
    } else {
      cardMessage = localizedMessage.messageList.cardMessageUnexpectedError + error.stack;
    }
  }
  return createMessageCard(cardMessage, config.userLocale);
}

/**
 * Send personalized email(s)
 */
function sendEmails(event) {
  const draftMode = false;
  const config = parseConfig_(event);
  return createMessageCard(mailMerge(draftMode, config), event.commonEventObject.userLocale);
}

/**
 * Post-process for mailMerge(); continues the process of mail merge
 * for executions that are expected to exceed the Google Workspace Add-ons' 30-second time limit.
 */
function postProcessMailMerge() {
  var userPropertiesValues = PropertiesService.getUserProperties().getProperties();
  var draftMode = (userPropertiesValues['draftMode'] === 'true');
  var config = JSON.parse(userPropertiesValues[UP_KEY_PREV_CONFIG]);
  config.MERGE_FIELD_MARKER = new RegExp(config.MERGE_FIELD_MARKER_TEXT, 'g');
  config.GROUP_FIELD_MARKER = new RegExp(config.GROUP_FIELD_MARKER_TEXT, 'g');
  var prevProperties = {
    completedRecipients: JSON.parse(userPropertiesValues.completedRecipients),
    createdDraftIds: JSON.parse(userPropertiesValues.createdDraftIds),
    templateDraftId: userPropertiesValues.templateDraftId
  };
  mailMerge(draftMode, config, prevProperties);
}

/**
 * Bulk send personalized emails based on a designated Gmail draft.
 * @param {boolean} draftMode Creates Gmail draft(s) instead of sending email. Defaults to true.
 * @param {Object} config Object returned by parseConfig_(eventObj)
 * @returns {string} 
 */
function mailMerge(draftMode = true, config = DEFAULT_CONFIG, prevProperties = {}) {
  var isPostProcess = (Object.keys(prevProperties).length > 0);
  var debugInfo = {
    completedRecipients: [],
    config: config,
    draftMode: draftMode,
    prevCompletedRecipients: (isPostProcess ? prevProperties.completedRecipients : []),
    processTime: [],
    start: (new Date()).getTime() //, templateDraftId: ''
  };
  var myEmail = Session.getActiveUser().getEmail();
  var localizedMessage = new LocalizedMessage(config.userLocale);
  var cardMessage = '';
  // Reset list of created drafts
  var createdDraftIds = (isPostProcess ? prevProperties.createdDraftIds : []);
  var userProperties = PropertiesService.getUserProperties().setProperty(UP_KEY_CREATED_DRAFT_IDS, JSON.stringify(createdDraftIds));
  // Save current settings in user property
  userProperties.setProperty(UP_KEY_PREV_CONFIG, JSON.stringify(config));
  // Designate name of fields without placeholders, i.e. values that can be skipped for the merge process later on
  var noPlaceholder = ['from', 'attachments', 'inLineImages', 'labels'];
  if (config.hostApp == 'SHEETS' && !isPostProcess) {
    var ui = SpreadsheetApp.getUi();
  }
  try {
    var ss = SpreadsheetApp.openByUrl(config.SPREADSHEET_URL);
    // Get data of field(s) to merge in form of 2d array
    let dataSheet = ss.getSheetByName(config.DATA_SHEET_NAME);
    if (dataSheet == null) {
      throw new Error(localizedMessage.messageList.errorSheetNotFound);
    }
    let mergeData = dataSheet.getDataRange().getDisplayValues();
    // Convert line breaks in the spreadsheet (in LF format, i.e., '\n')
    // to CRLF format ('\r\n') for merging into Gmail plain text
    let mergeDataEolReplaced = mergeData.map(element => element.map(value => value.replace(/\n|\r|\r\n/g, '\r\n')));
    debugInfo.processTime.push(`Retrieved recipient list data at ${(new Date()).getTime() - debugInfo.start} (millisec from start)`);
    if (config.hostApp == 'SHEETS' && !isPostProcess) {
      // Confirmation before sending email
      let confirmAccount = localizedMessage.replaceConfirmAccount(draftMode, myEmail);
      let answer = ui.alert(confirmAccount, ui.ButtonSet.OK_CANCEL);
      if (answer !== ui.Button.OK) {
        throw new Error(localizedMessage.messageList.errorMailMergeCanceled);
      }
      debugInfo.processTime.push(`Passed UI confirmation on SHEETS at ${(new Date()).getTime() - debugInfo.start} (millisec from start)`);
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
    // Get an array of GmailMessage class objects whose subject matches config.TEMPLATE_SUBJECT.
    let draftMessages = GmailApp.getDraftMessages().filter(element => element.getSubject() == config.TEMPLATE_SUBJECT);
    // Check for duplicates
    if (draftMessages.length > 1) {
      throw new Error(localizedMessage.messageList.errorTwoOrMoreDraftsWithSameSubject);
    }
    // Throw error if no draft template is found
    if (draftMessages.length == 0) {
      throw new Error(localizedMessage.messageList.errorNoMatchingTemplateDraft);
    }
    // Store template into an object
    let messageTo = draftMessages[0].getTo();
    let messageCc = draftMessages[0].getCc();
    let messageBcc = draftMessages[0].getBcc();
    let template = {
      'subject': config.TEMPLATE_SUBJECT,
      'plainBody': draftMessages[0].getPlainBody(),
      'htmlBody': draftMessages[0].getBody(),
      'from': draftMessages[0].getFrom(),
      'to': config.TO + (messageTo ? `,${messageTo}` : ''),
      'cc': config.CC + (messageCc ? `,${messageCc}` : ''),
      'bcc': config.BCC + (messageBcc ? `,${messageBcc}` : ''),
      'attachments': draftMessages[0].getAttachments({ 'includeInlineImages': false, 'includeAttachments': true }),
      'inLineImages': draftMessages[0].getAttachments({ 'includeInlineImages': true, 'includeAttachments': false }),
      'labels': draftMessages[0].getThread().getLabels(),
      'replyTo': (config.ENABLE_REPLY_TO ? config.REPLY_TO : '')
    };
    debugInfo.processTime.push(`Retrieved template draft data at ${(new Date()).getTime() - debugInfo.start} (millisec from start)`);
    // Check template format; plain or HTML text.
    let isPlainText = (template.plainBody === template.htmlBody);
    let inLineImageBlobs = {};
    if (!isPlainText) {
      // If the template is composed in HTML, check for in-line images
      // and create a mapping object of cid (content ID) and its corresponding in-line image blob.
      let regExpImgTag = /\<img data-surl\=\"cid\:(?<cidDataSurl>[^\"]+)\"[^\>]+src\=\"cid\:(?<cidSrc>[^\"]+)\"[^\>]+alt\=\"(?<blobName>[^\"]+)\"[^\>]+\>/g;
      let inLineImageTags = [...template.htmlBody.matchAll(regExpImgTag)];
      inLineImageBlobs = template.inLineImages.reduce((obj, blob) => {
        let cid = inLineImageTags.find(element => element.groups.blobName == blob.getName()).groups.cidSrc;
        obj[cid] = blob;
        return obj;
      }, {});
      debugInfo.processTime.push(`Template is composed in HTML. Retrieved ${template.inLineImages.length} in-line image data at ${(new Date()).getTime() - debugInfo.start} (millisec from start)`);
    }
    // Create draft or send email based on the template.
    // The process depends on the value of ENABLE_GROUP_MERGE
    let fillInTemplate_options = {
      'excludeFromTemplate': noPlaceholder,
      'asHtml': [(isPlainText ? '' : 'htmlBody')],
      'replaceValue': config.REPLACE_VALUE,
      'mergeFieldMarker': config.MERGE_FIELD_MARKER,
      'enableGroupMerge': config.ENABLE_GROUP_MERGE,
      'groupFieldMarker': config.GROUP_FIELD_MARKER,
      'rowIndexMarker': config.ROW_INDEX_MARKER
    };
    // Convert the 2d-array merge data into object, grouped by recipient(s) if group merge is enabled
    let mergeDataHeader = mergeDataEolReplaced.shift();
    let mergeDataKeyIndex = mergeDataHeader.indexOf(Tos[0][1]);
    if (mergeDataKeyIndex < 0) {
      // Throw error if there were no column header matching the placeholder value of To
      throw new Error(localizedMessage.messageList.errorInvalidTo);
    }
    let groupedMergeData = mergeDataEolReplaced.reduce((obj, dataRow, rowIndex) => {
      let key = dataRow[mergeDataKeyIndex];
      key += (config.ENABLE_GROUP_MERGE ? '' : `_${rowIndex}`);
      if (!obj[key]) {
        obj[key] = [];
      }
      obj[key].push(mergeDataHeader.reduce((o, k, i) => {
        o[k] = dataRow[i];
        return o;
      }, {}));
      return obj;
    }, {});
    debugInfo.processTime.push(`Group Merge is ${(config.ENABLE_GROUP_MERGE ? 'enabled. Grouped recipient list by recipient email address' : 'disabled. Data stored in groupedMergeData')} at ${(new Date()).getTime() - debugInfo.start} (millisec from start)`);
    // Create draft for each recipient and, depending on the value of draftMode, send it.
    for (let k in groupedMergeData) {
      if (debugInfo.prevCompletedRecipients.includes(k)) {
        continue;
      }
      let messageData = fillInTemplate_(template, groupedMergeData[k], fillInTemplate_options);
      let options = {
        'htmlBody': (isPlainText ? null : messageData.htmlBody),
        'from': messageData.from,
        'cc': messageData.cc,
        'bcc': messageData.bcc,
        'attachments': (messageData.attachments ? messageData.attachments : null),
        'inlineImages': (isPlainText ? null : inLineImageBlobs),
        'replyTo': (config.ENABLE_REPLY_TO ? messageData.replyTo : null)
      };
      let draft = GmailApp.createDraft(messageData.to, messageData.subject, messageData.plainBody, options);
      // Add the same Gmail labels as those on the template draft message.
      let draftThread = draft.getMessage().getThread();
      messageData.labels.forEach(label => draftThread.addLabel(label));
      if (!draftMode) {
        draft.send();
      } else {
        // List the created draft ID
        createdDraftIds.push(draft.getId());
      }
      // Determine current process time lapse.
      // If the remaining time limit is less than ACTION_LIMIT_TIME_OFFSET, break the process to complete by execution on time-based triggers.
      let timeLapsed = (new Date()).getTime() - debugInfo.start;
      debugInfo.processTime.push(`Message for ${k} drafted/sent at ${timeLapsed} (millisec from start)`);
      debugInfo.completedRecipients.push(k);
      if (ACTION_LIMIT_TIME - timeLapsed < ACTION_LIMIT_TIME_OFFSET && !isPostProcess) {
        debugInfo.completedRecipients = debugInfo.prevCompletedRecipients.concat(debugInfo.completedRecipients);
        userProperties.setProperties({
          completedRecipients: JSON.stringify(debugInfo.completedRecipients),
          createdDraftIds: JSON.stringify(createdDraftIds),
          draftMode: draftMode
        }, false);
        ScriptApp.newTrigger('postProcessMailMerge')
          .timeBased()
          .after(EXECUTE_TRIGGER_AFTER)
          .create();
        debugInfo.processTime.push(`Saved property and set trigger for post-process at ${(new Date()).getTime() - debugInfo.start} (millisec from start)`);
        throw new Error(localizedMessage.replaceProceedingToPostProcess(ACTION_LIMIT_TIME / 1000, myEmail, draftMode, debugInfo.completedRecipients.length));
      }
    }
    // Notification
    debugInfo.completedRecipients = debugInfo.prevCompletedRecipients.concat(debugInfo.completedRecipients);
    cardMessage = localizedMessage.replaceCompleteMessage(draftMode, debugInfo.completedRecipients.length);
  } catch (error) {
    let knownErrorMessages = [];
    for (let k in localizedMessage.messageList) {
      if (!(k.startsWith('error'))) {
        continue;
      }
      knownErrorMessages.push(localizedMessage.messageList[k]);
    }
    if (knownErrorMessages.includes(error.message) || error.message.startsWith(localizedMessage.messageList.proceedingToPostProcess.slice(0, 10))) {
      cardMessage = error.message;
    } else if (error.message.startsWith(localizedMessage.messageList.appsScriptMessageErrorOnOpenByUrl) || error.message.startsWith(localizedMessage.messageList.appsScriptMessageNoPermissionErrorStartsWith)) {
      cardMessage = localizedMessage.messageList.errorSpreadsheetNotFound;
    } else {
      cardMessage = localizedMessage.messageList.cardMessageUnexpectedError + error.stack;
    }
  }
  userProperties.setProperty(UP_KEY_CREATED_DRAFT_IDS, JSON.stringify(createdDraftIds));
  if (config.ENABLE_DEBUG_MODE) {
    let debugInfoText = localizedMessage.messageList.cardMessageDebugInfo;
    for (let k in debugInfo) {
      debugInfoText += `${k}: ${JSON.stringify(debugInfo[k])}\n`;
    }
    MailApp.sendEmail(myEmail, `[GROUP MERGE] Debug Info`, `${cardMessage}\n\n${debugInfoText}`);
    cardMessage += `\n\n${localizedMessage.messageList.cardMessageSentDebugInfo}`;
  }
  if (isPostProcess) {
    MailApp.sendEmail(myEmail, localizedMessage.messageList.subjectPostProcessUpdate, cardMessage);
  }
  return cardMessage;
}

/**
 * Retrieve configuration values for mail merge from the input values of the Card widget interactions.
 * @param {Object} eventObj Add-on event object. https://developers.google.com/workspace/add-ons/concepts/event-objects
 * @returns {Object}
 */
function parseConfig_(eventObj) {
  var targetSettings = [ // A complete list of the form input items in createMailMergeCard
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
    'ENABLE_DEBUG_MODE'
  ];
  let configObj = targetSettings.reduce((config, item) => {
    let input = eventObj.commonEventObject.formInputs[item] || { 'stringInputs': { 'value': [''] } };
    let value = input.stringInputs.value[0];
    if (item == 'ENABLE_GROUP_MERGE' || item == 'ENABLE_REPLY_TO' || item == 'ENABLE_DEBUG_MODE') {
      value = (value == 'enabled' || value == 'true');
    } else if (item == 'MERGE_FIELD_MARKER_TEXT' || item == 'GROUP_FIELD_MARKER_TEXT') {
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
 * @param {Object} options [Optional] Options, as described below.
 * @param {array} options.excludeFromTemplate [Optional] Array of the keys in template to exclude from the replacement process, in strings. 
 * @param {array} option.asHtml [Optional] Array of the keys in template to regard as HTML formatted.
 * @param {string} options.replaceValue [Optional] String to replace empty data of a marker. Defaults to 'NA' for Not Available. 
 * @param {RegExp} options.mergeFieldMarker [Optional] Regular expression for the merge field marker. Defaults to /\{\{[^\}]+\}\}/g e.g., {{field name}}
 * @param {boolean} options.enableGroupMerge [Optional] Merged texts are returned in concatenated form when true. Defaults to false.
 * @param {RegExp} options.groupFieldMarker [Optional] Regular expression for the group merge field marker. Defaults to /\[\[[^\]]+\]\]/g e.g., [[group merge]]
 * @param {string} options.rowIndexMarker [Optional] Marker for merging row index number in a group merge. Defaults to '{{i}}'
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
    options['mergeFieldMarker'] = new RegExp(DEFAULT_CONFIG.MERGE_FIELD_MARKER_TEXT, 'g');
  }
  if (!('enableGroupMerge' in options)) {
    options['enableGroupMerge'] = DEFAULT_CONFIG.ENABLE_GROUP_MERGE;
  }
  if (!('groupFieldMarker' in options)) {
    options['groupFieldMarker'] = new RegExp(DEFAULT_CONFIG.GROUP_FIELD_MARKER_TEXT, 'g');
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
        let groupTextsMerged = groupTexts.map(groupMergeField => {
          // Create an array of merge field markers within a group merge marker
          let fieldVars = [...groupMergeField[1].matchAll(options.mergeFieldMarker)];
          if (!fieldVars.length) {
            return groupMergeField; // return the text of group merge field itself if no merge field marker is found within the group merge marker.
          } else {
            let fieldMerged = [];
            data.forEach((datum, i) => {
              let rowIndex = i + 1;
              let fieldRowIndexed = groupMergeField[1].replace(options.rowIndexMarker, rowIndex);
              fieldVars.forEach((variable, ind) => {
                let replaceValue = datum[fieldVars[ind][1]] || options.replaceValue;
                replaceValue = (options.asHtml.includes(k) ? replaceValue.replace(/\r\n/g, '<br>') : replaceValue);
                fieldRowIndexed = fieldRowIndexed.replace(variable[0], replaceValue);
              });
              fieldMerged.push(fieldRowIndexed);
            });
            let fieldMergedText = fieldMerged.join('');
            return fieldMergedText;
          }
        });
        groupTexts.forEach((groupField, index) => text = text.replace(groupField[0], groupTextsMerged[index] || options.replaceValue));
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
        replaceValue = (options.asHtml.includes(k) ? replaceValue.replace(/\r\n/g, '<br>') : replaceValue);
        text = text.replace(variable[0], replaceValue);
      });
      messageData[k] = text;
    }
  }
  return messageData;
}
