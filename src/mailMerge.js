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
/* exported buildHomepage, buildHomepageRestoreUserConfig, buildHomepageRestoreDefault, saveUserConfig, createDraftEmails, sendDrafts, sendEmails */

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
  REPLY_TO: 'replyTo@email.com'
};
// Key(s) for storing and calling settings stored as user property
const UP_KEY_CREATED_DRAFT_IDS = 'createdDraftIds';
const UP_KEY_PREV_CONFIG = 'prevConfig';
const UP_KEY_USER_CONFIG = 'userConfig';

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
      .setText(localizedMessage.messageList.cardMessage))
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
      .setValue(config.ROW_INDEX_MARKER || DEFAULT_CONFIG.ROW_INDEX_MARKER)));
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
          .setUrl('https://www.scriptable-assets.page/add-ons/group-merge/')))
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
  for (let k in config) {
    cardMessage += `<b>${k}: ${config[k]}\n`;
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
  var localizedMessage = new LocalizedMessage(config.userLocale);
  var userProperties = PropertiesService.getUserProperties();
  var cardMessage = '';
  try {
    let messageCount = 0;
    if (config.hostApp == 'SHEETS') {
      // Confirmation before sending email
      var ui = SpreadsheetApp.getUi();
      let confirmSend = localizedMessage.replaceConfirmSendingOfDraft(Session.getActiveUser().getEmail());
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
    } else if (error.message.startsWith('Unexpected error while getting the method or property openByUrl') || error.message.startsWith('You do not have permission to access the requested document.')) {
      cardMessage = localizedMessage.messageList.errorSpreadsheetNotFound;
    } else {
      cardMessage = 'Unexpected Error:\n' + error.stack;
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
 * Bulk send personalized emails based on a designated Gmail draft.
 * @param {boolean} draftMode Creates Gmail draft(s) instead of sending email. Defaults to true.
 * @param {Object} config Object returned by parseConfig_(eventObj)
 * @returns {string} 
 */
function mailMerge(draftMode = true, config = DEFAULT_CONFIG) {
  var localizedMessage = new LocalizedMessage(config.userLocale);
  var cardMessage = '';
  var messageCount = 0;
  // Reset list of created drafts
  var createdDraftIds = [];
  var userProperties = PropertiesService.getUserProperties().setProperty(UP_KEY_CREATED_DRAFT_IDS, JSON.stringify(createdDraftIds));
  // Save current settings in user property
  userProperties.setProperty(UP_KEY_PREV_CONFIG, JSON.stringify(config));
  // Designate name of fields without placeholders, i.e. values that can be skipped for the merge process later on
  var noPlaceholder = ['from', 'attachments', 'inLineImages', 'labels'];
  if (config.hostApp == 'SHEETS') {
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
    if (config.hostApp == 'SHEETS') {
      // Confirmation before sending email
      let confirmAccount = localizedMessage.replaceConfirmAccount(draftMode, Session.getActiveUser().getEmail());
      let answer = ui.alert(confirmAccount, ui.ButtonSet.OK_CANCEL);
      if (answer !== ui.Button.OK) {
        throw new Error(localizedMessage.messageList.errorMailMergeCanceled);
      }
    }
    // Verify value of TO
    var Tos = [...config.TO.matchAll(config.MERGE_FIELD_MARKER)];
    if (config.TO == null || config.TO == '') {
      throw new Error(localizedMessage.message.errorNoToEntered);
    } else if (Tos.length > 1) {
      throw new Error(localizedMessage.message.errorMultipleToEntered);
    }
    // Verify value of TEMPLATE_SUBJECT
    if (config.TEMPLATE_SUBJECT == null || config.TEMPLATE_SUBJECT == '') {
      throw new Error(localizedMessage.message.errorNoSubjectTextEntered);
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
    let template = {
      'subject': config.TEMPLATE_SUBJECT,
      'plainBody': draftMessages[0].getPlainBody(),
      'htmlBody': draftMessages[0].getBody(),
      'from': draftMessages[0].getFrom(),
      'to': `${draftMessages[0].getTo()},${config.TO}`,
      'cc': `${draftMessages[0].getCc()},${config.CC}`,
      'bcc': `${draftMessages[0].getBcc()},${config.BCC}`,
      'attachments': draftMessages[0].getAttachments({ 'includeInlineImages': false, 'includeAttachments': true }),
      'inLineImages': draftMessages[0].getAttachments({ 'includeInlineImages': true, 'includeAttachments': false }),
      'labels': draftMessages[0].getThread().getLabels(),
      'replyTo': (config.ENABLE_REPLY_TO ? config.REPLY_TO : '')
    };
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
    }
    // Create draft or send email based on the template.
    // The process depends on the value of ENABLE_GROUP_MERGE
    if (config.ENABLE_GROUP_MERGE) {
      // Convert the 2d-array merge data into object grouped by recipient(s)
      let groupedMergeData = groupArray_(mergeDataEolReplaced, Tos[0][1]);
      // Validity check
      if (Object.keys(groupedMergeData).length == 0) {
        throw new Error(localizedMessage.messageList.errorInvalidTo);
      }
      // Create draft for each recipient and, depending on the value of draftMode, send it.
      for (let k in groupedMergeData) {
        let mergeDataObjArr = groupedMergeData[k];
        let fillInTemplate_options = {
          'excludeFromTemplate': noPlaceholder,
          'asHtml': ['htmlBody'],
          'replaceValue': config.REPLACE_VALUE,
          'mergeFieldMarker': config.MERGE_FIELD_MARKER,
          'enableGroupMerge': config.ENABLE_GROUP_MERGE,
          'groupFieldMarker': config.GROUP_FIELD_MARKER,
          'rowIndexMarker': config.ROW_INDEX_MARKER
        };
        let messageData = fillInTemplate_(template, mergeDataObjArr, fillInTemplate_options);
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
        messageCount += 1;
      }
    } else {
      // Convert the 2d-array merge data into object
      let groupedMergeData = groupArray_(mergeDataEolReplaced);
      // Create draft or send email for each recipient
      groupedMergeData.data.forEach(obj => {
        let mergeDataObjArr = [obj];
        let fillInTemplate_options = {
          'excludeFromTemplate': noPlaceholder,
          'asHtml': ['htmlBody'],
          'replaceValue': config.REPLACE_VALUE,
          'mergeFieldMarker': config.MERGE_FIELD_MARKER,
          'enableGroupMerge': config.ENABLE_GROUP_MERGE
        };
        let messageData = fillInTemplate_(template, mergeDataObjArr, fillInTemplate_options);
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
        let draftThread = draft.getMessage().getThread();
        messageData.labels.forEach(label => draftThread.addLabel(label));
        if (!draftMode) {
          draft.send();
        } else {
          // List the created draft ID
          createdDraftIds.push(draft.getId());
        }
        messageCount += 1;
      });
    }
    // Notification
    cardMessage = localizedMessage.replaceCompleteMessage(draftMode, messageCount);
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
    } else if (error.message.startsWith('Unexpected error while getting the method or property openByUrl') || error.message.startsWith('You do not have permission to access the requested document.')) {
      cardMessage = localizedMessage.messageList.errorSpreadsheetNotFound;
    } else {
      cardMessage = 'Unexpected Error:\n' + error.stack;
      console.error(cardMessage);
    }
  }
  userProperties.setProperty(UP_KEY_CREATED_DRAFT_IDS, JSON.stringify(createdDraftIds));
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
    'ROW_INDEX_MARKER'
  ];
  let configObj = targetSettings.reduce((config, item) => {
    let input = eventObj.commonEventObject.formInputs[item] || { 'stringInputs': { 'value': [''] } };
    let value = input.stringInputs.value[0];
    if (item == 'ENABLE_GROUP_MERGE' || item == 'ENABLE_REPLY_TO') {
      value = (value == 'enabled' || value == 'true');
    } else if (item == 'MERGE_FIELD_MARKER_TEXT' || item == 'GROUP_FIELD_MARKER_TEXT') {
      item = item.replace('_TEXT', '');
      value = new RegExp(value, 'g');
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
 * Create a Javascript object from a 2d array, grouped by a given property.
 * @param {array} data 2-dimensional array with a header as its first row.
 * @param {string} property [Optional] Name of field name in header to group by.
 * When property is not specified, this function will return an object with a key 'data', whose value is a simple array of objects converted from the given 2d array.
 * If the designated property is not included in the header, this function will return an empty object.
 * @return {object}
 */
function groupArray_(data, property = null) {
  let header = data.shift();
  let index = header.indexOf(property);
  if (property == null) {
    let groupedObj = {};
    groupedObj['data'] = data.map(values => header.reduce((o, k, i) => {
      o[k] = values[i];
      return o;
    }, {}));
    return groupedObj;
  } else if (index < 0) {
    return {};
  } else {
    let groupedObj = data.reduce((accObj, curArr) => {
      let key = curArr[index];
      if (!accObj[key]) {
        accObj[key] = [];
      }
      let rowObj = header.reduce((o, k, i) => {
        o[k] = curArr[i];
        return o;
      }, {});
      accObj[key].push(rowObj);
      return accObj;
    }, {});
    return groupedObj;
  }
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
      let groupTexts = [...text.matchAll(options.groupFieldMarker)].map(element => element[1]);
      // If the number of group field marker is not 0...
      if (groupTexts !== null) {
        let groupTextsMerged = groupTexts.map((field) => {
          // Get the text inside group field markers, e.g., [[group field]] => group field
          field = field.substring(2, field.length - 2); // assuming that the text length for opening and closing markers are 2 and 2, respectively
          // Create an array of merge field markers within a group merge marker
          let fieldVars = field.match(options.mergeFieldMarker);
          if (!fieldVars) {
            return field; // return the text of group merge field itself if no merge field marker is found within the group merge marker.
          } else {
            let fieldMerged = [];
            data.forEach((datum, i) => {
              let rowIndex = i + 1;
              let fieldRowIndexed = field.replace(options.rowIndexMarker, rowIndex);
              let fieldVarsCopy = fieldVars.slice();
              // Get the text inside markers, e.g., {{field name}} => field name
              let fieldMarkerText = fieldVarsCopy.map(value => value.substring(2, value.length - 2)); // assuming that the text length for opening and closing markers are 2 and 2, respectively 
              fieldVarsCopy.forEach((variable, ind) => {
                let replaceValue = datum[fieldMarkerText[ind]] || options.replaceValue;
                replaceValue = (options.asHtml.includes(k) ? replaceValue.replace(/\r\n/g, '<br>') : replaceValue);
                fieldRowIndexed = fieldRowIndexed.replace(variable, replaceValue);
              });
              fieldMerged.push(fieldRowIndexed);
            });
            let fieldMergedText = fieldMerged.join('');
            return fieldMergedText;
          }
        });
        groupTexts.forEach((groupField, index) => text = text.replace(groupField, groupTextsMerged[index] || options.replaceValue));
      }
    }
    // merging the rest of the field based on the first row of data object array.
    let mergeData = data[0];
    // Create an array of all merge fields to be replaced
    let textVars = text.match(options.mergeFieldMarker);
    if (!textVars) {
      messageData[k] = text; // return text itself if no marker is found
    } else {
      // Get the text inside markers, e.g., {{field name}} => field name 
      let markerText = textVars.map(value => value.substring(2, value.length - 2)); // assuming that the text length for opening and closing markers are 2 and 2, respectively
      // Replace variables in textVars with the actual values from the data object.
      // If no value is available, replace with replaceValue.
      textVars.forEach((variable, ind) => {
        let replaceValue = mergeData[markerText[ind]] || options.replaceValue;
        replaceValue = (options.asHtml.includes(k) ? replaceValue.replace(/\r\n/g, '<br>') : replaceValue);
        text = text.replace(variable, replaceValue);
      });
      messageData[k] = text;
    }
  }
  return messageData;
}
