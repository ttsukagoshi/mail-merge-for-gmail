// Copyright 2020 Taro TSUKAGOSHI
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

// Add spreadsheet menu
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Mail Merge')
    .addItem('Send Emails', 'sendEmails')
    .addSeparator()
    .addItem('Create Draft', 'createDraftEmails')
    .addToUi();
}

/**
 * Send personalized email(s) to recipient address listed in sheet 'SHEET_NAME_DATA'
 */
function sendEmails() {
  const draftMode = false;
  const config = getConfig_();
  sendPersonalizedEmails_(draftMode, config);
}

/**
 * Send test emails to myself the content of the first row in sheet 'List'
 */
function createDraftEmails() {
  const draftMode = true;
  const config = getConfig_();
  sendPersonalizedEmails_(draftMode, config);
}

/**
 * Bulk send personalized emails based on a designated Gmail draft.
 * The email(s) can be sent to multiple recipients, which will serve as an alternative for using BCC.
 * @param {boolean} draftMode Creates Gmail draft(s) instead of sending email. Defaults to true.
 * @param {Object} config Object returned by getConfig_()
 */
function sendPersonalizedEmails_(draftMode = true, config = CONFIG) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var myEmail = Session.getActiveUser().getEmail();

  // Get data of field(s) to merge
  //// Range of data
  const dataSheet = ss.getSheetByName(config.DATA_SHEET_NAME);
  var mergeDataRange = dataSheet.getDataRange().setNumberFormat('@'); // Convert all formatted dates and numbers into texts
  //// Get data
  var mergeData = mergeDataRange.getValues();
  //// Define first row of mergeData as header
  var header = mergeData.shift();



  ///////////////////////////////ここの代わりに
  
  //// Convert 2d array of mergeData into object array
  var mergeDataObjArr = mergeData.map(function (values) {
    return header.reduce(function (object, key, index) {
      object[key] = values[index];
      return object;
    }, {});
  })
  
  ///////////////////////////////

  try {
    // Confirmation before sending email
    let confirmAccount = (draftMode === true
      ? `Are you sure you want to create draft email(s) as ${myEmail}?`
      : `Are you sure you want to send email(s) as ${myEmail}?`);
    let answer = ui.alert(confirmAccount, ui.ButtonSet.OK_CANCEL);
    if (answer !== ui.Button.OK) {
      throw new Error('Canceled.');
    }

    // Email Template
    //// Prompt to enter the subject of the Gmail draft to use as template
    let promptMessage = 'Enter the subject text of Gmail draft to use as template.';
    let promptResult = ui.prompt(promptMessage, ui.ButtonSet.OK_CANCEL);
    let [selectedButton, subjectText] = [promptResult.getSelectedButton(), promptResult.getResponseText()];
    if (selectedButton !== ui.Button.OK) {
      throw new Error('Canceled.');
    } else if (subjectText == null) {
      throw new Error('No text entered.');
    }
    //// Get template
    let draftMessage = getDraftBySubject_(subjectText);
    if (draftMessage.length > 1) {
      throw new Error('There are 2 or more Gmail drafts with the subject you entered. Enter a unique subject text.');
    }
    let template = {
      'subject': subjectText,
      'plainBody': draftMessage[0].getPlainBody(),
      'htmlBody': draftMessage[0].getBody()
    };

    // Send or create draft of personalized email
    mergeDataObjArr.forEach(function(element){
      let messageData = fillInTemplateFromObject_(template, element, config.MERGE_FIELD_MARKER, config.REPLACE_VALUE);
      let options = {
        'htmlBody': messageData.htmlBody,
        'bcc': (config.BCC_TO_MYSELF === true ? myEmail : null)
      };
      draftMode === true
        ? GmailApp.createDraft(element[config.RECIPIENT_COL_NAME], messageData.subject, messageData.plainBody, options)
        : GmailApp.sendEmail(element[config.RECIPIENT_COL_NAME], messageData.subject, messageData.plainBody, options);
    });
    let completeMessage = (draftMode === true
      ? 'Complete: All draft(s) created.'
      : 'Complete: All mails sent.');
    ui.alert(completeMessage)
  } catch (e) {
    let message = errorMessage_(e);
    ui.alert(message);
  }
}

/**
 * Returns an object of configurations from spreadsheet.
 * The sheet should have a first row of headers, and its first column should include the following properties:
 * DATA_SHEET_NAME: {string} Name of sheet in which field(s) to merge in email are stored
 * RECIPIENT_COL_NAME: {string} Name of column in sheet 'DATA_SHEET_NAME' that designates the email address of the recipient
 * BCC_TO_MYSELF: {string} String boolean. When true, will send (or create draft) email with the sender's address set to BCC.
 * REPLACE_VALUE: {string} Text that will replace empty data of marker. 
 * MERGE_FIELD_MARKER: {string} Text to be processed in RegExp() constructor to define merge field marker. Note that the backslash itself does not need to be escaped, i.e., does not need to be repeated.
 * 
 * @param {string} configSheetName Name of sheet with configurations. Defaults to 'Config'.
 * @return {Object}
 */
function getConfig_(configSheetName = 'Config') {
  // Get values from spreadsheet
  let configValues = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(configSheetName).getDataRange().getValues();
  configValues.shift();

  // Convert the 2d array values into a Javascript object
  let configObj = {};
  configValues.forEach(element => configObj[element[0]] = element[1]);

  // Convert data types
  configObj.BCC_TO_MYSELF = toBoolean_(configObj.BCC_TO_MYSELF);
  configObj.MERGE_FIELD_MARKER = new RegExp(configObj.MERGE_FIELD_MARKER, 'g');

  return configObj;
}

/**
 * Convert string booleans into boolean data
 * @param {string} stringBoolean 
 * @return {boolean}
 */
function toBoolean_(stringBoolean) {
  return stringBoolean.toLowerCase === 'true';
}

/**
 * Get an array of Gmail message(s) with the designated subject
 * @param {string} subject Subject text of Gmail draft
 * @return {array} Array of GmailMessage class objects. https://developers.google.com/apps-script/reference/gmail/gmail-message
 */
function getDraftBySubject_(subject) {
  let draftMessages = GmailApp.getDraftMessages();
  let targetDrafts = draftMessages.filter(element => element.getSubject() == subject);
  return targetDrafts;
}

/**
 * Replaces markers in a template object with values defined in a JavaScript data object.
 * @param {Object} template Template object containing markers, as designated in regular expression in MERGE_FIELD_MARKER
 * @param {Object} mergeData Object with values to replace markers.
 * @param {string} mergeFieldMarker Regular expression for the merge field marker. Defaults to /\{\{[^\}]+\}\}/g e.g., {{field name}}
 * @param {string} replaceValue String to replace empty data of a marker. Defaults to 'NA' for Not Available. 
 * @return {Object} Returns object with markers replaced.
 */
function fillInTemplateFromObject_(template, mergeData, mergeFieldMarker = /\{\{[^\}]+\}\}/g, replaceValue = 'NA') {
  let messageData = {};
  for (let k in template) {
    let text = '';
    text = template[k];
    // Search for all the variables to be replaced
    let textVars = text.match(mergeFieldMarker);
    if (!textVars) {
      messageData[k] = text; // return text itself if no marker is found
    } else {
      // Get the text inside markers. e.g., {{field name}} => field name
      // assuming that the text length for opening and closing markers are 2 and 2, respectively 
      let markerText = textVars.map(value => value.substring(2, value.length - 2));
      // Replace variables in textVars with the actual values from the data object.
      // If no value is available, replace with the string with replaceValue.
      textVars.forEach(
        (variable, i) => text = text.replace(variable, mergeData[markerText[i]] || replaceValue)
      );
      messageData[k] = text;
    }
  }
  return messageData;
}

/**
 * Standarized error message
 * @param {Object} e Error object returned by try-catch
 * @return {string} Standarized error message
 */
function errorMessage_(e) {
  let message = `Error: line - ${e.lineNumber}\n[${e.name}] ${e.message}\n${e.stack}`
  return message;
}

// Default configurations
const CONFIG = {
  DATA_SHEET_NAME: 'List',
  RECIPIENT_COL_NAME: 'Email',
  BCC_TO_MYSELF: true,
  REPLACE_VALUE: 'NA',
  MERGE_FIELD_MARKER: /\{\{[^\}]+\}\}/g,
  ENABLE_NESTED_MERGE: false,
  NESTED_FIELD_MARKER: /\[\[[^\]]+\]\]/g
}
