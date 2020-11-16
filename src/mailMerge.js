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
//
// See https://github.com/ttsukagoshi/mail-merge-for-gmail for latest information.

// Default configurations
const CONFIG = {
  DATA_SHEET_NAME: 'List',
  RECIPIENT_COL_NAME: 'Email',
  BCC_TO_MYSELF: true,
  REPLACE_VALUE: 'NA',
  MERGE_FIELD_MARKER: /\{\{[^\}]+\}\}/g,
  ENABLE_GROUP_MERGE: false,
  GROUP_FIELD_MARKER: /\[\[[^\]]+\]\]/g,
  ROW_INDEX_MARKER: '{{i}}'
}

// Add spreadsheet menu
function onOpen() {
  var locale = Session.getActiveUserLocale();
  var localizedMessage = new LocalizedMessage(locale);
  SpreadsheetApp.getUi()
    .createMenu(localizedMessage.messageList.menuName)
    .addItem(localizedMessage.messageList.menuCreateDraft, 'createDraftEmails')
    .addSeparator()
    .addItem(localizedMessage.messageList.menuSendEmails, 'sendEmails')
    .addToUi();
}

/**
 * Create draft of personalized email(s)
 */
function createDraftEmails() {
  console.log('Initiating Mail Merge for Gmail on Draft Mode...'); // log
  var draftMode = true;
  console.log('Loading config...'); // log
  var config = getConfig_('Config');
  console.log(`config: ${JSON.stringify(config)}`); // log
  sendPersonalizedEmails(draftMode, config);
}

/**
 * Send personalized email(s)
 */
function sendEmails() {
  console.log('Initiating Mail Merge for Gmail on Send Mode...'); // log
  var draftMode = false;
  console.log('Loading config...'); // log
  var config = getConfig_('Config');
  console.log(`config: ${JSON.stringify(config)}`); // log
  sendPersonalizedEmails(draftMode, config);
}

/**
 * Bulk send personalized emails based on a designated Gmail draft.
 * The email(s) can be sent to multiple recipients, which will serve as an alternative for using BCC.
 * @param {boolean} draftMode Creates Gmail draft(s) instead of sending email. Defaults to true.
 * @param {Object} config Object returned by getConfig_()
 */
function sendPersonalizedEmails(draftMode = true, config = CONFIG) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var myEmail = Session.getActiveUser().getEmail();
  var locale = Session.getActiveUserLocale();
  var localizedMessage = new LocalizedMessage(locale);
  console.log(`Loaded spreadsheet. Language set to ${locale}`); // log
  try {
    // Get data of field(s) to merge in form of 2d array
    let dataSheet = ss.getSheetByName(config.DATA_SHEET_NAME);
    let mergeDataRange = dataSheet.getDataRange().setNumberFormat('@'); // Convert all formatted dates and numbers into texts
    let mergeData = mergeDataRange.getValues();
    // Convert line breaks in the spreadsheet (in LF format, i.e., '\n')
    // to CRLF format ('\r\n') for merging into Gmail plain text
    let mergeDataEolReplaced = mergeData.map(element => element.map(value => value.replace(/\n|\r|\r\n/g, '\r\n')));
    console.log('Loaded merge data.'); // log
    // Confirmation before sending email
    let confirmAccount = localizedMessage.replaceConfirmAccount(draftMode, myEmail);
    let answer = ui.alert(confirmAccount, ui.ButtonSet.OK_CANCEL);
    if (answer !== ui.Button.OK) {
      throw new Error(localizedMessage.messageList.errorCanceled);
    }
    console.log(`Account confirmed manually: ${myEmail}`); // log
    // Get template from Gmail draft
    let promptMessage = localizedMessage.messageList.promptEnterSubjectOfTemplateDraft;
    let promptResult = ui.prompt(promptMessage, ui.ButtonSet.OK_CANCEL);
    let [selectedButton, subjectText] = [promptResult.getSelectedButton(), promptResult.getResponseText()];
    if (selectedButton !== ui.Button.OK) {
      throw new Error(localizedMessage.messageList.errorCanceled);
    } else if (subjectText == null) {
      throw new Error(localizedMessage.message.errorNoTextEntered);
    }
    console.log(`Entered subject of draft template, loading template "${subjectText}"...`); // log
    let draftMessage = getDraftBySubject_(subjectText);
    // Check for duplicates
    if (draftMessage.length > 1) {
      throw new Error(localizedMessage.messageList.errorTwoOrMoreDraftsWithSameSubject);
    }
    // Throw error if no draft template is found
    if (draftMessage.length == 0) {
      throw new Error(localizedMessage.messageList.errorNoMatchingTemplateDraft);
    }
    // Store template into an object
    let template = {
      'subject': subjectText,
      'plainBody': draftMessage[0].getPlainBody(),
      'htmlBody': draftMessage[0].getBody(),
      'attachments': draftMessage[0].getAttachments()
    };
    // Check template format; plain or HTML text.
    let isPlainText = (template.plainBody === template.htmlBody);
    // Check for consistency between config.ENABLE_GROUP_MERGE and template
    console.log(`config.ENABLE_GROUP_MERGE is set to "${config.ENABLE_GROUP_MERGE}".`); // log
    if (!config.ENABLE_GROUP_MERGE) {
      let nmFieldCounter = 0;
      for (let k in template) {
        if (k == 'attachments') {
          continue; // Skip this process for attachment files
        }
        let nmField = template[k].match(config.GROUP_FIELD_MARKER);
        let nmFieldCount = (nmField === null ? 0 : nmField.length);
        nmFieldCounter += nmFieldCount;
      }
      if (nmFieldCounter > 0) {
        // If group merge field marker is detected in the template when ENABLE_GROUP_MERGE is set to false,
        // ask whether or not to enable this function, i.e., to change ENABLE_GROUP_MERGE to true.
        let confirmNM = localizedMessage.messageList.alertGroupMergeFieldMarkerDetected;
        let result = ui.alert(localizedMessage.messageList.alertTitleConfirmation, confirmNM, ui.ButtonSet.YES_NO);
        config.ENABLE_GROUP_MERGE = (result === ui.Button.YES);
        console.log(`config.ENABLE_GROUP_MERGE is switched to "${config.ENABLE_GROUP_MERGE}".`); // log
      }
    }
    // Create draft or send email based on the template.
    // The process depends on the value of ENABLE_GROUP_MERGE
    console.log('Composing personalized email...'); // log
    if (config.ENABLE_GROUP_MERGE) {
      // Convert the 2d-array merge data into object grouped by recipient(s)
      let groupedMergeData = groupArray_(mergeDataEolReplaced, config.RECIPIENT_COL_NAME);
      // Validity check
      if (Object.keys(groupedMergeData).length == 0) {
        throw new Error(localizedMessage.messageList.errorInvalidRECIPIENT_COL_NAME);
      }
      // Create draft or send email for each recipient
      for (let k in groupedMergeData) {
        let mergeDataObjArr = groupedMergeData[k];
        let fillInTemplate_options = {
          'excludeFromTemplate': ['attachments'],
          'replaceValue': config.REPLACE_VALUE,
          'mergeFieldMarker': config.MERGE_FIELD_MARKER,
          'enableGroupMerge': config.ENABLE_GROUP_MERGE,
          'groupFieldMarker': config.GROUP_FIELD_MARKER,
          'rowIndexMarker': config.ROW_INDEX_MARKER
        };
        let messageData = fillInTemplate_(template, mergeDataObjArr, fillInTemplate_options);
        let options = {
          'htmlBody': (isPlainText ? null : messageData.htmlBody),
          'attachments': (messageData.attachments ? messageData.attachments : null),
          'bcc': (config.BCC_TO_MYSELF ? myEmail : null)
        };
        if (draftMode) {
          GmailApp.createDraft(k, messageData.subject, messageData.plainBody, options);
          console.log(`Draft created for ${k} with group merge enabled.`); // log
        } else {
          GmailApp.sendEmail(k, messageData.subject, messageData.plainBody, options);
          console.log(`Mail sent to ${k} with group merge enabled.`); // log
        }
      }
    } else {
      // Convert the 2d-array merge data into object
      let groupedMergeData = groupArray_(mergeDataEolReplaced);
      // Create draft or send email for each recipient
      groupedMergeData.forEach(obj => {
        let mergeDataObjArr = [obj];
        let fillInTemplate_options = {
          'excludeFromTemplate': ['attachments'],
          'replaceValue': config.REPLACE_VALUE,
          'mergeFieldMarker': config.MERGE_FIELD_MARKER,
          'enableGroupMerge': config.ENABLE_GROUP_MERGE
        };
        let messageData = fillInTemplate_(template, mergeDataObjArr, fillInTemplate_options);
        let options = {
          'htmlBody': (isPlainText ? null : messageData.htmlBody),
          'attachments': (messageData.attachments ? messageData.attachments : null),
          'bcc': (config.BCC_TO_MYSELF ? myEmail : null)
        };
        if (draftMode) {
          GmailApp.createDraft(obj[config.RECIPIENT_COL_NAME], messageData.subject, messageData.plainBody, options);
          console.log(`Draft created for ${k} with group merge disabled.`); // log
        } else {
          GmailApp.sendEmail(obj[config.RECIPIENT_COL_NAME], messageData.subject, messageData.plainBody, options);
          console.log(`Mail sent to ${k} with group merge disabled.`); // log
        }
      });
    }
    // Notification
    let completeMessage = (draftMode ? localizedMessage.messageList.alertCompleteAllDraftsCreated : localizedMessage.messageList.alertCompleteAllMailsSent);
    console.log(`Processed all mails.`); // log
    ui.alert(completeMessage);
  } catch (e) {
    let message = errorMessage_(e);
    console.log(`Alert message: ${message}`); // log
    ui.alert(message);
  } finally {
    console.log('...Closing Mail Merge.') // log
  }
}

/**
 * Returns an object of configurations from spreadsheet.
 * @param {string} configSheetName Name of sheet with configurations. Defaults to 'Config'.
 * @return {Object}
 * The sheet should have a first row of headers, and its first column should include the following properties:
 * @property {string} DATA_SHEET_NAME Name of sheet in which field(s) to merge in email are stored
 * @property {string} RECIPIENT_COL_NAME Name of column in sheet 'DATA_SHEET_NAME' that designates the email address of the recipient
 * @property {string} BCC_TO_MYSELF String boolean. When true, will send (or create draft) email with the sender's address set to BCC.
 * @property {string} REPLACE_VALUE Text that will replace empty data of marker. 
 * @property {string} MERGE_FIELD_MARKER Text to be processed in RegExp() constructor to define merge field(s). Note that the backslash itself does not need to be escaped, i.e., does not need to be repeated.
 * @property {string} ENABLE_GROUP_MERGE String boolean. Enable group merge when true.
 * @property {string} GROUP_FIELD_MARKER Text to be processed in RegExp() constructor to define group merge field(s). Note that the backslash itself does not need to be escaped, i.e., does not need to be repeated.
 * @property {string} ROW_INDEX_MARKER Marker for merging row index number in a group merge.
 */
function getConfig_(configSheetName = 'Config') {
  // Get values from spreadsheet
  var configValues = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(configSheetName).getDataRange().getValues();
  configValues.shift();
  // Convert the 2d array values into a Javascript object
  var configObj = configValues.reduce((obj, element) => {
    obj[element[0]] = element[1];
    return obj
  }, {});
  // Convert data types
  configObj.BCC_TO_MYSELF = toBoolean_(configObj.BCC_TO_MYSELF);
  configObj.ENABLE_GROUP_MERGE = toBoolean_(configObj.ENABLE_GROUP_MERGE);
  configObj.MERGE_FIELD_MARKER = new RegExp(configObj.MERGE_FIELD_MARKER, 'g');
  configObj.GROUP_FIELD_MARKER = new RegExp(configObj.GROUP_FIELD_MARKER, 'g');
  return configObj;
}

/**
 * Convert string booleans into boolean data
 * @param {string} stringBoolean 
 * @return {boolean}
 */
function toBoolean_(stringBoolean) {
  return stringBoolean.toLowerCase() === 'true';
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
  if (!('replaceValue' in options)) {
    options['replaceValue'] = CONFIG.REPLACE_VALUE;
  }
  if (!('mergeFieldMarker' in options)) {
    options['mergeFieldMarker'] = CONFIG.MERGE_FIELD_MARKER;
  }
  if (!('enableGroupMerge' in options)) {
    options['enableGroupMerge'] = CONFIG.ENABLE_GROUP_MERGE;
  }
  if (!('groupFieldMarker' in options)) {
    options['groupFieldMarker'] = CONFIG.GROUP_FIELD_MARKER;
  }
  if (!('rowIndexMarker' in options)) {
    options['rowIndexMarker'] = CONFIG.ROW_INDEX_MARKER;
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
      let groupText = text.match(options.groupFieldMarker);
      // If the number of group field marker is not 0...
      if (groupText !== null) {
        let groupTextMerged = groupText.map(function (field) {
          // Get the text inside group field markers, e.g., [[group field]] => group field
          field = field.substring(2, field.length - 2); // assuming that the text length for opening and closing markers are 2 and 2, respectively
          // Create an array of merge field markers within a group merge marker
          let fieldVars = field.match(options.mergeFieldMarker);
          if (!fieldVars) {
            return field; // return the text of group merge field itself if no merge field marker is found within the group merge marker.
          } else {
            let fieldMerged = [];
            for (let i = 0; i < data.length; ++i) {
              let datum = data[i];
              let rowIndex = i + 1;
              let fieldRowIndexed = field.replace(options.rowIndexMarker, rowIndex);
              let fieldVarsCopy = fieldVars.slice();
              // Get the text inside markers, e.g., {{field name}} => field name
              let fieldMarkerText = fieldVarsCopy.map(value => value.substring(2, value.length - 2)); // assuming that the text length for opening and closing markers are 2 and 2, respectively 
              fieldVarsCopy.forEach(
                (variable, ind) => fieldRowIndexed = fieldRowIndexed.replace(variable, datum[fieldMarkerText[ind]] || options.replaceValue)
              );
              fieldMerged.push(fieldRowIndexed);
            }
            let fieldMergedText = fieldMerged.join('');
            return fieldMergedText;
          }
        });
        groupText.forEach(
          (groupField, index) => text = text.replace(groupField, groupTextMerged[index] || options.replaceValue)
        );
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
      textVars.forEach(
        (variable, i) => text = text.replace(variable, mergeData[markerText[i]] || options.replaceValue)
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
  let message = e.stack;
  return message;
}
