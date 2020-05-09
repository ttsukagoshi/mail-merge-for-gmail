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

// Default configurations
const CONFIG = {
  DATA_SHEET_NAME: 'List',
  RECIPIENT_COL_NAME: 'Email',
  BCC_TO_MYSELF: true,
  REPLACE_VALUE: 'NA',
  MERGE_FIELD_MARKER: /\{\{[^\}]+\}\}/g,
  ENABLE_NESTED_MERGE: false,
  NESTED_FIELD_MARKER: /\[\[[^\]]+\]\]/g,
  ROW_INDEX_MARKER: '{{i}}'
}

// Add spreadsheet menu
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Mail Merge')
    .addItem('Create Draft', 'createDraftEmails')
    .addSeparator()
    .addItem('Send Emails', 'sendEmails')
    .addToUi();
}

/**
 * Create draft of personalized email(s)
 */
function createDraftEmails() {
  const draftMode = true;
  const config = getConfig_('Config');
  sendPersonalizedEmails_(draftMode, config);
}

/**
 * Send personalized email(s)
 */
function sendEmails() {
  const draftMode = false;
  const config = getConfig_('Config');
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

  // Get data of field(s) to merge in form of 2d array
  var dataSheet = ss.getSheetByName(config.DATA_SHEET_NAME);
  var mergeDataRange = dataSheet.getDataRange().setNumberFormat('@'); // Convert all formatted dates and numbers into texts
  var mergeData = mergeDataRange.getValues();

  // Convert line breaks in the spreadsheet (in LF format, i.e., '\n')
  // to CRLF format ('\r\n') for merging into Gmail plain text
  let mergeDataEolReplaced = mergeData.map(element => arrReplace_(element, /\n|\r|\r\n/g, '\r\n'));

  try {
    // Confirmation before sending email
    let confirmAccount = (draftMode === true
      ? `Are you sure you want to create draft email(s) as ${myEmail}?`
      : `Are you sure you want to send email(s) as ${myEmail}?`);
    let answer = ui.alert(confirmAccount, ui.ButtonSet.OK_CANCEL);
    if (answer !== ui.Button.OK) {
      throw new Error('Mail merge canceled.');
    }
    // Get template from Gmail draft
    let promptMessage = 'Enter the subject text of Gmail draft to use as template.';
    let promptResult = ui.prompt(promptMessage, ui.ButtonSet.OK_CANCEL);
    let [selectedButton, subjectText] = [promptResult.getSelectedButton(), promptResult.getResponseText()];
    if (selectedButton !== ui.Button.OK) {
      throw new Error('Mail merge canceled.');
    } else if (subjectText == null) {
      throw new Error('No text entered.');
    }
    let draftMessage = getDraftBySubject_(subjectText);
    // Check for duplicates
    if (draftMessage.length > 1) {
      throw new Error('There are 2 or more Gmail drafts with the subject you entered. Enter a unique subject text.');
    }
    // Store template into an object
    let template = {
      'subject': subjectText,
      'plainBody': draftMessage[0].getPlainBody(),
      'htmlBody': draftMessage[0].getBody()
    };
    // Check for consistency between config.ENABLE_NESTED_MERGE and template
    if (config.ENABLE_NESTED_MERGE === false) {
      let nmFieldCounter = 0;
      for (let k in template) {
        let nmField = template[k].match(config.NESTED_FIELD_MARKER);
        let nmFieldCount = (nmField === null ? 0 : nmField.length)
        nmFieldCounter += nmFieldCount;
      }
      // If nested merge field marker is detected in the template when ENABLE_NESTED_MERGE is set to false,
      // ask whether or not to enable this function, i.e., to change ENABLE_NESTED_MERGE to true.
      if (nmFieldCounter > 0){
        let confirmNM = 'Nested merge field marker detected. Do you want to enable nested merge function?';
        let result = ui.alert('Confirmation', confirmNM, ui.ButtonSet.YES_NO);
        config.ENABLE_NESTED_MERGE = (result === ui.Button.YES ? true : config.ENABLE_NESTED_MERGE);
      }
    }
    // Create draft or send email based on the template.
    // The process depends on the value of ENABLE_NESTED_MERGE
    if (config.ENABLE_NESTED_MERGE === true) {
      // Convert the 2d-array merge data into object grouped by recipient(s)
      let groupedMergeData = groupArray_(mergeDataEolReplaced, config.RECIPIENT_COL_NAME);
      // Validity check
      if (Object.keys(groupedMergeData).length == 0) {
        throw new Error('Invalid RECIPIENT_COL_NAME. Check sheet "Config" to make sure it refers to an existing column name.');
      }
      // Create draft or send email for each recipient
      for (let k in groupedMergeData) {
        let mergeDataObjArr = groupedMergeData[k];
        let messageData = fillInTemplate_(template, mergeDataObjArr, config.REPLACE_VALUE, config.MERGE_FIELD_MARKER, config.ENABLE_NESTED_MERGE, config.NESTED_FIELD_MARKER, config.ROW_INDEX_MARKER);
        let options = {
          'htmlBody': messageData.htmlBody,
          'bcc': (config.BCC_TO_MYSELF === true ? myEmail : null)
        };
        draftMode === true
          ? GmailApp.createDraft(k, messageData.subject, messageData.plainBody, options)
          : GmailApp.sendEmail(k, messageData.subject, messageData.plainBody, options);
      }
    } else {
      // Convert the 2d-array merge data into object
      let groupedMergeData = groupArray_(mergeDataEolReplaced);
      // Create draft or send email for each recipient
      for (let i = 0; i < groupedMergeData.data.length; ++i) {
        let object = groupedMergeData.data[i]
        let mergeDataObjArr = [];
        mergeDataObjArr.push(object);
        let messageData = fillInTemplate_(template, mergeDataObjArr, config.REPLACE_VALUE, config.MERGE_FIELD_MARKER, config.ENABLE_NESTED_MERGE);
        let options = {
          'htmlBody': messageData.htmlBody,
          'bcc': (config.BCC_TO_MYSELF === true ? myEmail : null)
        };
        draftMode === true
          ? GmailApp.createDraft(object[config.RECIPIENT_COL_NAME], messageData.subject, messageData.plainBody, options)
          : GmailApp.sendEmail(object[config.RECIPIENT_COL_NAME], messageData.subject, messageData.plainBody, options);
      }
    }
    // Notification
    let completeMessage = (draftMode === true
      ? 'Complete: All draft(s) created.'
      : 'Complete: All mail(s) sent.');
    ui.alert(completeMessage)
  } catch (e) {
    let message = errorMessage_(e);
    ui.alert(message);
  }
}

/**
 * Process each element of an array with String.prototype.replace()
 * @param {array} array Array containing values to replace.
 * @param {string|RegExp} searchValue A RegExp object or literal, or string to be replaced by replaceValue
 * @param {string} replaceValue String to replace the searchValue.
 * @return {array} Array whose element(s) are replaced.
 */
function arrReplace_(array, searchValue, replaceValue) {
  let replacedArray = array.map(value => value.replace(searchValue, replaceValue));
  return replacedArray;
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
 * @property {string} ENABLE_NESTED_MERGE String boolean. Enable nested merge when true.
 * @property {string} NESTED_FIELD_MARKER Text to be processed in RegExp() constructor to define nested merge field(s). Note that the backslash itself does not need to be escaped, i.e., does not need to be repeated.
 * @property {string} ROW_INDEX_MARKER Marker for merging row index number in a nested merge.
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
  configObj.ENABLE_NESTED_MERGE = toBoolean_(configObj.ENABLE_NESTED_MERGE);
  configObj.MERGE_FIELD_MARKER = new RegExp(configObj.MERGE_FIELD_MARKER, 'g');
  configObj.NESTED_FIELD_MARKER = new RegExp(configObj.NESTED_FIELD_MARKER, 'g');
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
    groupedObj['data'] = data.map(function (values) {
      return header.reduce(function (obj, key, ind) {
        obj[key] = values[ind];
        return obj;
      }, {});
    });
    return groupedObj;
  } else if (index < 0) {
    let invalidProperty = {};
    return invalidProperty;
  } else {
    let groupedObj = data.reduce(
      function (accObj, curArr) {
        let key = curArr[index];
        if (!accObj[key]) {
          accObj[key] = [];
        }
        let rowObj = createObj_(header, curArr);
        accObj[key].push(rowObj);
        return accObj;
      }, {});
    return groupedObj;
  }
}

/**
 * Create a Javascript object from a set of keys and values
 * i.e., where keys = [key0, key1, ..., key[n]] and values = [value0, value1, ..., value[n]],
 * this function will return an object = {key0: value0, ..., key[n]: value[n]}
 * @param {array} keys 
 * @param {array} values
 * @return {Object} 
 */
function createObj_(keys, values) {
  let obj = {};
  for (let i = 0; i < keys.length; ++i) {
    obj[keys[i]] = values[i];
  }
  return obj;
}

/**
 * Replaces markers in a template object with values defined in a JavaScript data object.
 * @param {Object} template Template object containing markers, as designated in regular expression in mergeFieldMarker
 * @param {array} data Array of object(s) with values to replace markers.
 * @param {string} replaceValue [Optional] String to replace empty data of a marker. Defaults to 'NA' for Not Available. 
 * @param {RegExp} mergeFieldMarker [Optional] Regular expression for the merge field marker. Defaults to /\{\{[^\}]+\}\}/g e.g., {{field name}}
 * @param {boolean} enableNestedMerge [Optional] Merged texts are returned in concatenated form when true. Defaults to false.
 * @param {RegExp} nestedFieldMarker [Optional] Regular expression for the nested merge field marker. Defaults to /\[\[[^\]]+\]\]/g e.g., [[nested merge]]
 * @param {string} rowIndexMarker [Optional] Marker for merging row index number in a nested merge. Defaults to '{{i}}'
 * @return {Object} Returns the template object with markers replaced for personalized text.
 */
function fillInTemplate_(template, data, replaceValue = 'NA', mergeFieldMarker = /\{\{[^\}]+\}\}/g, enableNestedMerge = false, nestedFieldMarker = /\[\[[^\]]+\]\]/g, rowIndexMarker = '{{i}}') {
  let messageData = {};
  for (let k in template) {
    let text = '';
    text = template[k];
    // Nested merge
    if (enableNestedMerge === true) {
      // Create an array of nested field marker(s) in the text
      let nestedText = text.match(nestedFieldMarker);
      // If the number of nested field marker is not 0...
      if (nestedText !== null) {
        let nestedTextMerged = nestedText.map(function (field) {
          // Get the text inside nested field markers, e.g., [[nested field]] => nested field
          field = field.substring(2, field.length - 2); // assuming that the text length for opening and closing markers are 2 and 2, respectively
          // Create an array of merge field markers within a nested merge marker
          let fieldVars = field.match(mergeFieldMarker);
          if (!fieldVars) {
            return field; // return the text of nested merge field itself if no merge field marker is found within the nested merge marker.
          } else {
            let fieldMerged = [];
            for (let i = 0; i < data.length; ++i) {
              let datum = data[i];
              let rowIndex = i + 1;
              let fieldRowIndexed = field.replace(rowIndexMarker, rowIndex);
              let fieldVarsCopy = fieldVars.slice();
              // Get the text inside markers, e.g., {{field name}} => field name
              let fieldMarkerText = fieldVarsCopy.map(value => value.substring(2, value.length - 2)); // assuming that the text length for opening and closing markers are 2 and 2, respectively 
              fieldVarsCopy.forEach(
                (variable, ind) => fieldRowIndexed = fieldRowIndexed.replace(variable, datum[fieldMarkerText[ind]] || replaceValue)
              )
              fieldMerged.push(fieldRowIndexed);
            }
            let fieldMergedText = fieldMerged.join('');
            return fieldMergedText;
          }
        });
        nestedText.forEach(
          (nestedField, index) => text = text.replace(nestedField, nestedTextMerged[index] || replaceValue)
        );
      }
    }
    // merging the rest of the field based on the first row of data object array.
    let mergeData = data[0];
    // Create an array of all merge fields to be replaced
    let textVars = text.match(mergeFieldMarker);
    if (!textVars) {
      messageData[k] = text; // return text itself if no marker is found
    } else {
      // Get the text inside markers, e.g., {{field name}} => field name 
      let markerText = textVars.map(value => value.substring(2, value.length - 2)); // assuming that the text length for opening and closing markers are 2 and 2, respectively
      // Replace variables in textVars with the actual values from the data object.
      // If no value is available, replace with replaceValue.
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
