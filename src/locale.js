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

const MESSAGES = {
  'en_US': {
    'menuName': 'Mail Merge',
    'menuCreateDraft': 'Create Draft',
    'menuSendEmails': 'Send Emails'
  },
  'ja': {
  }
};

class LocalizedMessage {
  constructor(userLocale = 'en_US') {
    this.locale = (MESSAGES[userLocale] ? userLocale : 'en_US');
    this.messageList = MESSAGES[this.locale];
    Object.keys(MESSAGES.en_US).forEach(key => {
      if (!this.messageList[key]) {
        this.messageList[key] = MESSAGES.en_US[key];
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
   * Replace placeholder string in this.messageList.hogefuga
   * @param {*} var1
   * @param {*} var2
   * @returns {string} The replaced text.
   */
  replaceHogefuga(var1, var2) {
    let text = this.messageList.hogefuga;
    let placeholderValues = [
      {
        'regexp': '\{\{var1\}\}',
        'value': var1
      },
      {
        'regexp': '\{\{var2\}\}',
        'value': var2
      }
    ];
    text = this.replacePlaceholders_(text, placeholderValues);
    return text;
  }
}