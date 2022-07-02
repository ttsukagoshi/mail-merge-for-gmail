const EMAIL = 'myEmail@example.com';

/**
 * Mock class for MailApp
 */
class MailApp {
  /**
   * Mock sendMail method. Outputs the entered paramters on `console.log`.
   * Currently does not support:
   * - `options.attachments` and `options.inlineImages`
   * - Other forms of sendMail method like `sendMail(message)` and `sendMail(to, replyTo, subject, body)`
   * @param {String} recipient The addresses of the recipients, separated by commas
   * @param {String} subject The subject line
   * @param {String} body The body of the email
   * @param {Object} options A JavaScript object that specifies advanced parameters, as listed in the reference URL
   * @returns {String} Stringified object of the entered parameters
   * @see https://developers.google.com/apps-script/reference/mail/mail-app?hl=en#sendemailrecipient,-subject,-body,-options
   */
  sendMail(recipient, subject, body, options = {}) {
    console.log(
      JSON.stringify({
        recipient: recipient,
        subject: subject,
        body: body,
        options: options,
      })
    );
  }
}

class PropertiesService {
  static getUserProperties() {
    return new UserProperties({});
  }
}

class Session {
  static getActiveUser() {
    return new User();
  }
}

/**
 * Mock class for Sheet
 */
class Sheet {}

/**
 * Mock class for SpreadsheetApp
 */
class SpreadsheetApp {
  /**
   *
   * @param {*} properties
   * @returns
   */
  getActiveSpreadsheet(properties = {}) {
    return new Spreadsheet(properties);
  }
  openById(id) {
    return new Spreadsheet({ id: id });
  }
  openByUrl(url) {
    return new Spreadsheet({ url: url });
  }
}

/**
 * Mock class for Spreadsheet
 */
class Spreadsheet {
  constructor(properties) {
    this.name = properties.name || 'New Spreadsheet';
    this.id = properties.id || 'id_000000';
    this.url =
      properties.url ||
      `https://docs.google.com/spreadsheets/d/${this.id}/edit`;
    this.sheets = [new Sheet('Sheet 1')];
  }
  getId() {
    return this.id;
  }
  getName() {
    return this.name;
  }
  /**
   * Returns a sheet with the given name.
   * @param {String} sheetName
   * @returns Sheet object with the given name.
   * @see https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet?hl=en#getsheetbynamename
   */
  getSheetByName(sheetName) {
    return this.sheets.filter((sheet) => sheet.getName(sheetName))[0] || null;
  }
  getUrl() {
    return this.url;
  }
}

class User {
  constructor() {
    this.email = EMAIL;
  }
  getEmail() {
    return this.email;
  }
}

class UserProperties {
  constructor(propertiesString = '{}') {
    this.userProperties = JSON.parse(propertiesString);
  }
  getProperty(key) {
    if (Object.keys(this.userProperties).length > 0) {
      // 仕掛かり中
      console.log(key);
    }
  }
}

module.exports = {
  MailApp,
  PropertiesService,
  Session,
  SpreadsheetApp,
  Spreadsheet,
  Sheet,
};
