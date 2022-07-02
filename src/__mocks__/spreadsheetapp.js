// Mock GAS classes

const Spreadsheet = require('./spreadsheet');
const Ui = require('./ui');

/**
 * @see https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet-app?hl=en
 */
class SpreadsheetApp {
  static getActiveSpreadsheet() {
    return new Spreadsheet();
  }
  static getUi() {
    return new Ui();
  }
  /**
   * @param {String} id
   */
  static openById(id) {
    return new Spreadsheet({ id: id });
  }
  /**
   * @param {String} url
   */
  static openByUrl(url) {
    return new Spreadsheet({ url: url });
  }
}

module.exports = SpreadsheetApp;
