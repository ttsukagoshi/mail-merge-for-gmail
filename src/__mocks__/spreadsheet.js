// Mock GAS classes

const Sheet = require('./sheet');
const defaultSpreadsheetName = 'New Spreadsheet';
const defaultSpreadsheetId = 'id_000000';
const defaultSheetName = 'Sheet 1';

/**
 * @see https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet?hl=en
 */
class Spreadsheet {
  constructor(properties = {}) {
    this.name = properties.name || defaultSpreadsheetName;
    this.id = properties.id || defaultSpreadsheetId;
    this.url =
      properties.url ||
      `https://docs.google.com/spreadsheets/d/${this.id}/edit`;
    this.sheets = [new Sheet(defaultSheetName)];
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

module.exports = Spreadsheet;
