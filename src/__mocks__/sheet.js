// Mock GAS classes

/**
 * @see https://developers.google.com/apps-script/reference/spreadsheet/sheet?hl=en
 */
class Sheet {
  /**
   * @param {String} sheetName
   */
  constructor(sheetName) {
    this.name = sheetName;
  }
}

module.exports = Sheet;
