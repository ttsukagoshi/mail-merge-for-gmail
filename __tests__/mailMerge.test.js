jest.mock('../src/__mocks__/spreadsheetapp');

const groupMerge = require('../src/group-merge');
const { mockConfigs } = require('../src/__mocks__/mock-configs');
const SpreadsheetApp = require('../src/__mocks__/spreadsheetapp');
const Spreadsheet = require('../src/__mocks__/spreadsheet');
const Sheet = require('../src/__mocks__/sheet');
const draftModes = [true, false];

draftModes.forEach((draftMode) => {
  mockConfigs.forEach((mockConfig) => {
    SpreadsheetApp.openByUrl.mockReturnValue(
      new Spreadsheet({
        url: mockConfig.parsedConfig.SPREADSHEET_URL,
        sheets: [new Sheet(mockConfig.parsedConfig.DATA_SHEET_NAME)],
      })
    );
    test(`${mockConfig.eventName} with draftMode = ${
      draftMode ? 'true' : 'false'
    }`, () => {
      expect(groupMerge.mailMerge(draftMode, mockConfig.parsedConfig)).toBe(
        mockConfig.mailMergeOutput
      );
    });
  });
});
