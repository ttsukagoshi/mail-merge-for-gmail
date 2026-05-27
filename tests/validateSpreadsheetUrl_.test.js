// eslint-disable-next-line @typescript-eslint/no-var-requires
const groupMerge = require('../src/group-merge');

const localizedMessage = {
  messageList: groupMerge.MESSAGES['en'],
};

const VALID_URL =
  'https://docs.google.com/spreadsheets/d/1aBcDeFgHiJkLmNoPqRsTuVwXyZ/edit#gid=0';
const VALID_FILE_ID = '1aBcDeFgHiJkLmNoPqRsTuVwXyZ';

beforeEach(() => {
  global.DriveApp = {
    getFileById: jest.fn(),
  };
});

afterEach(() => {
  jest.restoreAllMocks();
});

test('should pass for a valid Google Sheets URL', () => {
  DriveApp.getFileById.mockReturnValue({
    getMimeType: () => 'application/vnd.google-apps.spreadsheet',
  });
  expect(() =>
    groupMerge.validateSpreadsheetUrl_(VALID_URL, localizedMessage),
  ).not.toThrow();
  expect(DriveApp.getFileById).toHaveBeenCalledWith(VALID_FILE_ID);
});

test('should throw errorSpreadsheetNotFound for a malformed URL', () => {
  expect(() =>
    groupMerge.validateSpreadsheetUrl_(
      'https://example.com/not-a-spreadsheet',
      localizedMessage,
    ),
  ).toThrow(localizedMessage.messageList.errorSpreadsheetNotFound);
});

test('should throw errorSpreadsheetNoAccess when DriveApp cannot access the file', () => {
  DriveApp.getFileById.mockImplementation(() => {
    throw new Error('No access');
  });
  expect(() =>
    groupMerge.validateSpreadsheetUrl_(VALID_URL, localizedMessage),
  ).toThrow(localizedMessage.messageList.errorSpreadsheetNoAccess);
});

test('should throw errorNotGoogleSheet for a non-Google Sheets file', () => {
  DriveApp.getFileById.mockReturnValue({
    getMimeType: () =>
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  expect(() =>
    groupMerge.validateSpreadsheetUrl_(VALID_URL, localizedMessage),
  ).toThrow(localizedMessage.messageList.errorNotGoogleSheet);
});
