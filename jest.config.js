/*
 * For a detailed explanation regarding each configuration property, visit:
 * https://jestjs.io/docs/configuration
 */

module.exports = {
  clearMocks: true,
  collectCoverage: true,
  coverageDirectory: 'coverage',
  coverageProvider: 'v8',
  globals: {
    PropertiesService: {},
    SpreadsheetApp: {},
    UrlFetchApp: {},
    Utilities: {},
  },
  moduleDirectories: ['node_modules'],
  moduleFileExtensions: ['js', 'ts', 'tsx', 'json'],
  preset: 'ts-jest',
  testEnvironment: 'node',
  testMatch: ['**/tests/**/*.[jt]s?(x)'],
  transform: {
    '^.+\\.tsx?$': 'ts-jest',
  },
  verbose: true,
};
