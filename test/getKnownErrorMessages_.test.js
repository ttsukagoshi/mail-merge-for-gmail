const mailMerge = require('../src/mailMerge');

Object.keys(mailMerge.MESSAGES).forEach((locale) => {
  let matchArray = Object.keys(mailMerge.MESSAGES[locale])
    .filter((messageKey) => messageKey.slice(0, 5) === 'error')
    .map((key) => mailMerge.MESSAGES[locale][key]);
  test('getKnownErrorMessages_', () => {
    expect(mailMerge.getKnownErrorMessages_(locale)).toEqual(matchArray);
  });
});
