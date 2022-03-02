const mailMerge = require('../src/mailMerge');
const i18n = require('../src/i18n');

Object.keys(i18n.MESSAGES).forEach((locale) => {
  let matchArray = Object.keys(i18n.MESSAGES[locale])
    .filter((messageKey) => messageKey.slice(0, 5) === 'error')
    .map((key) => i18n.MESSAGES[locale][key]);
  test('getKnownErrorMessages_', () => {
    expect(mailMerge.getKnownErrorMessages_(locale)).toEqual(matchArray);
  });
});
