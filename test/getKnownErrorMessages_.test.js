const groupMerge = require('../src/group-merge');

Object.keys(groupMerge.MESSAGES).forEach((locale) => {
  let matchArray = Object.keys(groupMerge.MESSAGES[locale])
    .filter((messageKey) => messageKey.slice(0, 5) === 'error')
    .map((key) => groupMerge.MESSAGES[locale][key]);
  test('getKnownErrorMessages_', () => {
    expect(groupMerge.getKnownErrorMessages_(locale)).toEqual(matchArray);
  });
});
