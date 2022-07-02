const groupMerge = require('../src/group-merge');
const { mockConfigs } = require('../src/__mocks__/mock-configs');

mockConfigs.forEach((mockConfig) => {
  test(mockConfig.eventName, () => {
    expect(groupMerge.parseConfig_(mockConfig.addonEvent)).toEqual(
      mockConfig.parsedConfig
    );
  });
});
