const groupMerge = require('../src/group-merge');
const { mockConfigs } = require('./mock-configs');

mockConfigs.forEach((mockConfig) => {
  test(mockConfig.eventName, () => {
    expect(groupMerge.parseConfig_(mockConfig.addonEvent)).toEqual(
      mockConfig.parsedConfig
    );
  });
});
