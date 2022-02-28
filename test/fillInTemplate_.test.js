const mailMerge = require('../src/mailMerge');
const testTemplate = {
  testBodyTemplateWithGroupMerge:
    '[Body]\n\n[[\nEntry Number: {{i}}\nEmail: {{Email}}\nName: {{Name}}\n]]',
  testBodyTemplateWithoutGroupMerge:
    '[Body]\n\nEmail: {{Email}}\nName: {{Name}}',
  testSubTemplateWithoutGroupMerge: 'Subject: {{Name}}',
};
const patterns = [
  {
    testTitle: 'fillInTemplate_ with Group Merge disabled',
    enableGroupMerge: false,
    groupedMergeData_k: [
      {
        ID: 'ID001',
        Email: 'sample1@email.com',
        Name: 'sample-name1',
      },
    ],
    expectedOutput: {
      testBodyTemplateWithGroupMerge:
        '[Body]\n\n[[\nEntry Number: NA\nEmail: sample1@email.com\nName: sample-name1\n]]',
      testBodyTemplateWithoutGroupMerge:
        '[Body]\n\nEmail: sample1@email.com\nName: sample-name1',
      testSubTemplateWithoutGroupMerge: 'Subject: sample-name1',
    },
  },
  {
    testTitle: 'fillInTemplate_ with Group Merge enabled',
    enableGroupMerge: true,
    groupedMergeData_k: [
      {
        ID: 'ID002',
        Email: 'sample2@example.com',
        Name: 'sample-name2',
      },
      {
        ID: 'ID004',
        Email: 'sample2@example.com',
        Name: 'sample-name 2',
      },
    ],
    expectedOutput: {
      testBodyTemplateWithGroupMerge:
        '[Body]\n\n\nEntry Number: 1\nEmail: sample2@example.com\nName: sample-name2\n\nEntry Number: 2\nEmail: sample2@example.com\nName: sample-name 2\n',
      testBodyTemplateWithoutGroupMerge:
        '[Body]\n\nEmail: sample2@example.com\nName: sample-name2',
      testSubTemplateWithoutGroupMerge: 'Subject: sample-name2',
    },
  },
];

patterns.forEach((pattern) => {
  test(pattern.testTitle, () => {
    expect(
      mailMerge.fillInTemplate_(testTemplate, pattern.groupedMergeData_k, {
        enableGroupMerge: pattern.enableGroupMerge,
      })
    ).toEqual(pattern.expectedOutput);
  });
});
