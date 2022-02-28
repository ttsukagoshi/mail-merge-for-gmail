/* global fillInTemplate_ */

function testAll() {
  let allTestSummary = [testFillInTemplate()].reduce(
    (summary, test) => {
      summary.allTestNames.push(test.name);
      if (!test.passed) {
        summary.allErrors.push(test);
      }
      return summary;
    },
    { allTestNames: [], allErrors: [] }
  );
  console.log(
    allTestSummary.allErrors.length === 0
      ? 'Passed all tests'
      : `${
          allTestSummary.allErrors.length
        } error(s) were found: ${JSON.stringify(allTestSummary.allErrors)}`
  );
}

/**
 * Test function fillInTemplate_ in src/mailMerge.js
 * @returns {Object}
 */
function testFillInTemplate() {
  let result = { name: 'fillInTemplate_', passed: true, errors: [] };
  const testTemplate = {
    testBodyTemplateWithGroupMerge:
      '[Body] [[ Entry Number: {{i}} Email: {{Email}} Name: {{Name}}]]',
    testBodyTemplateWithoutGroupMerge: '[Body] Email: {{Email}} Name: {{Name}}',
    testSubTemplateWithoutGroupMerge: 'Subject: {{Name}}',
  };
  const patterns = [
    {
      enableGroupMerge: false,
      groupedMergeData_k: [
        {
          ID: 'ID001',
          Email: 'sample1@email.com',
          Name: 'sample-name1',
        },
      ],
      expectedOutput: JSON.stringify({
        testBodyTemplateWithGroupMerge:
          '[Body] [[ Entry Number: NA Email: sample1@email.com Name: sample-name1]]',
        testBodyTemplateWithoutGroupMerge:
          '[Body] Email: sample1@email.com Name: sample-name1',
        testSubTemplateWithoutGroupMerge: 'Subject: sample-name1',
      }),
    },
    {
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
      expectedOutput: JSON.stringify({
        testBodyTemplateWithGroupMerge:
          '[Body]  Entry Number: 1 Email: sample2@example.com Name: sample-name2 Entry Number: 2 Email: sample2@example.com Name: sample-name 2',
        testBodyTemplateWithoutGroupMerge:
          '[Body] Email: sample2@example.com Name: sample-name2',
        testSubTemplateWithoutGroupMerge: 'Subject: sample-name2',
      }),
    },
  ];
  patterns.forEach((pattern) => {
    let mergedText = JSON.stringify(
      fillInTemplate_(testTemplate, pattern.groupedMergeData_k, {
        enableGroupMerge: pattern.enableGroupMerge,
      })
    );
    if (mergedText !== pattern.expectedOutput) {
      result.passed = false;
      result.errors.push(
        `fillInTemplate_ result does not match: ${mergedText}`
      );
    }
  });
  return result;
}

/**
 * Delete all user properties to reproduce what first-time users will see.
 */
function resetUserProperties() {
  PropertiesService.getUserProperties().deleteAllProperties();
}

if (typeof module === 'object') {
  module.exports = { resetUserProperties, testAll };
}
