// Mock configs for Group Merge

const hostApps = ['GMAIL']; // exclude 'SHEET' from test for now
const replyTo = 'replyTo@email.com';
const sheetName = 'mock sheet name 123';
const spreadsheetId = 'qwerty12345';
const mockConfigs = hostApps
  .map((hostApp) => [
    {
      eventName: `Default settings on ${hostApp}`,
      addonEvent: {
        commonEventObject: {
          userLocale: 'ja',
          formInputs: {
            GROUP_FIELD_MARKER_TEXT: {
              stringInputs: { value: ['[[]{2}([^[^\\]]+)[\\]]{2}'] },
            },
            MERGE_FIELD_MARKER_TEXT: {
              stringInputs: { value: ['[{]{2}([^{^}]+)[}]{2}'] },
            },
            REPLY_TO: { stringInputs: { value: [replyTo] } },
            REPLACE_VALUE: { stringInputs: { value: ['NA'] } },
            ENABLE_GROUP_MERGE: { stringInputs: { value: ['enabled'] } },
            ROW_INDEX_MARKER: { stringInputs: { value: ['{{i}}'] } },
          },
          timeZone: { offset: 32400000, id: 'Asia/Dili' },
          hostApp: hostApp,
          platform: 'WEB',
        },
      },
      parsedConfig: {
        SPREADSHEET_URL: '',
        DATA_SHEET_NAME: '',
        TO: '',
        CC: '',
        BCC: '',
        TEMPLATE_SUBJECT: '',
        ENABLE_GROUP_MERGE: true,
        ENABLE_REPLY_TO: false,
        REPLY_TO: replyTo,
        REPLACE_VALUE: 'NA',
        MERGE_FIELD_MARKER: /[{]{2}([^{^}]+)[}]{2}/g,
        MERGE_FIELD_MARKER_TEXT: '[{]{2}([^{^}]+)[}]{2}',
        GROUP_FIELD_MARKER: /[[]{2}([^[^\]]+)[\]]{2}/g,
        GROUP_FIELD_MARKER_TEXT: '[[]{2}([^[^\\]]+)[\\]]{2}',
        ROW_INDEX_MARKER: '{{i}}',
        ENABLE_DEBUG_MODE: false,
        hostApp: hostApp,
        userLocale: 'ja',
      },
      mailMergeOutput: '',
    },
    {
      eventName: `Default settings with SPREADSHEET_URL, DATA_SHEET_NAME, and TO values on ${hostApp}`,
      addonEvent: {
        commonEventObject: {
          userLocale: 'ja',
          formInputs: {
            DATA_SHEET_NAME: { stringInputs: { value: [sheetName] } },
            GROUP_FIELD_MARKER_TEXT: {
              stringInputs: { value: ['[[]{2}([^[^\\]]+)[\\]]{2}'] },
            },
            MERGE_FIELD_MARKER_TEXT: {
              stringInputs: { value: ['[{]{2}([^{^}]+)[}]{2}'] },
            },
            REPLY_TO: { stringInputs: { value: [replyTo] } },
            REPLACE_VALUE: { stringInputs: { value: ['NA'] } },
            TO: { stringInputs: { value: ['{{Email}}'] } },
            ENABLE_GROUP_MERGE: { stringInputs: { value: ['enabled'] } },
            SPREADSHEET_URL: {
              stringInputs: {
                value: [
                  `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`,
                ],
              },
            },
            ROW_INDEX_MARKER: { stringInputs: { value: ['{{i}}'] } },
          },
          timeZone: { offset: 32400000, id: 'Asia/Dili' },
          hostApp: hostApp,
          platform: 'WEB',
        },
      },
      parsedConfig: {
        SPREADSHEET_URL: `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`,
        DATA_SHEET_NAME: sheetName,
        TO: '{{Email}}',
        CC: '',
        BCC: '',
        TEMPLATE_SUBJECT: '',
        ENABLE_GROUP_MERGE: true,
        ENABLE_REPLY_TO: false,
        REPLY_TO: replyTo,
        REPLACE_VALUE: 'NA',
        MERGE_FIELD_MARKER: /[{]{2}([^{^}]+)[}]{2}/g,
        MERGE_FIELD_MARKER_TEXT: '[{]{2}([^{^}]+)[}]{2}',
        GROUP_FIELD_MARKER: /[[]{2}([^[^\]]+)[\]]{2}/g,
        GROUP_FIELD_MARKER_TEXT: '[[]{2}([^[^\\]]+)[\\]]{2}',
        ROW_INDEX_MARKER: '{{i}}',
        ENABLE_DEBUG_MODE: false,
        hostApp: hostApp,
        userLocale: 'ja',
      },
    },
    {
      eventName: `Settings all enabled on ${hostApp}`,
      addonEvent: {
        commonEventObject: {
          userLocale: 'ja',
          hostApp: hostApp,
          timeZone: { id: 'Asia/Dili', offset: 32400000 },
          platform: 'WEB',
          formInputs: {
            TEMPLATE_SUBJECT: {
              stringInputs: { value: ['test draft demo {{Name}}'] },
            },
            MERGE_FIELD_MARKER_TEXT: {
              stringInputs: { value: ['[{]{2}([^{^}]+)[}]{2}'] },
            },
            ENABLE_GROUP_MERGE: { stringInputs: { value: ['enabled'] } },
            ENABLE_REPLY_TO: { stringInputs: { value: ['enabled'] } },
            REPLACE_VALUE: { stringInputs: { value: ['NA'] } },
            ROW_INDEX_MARKER: { stringInputs: { value: ['{{i}}'] } },
            CC: { stringInputs: { value: ['{{cc}}'] } },
            GROUP_FIELD_MARKER_TEXT: {
              stringInputs: { value: ['[[]{2}([^[^\\]]+)[\\]]{2}'] },
            },
            TO: { stringInputs: { value: ['{{Email}}'] } },
            SPREADSHEET_URL: {
              stringInputs: {
                value: [
                  `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`,
                ],
              },
            },
            DATA_SHEET_NAME: { stringInputs: { value: [sheetName] } },
            REPLY_TO: { stringInputs: { value: [replyTo] } },
            ENABLE_DEBUG_MODE: { stringInputs: { value: ['enabled'] } },
            BCC: { stringInputs: { value: ['{{bcc}}'] } },
          },
        },
      },
      parsedConfig: {
        SPREADSHEET_URL: `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`,
        DATA_SHEET_NAME: sheetName,
        TO: '{{Email}}',
        CC: '{{cc}}',
        BCC: '{{bcc}}',
        TEMPLATE_SUBJECT: 'test draft demo {{Name}}',
        ENABLE_GROUP_MERGE: true,
        ENABLE_REPLY_TO: true,
        REPLY_TO: replyTo,
        REPLACE_VALUE: 'NA',
        MERGE_FIELD_MARKER: /[{]{2}([^{^}]+)[}]{2}/g,
        MERGE_FIELD_MARKER_TEXT: '[{]{2}([^{^}]+)[}]{2}',
        GROUP_FIELD_MARKER: /[[]{2}([^[^\]]+)[\]]{2}/g,
        GROUP_FIELD_MARKER_TEXT: '[[]{2}([^[^\\]]+)[\\]]{2}',
        ROW_INDEX_MARKER: '{{i}}',
        ENABLE_DEBUG_MODE: true,
        hostApp: hostApp,
        userLocale: 'ja',
      },
    },
    {
      eventName: `Settings all disabled on ${hostApp}`,
      addonEvent: {
        commonEventObject: {
          formInputs: {
            REPLY_TO: { stringInputs: { value: [replyTo] } },
            REPLACE_VALUE: { stringInputs: { value: ['NA'] } },
            ROW_INDEX_MARKER: { stringInputs: { value: ['{{i}}'] } },
            TEMPLATE_SUBJECT: {
              stringInputs: { value: ['test draft demo {{Name}}'] },
            },
            MERGE_FIELD_MARKER_TEXT: {
              stringInputs: { value: ['[{]{2}([^{^}]+)[}]{2}'] },
            },
            SPREADSHEET_URL: {
              stringInputs: {
                value: [
                  `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`,
                ],
              },
            },
            GROUP_FIELD_MARKER_TEXT: {
              stringInputs: { value: ['[[]{2}([^[^\\]]+)[\\]]{2}'] },
            },
            BCC: { stringInputs: { value: ['{{bcc}}'] } },
            TO: { stringInputs: { value: ['{{Email}}'] } },
            CC: { stringInputs: { value: ['{{cc}}'] } },
            DATA_SHEET_NAME: { stringInputs: { value: [sheetName] } },
          },
          userLocale: 'ja',
          platform: 'WEB',
          timeZone: { offset: 32400000, id: 'Asia/Dili' },
          hostApp: hostApp,
        },
      },
      parsedConfig: {
        SPREADSHEET_URL: `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`,
        DATA_SHEET_NAME: sheetName,
        TO: '{{Email}}',
        CC: '{{cc}}',
        BCC: '{{bcc}}',
        TEMPLATE_SUBJECT: 'test draft demo {{Name}}',
        ENABLE_GROUP_MERGE: false,
        ENABLE_REPLY_TO: false,
        REPLY_TO: replyTo,
        REPLACE_VALUE: 'NA',
        MERGE_FIELD_MARKER: /[{]{2}([^{^}]+)[}]{2}/g,
        MERGE_FIELD_MARKER_TEXT: '[{]{2}([^{^}]+)[}]{2}',
        GROUP_FIELD_MARKER: /[[]{2}([^[^\]]+)[\]]{2}/g,
        GROUP_FIELD_MARKER_TEXT: '[[]{2}([^[^\\]]+)[\\]]{2}',
        ROW_INDEX_MARKER: '{{i}}',
        ENABLE_DEBUG_MODE: false,
        hostApp: hostApp,
        userLocale: 'ja',
      },
    },
  ])
  .flat();
module.exports = { mockConfigs };
