const mockConfigs = [
  {
    eventName: 'Default settings',
    addonEvent: {
      commonEventObject: {
        userLocale: 'ja',
        formInputs: {
          DATA_SHEET_NAME: { stringInputs: { value: ['demo sheet name 123'] } },
          GROUP_FIELD_MARKER_TEXT: {
            stringInputs: { value: ['[[]{2}([^[^\\]]+)[\\]]{2}'] },
          },
          MERGE_FIELD_MARKER_TEXT: {
            stringInputs: { value: ['[{]{2}([^{^}]+)[}]{2}'] },
          },
          REPLY_TO: { stringInputs: { value: ['replyTo@email.com'] } },
          REPLACE_VALUE: { stringInputs: { value: ['NA'] } },
          TO: { stringInputs: { value: ['{{Email}}'] } },
          ENABLE_GROUP_MERGE: { stringInputs: { value: ['enabled'] } },
          SPREADSHEET_URL: {
            stringInputs: {
              value: [
                'https://docs.google.com/spreadsheets/d/qwerty12345/edit',
              ],
            },
          },
          ROW_INDEX_MARKER: { stringInputs: { value: ['{{i}}'] } },
        },
        timeZone: { offset: 32400000, id: 'Asia/Dili' },
        hostApp: 'GMAIL',
        platform: 'WEB',
      },
    },
    parsedConfig: {
      SPREADSHEET_URL:
        'https://docs.google.com/spreadsheets/d/qwerty12345/edit',
      DATA_SHEET_NAME: 'demo sheet name 123',
      TO: '{{Email}}',
      CC: '',
      BCC: '',
      TEMPLATE_SUBJECT: '',
      ENABLE_GROUP_MERGE: true,
      ENABLE_REPLY_TO: false,
      REPLY_TO: 'replyTo@email.com',
      REPLACE_VALUE: 'NA',
      MERGE_FIELD_MARKER: /[{]{2}([^{^}]+)[}]{2}/g,
      MERGE_FIELD_MARKER_TEXT: '[{]{2}([^{^}]+)[}]{2}',
      GROUP_FIELD_MARKER: /[[]{2}([^[^\]]+)[\]]{2}/g,
      GROUP_FIELD_MARKER_TEXT: '[[]{2}([^[^\\]]+)[\\]]{2}',
      ROW_INDEX_MARKER: '{{i}}',
      ENABLE_DEBUG_MODE: false,
      hostApp: 'GMAIL',
      userLocale: 'ja',
    },
  },
  {
    eventName: 'Settings all enabled',
    addonEvent: {
      commonEventObject: {
        userLocale: 'ja',
        hostApp: 'GMAIL',
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
                'https://docs.google.com/spreadsheets/d/qwerty12345/edit',
              ],
            },
          },
          DATA_SHEET_NAME: { stringInputs: { value: ['demo sheet name 123'] } },
          REPLY_TO: { stringInputs: { value: ['replyTo@email.com'] } },
          ENABLE_DEBUG_MODE: { stringInputs: { value: ['enabled'] } },
          BCC: { stringInputs: { value: ['{{bcc}}'] } },
        },
      },
    },
    parsedConfig: {
      SPREADSHEET_URL:
        'https://docs.google.com/spreadsheets/d/qwerty12345/edit',
      DATA_SHEET_NAME: 'demo sheet name 123',
      TO: '{{Email}}',
      CC: '{{cc}}',
      BCC: '{{bcc}}',
      TEMPLATE_SUBJECT: 'test draft demo {{Name}}',
      ENABLE_GROUP_MERGE: true,
      ENABLE_REPLY_TO: true,
      REPLY_TO: 'replyTo@email.com',
      REPLACE_VALUE: 'NA',
      MERGE_FIELD_MARKER: /[{]{2}([^{^}]+)[}]{2}/g,
      MERGE_FIELD_MARKER_TEXT: '[{]{2}([^{^}]+)[}]{2}',
      GROUP_FIELD_MARKER: /[[]{2}([^[^\]]+)[\]]{2}/g,
      GROUP_FIELD_MARKER_TEXT: '[[]{2}([^[^\\]]+)[\\]]{2}',
      ROW_INDEX_MARKER: '{{i}}',
      ENABLE_DEBUG_MODE: true,
      hostApp: 'GMAIL',
      userLocale: 'ja',
    },
  },
  {
    eventName: 'Settings all disabled',
    addonEvent: {
      commonEventObject: {
        formInputs: {
          REPLY_TO: { stringInputs: { value: ['replyTo@email.com'] } },
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
                'https://docs.google.com/spreadsheets/d/qwerty12345/edit',
              ],
            },
          },
          GROUP_FIELD_MARKER_TEXT: {
            stringInputs: { value: ['[[]{2}([^[^\\]]+)[\\]]{2}'] },
          },
          BCC: { stringInputs: { value: ['{{bcc}}'] } },
          TO: { stringInputs: { value: ['{{Email}}'] } },
          CC: { stringInputs: { value: ['{{cc}}'] } },
          DATA_SHEET_NAME: { stringInputs: { value: ['demo sheet name 123'] } },
        },
        userLocale: 'ja',
        platform: 'WEB',
        timeZone: { offset: 32400000, id: 'Asia/Dili' },
        hostApp: 'GMAIL',
      },
    },
    parsedConfig: {
      SPREADSHEET_URL:
        'https://docs.google.com/spreadsheets/d/qwerty12345/edit',
      DATA_SHEET_NAME: 'demo sheet name 123',
      TO: '{{Email}}',
      CC: '{{cc}}',
      BCC: '{{bcc}}',
      TEMPLATE_SUBJECT: 'test draft demo {{Name}}',
      ENABLE_GROUP_MERGE: false,
      ENABLE_REPLY_TO: false,
      REPLY_TO: 'replyTo@email.com',
      REPLACE_VALUE: 'NA',
      MERGE_FIELD_MARKER: /[{]{2}([^{^}]+)[}]{2}/g,
      MERGE_FIELD_MARKER_TEXT: '[{]{2}([^{^}]+)[}]{2}',
      GROUP_FIELD_MARKER: /[[]{2}([^[^\]]+)[\]]{2}/g,
      GROUP_FIELD_MARKER_TEXT: '[[]{2}([^[^\\]]+)[\\]]{2}',
      ROW_INDEX_MARKER: '{{i}}',
      ENABLE_DEBUG_MODE: false,
      hostApp: 'GMAIL',
      userLocale: 'ja',
    },
  },
];
module.exports = { mockConfigs };
