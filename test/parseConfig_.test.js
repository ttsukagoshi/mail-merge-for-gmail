const mailMerge = require('../src/mailMerge');
const mockEvents = [
  {
    eventTitle: 'Default settings',
    expectedConfig: {
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
      MERGE_FIELD_MARKER: /\{\{([^}]+)\}\}/g,
      MERGE_FIELD_MARKER_TEXT: '\\{\\{([^\\}]+)\\}\\}',
      GROUP_FIELD_MARKER: /\[\[([^\]]+)\]\]/g,
      GROUP_FIELD_MARKER_TEXT: '\\[\\[([^\\]]+)\\]\\]',
      ROW_INDEX_MARKER: '{{i}}',
      ENABLE_DEBUG_MODE: false,
      hostApp: 'GMAIL',
      userLocale: 'ja',
    },
    event: {
      commonEventObject: {
        userLocale: 'ja',
        formInputs: {
          DATA_SHEET_NAME: { stringInputs: { value: ['demo sheet name 123'] } },
          GROUP_FIELD_MARKER_TEXT: {
            stringInputs: { value: ['\\[\\[([^\\]]+)\\]\\]'] },
          },
          MERGE_FIELD_MARKER_TEXT: {
            stringInputs: { value: ['\\{\\{([^\\}]+)\\}\\}'] },
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
  },
  {
    eventTitle: 'Settings all enabled',
    expectedConfig: {
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
      MERGE_FIELD_MARKER: /\{\{([^}]+)\}\}/g,
      MERGE_FIELD_MARKER_TEXT: '\\{\\{([^}]+)\\}\\}',
      GROUP_FIELD_MARKER: /\[\[([^\]]+)\]\]/g,
      GROUP_FIELD_MARKER_TEXT: '\\[\\[([^\\]]+)\\]\\]',
      ROW_INDEX_MARKER: '{{i}}',
      ENABLE_DEBUG_MODE: true,
      hostApp: 'GMAIL',
      userLocale: 'ja',
    },
    event: {
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
            stringInputs: { value: ['\\{\\{([^}]+)\\}\\}'] },
          },
          ENABLE_GROUP_MERGE: { stringInputs: { value: ['enabled'] } },
          ENABLE_REPLY_TO: { stringInputs: { value: ['enabled'] } },
          REPLACE_VALUE: { stringInputs: { value: ['NA'] } },
          ROW_INDEX_MARKER: { stringInputs: { value: ['{{i}}'] } },
          CC: { stringInputs: { value: ['{{cc}}'] } },
          GROUP_FIELD_MARKER_TEXT: {
            stringInputs: { value: ['\\[\\[([^\\]]+)\\]\\]'] },
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
  },
  {
    eventTitle: 'Settings all disabled',
    expectedConfig: {
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
      MERGE_FIELD_MARKER: /\{\{([^}]+)\}\}/g,
      MERGE_FIELD_MARKER_TEXT: '\\{\\{([^}]+)\\}\\}',
      GROUP_FIELD_MARKER: /\[\[([^\]]+)\]\]/g,
      GROUP_FIELD_MARKER_TEXT: '\\[\\[([^\\]]+)\\]\\]',
      ROW_INDEX_MARKER: '{{i}}',
      ENABLE_DEBUG_MODE: false,
      hostApp: 'GMAIL',
      userLocale: 'ja',
    },
    event: {
      commonEventObject: {
        formInputs: {
          REPLY_TO: { stringInputs: { value: ['replyTo@email.com'] } },
          REPLACE_VALUE: { stringInputs: { value: ['NA'] } },
          ROW_INDEX_MARKER: { stringInputs: { value: ['{{i}}'] } },
          TEMPLATE_SUBJECT: {
            stringInputs: { value: ['test draft demo {{Name}}'] },
          },
          MERGE_FIELD_MARKER_TEXT: {
            stringInputs: { value: ['\\{\\{([^}]+)\\}\\}'] },
          },
          SPREADSHEET_URL: {
            stringInputs: {
              value: [
                'https://docs.google.com/spreadsheets/d/qwerty12345/edit',
              ],
            },
          },
          GROUP_FIELD_MARKER_TEXT: {
            stringInputs: { value: ['\\[\\[([^\\]]+)\\]\\]'] },
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
  },
];

mockEvents.forEach((mockEvent) => {
  test(mockEvent.eventTitle, () => {
    expect(mailMerge.parseConfig_(mockEvent.event)).toEqual(
      mockEvent.expectedConfig
    );
  });
});
