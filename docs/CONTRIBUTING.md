# Contributing Guideline for Group Merge

## 🎉 Welcome!

First of all, thank you for taking interest in Group Merge! We would love to have you onboard this open-sourced project. This guideline serves as the documentation for anyone who has the spirit of making Group Merge a better add-on.

## 💬 Concept of Group Merge

Group Merge is designed to make **basic** mail merge operations on Gmail easy-to-use, open, and free for every Gmail users. As such, features focusing on commercial use of mail merge, like tracking link clicks for each emails sent, are considered excessive and will not be target of additional development or pull request reviews.

## 📝 Contributing

Contributing isn't just about writing codes. It could be from reporting a bug to coding a new feature on your own, or simply making suggestions for improving the add-on. Just make sure you adhere to our [Code of Conduct](https://github.com/ttsukagoshi/mail-merge-for-gmail/blob/main/docs/CODE_OF_CONDUCT.md).

### Open an Issue

To report a bug or request a new feature, [create a new issue](https://github.com/ttsukagoshi/mail-merge-for-gmail/issues/new/choose). If you have a question you want answered, please consider posting them on the [Discussions](https://github.com/ttsukagoshi/mail-merge-for-gmail/discussions) page instead, where we have the [Q & A](https://github.com/ttsukagoshi/mail-merge-for-gmail/discussions/categories/q-a) category specifically for that purpose.

Please take a look at the existing open issues & discussions before actually posting them to avoid duplications.

### Code on your own. Make pull-requests.

Take a look at [firstcontributions/first-contributions: 🚀✨ Help beginners to contribute to open source projects](https://github.com/firstcontributions/first-contributions) for the basics of contributing code to this project. We recommend creating an issue beforehand and referencing it in your pull request to keep the review process simple.

#### Coding Style

- Install [Node.js](https://nodejs.org/)
- We use [clasp](https://github.com/google/clasp), [ESLint](https://eslint.org/), and [Prettier](https://prettier.io/) to code Group Merge, and [Jest](https://jestjs.io/) to conduct unit testing on the code. Once you've forked and cloned this repository, execute `npm ci` at the cloned directory.
- [Visual Studio Code](https://code.visualstudio.com/) is the preferred editor.

#### \[Help Needed\] Internationalization/Localization

Want to contribute, but not sure where to start with? Localization could be a great starter! Group Merge currently supports the following languages:

- English (US)
- Japanese

Adding other languages to this list is always welcome. Please follow the steps below to make Group Merge more inclusive:

1. [Create a new issue](https://github.com/ttsukagoshi/mail-merge-for-gmail/issues/new?assignees=ttsukagoshi&labels=enhancement&template=feature_request.md&title=%5BNEW+FEAT%5D) to specify the language you will be adding, and to let everyone in the community know that you will be working on it. The issue title should be something like `[NEW FEAT] Add Spanish language`. If there is an existing issue for the language of your choice, avoid creating a duplicate issue and see if you can collaborate with the user who created the original issue.
2. Fork the repository and start working on the translation. All UI messages are defined in `../src/group-merge.js` under the const `MESSAGES`:

```javascript
const MESSAGES = {
  en: {
    '<messageName>': '<messageString>',
  },
  ja: {...}
};
```

The keys directly under MESSAGES correspond with the user locale value in the Google Workspace Add-on event object.

> `event.commonEventObject.userLocale`: The user's language and country/region identifier in the format of [ISO 639](https://en.wikipedia.org/wiki/ISO_639_macrolanguage) language code-[ISO 3166](https://en.wikipedia.org/wiki/ISO_3166) country/region code. For example, `en-US`.
>
> Source: https://developers.google.com/apps-script/add-ons/concepts/event-objects#common_event_object

Contrary to the description in the official document, it seems `en-US` is expressed simply as `en`, whereas English (United Kingdom) is expressed `en-GB`. Moreover, as in the above example `ja`, it seems languages that do not have country/region variations are expressed simply by the two-letter, lowercase ISO 639 language code. Add the locale code of your choice, copy the values from `en`, and translate the messages.

Note that some messages have placeholders like `{{myEmail}}`. Translate such messages so that these placeholders are kept in context.

## 📗 References

- [Code of Conduct](https://github.com/ttsukagoshi/mail-merge-for-gmail/blob/main/docs/CODE_OF_CONDUCT.md)
