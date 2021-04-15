# Mail Merge for Gmail (English / [日本語](https://github.com/ttsukagoshi/mail-merge-for-gmail/blob/main/README.ja.md))
[![GitHub Super-Linter](https://github.com/ttsukagoshi/mail-merge-for-gmail/workflows/Lint%20Code%20Base/badge.svg)](https://github.com/marketplace/actions/super-linter) [![Total alerts](https://img.shields.io/lgtm/alerts/g/ttsukagoshi/mail-merge-for-gmail.svg?logo=lgtm&logoWidth=18)](https://lgtm.com/projects/g/ttsukagoshi/mail-merge-for-gmail/alerts/)  
Send personalized emails based on Gmail template to multiple recipients using Google Sheets and Google Apps Script. The **Group Merge** feature, which allows the sender to group the contents of two or more rows into one row for a single recipient, is available.

> **This is a Legacy Version**  
> This latest version of the v1 series, a legacy version preceding the current Google Workspace Add-on versions, is no longer maintained. For latest information on the add-on, see [this repository's top page](https://github.com/ttsukagoshi/mail-merge-for-gmail).

## Overview
Similar to [the mail merge feature available in Microsoft Word](https://support.office.com/en-us/article/use-mail-merge-for-bulk-email-letters-labels-and-envelopes-f488ed5b-b849-4c11-9cff-932c49474705), this Mail Merge for Gmail allows Gmail/Google Workspace users to send personalized emails to the recipients listed in the spreadsheet. Some notable features are:  
- Use Gmail drafts as template for mail merge. HTML styling and file attachments are preserved in the personalized emails.
- The **Group Merge** feature available for combining two or more entries with the same recipient.

## How to Use
### 1. Prepare
Copy [the sample spreadsheet](https://docs.google.com/spreadsheets/d/1pVoKzoldYOaEXhbEmpsLJAZqmkB1IDncQ6rTXlbqETY/edit?usp=sharing) to your Google Drive by `File` -> `Make a copy`

### 2. Make your List
Edit your spreadsheet in any way you want to. If you want to change the sheet name `List`, make sure to change the value of `DATA_SHEET_NAME` in sheet `Config`. 

<blockquote>
<h4>Notes</h4>
<ul>
    <li>Keep the first row of the spreadsheet as a header, i.e., the first row should be the name of each column, and nothing else.</li>
    <li>The default value for the field (column) name of recipient email address is set to <code>Email</code>; change the value of <code>RECIPIENT_COL_NAME</code> in sheet <code>Config</code> to suit your needs</li>
    <li>Changing sheet name of <code>Config</code> is not recommended unless you are familiar with Google Apps Script and can edit the relevant section of the script.</li>
    <li>The lower-case letter <code>i</code> is reserved as part of the group merge function, as described below, and cannot be used for a column name.</li>
    <li>Line breaks within a spreadsheet cell, but not the text stylings like color and bold fonts, will be reflected in the merged mail.</li>
</ul>
</blockquote>

### 3. Create a template draft on Gmail
Create a Gmail draft to serve as the template. By default, the merge fields are specified by double curly brackets, i.e., `Dear {{Name}},... `. The field names should correspond with the column names of the spreadsheet (case-sensitive).

#### Properties Reflected on the Personalized Emails
The below lists the properties of the template (Gmail draft) that are reflected as-is in the personalized email(s):  
- Fixed CC and BCC recipients
- File attachments
- In-line image attachments (when HTML styling is enabled)
- Gmail labels

#### Group Merge
In a case where there are two or more entries in your list with the same recipient, you might want to group the entries into a single email rather than sending the recipient similar emails more than once. Group merge enables you to specify which field to list individually and which to combine in an email, as shown in [the example below](https://github.com/ttsukagoshi/mail-merge-for-gmail#example-of-group-merge).

The group merge field is, by default, marked by double square brackets, i.e., `[[Meeting ID: {{Meeting ID}}]]`. The merge fields (the curly brackets) nested inside this group merge field will be merged reclusively if there are two or more rows for the same recipient. A special index field `{{i}}` can be used inside the group merge field to indicate the index number within the group merge. To enable the group merge function, change the value of `ENABLE_GROUP_MERGE` to `true`.

<blockquote>
    <h4>Notes</h4>
    <ul>
        <li>The subject of the template Gmail draft must be unique. An error will be returned during the process of Step 4 below if there are two or more Gmail templates with the designated subject.</li>
        <li>You can use merge fields and group merge fields in the subject line, too.</li>
        <li>If an invalid field name (e.g., a field name that does not match the column names) is designated, the field is replace by <code>NA</code> or the text value you entered for <code>REPLACE_VALUE</code> in sheet <code>Config</code>.</li>
    </ul>
</blockquote>

### 4. Create Personalized Gmail Drafts or Send Merged Emails
From the spreadsheet menu `Mail Merge`, you can choose either to `Create Draft` or directly `Send Emails` based on the template. You will be prompted to enter the subject of the template mail that you created in the above Step 3.

<blockquote>
    <h4>Notes</h4>
    <ul>
        <li>Attachments in the draft template, including in-line images for HTML drafts, will be preserved in the merge process and sent to each recipient.</li>
        <li>If you chose <code>Create Draft</code>, you can send the created drafts by selecting <code>Send the Created Drafts</code>. This option will send <strong>only the drafts created by the latest <code>Create Draft</code></strong></li>
    </ul>
</blockquote>

## Example of Group Merge
Given a list of email addresses below:
|Email|Name|Meeting ID|Date|Start Time|End Time|
| --- | --- | --- | --- | --- | --- |
|`john@example.com`|John|00001|May 7, 2020|13:00|14:00|
|`mary@sample.com`|Mary|00002|May 7, 2020|14:30|15:30|
|`john@example.com`|John|00003|May 8, 2020|9:00|10:00|

and a Gmail draft with the below text as its body:
```
Dear {{Name}},

Thank you for your application.
Details of your meeting are as below:

[[
Meeting No. {{i}}
Date: {{Date}}
Time Slot: {{Start Time}} – {{End Time}}
Meeting ID: {{Meeting ID}}

]]

We look forward to seeing you!
```

The personlized emails using group merge will look like this:  
**Email to John:**
```
Dear John,

Thank you for your application.
Details of your meeting are as below:

Meeting No. 1
Date: May 7, 2020
Time Slot: 13:00 – 14:00
Meeting ID: 00001

Meeting No. 2
Date: May 8, 2020
Time Slot: 9:00 – 10:00
Meeting ID: 00003

We look forward to seeing you!
```
**Email to Mary:**
```
Dear Mary,

Thank you for your application.
Details of your meeting are as below:

Meeting No. 1
Date: May 7, 2020
Time Slot: 14:30 – 15:30
Meeting ID: 00002

We look forward to seeing you!
```

## Advanced Settings
### Field Markers (Placeholders)
- The markers for merge fields and group merge fields can be adjusted via the values `MERGE_FIELD_MARKER` and `GROUP_FIELD_MARKER`, respectively, in the sheet `Config`. You will need to be familiar with [the regular expressions of JavaScript](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Guide/Regular_Expressions).
- The index field marker for group merge `{{i}}` can also be modified through the value `ROW_INDEX_MARKER` in sheet `Config`.
- If HTML is enabled in your Gmail, make sure that your modified markers can still be detected in the HTML string.

### Reply-To Settings
- Reply-To values for each personalized mails can be set by switching `ENABLE_REPLY_TO` to `true` and entering the appropriate value for `REPLY_TO` in the sheet `Config` (`ENABLE_REPLY_TO` is set to `false` by default). 
- The `REPLY_TO` value can either be a fixed value like `contact@example.com` or a value with placeholder(s) like `{{replyTo}}@example.com`. In the latter case, the respective values of the field name `replyTo` will be set for each recipient.
- IMPORTANT: If you want to combine the Reply-To settings with `Create Drafts` rather than sending the personalized mails directly via `Send Emails`, note that [Reply-To settings will NOT be preserved when sending the emails via the `Send` button in the Gmail website](https://stackoverflow.com/questions/65878696/how-can-i-keep-the-reply-to-setting-in-gmail-drafts-created-by-gmailapp-createdr). You have to instead use the `Send the Created Drafts` to send the drafts.

## Terms and Conditions
You must agree to the [Terms and Conditions](https://ttsukagoshi.github.io/scriptable-assets/terms-and-conditions/) to use this solution from [the sample spreadsheet](https://docs.google.com/spreadsheets/d/1pVoKzoldYOaEXhbEmpsLJAZqmkB1IDncQ6rTXlbqETY/edit?usp=sharing).

## Acknowledgements
This work was inspired by [Tutorial: Simple Mail Merge (Google Apps Script Tutorial)](https://developers.google.com/apps-script/articles/mail_merge).

