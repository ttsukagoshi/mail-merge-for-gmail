# Group Merge: Gmail のための差し込みメール作成 ([English](https://github.com/ttsukagoshi/mail-merge-for-gmail) / 日本語)

[![Get this add-on from Google Workspace Marketplace](https://img.shields.io/badge/Google%20Workspace%20Add--on-Available-green?style=flat-square)](https://workspace.google.com/marketplace/app/group_merge_mail_merge_for_gmail/586770229603) [![clasp](https://img.shields.io/badge/built%20with-clasp-4285f4.svg?style=flat-square)](https://github.com/google/clasp) [![code style: prettier](https://img.shields.io/badge/code_style-prettier-ff69b4.svg?style=flat-square)](https://github.com/prettier/prettier)  
[![Accessibility-alt-text-bot](https://github.com/ttsukagoshi/mail-merge-for-gmail/actions/workflows/a11y_alt_text_bot.yml/badge.svg)](https://github.com/ttsukagoshi/mail-merge-for-gmail/actions/workflows/a11y_alt_text_bot.yml) [![CodeQL](https://github.com/ttsukagoshi/mail-merge-for-gmail/actions/workflows/codeql-analysis.yml/badge.svg)](https://github.com/ttsukagoshi/mail-merge-for-gmail/actions/workflows/codeql-analysis.yml) [![Deploy](https://github.com/ttsukagoshi/mail-merge-for-gmail/actions/workflows/deploy.yml/badge.svg)](https://github.com/ttsukagoshi/mail-merge-for-gmail/actions/workflows/deploy.yml) [![Labeler](https://github.com/ttsukagoshi/mail-merge-for-gmail/actions/workflows/label.yml/badge.svg)](https://github.com/ttsukagoshi/mail-merge-for-gmail/actions/workflows/label.yml) [![Lint Code Base](https://github.com/ttsukagoshi/mail-merge-for-gmail/actions/workflows/linter.yml/badge.svg)](https://github.com/ttsukagoshi/mail-merge-for-gmail/actions/workflows/linter.yml)

Gmail と Google スプレッドシートを使って、受信者一人ひとり向けに宛名などを差し込んで作成。宛先リストで、宛先（メールアドレス）に重複がある場合、内容を 1 通のメールにまとめて送信できる **「まとめ差し込み（Group Merge）」** 機能つき。

> **レガシー（スプレッドシート）版**  
> Google Workspace アドオンとして提供される前に公開していたレガシー版（スプレッドシート版）は現在、開発が停止しています。今後、Google 側での仕様変更等により使用できなくなる可能性もありますので、ご注意ください。ソースコードやサンプルスプレッドシートを確認したい方のために[GitHub 上のブランチ](https://github.com/ttsukagoshi/mail-merge-for-gmail/tree/legacy-v1.8.0-spreadsheet)を残していますので、必要に応じてご参照ください。

## 概要

![Group Mergeのアイコン](https://lh3.googleusercontent.com/pw/ACtC-3eZPKFkzQJvMs2P_HgJIwNRSy1OGklUpOr0gm9ncC3OGcJw-nVvNUDYta6mMWo3d57gEc9KD_KV-UNOJvcTCBjGru3MG1KUpzP3z15I-bjEfT3u1V12mzRQrcA89pzb_RoIbINO3B1WxT4qP0KefNs=s96-no)  
Gmail と Google スプレッドシートを使って、宛名や会議の日時などの情報を宛先ごとに個別化したメールを作成して送信するための、オープンソースの Google Workspace アドオン。

- 宛先リスト内に同じメールアドレスの宛先が 2 つ以上ある場合に、内容を 1 通のメールにまとめられる「**まとめ差し込み（Group Merge）**」機能つき。
- Gmail の下書きをテンプレートとして利用。**書式設定**（文字の色など）や**添付ファイル**、**CC**及び**BCC**、そして**Gmail ラベル**が差し込み後のメールにも引き継がれる。

## 使い方

本アドオンの詳細は、[ウェブサイト](https://www.scriptable-assets.page/ja/add-ons/group-merge/)をご覧ください。

## 利用規約

本アドオンの使用には、[利用規約（英）](https://www.scriptable-assets.page/ja/add-ons/group-merge/#%E5%88%A9%E7%94%A8%E8%A6%8F%E7%B4%84)への同意が必要です。

## 開発に参加

オープンソースである本アドオン開発はどなたでも参加いただけます。[Contributing Guideline](https://github.com/ttsukagoshi/mail-merge-for-gmail/blob/main/docs/CONTRIBUTING.md) をご一読の上、ご参加ください！

## アイコンの出典

The original icon of this add-on was made by [Freepik](https://www.freepik.com) from [www.flaticon.com](https://www.flaticon.com/) and its colors are modified to fit the theme color of the app.

## 謝辞

This work was inspired by [Tutorial: Simple Mail Merge (Google Apps Script Tutorial)](https://developers.google.com/apps-script/articles/mail_merge).
