# Slack User Exporter

## 概要

Slackのチームに所属するユーザを取得し、Google SpreadSheetに書き込むGAS(Google Apps Script)スクリプトです。

## 使い方

1. このスクリプトをGoogle Apps Scriptにアップロードする
1. スクリプトのプロパティに、Slack API TokenとGoogle SpreadSheetのIDを設定する
1. スクリプトを実行する

### Slack API Tokenの取得方法

1. <https://api.slack.com/apps/> にアクセスし、Slack Appを作成する
1. `OAuth & Permissions` タブを選択し、`Bot Token Scopes` に 以下のスコープを追加する
    - `users:read`
    - `users:read.email`
1. `OAuth & Permissions` タブの`Install App to [Workspace]` ボタンを押下し、、インストールする
1. インストールが完了したら、`OAuth & Permissions` タブの`Bot User OAuth Token` をコピーする
    - `xoxb-` から始まります
1. このトークンをスクリプトプロパティの`SLACK_API_TOKEN`に設定します

### Google SpreadSheetのIDの取得方法

1. GASファイルの作成に利用したGoogle SpreadSheetを開く
![image](https://github.com/user-attachments/assets/66971cad-5ed5-40ac-8b4d-cee36dbe99aa)
1. ドキュメントのURLをコピーする
    - `https://docs.google.com/spreadsheets/d/1234567890/edit#gid=0` の `1234567890` の部分がIDです
1. このIDをスクリプトプロパティの`GOOGLE_SPREADSHEET_ID`に設定します

最終的に以下のようになります
![image](https://github.com/user-attachments/assets/8c3ede51-b0ed-476d-9079-7196c9f848b8)

## 注意点

- スクリプトを実行すると、Google SpreadSheetの既存のデータが削除されます。
- 1000人以上の場合はページネーションを実装する必要があります。
