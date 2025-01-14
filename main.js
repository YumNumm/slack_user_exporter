/* **Slack User Exporter**
 * このスクリプトでは、Slackのチームに所属するユーザを取得し、Google SpreadSheetに書き込むことができます。
 *
 * このスクリプトを実行するには、以下の手順を行ってください。
 * 1. このスクリプトをGoogle Apps Scriptにアップロードする
 * 2. スクリプトのプロパティに、Slack API TokenとGoogle SpreadSheetのIDを設定する
 * 3. スクリプトを実行する
 *
 * スクリプトを実行すると、Slackのチームに所属するユーザがGoogle SpreadSheetに書き込まれます。
 */


class SlackApi {
  constructor(token) {
    this.token = token;
    // tokenは空ではない文字列 である必要がある
    if (typeof token !== "string" || token.length === 0) {
      throw new Error("token must be a string which is NOT empty");
    }
  }

  /// Teamに所属するユーザを取得
  /// cf. User Object https://api.slack.com/types/user
  /// User Objectの配列を返します
  async listUsers() {
    // Slack APIのusers.list methodを利用して チームのユーザを取得します
    // cf. https://api.slack.com/methods/users.list
    // 1000人以上の場合はページネーションを実装する必要があります

    // HTTP APIのリクエスト送信については、UrlFetchAppを使用します
    // cf. https://developers.google.com/apps-script/reference/url-fetch/url-fetch-app?hl=ja
    const response = await UrlFetchApp.fetch(
      "https://slack.com/api/users.list?limit=1000",
      {
        headers: {
          Authorization: `Bearer ${this.token}`,
        },
      }
    );
    const data = JSON.parse(response.getContentText());
    if (!data.ok) {
      throw new Error(
        `Error while fetching users from Slack API: ${data.error}`
      );
    }
    return data.members;
  }
}

/// SpreadSheetの操作を行うクラス
/// `https://docs.google.com/spreadsheets/d/***/edit` の *** の部分を sheetId としてプロパティに指定してください
class SpreadSheetApi {
  constructor(sheetId) {
    this.sheetId = sheetId;
    // sheetIdは空ではない文字列 である必要がある
    if (typeof sheetId !== "string" || sheetId.length === 0) {
      throw new Error("sheetId must be a string which is NOT empty");
    }
  }

  async writeUsersToSheet(users) {
    // users変数は配列 かつ その要素はObjectである必要がある
    if (
      !Array.isArray(users) ||
      !users.every((user) => typeof user === "object")
    ) {
      throw new Error("users must be an array of User Object");
    }

    const spreadsheet = SpreadsheetApp.openById(this.sheetId);
    // 0番目のSheet
    const sheet = spreadsheet.getSheets()[0];
    // 既存のデータを削除
    sheet.clear();
    // ヘッダー行を作成
    const headerRow = ["id", "name", "real_name", "display_name", "email"];
    sheet.appendRow(headerRow);
    // ユーザを追加
    const rows = [];
    for (const user of users) {
      rows.push([
        user.id,
        user.name,
        user.profile.real_name,
        user.profile.display_name,
        user.profile.email,
      ]);
    }
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
}

async function main() {
  const properties = PropertiesService.getScriptProperties();
  const slackApi = new SlackApi(properties.getProperty("SLACK_API_TOKEN"));
  console.log("Slack API initialized. Fetching users...");
  const users = await slackApi.listUsers();
  console.log(`${users.length} users fetched from Slack API`);
  const sheetApi = new SpreadSheetApi(
    properties.getProperty("SPREAD_SHEET_ID")
  );
  await sheetApi.writeUsersToSheet(users);
  console.log("Users written to SpreadSheet");
}
