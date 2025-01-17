/*
 * **Slack User Exporter**
 * このスクリプトでは、Slackのチームに所属するユーザを取得し、Google SpreadSheetに書き込むことができます。
 *
 * このスクリプトを実行するには、以下の手順を行ってください。
 * 1. このスクリプトをGoogle Apps Scriptにアップロードする
 * 2. スクリプトのプロパティに、以下の値を設定する
 *    - SLACK_API_TOKEN: SlackのAPIトークン
 *    - SLACK_CHANNEL_ID: 対象のチャンネルID
 *    - SPREAD_SHEET_ID: 出力先のスプレッドシートID
 * 3. スクリプトを実行する
 *
 * スクリプトを実行すると、指定されたSlackチャンネルに所属するユーザ情報がGoogle SpreadSheetに書き込まれます。
 * ユーザ情報には、SlackのユーザID、表示名から抽出したIDと名前が含まれます。
 */

/**
 * Slack APIとの通信を担当するクラス
 * 認証やリクエストの処理を一元管理します
 */
class SlackApi {
  /**
   * @param {string} token - Slack APIトークン
   * @throws {Error} トークンが空または文字列でない場合
   */
  constructor(token) {
    if (typeof token !== "string" || token.length === 0) {
      throw new Error("token must be a string which is NOT empty");
    }
    this.token = token;
    this.baseUrl = "https://slack.com/api";
  }

  /**
   * Slack APIにリクエストを送信する共通メソッド
   * @param {string} endpoint - APIエンドポイント（例: 'users.list'）
   * @param {string} method - HTTPメソッド
   * @param {Object} params - リクエストパラメータ
   * @returns {Promise<Object>} APIレスポンス
   * @throws {Error} APIリクエストが失敗した場合
   */
  async request(endpoint, method = "GET", params = {}) {
    let url = `${this.baseUrl}/${endpoint}`;

    // GETリクエストの場合、パラメータをURLに追加
    if (method === "GET" && Object.keys(params).length > 0) {
      const queryParams = Object.entries(params)
        .map(
          ([key, value]) =>
            `${encodeURIComponent(key)}=${encodeURIComponent(value)}`
        )
        .join("&");
      url = `${url}?${queryParams}`;
    }

    const response = await UrlFetchApp.fetch(url, {
      method: method,
      headers: {
        Authorization: `Bearer ${this.token}`,
        "Content-Type": "application/x-www-form-urlencoded",
      },
      ...(method !== "GET" && { payload: params }),
    });

    const result = JSON.parse(response.getContentText());
    if (!result.ok) {
      throw new Error(`Slack API Error: ${result.error}`);
    }
    return result;
  }

  /**
   * チーム全体のユーザ一覧を取得
   * @returns {Promise<Array>} ユーザオブジェクトの配列
   */
  async listUsersInTeam() {
    const response = await this.request("users.list", "GET", { limit: 1000 });
    return response.members;
  }

  /**
   * 特定のチャンネルに所属するユーザIDの一覧を取得
   * ページネーション対応済み
   * @param {string} channelId - チャンネルID
   * @returns {Promise<Array<string>>} ユーザIDの配列
   * @throws {Error} チャンネルIDが無効な場合
   */
  async listUsersInChannel(channelId) {
    if (typeof channelId !== "string" || channelId.length === 0) {
      throw new Error("channelId must be a string which is NOT empty");
    }

    const members = [];
    let nextCursor;

    do {
      const params = {
        channel: channelId,
        limit: 200,
        ...(nextCursor && { cursor: nextCursor }),
      };

      const response = await this.request(
        "conversations.members",
        "GET",
        params
      );
      members.push(...response.members);

      nextCursor = response.response_metadata?.next_cursor;
      if (nextCursor) {
        console.log(`次のページのユーザを取得中... (カーソル: ${nextCursor})`);
      }
    } while (nextCursor);

    return members;
  }

  /**
   * 特定のユーザの詳細情報を取得
   * @param {string} userId - ユーザID
   * @returns {Promise<Object>} ユーザ情報
   * @throws {Error} ユーザIDが無効な場合
   */
  async userInfo(userId) {
    if (typeof userId !== "string" || userId.length === 0) {
      throw new Error("userId must be a string which is NOT empty");
    }

    const response = await this.request("users.info", "GET", { user: userId });
    return response.user;
  }
}

/**
 * Google Spreadsheetとの操作を担当するクラス
 * シートの作成、データの読み書きを一元管理します
 */
class SpreadSheetApi {
  /**
   * @param {string} sheetId - SpreadsheetのID
   * @throws {Error} シートIDが無効な場合
   */
  constructor(sheetId) {
    if (typeof sheetId !== "string" || sheetId.length === 0) {
      throw new Error("sheetId must be a string which is NOT empty");
    }
    this.sheetId = sheetId;
    this.spreadsheet = SpreadsheetApp.openById(sheetId);
  }

  /**
   * シートを取得または作成
   * @param {string} sheetName - シート名
   * @returns {GoogleAppsScript.Spreadsheet.Sheet} シートオブジェクト
   */
  getOrCreateSheet(sheetName) {
    let sheet = this.spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      sheet = this.spreadsheet.insertSheet(sheetName);
    }
    return sheet;
  }

  /**
   * データの形式を検証
   * @param {Array<Array>} data - 2次元配列のデータ
   * @throws {Error} データ形式が不正な場合
   */
  validateData(data) {
    if (!Array.isArray(data) || !data.every((row) => Array.isArray(row))) {
      throw new Error("Data must be a 2D array");
    }

    const rowLength = data[0].length;
    if (!data.every((row) => row.length === rowLength)) {
      throw new Error("All rows must have the same length");
    }
  }

  /**
   * シートにデータを書き込む
   * @param {Array<Array>} data - 書き込むデータ（2次元配列）
   * @param {string} sheetName - シート名
   */
  async writeMapToSheet(data, sheetName) {
    this.validateData(data);

    const sheet = this.getOrCreateSheet(sheetName);
    sheet.clear();

    if (data.length === 0) {
      return;
    }

    // ヘッダー行を書き込み
    sheet.appendRow(data[0]);

    // データ行を書き込み
    if (data.length > 1) {
      sheet
        .getRange(2, 1, data.length - 1, data[0].length)
        .setValues(data.slice(1));
    }
  }

  /**
   * ユーザ情報をシートに書き込む
   * @param {Array<User>} users - ユーザオブジェクトの配列
   */
  async writeUsersToSheet(users) {
    if (!Array.isArray(users)) {
      throw new Error("users must be an array");
    }

    const rows = [
      ["userId", "id", "name"], // ヘッダー行
    ];

    for (const user of users) {
      if (!user.userId || !user.name?.id || !user.name?.name) {
        console.warn(
          `無効なユーザデータをスキップしました: ${JSON.stringify(user)}`
        );
        continue;
      }
      rows.push([user.userId, user.name.id, user.name.name]);
    }

    await this.writeMapToSheet(rows, "users");
  }

  /**
   * シートからユーザ情報を読み込む
   * @returns {Promise<Array<User>>} ユーザオブジェクトの配列
   */
  async readUsersFromSheet() {
    const sheet = this.getOrCreateSheet("users");
    const rows = sheet.getDataRange().getValues();

    if (rows.length <= 1) {
      // ヘッダーのみまたは空の場合
      return [];
    }

    // ヘッダー行をスキップ
    return rows.slice(1).map((row) => ({
      userId: row[0],
      name: { id: row[1], name: row[2] },
    }));
  }
}

/**
 * Slackユーザ情報を扱うクラス
 * ユーザ情報のパースと保持を担当
 */
class User {
  /**
   * @param {string} userId - SlackのユーザID
   * @param {Object} name - ユーザの表示名情報
   * @param {string} name.id - 表示名から抽出したID
   * @param {string} name.name - 表示名から抽出した名前
   */
  constructor(userId, name) {
    this.userId = userId;
    this.name = name;
  }

  /**
   * Slackのユーザオブジェクトからユーザ情報をパース
   * @param {Object} user - Slackのユーザオブジェクト
   * @returns {User} パースされたユーザ情報
   * @throws {Error} ユーザ情報が不正な場合
   */
  static parseFromSlackUser(user) {
    if (!user || !user.profile) {
      throw new Error("Invalid user object provided");
    }

    const displayName = user.profile.display_name;
    if (!displayName) {
      throw new Error(
        `displayName is not found: ${JSON.stringify(user.profile)}`
      );
    }

    const [id, name, ...others] = displayName.split(" ");
    if (!id || !name || others.length > 0) {
      throw new Error(`Invalid display name format: ${displayName}`);
    }

    return new User(user.id, { id, name });
  }

  /**
   * ユーザ情報をJSON形式に変換
   * @returns {Object} JSON形式のユーザ情報
   */
  toJSON() {
    return {
      userId: this.userId,
      name: this.name,
    };
  }
}

/**
 * ユーザ情報の取得とパースを行う
 * @param {SlackApi} slackApi - SlackAPIインスタンス
 * @param {string} channelId - チャンネルID
 * @returns {Promise<Array<User>>} パース済みのユーザ情報配列
 */
async function fetchAndParseUsers(slackApi, channelId) {
  const userIds = await slackApi.listUsersInChannel(channelId);
  console.log(`${userIds.length}人のユーザ情報を取得中...`);

  const users = [];
  const errors = [];

  for (const userId of userIds) {
    try {
      const userInfo = await slackApi.userInfo(userId);
      const parsedUser = User.parseFromSlackUser(userInfo);
      users.push(parsedUser);
    } catch (e) {
      errors.push({ userId, error: e.message });
      console.error(
        `ユーザ(${userId})の処理中にエラーが発生しました: ${e.message}`
      );
    }
  }

  if (errors.length > 0) {
    console.warn(`${errors.length}人のユーザの処理に失敗しました`);
  }

  return users;
}

/**
 * 新規ユーザと既存ユーザの情報をマージ
 * @param {SpreadSheetApi} sheetApi - SpreadsheetAPIインスタンス
 * @param {Array<User>} newUsers - 新規ユーザ情報
 * @returns {Promise<Array<User>>} マージ済みのユーザ情報
 */
async function mergeAndUpdateUsers(sheetApi, newUsers) {
  const existingUsers = await sheetApi.readUsersFromSheet();
  console.log(`既存のユーザ数: ${existingUsers.length}人`);

  const uniqueUsers = [...existingUsers, ...newUsers].reduce((acc, user) => {
    acc.set(user.userId, user);
    return acc;
  }, new Map());

  const mergedUsers = Array.from(uniqueUsers.values());
  await sheetApi.writeUsersToSheet(mergedUsers);
  console.log(`スプレッドシートを更新しました (合計: ${mergedUsers.length}人)`);

  return mergedUsers;
}

/**
 * メイン処理
 * 必要な環境変数を確認し、ユーザ情報の取得と保存を実行
 * @returns {Promise<Array<User>>} 処理済みのユーザ情報
 * @throws {Error} 必要な環境変数が不足している場合やその他のエラー
 */
async function main() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const requiredProps = [
      "SLACK_API_TOKEN",
      "SLACK_CHANNEL_ID",
      "SPREAD_SHEET_ID",
    ];

    for (const prop of requiredProps) {
      if (!properties.getProperty(prop)) {
        throw new Error(`必要な環境変数が設定されていません: ${prop}`);
      }
    }

    const slackApi = new SlackApi(properties.getProperty("SLACK_API_TOKEN"));
    const sheetApi = new SpreadSheetApi(
      properties.getProperty("SPREAD_SHEET_ID")
    );

    console.log("ユーザ情報のエクスポートを開始します...");

    const parsedUsers = await fetchAndParseUsers(
      slackApi,
      properties.getProperty("SLACK_CHANNEL_ID")
    );

    const finalUsers = await mergeAndUpdateUsers(sheetApi, parsedUsers);

    console.log("エクスポートが正常に完了しました");
    return finalUsers;
  } catch (error) {
    console.error("実行中に致命的なエラーが発生しました:", error.message);
    throw error;
  }
}
