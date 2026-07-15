// ===============================================================
// ★★★ 基本設定 ★★★
// このセクションの値を、ご自身の環境に合わせて設定してください。
// ===============================================================

/** @constant {Properties} スクリプトプロパティのインスタンス */
const PROPS_ = PropertiesService.getScriptProperties();

/** @constant {string} スプレッドシートのID */
const MY_SPREADSHEET_ID = PROPS_.getProperty('SPREADSHEET_ID');

/** @constant {string} Gemini APIキー */
const GEMINI_API_KEY = PROPS_.getProperty('GEMINI_API_KEY');

// --- シート名設定 ---
const MY_SHEET_NAME = 'デイサービス連絡版';
const USER_SHEET_NAME = 'ユーザー';
const HISTORY_SHEET_NAME = 'チェックイン履歴';
const SCHEDULE_SHEET_NAME = '週間スケジュール';
const CHARACTER_SHEET_NAME = 'キャラクターリスト';
const ACTIVITY_LOG_SHEET_NAME = '操作ログ';
const ERROR_LOG_SHEET_NAME = 'エラーログ';
const AI_CONFIG_SHEET_NAME = 'AI設定';

/** @constant {string} QRコード画像を保存するGoogleドライブのフォルダID */
const QR_CODE_FOLDER_ID = PROPS_.getProperty('QR_CODE_FOLDER_ID');

/** @constant {string} (任意) QRコード読み取りエラー時に再生する音声ファイルのID */
const ERROR_SOUND_ID = PROPS_.getProperty('ERROR_SOUND_ID');

// AIプロンプトのセル直接指定(MASTER_PROMPT_CELL)は廃止し、動的検索に変更しました

// ===============================================================
// ★★★ 詳細設定（通常は変更不要） ★★★
// ===============================================================

/**
 * QRコードIDの接頭辞を取得（スクリプトプロパティから）
 */
function getQrIdPrefix_() {
  return PROPS_.getProperty('QR_ID_PREFIX') || 'R07DS';
}

/** @constant {string} サーバーキャッシュのキー */
const CACHE_KEY = 'serviceDataCache';

/** @constant {number} キャッシュの有効期間（秒） */
const CACHE_EXPIRATION_SECONDS = 300; // 5分

/** @constant {string} スクリプトのバージョン情報 */
const SCRIPT_VERSION = 'v9.5.Pro';

// ===============================================================
// ★★★ メインのサーバー処理 (doGet) ★★★
// ===============================================================

/**
 * ウェブアプリケーションのメインエントリーポイント
 * @param {object} e - URLパラメータなどを含むイベントオブジェクト
 * @returns {HtmlOutput} 生成されたHTMLページ
 */

// ===============================================================
// ★★★ 認証・認可モジュール (Auth) ★★★
// ===============================================================
const Auth = {
  /**
   * 現在のユーザーの権限情報を取得する（キャッシュ対応）
   * @returns {object} { email: string, role: string|null, isAllowed: boolean, authMethod: string }
   */
  getUserContext: function () {
    try {
      const email = Session.getActiveUser().getEmail().toLowerCase();
      if (!email) return { email: '', role: null, isAllowed: false, authMethod: 'Unknown' };

      // 1. キャッシュチェック (UserCacheはユーザーごとに独立)
      const cache = CacheService.getUserCache();
      const cachedData = cache.get('user_context');

      if (cachedData) {
        return JSON.parse(cachedData);
      }

      // 2. スプレッドシートから判定 (キャッシュミス時)
      const user = this._resolveUserFromSheet(email);

      if (user.isAllowed) {
        // キャッシュに保存 (20分間有効)
        cache.put('user_context', JSON.stringify(user), 1200);
      }

      return user;
    } catch (e) {
      console.error('Auth Error:', e);
      return { email: '', role: null, isAllowed: false, authMethod: 'Error' };
    }
  },

  /**
   * 内部関数: シートとグループを確認して権限を解決する
   */
  _resolveUserFromSheet: function (userEmail) {
    const ss = SpreadsheetApp.openById(MY_SPREADSHEET_ID);
    const userSheet = ss.getSheetByName(USER_SHEET_NAME);
    if (!userSheet) return { email: userEmail, role: null, isAllowed: false };

    const data = userSheet.getDataRange().getValues();
    const headers = mapHeaders_(userSheet);

    // ヘッダー行を除外
    const rows = data.slice(1);

    // 1. データを権限・種別ごとに分類（シートのパースは1回のみ）
    const categorized = {
      admin: { users: new Set(), groups: [] },
      teacher: { users: new Set(), groups: [] },
      general: { users: new Set(), groups: [] }
    };

    for (const row of rows) {
      const email = String(row[headers['email']]).trim().toLowerCase();
      if (!email) continue;

      const type = String(row[headers['type']] || 'user').trim().toLowerCase();
      let role = String(row[headers['role']] || 'general').trim().toLowerCase();

      // 予期せぬroleに対するフォールバック
      if (!categorized[role]) role = 'general';

      if (type === 'group' && email.includes('@')) {
        categorized[role].groups.push(email);
      } else {
        // 個人メールは検索が高速なSet(O(1))に格納
        categorized[role].users.add(email);
      }
    }

    // 2. まず「個人メール」の最高権限をチェック（外部通信なし、一瞬で完了）
    if (categorized['admin'].users.has(userEmail)) {
      // 最強権限のadminなら即時確定（最も多いケース）
      return { email: userEmail, role: 'admin', isAllowed: true, authMethod: 'Email Match' };
    }

    let bestMatch = null;
    if (categorized['teacher'].users.has(userEmail)) {
      bestMatch = { email: userEmail, role: 'teacher', isAllowed: true, authMethod: 'Email Match' };
    } else if (categorized['general'].users.has(userEmail)) {
      bestMatch = { email: userEmail, role: 'general', isAllowed: true, authMethod: 'Email Match' };
    }

    // 3. グループのチェック（重い処理）
    // 既に個人で持っている権限より「強い」権限のグループのみチェックする
    const roleWeight = { 'admin': 3, 'teacher': 2, 'general': 1 };
    const currentWeight = bestMatch ? roleWeight[bestMatch.role] : 0;

    const rolesToCheck = ['admin', 'teacher', 'general'];
    for (const role of rolesToCheck) {
      // 評価中の権限が、既に持っている個人権限以下なら、これ以上探す意味がないため即時終了
      if (roleWeight[role] <= currentWeight) {
        return bestMatch;
      }

      for (const groupEmail of categorized[role].groups) {
        if (isUserInGroup_(userEmail, groupEmail)) {
          return { email: userEmail, role: role, isAllowed: true, authMethod: `Group Match (${groupEmail})` };
        }
      }
    }

    // 4. どの条件にも合致しなかった場合は弾く
    return bestMatch || { email: userEmail, role: null, isAllowed: false, authMethod: 'None' };
  },

  /** ガード句: 管理者(admin)権限が必要 */
  assertAdmin: function () {
    const context = this.getUserContext();
    if (!context.isAllowed || context.role !== 'admin') {
      throw new Error(`権限エラー: この操作には管理者権限が必要です。(User: ${context.email})`);
    }
    return context;
  },

  /** ガード句: 教職員(teacher)または管理者(admin)権限が必要 */
  assertTeacherOrAdmin: function () {
    const context = this.getUserContext();
    const allowedRoles = ['admin', 'teacher'];
    if (!context.isAllowed || !allowedRoles.includes(context.role)) {
      throw new Error(`権限エラー: この操作を行う権限がありません。(User: ${context.email})`);
    }
    return context;
  },

  /** ガード句: ログイン済みユーザーなら誰でもOK */
  assertLogin: function () {
    const context = this.getUserContext();
    if (!context.isAllowed) {
      throw new Error('権限エラー: ユーザー登録されていません。');
    }
    return context;
  }
};

// ===============================================================
// ★★★ 外部スキャナー連携 (doPost) ★★★
// GitHub Pages上のscanner.htmlから直接呼び出されるエンドポイント。
// 認証は「当日限定キー」（日付＋秘密ソルトのSHA-256ハッシュ）で行う。
// ===============================================================

/**
 * 当日限定キーを生成する（内部関数）
 * スクリプトプロパティ DAILY_KEY_SALT + 今日の日付（JST）をSHA-256ハッシュ化した
 * 24文字の16進数文字列を返す。キーは日付が変わると自動的に無効になる。
 * @returns {string} 当日限定キー
 */
function generateDailyKey_() {
  const salt = PROPS_.getProperty('DAILY_KEY_SALT');
  if (!salt) throw new Error('スクリプトプロパティ DAILY_KEY_SALT が設定されていません。');
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  const raw = today + ':' + salt;
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    raw,
    Utilities.Charset.UTF_8
  );
  return digest.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('').substring(0, 24);
}

/**
 * クライアント（Index.html）から呼び出して当日キーを取得する
 * 認証済みユーザー（Teacher以上）のみ利用可能。
 * @returns {string} 当日限定キー
 */
function getDailyKey() {
  Auth.assertTeacherOrAdmin();
  return generateDailyKey_();
}

/**
 * 外部スキャナーからのHTTPポストを受け付けてチェックイン処理を行う
 * @param {object} e - POSTリクエストのイベントオブジェクト
 * @returns {TextOutput} JSON形式の処理結果
 */
function doPost(e) {
  const headers = {
    'Access-Control-Allow-Origin': 'https://mikage-mmf.github.io',
    'Access-Control-Allow-Methods': 'POST',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Content-Type': 'application/json'
  };

  try {
    // ① 営業時間チェック（JST 7:00〜19:00 のみ受け付ける）
    const nowHour = parseInt(Utilities.formatDate(new Date(), 'Asia/Tokyo', 'HH'), 10);
    if (nowHour < 7 || nowHour >= 19) {
      logError({ functionName: 'doPost', message: `営業時間外アクセス（${nowHour}時）`, type: 'Security' });
      return buildJsonResponse_({ success: false, message: '営業時間外のアクセスです。' }, headers);
    }

    // ② リクエストボディのパース
    let payload;
    try {
      payload = JSON.parse(e.postData.contents);
    } catch (parseErr) {
      return buildJsonResponse_({ success: false, message: 'リクエストの形式が正しくありません。' }, headers);
    }

    const { apiKey, qrCodeId } = payload;

    // ③ 当日限定キー認証
    let validKey;
    try {
      validKey = generateDailyKey_();
    } catch (keyErr) {
      logError({ functionName: 'doPost', message: keyErr.message, type: 'Config' });
      return buildJsonResponse_({ success: false, message: 'サーバー設定エラーです。管理者に連絡してください。' }, headers);
    }

    if (!apiKey || apiKey !== validKey) {
      logError({ functionName: 'doPost', message: '当日キー不一致による不正アクセス', type: 'Security' });
      return buildJsonResponse_({ success: false, message: '認証に失敗しました。ページを再読み込みしてください。' }, headers);
    }

    // ④ QRコードIDの検証
    if (!qrCodeId || typeof qrCodeId !== 'string' || qrCodeId.trim() === '') {
      return buildJsonResponse_({ success: false, message: 'QRコードIDが指定されていません。' }, headers);
    }

    // ⑤ チェックイン処理（内部関数を直接呼び出し。Auth不要）
    const result = updateStatusInSheet(qrCodeId.trim(), '到着');

    // ⑥ 到着処理成功時は遅延メモもクリア
    if (result.success) {
      try {
        const ss = SpreadsheetApp.openById(MY_SPREADSHEET_ID);
        const sheet = ss.getSheetByName(MY_SHEET_NAME);
        const rowNum = findRowByQrCodeId_(sheet, qrCodeId.trim());
        if (rowNum !== -1) {
          const hdrs = mapHeaders_(sheet);
          if (hdrs['遅延連絡'] !== undefined) {
            sheet.getRange(rowNum, hdrs['遅延連絡'] + 1).setValue('');
          }
        }
      } catch (cleanupErr) {
        // 遅延連絡クリアは非致命的。ログだけ残す
        console.warn('doPost: 遅延連絡クリア中にエラー:', cleanupErr.message);
      }
    }

    return buildJsonResponse_(result, headers);

  } catch (err) {
    logError({ functionName: 'doPost', message: err.message, stack: err.stack, type: 'Server' });
    return buildJsonResponse_({ success: false, message: 'サーバーエラーが発生しました。' }, headers);
  }
}

/**
 * JSON形式のレスポンスを生成するヘルパー
 * @param {object} data - レスポンスデータ
 * @param {object} headers - レスポンスヘッダー
 * @returns {TextOutput}
 */
function buildJsonResponse_(data, headers) {
  const output = ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  return output;
}

function doGet(e) {

  // カメラの独立テストルーティング
  if (e && e.parameter && e.parameter.page === 'test_camera') {
    return HtmlService.createTemplateFromFile('TestCamera').evaluate()
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .setTitle("Camera Test");
  }

  // Authモジュールを使ってコンテキストを取得
  const userContext = Auth.getUserContext();

  // --- 認証失敗時の表示 ---
  if (!userContext.isAllowed) {
    const currentUser = Session.getActiveUser().getEmail() || '未ログイン';
    return HtmlService.createHtmlOutput(
      `<h1>アクセス権限がありません</h1>
       <p>現在のアカウント: <b>${currentUser}</b></p>
       <p>ユーザーリストに登録されていないか、権限がありません。<br>
       管理者に連絡して、アカウントの登録を依頼してください。</p>`
    );
  }

  const page = e.parameter.page || 'index';
  const userRole = userContext.role;

  // 各ページにアクセスするために必要な役割定義
  const pageAccessRoles = {
    'index': ['admin'],         // QRスキャンページ
    'kiosk': ['admin', 'teacher'],
    'kioskarrived': ['admin', 'teacher'],
    'viewer': ['admin', 'teacher'],
    'admin': ['admin'],         // 管理ページ
    'report': ['admin', 'teacher'],
    'card': ['admin'],
    'bulkcard': ['admin'],
    'businesscard': ['admin'],
    'help': ['admin', 'teacher']
  };

  const requiredRoles = pageAccessRoles[page];

  if (requiredRoles && !requiredRoles.includes(userRole)) {
    return HtmlService.createHtmlOutput(
      `<h1>アクセス許可エラー</h1>
       <p>あなたのアカウント (${userContext.email}) の役割 [${userRole}] では、<br>
       ページ "${page}" を表示する権限がありません。</p>`
    );
  }

  // --- ページ生成 ---
  let template;
  let title;
  const pageMap = {
    'index': { file: 'Index', title: 'デイサービス チェックイン' },
    'viewer': { file: 'Viewer', title: 'デイサービス 到着状況' },
    'admin': { file: 'Admin', title: '管理者用ページ' },
    'kiosk': { file: 'Kiosk', title: '到着状況ボード' },
    'kiosk_arrived': { file: 'KioskArrived', title: '到着状況ボード（小）' },
    'report': { file: 'Report', title: '到着傾向レポート' },
    'card': { file: 'Card', title: '事業所カード' },
    'bulkcard': { file: 'BulkCard', title: 'カード一括印刷' },
    'businesscard': { file: 'BusinessCard', title: '名刺サイズ一括印刷' },
    'help': { file: 'Help', title: 'ヘルプ' },
  };

  const pageInfo = pageMap[page] || pageMap['index'];
  template = HtmlService.createTemplateFromFile(pageInfo.file);
  title = pageInfo.title;

  // 複数IDを受け取る処理
  if (e.parameter.ids) {
    const ids = e.parameter.ids.split(',');
    template.serviceCenters = getServiceCenterDetailsByIds_(ids);
  } else if (page === 'card') {
    template.qrCodeId = e.parameter.id;
  }

  template.currentPage = page;
  template.userRole = userRole;
  template.userEmail = userContext.email;
  template.authMethod = userContext.authMethod;
  template.scriptVersion = SCRIPT_VERSION;
  template.nonce = Utilities.getUuid();

  const htmlOutput = template.evaluate();
  htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  htmlOutput.setTitle(title);
  htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return htmlOutput;
}

// --------------------------------
// メイン機能
// --------------------------------

/**
 * 到着/未到着の事業所リストと、全事業所名のリストを取得する
 * @param {object} options - (オプション) { sortByTimestamp: boolean } を含むオブジェクト
 * @returns {object} { arrived: [], notArrived: [], centers: [], isKioskTouchEnabled: boolean }
 */
function getServiceData(options) {
  // 読み取り専用なのでログインチェックのみ
  Auth.assertLogin();

  // デバッグのため一時的にキャッシュ無効化（必要に応じて有効化してください）
  CacheService.getScriptCache().remove(CACHE_KEY);
  let debugLog = "--- サーバーサイドデバッグログ ---\n";

  try {
    const sortByTimestamp = options && options.sortByTimestamp;
    if (!sortByTimestamp) {
      const cache = CacheService.getScriptCache();
      const cachedData = cache.get(CACHE_KEY);
      if (cachedData) {
        const parsedData = JSON.parse(cachedData);
        parsedData.isKioskTouchEnabled = getKioskTouchMode();
        return parsedData;
      }
    }

    const ss = SpreadsheetApp.openById(MY_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(MY_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) {
      return { arrived: [], notArrived: [], centers: [], isKioskTouchEnabled: getKioskTouchMode() };
    }

    const headers = mapHeaders_(sheet);

    // ★追加: 必須ヘッダーのチェック (1行目が正しいか確認するため)
    if (headers['QRコードID'] === undefined || headers['デイサービス名'] === undefined) {
      throw new Error("必須項目（QRコードID、デイサービス名）がシートの1行目に見つかりません。1行目がタイトル行になっていないか、項目名が正しいか確認してください。");
    }

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    const arrived = [];
    const notArrived = [];
    const centerList = new Set();
    const centerMap = new Map(); // 事業所名と読み仮名をペアで保持

    debugLog += "認識されたヘッダー: " + JSON.stringify(headers) + "\n";
    const displayStatusIndex = headers['表示設定'];

    data.forEach((row, index) => {
      // 列が存在しない場合、または空欄の場合はデフォルトで '表示' とする
      const rawDisplayStatus = (displayStatusIndex !== undefined) ? row[displayStatusIndex] : '表示';
      const displayStatus = rawDisplayStatus ? String(rawDisplayStatus).replace(/\s/g, '') : '';

      // '表示' または空文字（未設定）の場合にリストに含める
      if (displayStatus === '表示' || displayStatus === '') {
        const qrCodeId = String(row[headers['QRコードID']] || '').trim();
        const name = String(row[headers['デイサービス名']] || '').trim();
        const readingName = String(row[headers['読み仮名']] || '').trim();

        if (qrCodeId && name) {
          centerList.add(name);
          if (!centerMap.has(name)) centerMap.set(name, readingName || name);

          const service = {
            qrCodeId: qrCodeId,
            name: name,
            readingName: readingName,
            timestamp: row[headers['到着時刻']] ? new Date(row[headers['到着時刻']]).getTime() : null,
            lateNote: row[headers['遅延連絡']] || ''
          };

          const status = String(row[headers['状態']] || '').trim();
          if (status === '到着') {
            arrived.push(service);
          } else {
            notArrived.push(service);
          }
        }
      }
    });

    const nameSorter = (a, b) => {
      const nameA = a.readingName || a.name;
      const nameB = b.readingName || b.name;
      return nameA.localeCompare(nameB, 'ja');
    };

    const sortedArrived = sortByTimestamp
      ? arrived.sort((a, b) => a.timestamp - b.timestamp)
      : arrived.sort(nameSorter);

    const resultData = {
      arrived: sortedArrived,
      notArrived: notArrived.sort(nameSorter),
      centers: Array.from(centerList).sort((a, b) => {
        const readingA = centerMap.get(a);
        const readingB = centerMap.get(b);
        return readingA.localeCompare(readingB, 'ja');
      }),
      isKioskTouchEnabled: getKioskTouchMode(),
      debugLog: debugLog
    };

    if (!sortByTimestamp) {
      CacheService.getScriptCache().put(CACHE_KEY, JSON.stringify(resultData), 21600);
    }

    return resultData;

  } catch (e) {
    logError({ functionName: 'getServiceData', message: e.message, stack: e.stack });
    throw new Error('事業所データの取得中にエラーが発生しました。');
  }
}

/**
 * 指定されたQRコードIDの事業所を「到着」として記録する
 * @param {string} qrCodeId - 事業所のQRコードID
 * @returns {object} 処理結果
 */
function checkInService(qrCodeId) {
  // Teacher以上
  Auth.assertTeacherOrAdmin();

  try {
    const result = updateStatusInSheet(qrCodeId, '到着');

    // 到着処理が成功したら遅延メモをクリアする
    if (result.success) {
      const ss = SpreadsheetApp.openById(MY_SPREADSHEET_ID);
      const sheet = ss.getSheetByName(MY_SHEET_NAME);
      const rowNum = findRowByQrCodeId_(sheet, qrCodeId);

      if (rowNum !== -1) {
        // ★修正：固定列(5)ではなくヘッダーから動的に列を取得
        const headers = mapHeaders_(sheet);
        const noteColIndex = headers['遅延連絡'] + 1;
        sheet.getRange(rowNum, noteColIndex).setValue('');
      }
    }

    return result;
  } catch (e) {
    logError({ functionName: 'checkInService', message: e.message, stack: e.stack, type: 'Server' });
    throw e;
  }
}

/**
 * 指定されたQRコードIDの事業所を「未到着」として記録する
 * @param {string} qrCodeId - 事業所のQRコードID
 * @returns {object} 処理結果
 */
function checkOutService(qrCodeId) {
  // Teacher以上
  Auth.assertTeacherOrAdmin();

  try {
    return updateStatusInSheet(qrCodeId, '未到着');
  } catch (e) {
    logError({ functionName: 'checkOutService', message: e.message, stack: e.stack, type: 'Server' });
    throw e;
  }
}

/**
 * スプレッドシートの指定されたIDのステータスを更新する（内部関数）
 */
function updateStatusInSheet(qrCodeId, newStatus) {
  // ★★★ 排他制御の開始 ★★★
  const lock = LockService.getScriptLock();
  try {
    // ロックを取得（最大10秒待機）。他が処理中ならここで待たされる。
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, message: 'サーバーが混み合っています。もう一度お試しください。' };
  }

  try {
    const ss = SpreadsheetApp.openById(MY_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(MY_SHEET_NAME);
    if (!sheet) {
      return { success: false, message: `シートが見つかりません。` };
    }

    const historySheet = ss.getSheetByName(HISTORY_SHEET_NAME);
    if (!historySheet) {
      console.warn("「チェックイン履歴」シートが見つかりません。記録はスキップされます。");
    }

    const headers = mapHeaders_(sheet);
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    const trimmedQrCodeId = String(qrCodeId).trim();

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (String(row[headers['QRコードID']]).trim() === trimmedQrCodeId) {
        const rowNum = i + 2;
        const name = String(row[headers['デイサービス名']]).trim();
        const readingName = String(row[headers['読み仮名']] || '').trim();
        const currentStatus = String(row[headers['状態']]).trim();

        if (currentStatus === newStatus) {
          const statusText = newStatus === '到着' ? '受付済み' : '未到着';
          return { success: false, message: `${name}は既に${statusText}です。` };
        }

        // ヘッダーマップを使って動的に列を指定
        sheet.getRange(rowNum, headers['状態'] + 1).setValue(newStatus);
        const now = new Date();

        if (newStatus === '到着') {
          sheet.getRange(rowNum, headers['到着時刻'] + 1).setValue(now);
          if (historySheet) {
            historySheet.appendRow([now, trimmedQrCodeId, name]);
          }
        } else {
          sheet.getRange(rowNum, headers['到着時刻'] + 1).setValue('');
        }

        SpreadsheetApp.flush();
        CacheService.getScriptCache().remove(CACHE_KEY);

        const message = newStatus === '到着' ? `${name}、到着しました。` : `${name}を未到着にしました。`;
        logActivity({ action: 'ステータス変更', details: `事業所: ${name} (ID: ${trimmedQrCodeId}), 新ステータス: ${newStatus}` });

        return {
          success: true,
          message: message,
          updatedService: {
            qrCodeId: trimmedQrCodeId,
            name: name,
            textToSpeak: readingName ? readingName : name
          }
        };
      }
    }

    const errorMessage = '登録されていないか、古いQRコードです。新しいカードを使用してください。';
    throw new Error(errorMessage);

  } catch (e) {
    const errorAudioId = ERROR_SOUND_ID;
    logError({ functionName: 'updateStatusInSheet', message: e.message, stack: e.stack, type: 'Server' });
    return { success: false, message: e.message, audioId: errorAudioId };
  } finally {
    // ★★★ ロックの解除（必須） ★★★
    lock.releaseLock();
  }
}

// --------------------------------
// 管理者用・開発者用・トリガー用機能
// --------------------------------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('アプリツール')
    .addItem('選択範囲の事業所カードを生成', 'generateSelectedCardsView')
    .addItem('選択範囲のQRコードを生成', 'generateQRCodesForSelection')
    .addSeparator()
    .addItem('キャッシュを強制クリア', 'forceClearCache')
    .addToUi();
}

function addServiceCenter(name, reading = '', address = '', phone = '') {
  // Admin権限必須
  Auth.assertAdmin();

  // ★★★ 排他制御 ★★★
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error('サーバーが混み合っています。再試行してください。'); }

  try {
    if (!name || name.trim() === '') { throw new Error('事業所名を入力してください。'); }
    const newName = name.trim();
    const ss = SpreadsheetApp.openById(MY_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(MY_SHEET_NAME);
    const headers = mapHeaders_(sheet);
    const lastRow = sheet.getLastRow();
    const prefix = getQrIdPrefix_();
    const existingQrIds = lastRow > 1 ? sheet.getRange(2, headers['QRコードID'] + 1, lastRow - 1, 1).getValues().flat() : [];
    let maxNumber = 0;
    existingQrIds.forEach(id => {
      if (id && String(id).startsWith(prefix)) {
        const numPart = parseInt(String(id).substring(prefix.length), 10);
        if (!isNaN(numPart) && numPart > maxNumber) { maxNumber = numPart; }
      }
    });
    const newNumber = maxNumber + 1;
    const newQrId = prefix + String(newNumber).padStart(3, '0');

    if (!QR_CODE_FOLDER_ID || QR_CODE_FOLDER_ID === 'ここにQRコード保存用フォルダのIDをペースト') {
      throw new Error('スクリプト内のQR_CODE_FOLDER_IDが設定されていません。');
    }
    const folder = DriveApp.getFolderById(QR_CODE_FOLDER_ID);

    const apiUrl = `https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=${encodeURIComponent(newQrId)}&ecc=H`;
    const response = UrlFetchApp.fetch(apiUrl);
    const blob = response.getBlob().setName(`${newQrId}.png`);
    folder.createFile(blob);

    const newRow = new Array(sheet.getLastColumn()).fill('');
    newRow[headers['QRコードID']] = newQrId;
    newRow[headers['デイサービス名']] = newName;
    if (headers['読み仮名'] !== undefined) newRow[headers['読み仮名']] = String(reading).trim();
    if (headers['住所'] !== undefined) newRow[headers['住所']] = String(address).trim();
    if (headers['電話番号'] !== undefined) newRow[headers['電話番号']] = String(phone).trim();
    newRow[headers['状態']] = '未到着';
    newRow[headers['表示設定']] = '表示';

    sheet.appendRow(newRow);

    CacheService.getScriptCache().remove(CACHE_KEY);
    logActivity({ action: '事業所追加', details: `追加された事業所: ${newName} (ID: ${newQrId})` });

    return { success: true, name: newName, qrCodeId: newQrId };

  } catch (e) {
    logError({ functionName: 'addServiceCenter', message: e.message, stack: e.stack, type: 'Server' });
    throw new Error('事業所の追加中にエラーが発生しました: ' + e.message);
  } finally {
    lock.releaseLock();
  }
}

function deleteServiceCenter(qrCodeId) {
  // Admin権限必須
  Auth.assertAdmin();

  // ★★★ 排他制御 ★★★
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error('サーバーが混み合っています。再試行してください。'); }

  try {
    if (!qrCodeId) { throw new Error('削除対象のQRコードIDが指定されていません。'); }

    const ss = SpreadsheetApp.openById(MY_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(MY_SHEET_NAME);
    const headers = mapHeaders_(sheet);
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    let rowToDelete = -1;

    for (let i = 0; i < data.length; i++) {
      if (data[i][headers['QRコードID']] === qrCodeId) {
        rowToDelete = i + 2;
        break;
      }
    }

    if (rowToDelete !== -1) {
      const deletedName = sheet.getRange(rowToDelete, headers['デイサービス名'] + 1).getValue();
      sheet.deleteRow(rowToDelete);
      CacheService.getScriptCache().remove(CACHE_KEY);

      logActivity({ action: '事業所削除', details: `削除された事業所: ${deletedName} (ID: ${qrCodeId})` });

      return true;

    } else {
      throw new Error(`ID「${qrCodeId}」が見つかりませんでした。`);
    }
  } catch (e) {
    logError({ functionName: 'deleteServiceCenter', message: e.message, stack: e.stack, type: 'Server' });
    throw new Error('事業所の削除中にエラーが発生しました: ' + e.message);
  } finally {
    lock.releaseLock();
  }
}

function getAudioAsBase64(fileId) {
  // 内部利用または読み取りのみ
  try {
    if (!fileId) return null;
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    return {
      mimeType: blob.getContentType(),
      data: Utilities.base64Encode(blob.getBytes())
    };
  } catch (e) {
    logError({ functionName: 'getAudioAsBase64', message: e.message, stack: e.stack, type: 'Server' });
    return null;
  }
}

function getAudioDataForService(qrCodeId, direction) {
  // 読み取りのみ
  console.log(`--- getAudioDataForService 開始: QR ID [${qrCodeId}], 方向 [${direction}] ---`);
  try {
    const soundConfig = getSoundConfigForToday();
    if (!soundConfig || !soundConfig.folderId) {
      console.error("デバッグ情報: 今日の担当キャラクターまたはフォルダIDが設定されていません。");
      return null;
    }

    const folderId = soundConfig.folderId;
    const folder = DriveApp.getFolderById(folderId);
    const cleanQrCodeId = String(qrCodeId || '').trim();
    const filenameMp3 = `${cleanQrCodeId}_${direction}.mp3`;
    const filenameWav = `${cleanQrCodeId}_${direction}.wav`;

    let file = null;
    const filesMp3 = folder.getFilesByName(filenameMp3);
    if (filesMp3.hasNext()) {
      file = filesMp3.next();
    } else {
      const filesWav = folder.getFilesByName(filenameWav);
      if (filesWav.hasNext()) {
        file = filesWav.next();
      }
    }

    if (file) {
      const blob = file.getBlob();
      return {
        mimeType: blob.getContentType(),
        data: Utilities.base64Encode(blob.getBytes())
      };
    } else {
      console.error(`デバッグ情報: フォルダ「${folder.getName()}」の中に、該当する音声ファイルが見つかりませんでした。`);
      return null;
    }
  } catch (e) {
    logError({ functionName: 'getAudioDataForService', message: e.message, stack: e.stack });
    return null;
  }
}

function getSoundConfigForToday() {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'soundConfigCache';
  const cachedConfig = cache.get(cacheKey);
  if (cachedConfig) {
    return JSON.parse(cachedConfig);
  }

  const dayIndex = new Date().getDay();
  const dayKeys = ['daySunday', 'dayMonday', 'dayTuesday', 'dayWednesday', 'dayThursday', 'dayFriday', 'daySaturday'];
  const todayKey = dayKeys[dayIndex];

  const ss = SpreadsheetApp.openById(MY_SPREADSHEET_ID);
  const scheduleSheet = ss.getSheetByName(SCHEDULE_SHEET_NAME);
  if (!scheduleSheet || scheduleSheet.getLastRow() < 2) return null;
  const scheduleHeaders = mapHeaders_(scheduleSheet);
  const numRows = scheduleSheet.getLastRow() - 1;
  const scheduleData = scheduleSheet.getRange(2, 1, numRows, scheduleSheet.getLastColumn()).getValues();

  let todaysCharacterName = '';
  for (const row of scheduleData) {
    if (String(row[scheduleHeaders['曜日']]).trim() === String(todayKey).trim()) {
      todaysCharacterName = String(row[scheduleHeaders['担当キャラクター']]).trim();
      break;
    }
  }
  if (!todaysCharacterName) return null;

  const charSheet = ss.getSheetByName(CHARACTER_SHEET_NAME);
  const lastRow = charSheet.getLastRow();
  if (lastRow < 2) return null;
  const charHeaders = mapHeaders_(charSheet);
  const charData = charSheet.getRange(2, 1, lastRow - 1, charSheet.getLastColumn()).getValues();
  let todaysFolderId = '';
  for (const row of charData) {
    if (String(row[charHeaders['キャラクター名']]).trim() === todaysCharacterName) {
      todaysFolderId = String(row[charHeaders['音声フォルダID']]).trim();
      break;
    }
  }
  if (!todaysFolderId) return null;


  const config = {
    characterName: todaysCharacterName,
    folderId: todaysFolderId
  };

  cache.put(cacheKey, JSON.stringify(config), 21600);
  return config;
}

function forceClearCache() {
  // Admin権限必須
  Auth.assertAdmin();
  CacheService.getScriptCache().remove(CACHE_KEY);
  logActivity({ action: 'キャッシュクリア', details: '管理者が手動で実行' });
  return true;
}

/**
 * 全ての事業所のステータスを「未到着」にリセットする（高速・安全版）
 * ★修正：ループを使わず、列単位で一括クリアすることで計算式破壊を防ぎ高速化
 */
function resetAllStatusesToNotArrived() {
  // Admin権限必須
  Auth.assertAdmin();
  return resetAllStatusesToNotArrivedInternal_(false);
}

/**
 * トリガーからの実行用関数
 */
function dailyResetTask() {
  try {
    resetAllStatusesToNotArrivedInternal_(true);
  } catch (e) {
    logError({ functionName: 'dailyResetTask', message: e.message, stack: e.stack, type: 'Trigger' });
  }
}

/**
 * 全ステータスリセットの内部処理
 * @param {boolean} isAuto - トリガーによる自動実行かどうかのフラグ
 */
function resetAllStatusesToNotArrivedInternal_(isAuto) {

  // ★★★ 排他制御 ★★★
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error('サーバーが混み合っています。再試行してください。'); }

  try {
    const ss = SpreadsheetApp.openById(MY_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(MY_SHEET_NAME);
    const lastRow = sheet.getLastRow();

    if (lastRow > 1) {
      const headers = mapHeaders_(sheet);
      const numRows = lastRow - 1;

      // ヘッダー名から列番号を特定（+1 で1-based indexにする）
      const statusCol = headers['状態'] !== undefined ? headers['状態'] + 1 : null;
      const timeCol = headers['到着時刻'] !== undefined ? headers['到着時刻'] + 1 : null;
      const noteCol = headers['遅延連絡'] !== undefined ? headers['遅延連絡'] + 1 : null;

      // 1. 状態列をすべて '未到着' にする
      if (statusCol) {
        sheet.getRange(2, statusCol, numRows, 1).setValue('未到着');
      }

      // 2. 到着時刻列をクリアする
      if (timeCol) {
        sheet.getRange(2, timeCol, numRows, 1).clearContent();
      }

      // 3. 遅延連絡列をクリアする
      if (noteCol) {
        sheet.getRange(2, noteCol, numRows, 1).clearContent();
      }
    }

    CacheService.getScriptCache().remove(CACHE_KEY);
    const executor = isAuto ? 'システム (自動実行)' : '管理者が手動で実行';
    logActivity({ action: '全ステータスリセット', details: executor });
    return true;
  } catch (e) {
    logError({ functionName: 'resetAllStatusesToNotArrivedInternal_', message: e.message, stack: e.stack, type: 'Server' });
    throw e;
  } finally {
    lock.releaseLock();
  }
}

function generateQRCodesForSelection() {
  // スプレッドシートUIからの実行を想定（Authチェックはコンテキストによるが、Webアプリ経由ならAdmin）
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    if (sheet.getName() !== MY_SHEET_NAME) {
      SpreadsheetApp.getUi().alert(`この機能は「${MY_SHEET_NAME}」シートでのみ実行できます。`);
      return;
    }

    const selection = sheet.getActiveRange();
    if (!selection) {
      SpreadsheetApp.getUi().alert('セルを選択してください。');
      return;
    }

    const startRow = selection.getRow();
    const numRows = selection.getNumRows();

    if (startRow === 1 && numRows === 1 && selection.getWidth() === sheet.getMaxColumns()) {
      SpreadsheetApp.getUi().alert('データ行を選択してください。（1行目のヘッダーは選択できません）');
      return;
    }

    const headers = mapHeaders_(sheet);
    const idData = sheet.getRange(startRow, headers['QRコードID'] + 1, numRows, 1).getValues();
    const qrCodeIds = idData.flat().filter(id => id.toString().trim() !== '');

    if (qrCodeIds.length === 0) {
      SpreadsheetApp.getUi().alert('選択範囲に有効なQRコードIDが見つかりません。');
      return;
    }

    const folder = DriveApp.getFolderById(QR_CODE_FOLDER_ID);
    let count = 0;
    let errors = [];

    qrCodeIds.forEach(qrCodeId => {
      try {
        const apiUrl = `https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=${encodeURIComponent(qrCodeId)}&ecc=H`;
        const response = UrlFetchApp.fetch(apiUrl);
        const blob = response.getBlob().setName(`${qrCodeId}.png`);
        folder.createFile(blob);
        count++;
      } catch (e) {
        errors.push(qrCodeId);
      }
    });

    let message = `${count}件のQRコードをGoogleドライブに保存しました。`;
    if (errors.length > 0) {
      message += `\n\n以下のIDのQRコード生成に失敗しました:\n${errors.join('\n')}`;
      logError({ functionName: 'generateQRCodesForSelection', message: `QRコード生成失敗: ${errors.join(',')}`, stack: '', type: 'Server' });
    }
    SpreadsheetApp.getUi().alert(message);

  } catch (e) {
    logError({ functionName: 'generateQRCodesForSelection', message: e.message, stack: e.stack, type: 'Server' });
    SpreadsheetApp.getUi().alert('QRコードの生成中にエラーが発生しました。エラーログを確認してください。');
  }
}

function getAllowedUsers() {
  // Admin権限必須
  Auth.assertAdmin();

  // ★★★ 排他制御 ★★★
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error('サーバーが混み合っています。再試行してください。'); }

  try {
    const ss = SpreadsheetApp.openById(MY_SPREADSHEET_ID);
    const userSheet = ss.getSheetByName(USER_SHEET_NAME);
    if (!userSheet || userSheet.getLastRow() < 2) {
      return [];
    }
    const headers = mapHeaders_(userSheet);
    const data = userSheet.getRange(2, 1, userSheet.getLastRow() - 1, userSheet.getLastColumn()).getValues();

    const users = data.map(row => ({
      email: String(row[headers['email']]).trim(),
      role: String(row[headers['role']] || '').trim(),
      type: String(row[headers['type']] || 'user').trim()
    }));
    return users;
  } catch (e) {
    logError({ functionName: 'getAllowedUsers', message: e.message, stack: e.stack });
    throw new Error('許可ユーザーの取得に失敗しました。');
  }
}

function addAllowedUser(userData) {
  // Admin権限必須
  Auth.assertAdmin();

  // ★★★ 排他制御 ★★★
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error('サーバーが混み合っています。再試行してください。'); }

  try {
    const { email, type, role } = userData;
    if (!email || !email.includes('@')) {
      throw new Error('有効なメールアドレスを入力してください。');
    }
    const newEmail = email.trim().toLowerCase();
    const newType = type || 'user';
    const newRole = role || '';
    const ss = SpreadsheetApp.openById(MY_SPREADSHEET_ID);
    const userSheet = ss.getSheetByName(USER_SHEET_NAME);
    const headers = mapHeaders_(userSheet);
    const lastRow = userSheet.getLastRow();
    const existingEmails = lastRow > 1 ? userSheet.getRange(2, headers['email'] + 1, lastRow - 1, 1).getValues().flat().map(e => e.toLowerCase()) : [];
    if (existingEmails.includes(newEmail)) {
      throw new Error(`アドレス「${newEmail}」は既に追加されています。`);
    }
    const newRow = new Array(userSheet.getLastColumn()).fill('');
    newRow[headers['email']] = newEmail;
    newRow[headers['type']] = newType;
    newRow[headers['role']] = newRole;
    userSheet.appendRow(newRow);

    // ユーザー追加時はキャッシュをクリアして即時反映させる（ただし自分自身のキャッシュのみ）
    CacheService.getUserCache().remove('user_context');

    logActivity({ action: 'ユーザー追加', details: `追加されたアドレス: ${newEmail}` });
    return { success: true, email: newEmail };
  } catch (e) {
    logError({ functionName: 'addAllowedUser', message: e.message, stack: e.stack, type: 'Server' });
    throw new Error(e.message);
  } finally {
    lock.releaseLock();
  }
}

function setUserType(email, type) {
  // Admin権限必須
  Auth.assertAdmin();

  // ★★★ 排他制御 ★★★
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error('サーバーが混み合っています。再試行してください。'); }

  try {
    const ss = SpreadsheetApp.openById(MY_SPREADSHEET_ID);
    const userSheet = ss.getSheetByName(USER_SHEET_NAME);
    if (!userSheet || userSheet.getLastRow() < 2) {
      throw new Error("ユーザーシートが見つからないか、データがありません。");
    }

    const headers = mapHeaders_(userSheet);
    const rowToUpdate = findRowByEmail_(userSheet, email);

    if (rowToUpdate !== -1) {
      userSheet.getRange(rowToUpdate, headers['type'] + 1).setValue(type);
      logActivity({ action: '種別変更', details: `ユーザー: ${email}, 新しい種別: ${type}` });
      return { success: true, message: `ユーザー「${email}」の種別を更新しました。` };
    } else {
      throw new Error(`ユーザー「${email}」が見つかりませんでした。`);
    }
  } catch (e) {
    logError({ functionName: 'setUserType', message: e.message, stack: e.stack, type: 'Server' });
    throw new Error('種別の設定に失敗しました。');
  } finally {
    lock.releaseLock();
  }
}

function deleteAllowedUser(email) {
  // Admin権限必須
  Auth.assertAdmin();

  if (!email) {
    throw new Error('削除対象のアドレスが指定されていません。');
  }
  const emailToDelete = email.trim().toLowerCase();

  // ★★★ 排他制御 ★★★
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error('サーバーが混み合っています。再試行してください。'); }

  try {
    const ss = SpreadsheetApp.openById(MY_SPREADSHEET_ID);
    const userSheet = ss.getSheetByName(USER_SHEET_NAME);
    const headers = mapHeaders_(userSheet);
    const data = userSheet.getRange(2, headers['email'] + 1, userSheet.getLastRow() - 1, 1).getValues();

    let rowToDelete = -1;
    for (let i = 0; i < data.length; i++) {
      if (data[i][0].toLowerCase() === emailToDelete) {
        rowToDelete = i + 2;
        break;
      }
    }

    if (rowToDelete !== -1) {
      userSheet.deleteRow(rowToDelete);
      logActivity({ action: 'ユーザー削除', details: `削除されたアドレス: ${emailToDelete}` });
      return { success: true, email: emailToDelete };
    } else {
      throw new Error(`アドレス「${emailToDelete}」が見つかりませんでした。`);
    }
  } catch (e) {
    logError({ functionName: 'deleteAllowedUser', message: e.message, stack: e.stack, type: 'Server' });
    throw new Error(e.message);
  } finally {
    lock.releaseLock();
  }
}

function bulkDeleteAllowedUsers(emails) {
  // Admin権限必須
  Auth.assertAdmin();

  // ★★★ 排他制御 ★★★
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error('サーバーが混み合っています。再試行してください。'); }

  try {
    const ss = SpreadsheetApp.openById(MY_SPREADSHEET_ID);
    const userSheet = ss.getSheetByName(USER_SHEET_NAME);
    const headers = mapHeaders_(userSheet);
    const data = userSheet.getRange(2, headers['email'] + 1, userSheet.getLastRow() - 1, 1).getValues();

    const emailSet = new Set(emails.map(e => e.toLowerCase()));
    let count = 0;

    // 下からループして削除時の行ズレを防ぐ
    for (let i = data.length - 1; i >= 0; i--) {
      if (emailSet.has(String(data[i][0]).toLowerCase())) {
        userSheet.deleteRow(i + 2);
        count++;
      }
    }

    logActivity({ action: 'ユーザー一括削除', details: `${count}件のユーザーを削除` });
    return count;
  } catch (e) {
    logError({ functionName: 'bulkDeleteAllowedUsers', message: e.message, stack: e.stack, type: 'Server' });
    throw new Error('ユーザーの一括削除に失敗しました。');
  } finally {
    lock.releaseLock();
  }
}

function getKioskTouchMode() {
  // 読み取りのみ
  Auth.assertLogin();
  try {
    return PropertiesService.getScriptProperties().getProperty('kioskTouchEnabled') === 'true';
  } catch (e) {
    logError({ functionName: 'getKioskTouchMode', message: e.message, stack: e.stack, type: 'Server' });
    throw new Error('キオスク設定の取得に失敗しました。');
  }
}

function setKioskTouchMode(isEnabled) {
  // Admin権限必須
  Auth.assertAdmin();
  // ★★★ 排他制御 ★★★
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error('サーバーが混み合っています。再試行してください。'); }

  try {
    PropertiesService.getScriptProperties().setProperty('kioskTouchEnabled', isEnabled);
    forceClearCache();
    const message = `キオスクのタッチ操作を「${isEnabled ? '有効' : '無効'}」に設定しました。`;
    logActivity({ action: 'キオスク設定変更', details: message });
    return message;
  } catch (e) {
    logError({ functionName: 'setKioskTouchMode', message: e.message, stack: e.stack, type: 'Server' });
    throw new Error('キオスク設定の保存に失敗しました。');
  } finally {
    lock.releaseLock();
  }
}

function getDailyResetTriggerStatus() {
  Auth.assertAdmin();
  const triggers = ScriptApp.getProjectTriggers();
  return triggers.some(t => t.getHandlerFunction() === 'dailyResetTask');
}

function setDailyResetTrigger(isEnabled) {
  Auth.assertAdmin();
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error('サーバーが混み合っています。再試行してください。'); }

  try {
    const triggers = ScriptApp.getProjectTriggers();
    // 既存のトリガーを削除
    triggers.forEach(t => {
      if (t.getHandlerFunction() === 'dailyResetTask') {
        ScriptApp.deleteTrigger(t);
      }
    });

    if (isEnabled) {
      // 毎日午前0時〜1時の間に実行するトリガーを作成
      ScriptApp.newTrigger('dailyResetTask')
        .timeBased()
        .everyDays(1)
        .atHour(0)
        .create();
      logActivity({ action: '自動リセット設定', details: '毎日深夜0時の自動リセットを有効化しました。' });
    } else {
      logActivity({ action: '自動リセット設定', details: '自動リセットを無効化しました。' });
    }
    return { success: true, isEnabled: isEnabled };
  } catch (e) {
    logError({ functionName: 'setDailyResetTrigger', message: e.message, stack: e.stack, type: 'Server' });
    throw new Error('自動リセット設定の保存に失敗しました。');
  } finally {
    lock.releaseLock();
  }
}

function getArrivalReportData(serviceName, startDateString, endDateString) {
  // Teacher以上
  Auth.assertTeacherOrAdmin();
  try {
    const INTERVAL_MINUTES = 10;
    const ss = SpreadsheetApp.openById(MY_SPREADSHEET_ID);
    const historySheet = ss.getSheetByName(HISTORY_SHEET_NAME);
    if (!historySheet || historySheet.getLastRow() < 2) { return {}; }
    const headers = mapHeaders_(historySheet);
    const data = historySheet.getRange(2, 1, historySheet.getLastRow() - 1, historySheet.getLastColumn()).getValues();
    const report = {};
    let startDate, endDate;
    if (startDateString) { startDate = new Date(startDateString); startDate.setHours(0, 0, 0, 0); }
    if (endDateString) { endDate = new Date(endDateString); endDate.setHours(23, 59, 59, 999); }
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const timestamp = new Date(row[headers['タイムスタンプ']]);
      const name = row[headers['事業所名']];
      if (!name || isNaN(timestamp.getTime())) continue;
      if (startDate && timestamp < startDate) continue;
      if (endDate && timestamp > endDate) continue;
      if (serviceName && serviceName !== 'all' && name !== serviceName) continue;
      const hour = timestamp.getHours();
      const minute = timestamp.getMinutes();
      const startMinute = Math.floor(minute / INTERVAL_MINUTES) * INTERVAL_MINUTES;
      const endMinute = startMinute + INTERVAL_MINUTES - 1;
      const timeSlot = String(hour).padStart(2, '0') + ':' + String(startMinute).padStart(2, '0') + '-' + String(hour).padStart(2, '0') + ':' + String(endMinute).padStart(2, '0');
      if (!report[name]) { report[name] = {}; }
      if (!report[name][timeSlot]) { report[name][timeSlot] = 0; }
      report[name][timeSlot]++;
    }
    return report;
  } catch (e) {
    logError({ functionName: 'getArrivalReportData', message: e.message, stack: e.stack, type: 'Server' });
    throw e;
  }
}

function getServiceCenterList() {
  // 読み取りのみ
  Auth.assertLogin();

  const data = getServiceData();
  return data.centers;
}

function editServiceCenter(qrCodeId, updateData) {
  // Admin権限必須
  Auth.assertAdmin();

  // ★★★ 排他制御 ★★★
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error('サーバーが混み合っています。再試行してください。'); }

  try {
    const newName = updateData.name;
    const newReading = updateData.reading || '';

    if (!newName || newName.trim() === '') {
      throw new Error('事業所名が空です。');
    }
    const trimmedNewName = newName.trim();
    const trimmedNewReading = newReading.trim();

    const ss = SpreadsheetApp.openById(MY_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(MY_SHEET_NAME);
    if (!sheet) { throw new Error('デイサービス連絡版シートが見つかりません。'); }
    const headers = mapHeaders_(sheet);
    const rowNum = findRowByQrCodeId_(sheet, qrCodeId);

    if (rowNum !== -1) {
      const oldName = sheet.getRange(rowNum, headers['デイサービス名'] + 1).getValue();
      const oldReading = sheet.getRange(rowNum, headers['読み仮名'] + 1).getValue();

      // ★修正：マジックナンバー(2,6)を排除し、ヘッダーマップを使用
      sheet.getRange(rowNum, headers['デイサービス名'] + 1).setValue(trimmedNewName);
      sheet.getRange(rowNum, headers['読み仮名'] + 1).setValue(trimmedNewReading);

      CacheService.getScriptCache().remove(CACHE_KEY);

      const logDetails = `ID: ${qrCodeId}, 名前: [${oldName}]->[${trimmedNewName}], 読み仮名: [${oldReading}]->[${trimmedNewReading}]`;
      logActivity({ action: '事業所情報編集', details: logDetails });

      return { success: true, message: '事業所情報を更新しました。' };
    } else {
      throw new Error(`ID「${qrCodeId}」の事業所が見つかりませんでした。`);
    }
  } catch (e) {
    logError({ functionName: 'editServiceCenter', message: e.message, stack: e.stack, type: 'Server' });
    throw new Error(e.message);
  } finally {
    lock.releaseLock();
  }
}

function getArrivalHistoryRaw(serviceName, startDateString, endDateString) {
  // Teacher以上
  Auth.assertTeacherOrAdmin();
  try {
    if (!serviceName || serviceName === 'all') { return []; }
    const ss = SpreadsheetApp.openById(MY_SPREADSHEET_ID);
    const historySheet = ss.getSheetByName(HISTORY_SHEET_NAME);
    if (!historySheet || historySheet.getLastRow() < 2) { return []; }
    const headers = mapHeaders_(historySheet);
    const data = historySheet.getRange(2, 1, historySheet.getLastRow() - 1, historySheet.getLastColumn()).getValues();
    const rawData = [];
    let startDate, endDate;
    if (startDateString) { startDate = new Date(startDateString); startDate.setHours(0, 0, 0, 0); }
    if (endDateString) { endDate = new Date(endDateString); endDate.setHours(23, 59, 59, 999); }
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const recordName = row[headers['事業所名']];
      if (recordName !== serviceName) continue;
      const timestamp = new Date(row[headers['タイムスタンプ']]);
      if (!recordName || isNaN(timestamp.getTime())) continue;
      if (startDate && timestamp < startDate) continue;
      if (endDate && timestamp > endDate) continue;
      rawData.push([timestamp.toISOString(), timestamp.toLocaleTimeString('ja-JP', { hour12: false })]);
    }
    return rawData;
  } catch (e) {
    logError({ functionName: 'getArrivalHistoryRaw', message: e.message, stack: e.stack, type: 'Server' });
    throw e;
  }
}

function getReportForExport(serviceName, startDateString, endDateString) {
  // Teacher以上
  Auth.assertTeacherOrAdmin();

  const rawData = getArrivalHistoryRaw(serviceName, startDateString, endDateString);

  if (!rawData || rawData.length === 0) {
    throw new Error("指定された条件のエクスポート対象データが見つかりません。");
  }

  const ss = SpreadsheetApp.openById(MY_SPREADSHEET_ID);
  const masterSheet = ss.getSheetByName(MY_SHEET_NAME);
  const headers = mapHeaders_(masterSheet);
  const masterData = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, masterSheet.getLastColumn()).getValues();
  let qrCodeId = 'ID_UNKNOWN';
  for (const row of masterData) {
    if (row[headers['デイサービス名']] === serviceName) {
      qrCodeId = row[headers['QRコードID']];
      break;
    }
  }

  const header = '"到着日","到着時刻"\n';
  const csvRows = rawData.map(row => {
    const date = new Date(row[0]);
    const dateString = `${date.getFullYear()}/${(date.getMonth() + 1).toString().padStart(2, '0')}/${date.getDate().toString().padStart(2, '0')}`;
    const timeString = row[1];
    return `"${dateString}","${timeString}"`;
  });

  const csvContent = header + csvRows.join('\n');

  return {
    csvContent: csvContent,
    qrCodeId: qrCodeId,
    serviceName: serviceName
  };
}

function editAllowedUser(oldEmail, newEmail) {
  // Admin権限必須
  Auth.assertAdmin();

  // ★★★ 排他制御 ★★★
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error('サーバーが混み合っています。再試行してください。'); }

  try {
    if (!newEmail || !newEmail.includes('@')) {
      throw new Error('有効なメールアドレスを入力してください。');
    }
    const trimmedNewEmail = newEmail.trim().toLowerCase();
    const trimmedOldEmail = oldEmail.trim().toLowerCase();

    const ss = SpreadsheetApp.openById(MY_SPREADSHEET_ID);
    const userSheet = ss.getSheetByName(USER_SHEET_NAME);
    if (!userSheet) { throw new Error('「ユーザー」シートが見つかりません。'); }

    const headers = mapHeaders_(userSheet);
    const data = userSheet.getRange(2, headers['email'] + 1, userSheet.getLastRow() - 1, 1).getValues();

    const existingEmails = data.flat().map(e => e.toLowerCase());
    if (existingEmails.includes(trimmedNewEmail) && trimmedNewEmail !== trimmedOldEmail) {
      throw new Error(`アドレス「${trimmedNewEmail}」は既に追加されています。`);
    }

    let found = false;
    for (let i = 0; i < data.length; i++) {
      if (data[i][0].toLowerCase() === trimmedOldEmail) {
        const rowNum = i + 2;
        // ★修正：マジックナンバー(1)を排除し、ヘッダーマップを使用
        userSheet.getRange(rowNum, headers['email'] + 1).setValue(trimmedNewEmail);
        found = true;
        break;
      }
    }

    if (found) {
      return { success: true, message: 'ユーザーアドレスを更新しました。' };
    } else {
      throw new Error(`アドレス「${trimmedOldEmail}」が見つかりませんでした。`);
    }
  } catch (e) {
    logError({ functionName: 'editAllowedUser', message: e.message, stack: e.stack, type: 'Server' });
    throw e;
  } finally {
    lock.releaseLock();
  }
}

function setUserRole(email, role) {
  // Admin権限必須
  Auth.assertAdmin();

  // ★★★ 排他制御 ★★★
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error('サーバーが混み合っています。再試行してください。'); }

  try {
    const ss = SpreadsheetApp.openById(MY_SPREADSHEET_ID);
    const userSheet = ss.getSheetByName(USER_SHEET_NAME);
    if (!userSheet || userSheet.getLastRow() < 2) {
      throw new Error("ユーザーシートが見つからないか、データがありません。");
    }

    const headers = mapHeaders_(userSheet);
    const data = userSheet.getRange(2, headers['email'] + 1, userSheet.getLastRow() - 1, 1).getValues();
    let rowToUpdate = -1;
    for (let i = 0; i < data.length; i++) {
      if (data[i][0].toLowerCase() === email.toLowerCase()) {
        rowToUpdate = i + 2;
        break;
      }
    }

    if (rowToUpdate !== -1) {
      userSheet.getRange(rowToUpdate, headers['role'] + 1).setValue(role);
      logActivity({ action: '役割変更', details: `ユーザー: ${email}, 新しい役割: ${role || '一般'}` });
      return { success: true, message: `ユーザー「${email}」の権限を更新しました。` };
    } else {
      throw new Error(`ユーザー「${email}」が見つかりませんでした。`);
    }
  } catch (e) {
    logError({ functionName: 'setUserRole', message: e.message, stack: e.stack, type: 'Server' });
    throw new Error('役割の設定に失敗しました。');
  } finally {
    lock.releaseLock();
  }
}

function logActivity(activityInfo) {
  try {
    const ss = SpreadsheetApp.openById(MY_SPREADSHEET_ID);
    const logSheet = ss.getSheetByName(ACTIVITY_LOG_SHEET_NAME);
    if (!logSheet) {
      console.warn("「操作ログ」シートが見つかりません。");
      return;
    }
    const headers = mapHeaders_(logSheet);
    const timestamp = new Date();
    const user = Session.getActiveUser().getEmail();

    // ヘッダーに基づいて正しい順序で行データを作成
    const newRow = new Array(logSheet.getLastColumn()).fill('');
    newRow[headers['発生日時']] = timestamp;
    newRow[headers['ユーザー']] = user;
    newRow[headers['操作内容']] = activityInfo.action || '不明な操作';
    newRow[headers['詳細']] = activityInfo.details || '';

    logSheet.appendRow(newRow);
  } catch (e) {
    console.error(`操作ログの記録中にエラー: ${e.toString()}`);
  }
}

function getServiceCenterDetails(qrCodeId) {
  // 読み取りのみ
  if (!qrCodeId) return null;

  const ss = SpreadsheetApp.openById(MY_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(MY_SHEET_NAME);
  if (!sheet) return null;

  const headers = mapHeaders_(sheet);
  const rowNum = findRowByQrCodeId_(sheet, qrCodeId);

  if (rowNum !== -1) {
    const row = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
    return {
      qrCodeId: row[headers['QRコードID']],
      name: row[headers['デイサービス名']],
      readingName: row[headers['読み仮名']] || '',
      address: row[headers['住所']] || '',
      phone: row[headers['電話番号']] || ''
    };
  }
  return null;
}

function mapHeaders_(sheet) {
  try {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    const headerMap = {};
    headers.forEach((header, index) => {
      if (header && String(header).trim() !== '') {
        const cleanedHeader = String(header).replace(/　/g, " ").replace(/\s/g, "");
        headerMap[cleanedHeader] = index;
        // ★追加: 大文字・小文字の違いを吸収するため、小文字化したキーも登録する
        headerMap[cleanedHeader.toLowerCase()] = index;
      }
    });
    return headerMap;
  } catch (e) {
    throw new Error(`ヘッダーの解析中にエラーが発生しました (シート名: ${sheet.getName()})。詳細: ${e.message}`);
  }
}

function generateSelectedCardsView() {
  // Admin権限必須
  Auth.assertAdmin();
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    if (sheet.getName() !== MY_SHEET_NAME) {
      SpreadsheetApp.getUi().alert(`この機能は「${MY_SHEET_NAME}」シートでのみ実行できます。`);
      return;
    }

    const selection = sheet.getActiveRange();
    if (!selection) {
      SpreadsheetApp.getUi().alert('セルを選択してください。');
      return;
    }

    const startRow = selection.getRow();
    const numRows = selection.getNumRows();

    if (startRow === 1 && numRows === 1) {
      SpreadsheetApp.getUi().alert('データ行を選択してください。（1行目のヘッダーは選択できません）');
      return;
    }

    const headers = mapHeaders_(sheet);
    const fullRowData = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getValues();

    const serviceCenters = fullRowData.map(row => {
      return {
        qrCodeId: row[headers['QRコードID']],
        name: row[headers['デイサービス名']],
        address: row[headers['住所']] || '',
        phone: row[headers['電話番号']] || ''
      };
    }).filter(center => center.qrCodeId);

    if (serviceCenters.length === 0) {
      SpreadsheetApp.getUi().alert('選択範囲に有効な事業所データが見つかりません。');
      return;
    }

    const template = HtmlService.createTemplateFromFile('BulkCard');
    template.serviceCenters = serviceCenters;

    const htmlOutput = template.evaluate().setWidth(1200).setHeight(800);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, '選択範囲の事業所カード プレビュー');

  } catch (e) {
    logError({ functionName: 'generateSelectedCardsView', message: e.message, stack: e.stack, type: 'Server' });
    SpreadsheetApp.getUi().alert('カードの生成中にエラーが発生しました。エラーログを確認してください。');
  }
}

function logError(errorInfo) {
  try {
    const ss = SpreadsheetApp.openById(MY_SPREADSHEET_ID);
    const logSheet = ss.getSheetByName(ERROR_LOG_SHEET_NAME);
    if (!logSheet) {
      console.error("致命的エラー: 「エラーログ」シートが見つかりません。");
      return;
    }
    const headers = mapHeaders_(logSheet);
    const timestamp = new Date();
    const user = Session.getActiveUser().getEmail();

    // ヘッダーに基づいて正しい順序で行データを作成
    const newRow = new Array(logSheet.getLastColumn()).fill('');
    newRow[headers['発生日時']] = timestamp;
    newRow[headers['ユーザー']] = user;
    newRow[headers['エラー種別']] = errorInfo.type || 'Unknown';
    newRow[headers['発生関数/場所']] = errorInfo.functionName || 'N/A';
    newRow[headers['エラーメッセージ']] = errorInfo.message || 'No message';
    newRow[headers['スタックトレース']] = errorInfo.stack || 'No stack trace';

    logSheet.appendRow(newRow);
  } catch (e) {
    console.error(`ログ記録関数自体でエラーが発生しました: ${e.toString()}`);
    console.error(`記録しようとした元のエラー情報: ${JSON.stringify(errorInfo)}`);
  }
}

function getActivityLog(limit = 50, filters = {}) {
  // Admin権限必須
  Auth.assertAdmin();
  try {
    const ss = SpreadsheetApp.openById(MY_SPREADSHEET_ID);
    const logSheet = ss.getSheetByName(ACTIVITY_LOG_SHEET_NAME);
    if (!logSheet || logSheet.getLastRow() < 2) { return []; }

    const headers = mapHeaders_(logSheet);
    const allData = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, logSheet.getLastColumn()).getValues();

    const filteredData = allData.filter(row => {
      const userFilter = filters.user ? row[headers['ユーザー']] === filters.user : true;
      const actionFilter = filters.action ? row[headers['操作内容']] === filters.action : true;
      return userFilter && actionFilter;
    });

    const timestampIndex = headers['発生日時'];
    const formattedData = filteredData.map(row => {
      if (timestampIndex !== undefined && row[timestampIndex] instanceof Date) {
        row[timestampIndex] = Utilities.formatDate(row[timestampIndex], 'JST', 'yyyy/MM/dd HH:mm:ss');
      }
      return row;
    });

    return formattedData.reverse().slice(0, limit);

  } catch (e) {
    logError({ functionName: 'getActivityLog', message: e.message, stack: e.stack, type: 'Server' });
    throw new Error('操作ログの取得に失敗しました。');
  }
}

function getErrorLog(limit = 50, filters = {}) {
  // Admin権限必須
  Auth.assertAdmin();
  try {
    const ss = SpreadsheetApp.openById(MY_SPREADSHEET_ID);
    const logSheet = ss.getSheetByName(ERROR_LOG_SHEET_NAME);
    if (!logSheet || logSheet.getLastRow() < 2) { return []; }

    const headers = mapHeaders_(logSheet);
    const allData = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, logSheet.getLastColumn()).getValues();

    const filteredData = allData.filter(row => {
      const userFilter = filters.user ? row[headers['ユーザー']] === filters.user : true;
      const typeFilter = filters.type ? row[headers['エラー種別']] === filters.type : true;
      return userFilter && typeFilter;
    });

    const timestampIndex = headers['発生日時'];
    const formattedData = filteredData.map(row => {
      if (timestampIndex !== undefined && row[timestampIndex] instanceof Date) {
        row[timestampIndex] = Utilities.formatDate(row[timestampIndex], 'JST', 'yyyy/MM/dd HH:mm:ss');
      }
      return row;
    });

    return formattedData.reverse().slice(0, limit);

  } catch (e) {
    logError({ functionName: 'getErrorLog', message: e.message, stack: e.stack, type: 'Server' });
    throw new Error('エラーログの取得に失敗しました。');
  }
}

function clearLogSheets(archive) {
  // Admin権限必須
  Auth.assertAdmin();
  // ★★★ 排他制御 ★★★
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error('サーバーが混み合っています。再試行してください。'); }

  try {
    const ss = SpreadsheetApp.openById(MY_SPREADSHEET_ID);
    const sheetNames = [ACTIVITY_LOG_SHEET_NAME, ERROR_LOG_SHEET_NAME];
    const timestamp = Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmm');

    sheetNames.forEach(name => {
      const sheet = ss.getSheetByName(name);
      if (sheet && sheet.getLastRow() > 1) {
        // アーカイブが指定されている場合、シートをコピーしてリネーム保存
        if (archive) {
          const archiveName = `${name}_${timestamp}`;
          sheet.copyTo(ss).setName(archiveName);
        }
        sheet.deleteRows(2, sheet.getLastRow() - 1);
      }
    });

    logActivity({ action: 'ログ削除', details: `管理者がログを手動でクリアしました (アーカイブ: ${archive ? 'あり' : 'なし'})` });
    return true;
  } catch (e) {
    logError({ functionName: 'clearLogSheets', message: e.message, stack: e.stack });
    throw new Error('ログの削除に失敗しました: ' + e.message);
  } finally {
    lock.releaseLock();
  }
}

function getLogFilterOptions() {
  // Admin権限必須
  Auth.assertAdmin();
  try {
    const ss = SpreadsheetApp.openById(MY_SPREADSHEET_ID);
    const activitySheet = ss.getSheetByName(ACTIVITY_LOG_SHEET_NAME);
    const errorSheet = ss.getSheetByName(ERROR_LOG_SHEET_NAME);

    const uniqueValues = (sheet, columnIndex) => {
      const lastRow = sheet.getLastRow();
      if (!sheet || lastRow < 2) return [];
      const data = sheet.getRange(2, columnIndex, lastRow - 1, 1).getValues();
      return [...new Set(data.flat().filter(String))].sort();
    };

    return {
      activityUsers: uniqueValues(activitySheet, mapHeaders_(activitySheet)['ユーザー'] + 1),
      activityActions: uniqueValues(activitySheet, mapHeaders_(activitySheet)['操作内容'] + 1),
      errorUsers: uniqueValues(errorSheet, mapHeaders_(errorSheet)['ユーザー'] + 1),
      errorTypes: uniqueValues(errorSheet, mapHeaders_(errorSheet)['エラー種別'] + 1)
    };
  } catch (e) {
    logError({ functionName: 'getLogFilterOptions', message: e.message, stack: e.stack, type: 'Server' });
    return { activityUsers: [], activityActions: [], errorUsers: [], errorTypes: [] };
  }
}

/**
 * AI設定シートから、キー名（A列）に対応する値（B列）を取得する
 * ★修正：createTextFinderをやめ、配列処理に変更して確実性を向上
 */
function getAiConfigValue(keyName) {
  console.log(`[getAiConfigValue] 検索キー: "${keyName}"`);
  try {
    const ss = SpreadsheetApp.openById(MY_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(AI_CONFIG_SHEET_NAME);
    if (!sheet) {
      console.warn(`[getAiConfigValue] シート "${AI_CONFIG_SHEET_NAME}" が見つかりません。`);
      return '';
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 1) {
      console.warn(`[getAiConfigValue] シート "${AI_CONFIG_SHEET_NAME}" は空です。`);
      return '';
    }

    // A列とB列の値を一括取得
    const data = sheet.getRange(1, 1, lastRow, 2).getValues();

    // シート上の全キーをログ出力（デバッグ用）
    const existingKeys = data.map(row => String(row[0]).trim());
    console.log(`[getAiConfigValue] シート上のキー一覧: ${JSON.stringify(existingKeys)}`);

    // ★修正: 保存ロジック（後勝ち）に合わせて、下から上に検索して最新の設定を採用する
    for (let i = data.length - 1; i >= 0; i--) {
      // キーの前後の空白を削除して比較
      if (String(data[i][0]).trim() === keyName) {
        console.log(`[getAiConfigValue] ヒットしました: 行=${i + 1}, 値(先頭20文字)="${String(data[i][1]).substring(0, 20)}..."`);
        return data[i][1]; // B列の値を返す
      }
    }
    console.warn(`[getAiConfigValue] キー "${keyName}" は見つかりませんでした。`);
    return '';
  } catch (e) {
    console.error(`AI設定取得エラー (${keyName}): ${e.message}`);
    return '';
  }
}

/**
 * AIに分析コメントをリクエストする関数
 * ★修正：動的取得(getAiConfigValue)とセカンダリープロンプト対応
 */
function getAiAnalysisComment(reportData, lang) {
  // Teacher以上
  Auth.assertTeacherOrAdmin();

  let primaryPrompt = getAiConfigValue("プライマリープロンプト");
  const secondaryPrompt = getAiConfigValue("セカンダリープロンプト");

  const defaultPrompt = `あなたはデータ分析の専門家です。
提供されたデイサービスの到着時刻データを分析し、傾向や特徴を要約してください。
特に、到着が集中する時間帯や、遅れがちな傾向があれば指摘してください。
また、業務改善に向けた簡単なアドバイスもあれば含めてください。`;

  if (!primaryPrompt) {
    console.warn('プライマリープロンプトが見つかりません。デフォルト値を使用します。');
    primaryPrompt = defaultPrompt;
  }

  const apiKey = GEMINI_API_KEY;
  if (!apiKey) {
    throw new Error('AIのAPIキーが設定されていません。');
  }

  // ★修正: シートからモデル名を取得するように変更
  let modelId = getAiConfigValue("使用するAIモデル名");
  if (!modelId) {
    console.warn('AIモデル名が指定されていません。デフォルト(gemini-2.5-flash)を使用します。');
    modelId = "gemini-2.5-flash";
  }
  // APIの形式に合わせて "models/" がなければ付与する
  const modelName = modelId.startsWith("models/") ? modelId : `models/${modelId}`;

  const url = `https://generativelanguage.googleapis.com/v1beta/${modelName}:generateContent?key=${apiKey}`;

  const translations = {
    ja: {
      provider: '事業所名',
      arrivals: '回',
      noData: '該当期間のデータはありませんでした。\n',
      instruction: '日本語で回答してください。\n'
    },
    en: {
      provider: 'Provider',
      arrivals: ' arrivals',
      noData: 'No data for the selected period.\n',
      instruction: 'Please respond in English. In your response, please translate the word "デイサービス" as "day service".\n'
    },
    es: {
      provider: 'Proveedor',
      arrivals: ' llegadas',
      noData: 'No hay datos para el período seleccionado.\n',
      instruction: 'Por favor, responde en español. En tu respuesta, traduce la palabra "デイサービス" como "servicio de día".\n'
    },
    zh: {
      provider: '服务提供商',
      arrivals: '次到达',
      noData: '选定期间内无数据。\n',
      instruction: '请用中文回答。在回答中，请将“デイサービス”这个词翻译成“日间照料服务”。\n'
    }
  };

  const dict = translations[lang] || translations.en;

  // プロンプトの組み立て
  let finalPrompt = primaryPrompt;

  let dataText = '';
  for (const serviceName in reportData) {
    dataText += `■ ${dict.provider}: ${serviceName}\n`;
    const timeSlots = reportData[serviceName];
    const sortedSlots = Object.keys(timeSlots).sort();
    for (const slot of sortedSlots) {
      dataText += `- ${slot}: ${timeSlots[slot]}${dict.arrivals}\n`;
    }
  }

  finalPrompt += "\n\n--- 分析対象データ ---\n";
  finalPrompt += (dataText === '') ? dict.noData : dataText;
  finalPrompt += "------------------\n\n";

  // セカンダリープロンプトがあれば追加
  if (secondaryPrompt) {
    finalPrompt += "\n" + secondaryPrompt + "\n";
  }

  finalPrompt += dict.instruction;

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify({ "contents": [{ "parts": [{ "text": finalPrompt }] }] }),
    'muteHttpExceptions': true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const resultText = response.getContentText();
    if (response.getResponseCode() === 200) {
      const result = JSON.parse(resultText);
      if (result.candidates && result.candidates[0] && result.candidates[0].content && result.candidates[0].content.parts[0]) {
        return result.candidates[0].content.parts[0].text;
      } else {
        throw new Error(`AIからの応答形式が不正です: ${resultText}`);
      }
    } else {
      throw new Error(`AIからの応答が不正です - ${resultText}`);
    }
  } catch (e) {
    logError({ functionName: 'getAiAnalysisComment', message: e.message, stack: e.stack, type: 'Server' });
    throw new Error(`AIによる分析コメントの生成に失敗しました。詳細: ${e.message}`);
  }
}

/**
 * 管理画面用に、保存されているAIマスタープロンプトを取得する
 * ★修正：マスターとセカンダリーの両方を返すように変更
 */
function getAiPrompt() {
  // Admin権限必須
  Auth.assertAdmin();

  console.log("[getAiPrompt] プロンプト取得開始");
  let masterPrompt = getAiConfigValue("プライマリープロンプト");
  const secondaryPrompt = getAiConfigValue("セカンダリープロンプト");

  console.log(`[getAiPrompt] 取得結果 - Master: ${Boolean(masterPrompt)}, Secondary: ${Boolean(secondaryPrompt)}`);

  // マスタープロンプトが空の場合、デフォルト値を設定する
  if (!masterPrompt) {
    masterPrompt = `あなたはデータ分析の専門家です。
提供されたデイサービスの到着時刻データを分析し、傾向や特徴を要約してください。
特に、到着が集中する時間帯や、遅れがちな傾向があれば指摘してください。
また、業務改善に向けた簡単なアドバイスもあれば含めてください。`;
  }

  return {
    master: masterPrompt,
    secondary: secondaryPrompt
  };
}

/**
 * 管理画面から送信されたAIマスタープロンプトをスプレッドシートに保存する
 * ★修正：配列処理に変更して確実性を向上
 */
function saveAiPrompt(prompts) {
  // ★★★ 排他制御 ★★★
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error('サーバーが混み合っています。再試行してください。'); }

  try {
    const userEmail = Session.getActiveUser().getEmail();
    const user = findUserByEmail(userEmail);

    if (!user || user.role !== 'admin') {
      throw new Error('この操作を行う権限がありません。');
    }

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName(AI_CONFIG_SHEET_NAME);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(AI_CONFIG_SHEET_NAME);
    }

    const lastRow = sheet.getLastRow();
    // A列の値を一括取得して検索用マップを作成
    const keyMap = new Map();
    if (lastRow > 0) {
      const keys = sheet.getRange(1, 1, lastRow, 1).getValues();
      keys.forEach((row, index) => {
        keyMap.set(String(row[0]).trim(), index + 1); // 1-based index
      });
    }

    // 値を保存するヘルパー関数
    const saveValue = (key, value) => {
      if (keyMap.has(key)) {
        const rowIndex = keyMap.get(key);
        sheet.getRange(rowIndex, 2).setValue(value);
      } else {
        sheet.appendRow([key, value]);
        // 新しく追加された行をマップにも反映（念のため）
        keyMap.set(key, sheet.getLastRow());
      }
    };

    saveValue("プライマリープロンプト", prompts.master);
    saveValue("セカンダリープロンプト", prompts.secondary);

  } catch (e) {
    logError({ functionName: 'saveAiPrompt', message: e.message, stack: e.stack, type: 'Server' });
    throw new Error(`プロンプトの保存に失敗しました: ${e.message}`);
  } finally {
    lock.releaseLock();
  }
}

function findUserByEmail(email) {
  // Authモジュールのラッパーとして残す（互換性のため）
  const result = Auth._resolveUserFromSheet(email);
  return result.isAllowed ? { email: result.email, role: result.role } : null;
}

function getCharacters() {
  // 読み取りのみ
  Auth.assertLogin();
  try {
    const sheet = SpreadsheetApp.openById(MY_SPREADSHEET_ID).getSheetByName(CHARACTER_SHEET_NAME);
    const headers = mapHeaders_(sheet);
    if (!sheet || sheet.getLastRow() < 2) return [];
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    return data.map(row => ({ name: row[headers['キャラクター名']], folderId: row[headers['音声フォルダID']] })).filter(c => c.name && c.folderId);
  } catch (e) {
    logError({ functionName: 'getCharacters', message: e.message, stack: e.stack });
    throw new Error('キャラクターリストの取得に失敗しました。');
  }
}

function saveCharacter(characterData) {
  // Admin権限必須
  Auth.assertAdmin();
  // ★★★ 排他制御 ★★★
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error('サーバーが混み合っています。再試行してください。'); }

  try {
    const { name, folderId } = characterData;
    if (!name || !folderId) throw new Error('キャラクター名とフォルダIDは必須です。');

    const sheet = SpreadsheetApp.openById(MY_SPREADSHEET_ID).getSheetByName(CHARACTER_SHEET_NAME);
    const headers = mapHeaders_(sheet);
    const names = sheet.getRange(2, headers['キャラクター名'] + 1, sheet.getLastRow() - 1, 1).getValues().flat();
    const rowIndex = names.findIndex(n => n === name);

    if (rowIndex !== -1) {
      sheet.getRange(rowIndex + 2, headers['音声フォルダID'] + 1).setValue(folderId);
    } else {
      const newRow = new Array(sheet.getLastColumn()).fill('');
      newRow[headers['キャラクター名']] = name; newRow[headers['音声フォルダID']] = folderId;
      sheet.appendRow(newRow);
    }
    logActivity({ action: 'キャラクター保存', details: `名前: ${name}` });
    return true;
  } catch (e) {
    logError({ functionName: 'saveCharacter', message: e.message, stack: e.stack });
    throw new Error('キャラクターの保存に失敗しました。');
  } finally {
    lock.releaseLock();
  }
}

function deleteCharacter(characterName) {
  // Admin権限必須
  Auth.assertAdmin();
  // ★★★ 排他制御 ★★★
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error('サーバーが混み合っています。再試行してください。'); }

  try {
    const sheet = SpreadsheetApp.openById(MY_SPREADSHEET_ID).getSheetByName(CHARACTER_SHEET_NAME);
    const headers = mapHeaders_(sheet);
    const names = sheet.getRange(2, headers['キャラクター名'] + 1, sheet.getLastRow() - 1, 1).getValues().flat();
    const rowIndex = names.findIndex(n => n === characterName);

    if (rowIndex !== -1) {
      sheet.deleteRow(rowIndex + 2);
      logActivity({ action: 'キャラクター削除', details: `名前: ${characterName}` });
      return true;
    } else {
      throw new Error('削除対象のキャラクターが見つかりません。');
    }
  } catch (e) {
    logError({ functionName: 'deleteCharacter', message: e.message, stack: e.stack });
    throw new Error('キャラクターの削除に失敗しました。');
  } finally {
    lock.releaseLock();
  }
}

function getWeeklySchedule() {
  // 読み取りのみ
  Auth.assertLogin();
  try {
    const sheet = SpreadsheetApp.openById(MY_SPREADSHEET_ID).getSheetByName(SCHEDULE_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) return {};
    const headers = mapHeaders_(sheet);
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    const schedule = {};

    data.forEach(row => {
      const dayKey = row[headers['曜日']];
      const characterName = row[headers['担当キャラクター']];
      if (dayKey) schedule[dayKey] = characterName || '';
    });
    return schedule;
  } catch (e) {
    logError({ functionName: 'getWeeklySchedule', message: e.message, stack: e.stack });
    throw new Error('週間スケジュールの取得に失敗しました。');
  }
}

function saveWeeklySchedule(scheduleData) {
  // Admin権限必須
  Auth.assertAdmin();
  // ★★★ 排他制御 ★★★
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error('サーバーが混み合っています。再試行してください。'); }

  try {
    const sheet = SpreadsheetApp.openById(MY_SPREADSHEET_ID).getSheetByName(SCHEDULE_SHEET_NAME);
    if (!sheet) throw new Error(`シート「${SCHEDULE_SHEET_NAME}」が見つかりません。`);

    const headers = mapHeaders_(sheet);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return true;

    const range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
    const values = range.getValues();

    for (let i = 0; i < values.length; i++) {
      const dayKey = values[i][headers['曜日']];
      if (scheduleData.hasOwnProperty(dayKey)) {
        values[i][headers['担当キャラクター']] = scheduleData[dayKey] || '';
      }
    }
    range.setValues(values);

    logActivity({ action: 'スケジュール保存', details: '週間スケジュールが更新されました。' });
    CacheService.getScriptCache().remove('soundConfigCache');
    return true;
  } catch (e) {
    logError({ functionName: 'saveWeeklySchedule', message: e.message, stack: e.stack });
    throw new Error('週間スケジュールの保存に失敗しました。');
  } finally {
    lock.releaseLock();
  }
}

function getAllServiceDataForAdmin() {
  // Admin権限必須
  Auth.assertAdmin();
  const sheet = SpreadsheetApp.openById(MY_SPREADSHEET_ID).getSheetByName(MY_SHEET_NAME);
  const headers = mapHeaders_(sheet);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  return data.map(row => ({
    qrCodeId: row[headers['QRコードID']],
    name: row[headers['デイサービス名']],
    reading: row[headers['読み仮名']] || '',
    displayStatus: row[headers['表示設定']] || '表示'
  })).filter(s => s.qrCodeId && s.name);
}

function updateDisplayStatus(qrCodeId, newStatus) {
  // Admin権限必須
  Auth.assertAdmin();

  // ★★★ 排他制御 ★★★
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error('サーバーが混み合っています。再試行してください。'); }

  try {
    const sheet = SpreadsheetApp.openById(MY_SPREADSHEET_ID).getSheetByName(MY_SHEET_NAME);
    const headers = mapHeaders_(sheet);
    const data = sheet.getRange(2, headers['QRコードID'] + 1, sheet.getLastRow() - 1, 1).getValues();

    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === qrCodeId) {
        const rowNum = i + 2;
        sheet.getRange(rowNum, headers['表示設定'] + 1).setValue(newStatus);

        logActivity({
          action: '表示設定変更',
          details: `ID: ${qrCodeId}, 新ステータス: ${newStatus}`
        });

        CacheService.getScriptCache().remove(CACHE_KEY);
        return true;
      }
    }
    throw new Error('対象の事業所が見つかりませんでした。');
  } catch (e) {
    logError({ functionName: 'updateDisplayStatus', message: e.message, stack: e.stack });
    throw new Error('表示設定の更新に失敗しました。');
  } finally {
    lock.releaseLock();
  }
}

function importFromCsv(fileContent, importType) {
  // Admin権限必須
  Auth.assertAdmin();

  // ★★★ 排他制御 ★★★
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error('サーバーが混み合っています。再試行してください。'); }

  try {
    if (!fileContent) throw new Error("ファイルの中身が空です。");

    const rows = Utilities.parseCsv(fileContent);
    const dataRows = rows.filter(row => row.length > 0 && !String(row[0]).startsWith('#'));
    const headerRow = dataRows.shift();
    if (!headerRow) throw new Error("CSVにヘッダー行がありません。");

    const headerMap = {};
    headerRow.forEach((h, i) => { headerMap[h.trim()] = i; });

    let results = { totalRows: dataRows.length, successCount: 0, failureCount: 0, errors: [] };

    dataRows.forEach((row, index) => {
      try {
        if (importType === 'users') {
          const email = row[headerMap['email']];
          const type = row[headerMap['type']];
          const role = row[headerMap['role']];
          if (!email || email.trim() === '') return;
          addAllowedUser({ email: email, type: type, role: role });
        } else if (importType === 'services') {
          const name = row[headerMap['デイサービス名']];
          if (!name || name.trim() === '') return;
          const reading = headerMap['読み仮名'] !== undefined ? row[headerMap['読み仮名']] : '';
          const address = headerMap['住所'] !== undefined ? row[headerMap['住所']] : '';
          const phone = headerMap['電話番号'] !== undefined ? row[headerMap['電話番号']] : '';
          addServiceCenter(name, reading, address, phone);
        }
        results.successCount++;
      } catch (e) {
        results.failureCount++;
        results.errors.push(`CSVの${index + 2}行目: ${e.message}`);
      }
    });

    logActivity({ action: 'CSV一括登録', details: `${importType} - ${results.successCount}/${results.totalRows}件成功` });
    return results;
  } catch (e) {
    logError({ functionName: 'importFromCsv', message: e.message, stack: e.stack, type: 'Server' });
    throw e;
  } finally {
    lock.releaseLock();
  }
}

function getUserCsvTemplate() {
  const bom = "\uFEFF";
  const header = "email,type,role,※1行目の項目名はそのまま。";
  const examples = [
    "user1@example.com,user,admin,※2行目以降に入力（この例は消してください）。email必須。typeはuser或いはgroup。roleはadmin/teacher(一般は空欄)。",
    "teachers@example.com,group,teacher,",
    "user2@example.com,user,,",
    "parents@example.com,group,,"
  ];
  return bom + header + "\n" + examples.join("\n");
}

function getServiceCsvTemplate() {
  const bom = "\uFEFF";
  const header = "デイサービス名,読み仮名,住所,電話番号,※1行目の項目名はそのまま。";
  const example = "〇〇デイサービス,まるまるでいさーびす,大阪府〇〇市1-2-3,06-0000-0000,※2行目以降に入力（この例は消してください）。デイサービス名のみ必須です。";

  return bom + header + "\n" + example;
}

/**
 * 【移行用】スクリプトプロパティに初期値を一括設定する関数
 * この関数をエディタ上で一度だけ実行してください。
 * 実行後、この関数は削除しても構いません。
 */
function setupInitialProperties() {
  const props = PropertiesService.getScriptProperties();
  const currentProps = props.getProperties();

  // コードから抽出できた既知のID（これらは強制的に設定・更新します）
  const updates = {
    'SPREADSHEET_ID': '1rpviVe7dwM1pa8XblpTWSKrp7c2pi5ILAgXFPpHRqP0',
    'QR_CODE_FOLDER_ID': '1ewReiKtX1tUiwBSsI47cB2R_vBRO0NqF'
  };

  // 以下の項目は、既存の設定がない場合のみプレースホルダーまたは初期値を設定します
  if (!currentProps['GEMINI_API_KEY']) {
    updates['GEMINI_API_KEY'] = '【ここにGemini APIキーを入力してください】';
  }
  if (!currentProps['ERROR_SOUND_ID']) {
    updates['ERROR_SOUND_ID'] = '【ここにエラー音ファイルのIDを入力してください】';
  }
  if (!currentProps['kioskTouchEnabled']) {
    updates['kioskTouchEnabled'] = 'false'; // デフォルトは無効
  }

  // 既存のプロパティを維持しつつ更新
  props.setProperties(updates, false);

  console.log('以下の設定をスクリプトプロパティに保存しました:\n' + JSON.stringify(updates, null, 2));
  console.log('重要: GEMINI_API_KEY と ERROR_SOUND_ID はプレースホルダーが設定された可能性があります。スクリプトプロパティ画面で正しい値を確認・編集してください。');
}

function updateCharacter(oldName, newData) {
  // Admin権限必須
  Auth.assertAdmin();

  // ★★★ 排他制御 ★★★
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error('サーバーが混み合っています。再試行してください。'); }

  try {
    const { newName, newFolderId } = newData;
    if (!oldName || !newName || !newFolderId) {
      throw new Error("必要な情報が不足しています。");
    }

    const ss = SpreadsheetApp.openById(MY_SPREADSHEET_ID).getSheetByName(CHARACTER_SHEET_NAME);
    const headers = mapHeaders_(ss);
    const data = ss.getRange(2, headers['キャラクター名'] + 1, ss.getLastRow() - 1, 1).getValues().flat();

    const rowIndex = data.findIndex(name => name === oldName);

    if (rowIndex === -1) {
      throw new Error("更新対象のキャラクターが見つかりません。");
    }

    ss.getRange(rowIndex + 2, headers['キャラクター名'] + 1).setValue(newName);
    ss.getRange(rowIndex + 2, headers['音声フォルダID'] + 1).setValue(newFolderId);

    const scheduleSheet = ss.getSheetByName(SCHEDULE_SHEET_NAME);
    const scheduleHeaders = mapHeaders_(scheduleSheet);
    const scheduleData = scheduleSheet.getRange(2, scheduleHeaders['担当キャラクター'] + 1, 7, 1).getValues();
    scheduleData.forEach((row, index) => {
      if (row[0] === oldName) {
        scheduleSheet.getRange(index + 2, scheduleHeaders['担当キャラクター'] + 1).setValue(newName);
      }
    });

    logActivity({ action: 'キャラクター更新', details: `変更前: ${oldName}, 変更後: ${newName}` });
    CacheService.getScriptCache().remove('soundConfigCache');
    return true;

  } catch (e) {
    logError({ functionName: 'updateCharacter', message: e.message, stack: e.stack });
    throw new Error('キャラクターの更新に失敗しました。');
  } finally {
    lock.releaseLock();
  }
}

function findRowByQrCodeId_(sheet, qrCodeId) {
  if (!sheet || !qrCodeId) return -1;
  const headers = mapHeaders_(sheet);
  const data = sheet.getRange(2, headers['QRコードID'] + 1, sheet.getLastRow() - 1, 1).getValues();
  const trimmedQrCodeId = String(qrCodeId).trim();
  for (let i = 0; i < data.length; i++) {
    const currentId = String(data[i][0]).trim();
    if (currentId === trimmedQrCodeId) {
      return i + 2;
    }
  }
  return -1;
}

function isUserInGroup_(userEmail, groupEmail) {
  try {
    const group = GroupsApp.getGroupByEmail(groupEmail);
    return group.hasUser(userEmail);
  } catch (e) {
    console.warn(`isUserInGroup_ Error: Failed to check group '${groupEmail}' for user '${userEmail}'. Details: ${e.message}`);
    return false;
  }
}

function findRowByEmail_(sheet, email) {
  if (!sheet || !email) return -1;
  const headers = mapHeaders_(sheet);
  const data = sheet.getRange(2, headers['email'] + 1, sheet.getLastRow() - 1, 1).getValues();
  const lowerCaseEmail = email.toLowerCase();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === lowerCaseEmail) {
      return i + 1;
    }
  }
  return -1;
}

function addLateArrivalNote(qrCodeId, note) {
  // Teacher以上
  Auth.assertTeacherOrAdmin();

  // ★★★ 排他制御 ★★★
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error('サーバーが混み合っています。再試行してください。'); }

  try {
    if (!qrCodeId || !note) {
      throw new Error("事業所と連絡内容の両方が必要です。");
    }

    const sheet = SpreadsheetApp.openById(MY_SPREADSHEET_ID).getSheetByName(MY_SHEET_NAME);
    const headers = mapHeaders_(sheet);
    const qrIdColumn = headers['QRコードID'] + 1;
    const lateNoteColumn = headers['遅延連絡'] + 1;

    const finder = sheet.getRange(2, qrIdColumn, sheet.getLastRow() - 1, 1)
      .createTextFinder(qrCodeId)
      .matchEntireCell(true);

    const foundCell = finder.findNext();

    if (foundCell) {
      sheet.getRange(foundCell.getRow(), lateNoteColumn).setValue(note);
      logActivity({ action: '遅延連絡入力', details: `ID: ${qrCodeId}, 内容: ${note}` });
      CacheService.getScriptCache().remove(CACHE_KEY);
      return true;
    }
    throw new Error('対象の事業所が見つかりませんでした。');
  } catch (e) {
    logError({ functionName: 'addLateArrivalNote', message: e.message, stack: e.stack });
    throw new Error('遅延連絡の保存に失敗しました。');
  } finally {
    lock.releaseLock();
  }
}

function getLateNote(qrCodeId) {
  try {
    if (!qrCodeId) return '';
    const sheet = SpreadsheetApp.openById(MY_SPREADSHEET_ID).getSheetByName(MY_SHEET_NAME);
    const rowNum = findRowByQrCodeId_(sheet, qrCodeId);
    if (rowNum !== -1) {
      const headers = mapHeaders_(sheet);
      return sheet.getRange(rowNum, headers['遅延連絡'] + 1).getValue();
    }
    return '';
  } catch (e) {
    logError({ functionName: 'getLateNote', message: e.message, stack: e.stack });
    return '';
  }
}

function getServiceCenterDetailsByIds_(qrCodeIds) {
  if (!qrCodeIds || qrCodeIds.length === 0) return [];

  const ss = SpreadsheetApp.openById(MY_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(MY_SHEET_NAME);
  if (!sheet) return [];

  const headers = mapHeaders_(sheet);
  const allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

  const results = [];
  const idSet = new Set(qrCodeIds);

  for (const row of allData) {
    const currentId = row[headers['QRコードID']];
    if (idSet.has(currentId)) {
      results.push({
        qrCodeId: currentId,
        name: row[headers['デイサービス名']] || '',
        readingName: row[headers['読み仮名']] || '',
        address: row[headers['住所']] || '',
        phone: row[headers['電話番号']] || ''
      });
    }
  }

  return qrCodeIds.map(id => results.find(item => item.qrCodeId === id)).filter(Boolean);
}

// ===============================================================
// ★★★ 年度更新・一括処理機能 ★★★
// ===============================================================

function getQrPrefix() {
  Auth.assertAdmin();
  return getQrIdPrefix_();
}

function setQrPrefix(prefix) {
  Auth.assertAdmin();
  if (!prefix || prefix.trim() === '') throw new Error('接頭辞を入力してください。');
  PROPS_.setProperty('QR_ID_PREFIX', prefix.trim());
  logActivity({ action: '設定変更', details: `QRコードID接頭辞を「${prefix.trim()}」に変更` });
  return true;
}

function bulkUpdateDisplayStatus(qrCodeIds, newStatus) {
  Auth.assertAdmin();
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error('サーバーが混み合っています。再試行してください。'); }
  try {
    const sheet = SpreadsheetApp.openById(MY_SPREADSHEET_ID).getSheetByName(MY_SHEET_NAME);
    const headers = mapHeaders_(sheet);
    const data = sheet.getRange(2, headers['QRコードID'] + 1, sheet.getLastRow() - 1, 1).getValues();
    const idSet = new Set(qrCodeIds);
    let count = 0;
    for (let i = 0; i < data.length; i++) {
      if (idSet.has(String(data[i][0]))) {
        sheet.getRange(i + 2, headers['表示設定'] + 1).setValue(newStatus);
        count++;
      }
    }
    logActivity({ action: '事業所一括設定', details: `${count}件の表示設定を「${newStatus}」に変更` });
    CacheService.getScriptCache().remove(CACHE_KEY);
    return count;
  } catch (e) {
    logError({ functionName: 'bulkUpdateDisplayStatus', message: e.message, stack: e.stack });
    throw new Error('一括更新に失敗しました。');
  } finally { lock.releaseLock(); }
}

function bulkDeleteServiceCenters(qrCodeIds) {
  Auth.assertAdmin();
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error('サーバーが混み合っています。再試行してください。'); }
  try {
    const sheet = SpreadsheetApp.openById(MY_SPREADSHEET_ID).getSheetByName(MY_SHEET_NAME);
    const headers = mapHeaders_(sheet);
    const data = sheet.getRange(2, headers['QRコードID'] + 1, sheet.getLastRow() - 1, 1).getValues();
    const idSet = new Set(qrCodeIds);
    let count = 0;
    // 削除時のズレを防ぐため下からループ
    for (let i = data.length - 1; i >= 0; i--) {
      if (idSet.has(String(data[i][0]))) { sheet.deleteRow(i + 2); count++; }
    }
    logActivity({ action: '事業所一括削除', details: `${count}件の事業所を削除` });
    CacheService.getScriptCache().remove(CACHE_KEY);
    return count;
  } catch (e) {
    logError({ functionName: 'bulkDeleteServiceCenters', message: e.message, stack: e.stack });
    throw new Error('一括削除に失敗しました。');
  } finally { lock.releaseLock(); }
}

function archiveHiddenServices() {
  Auth.assertAdmin();
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error('サーバーが混み合っています。再試行してください。'); }

  try {
    const ss = SpreadsheetApp.openById(MY_SPREADSHEET_ID);
    const masterSheet = ss.getSheetByName(MY_SHEET_NAME);
    let archiveSheet = ss.getSheetByName('事業所アーカイブ');

    if (!archiveSheet) {
      archiveSheet = ss.insertSheet('事業所アーカイブ');
      const headers = masterSheet.getRange(1, 1, 1, masterSheet.getLastColumn()).getValues();
      const newHeaders = [...headers[0], 'アーカイブ日時'];
      archiveSheet.appendRow(newHeaders);
    } else {
      const lastCol = archiveSheet.getLastColumn();
      if (lastCol > 0 && archiveSheet.getRange(1, lastCol).getValue() !== 'アーカイブ日時') {
        archiveSheet.getRange(1, lastCol + 1).setValue('アーカイブ日時');
      }
    }

    const headers = mapHeaders_(masterSheet);
    const displayColIndex = headers['表示設定'] + 1; // 1-based index
    const data = masterSheet.getDataRange().getValues();
    const now = new Date();

    let archiveCount = 0;
    // 下からループして削除時の行ズレを防ぐ
    for (let i = data.length - 1; i >= 1; i--) { // i=0はヘッダー
      if (String(data[i][displayColIndex - 1]).trim() === '非表示') {
        const rowToArchive = [...data[i], now];
        archiveSheet.appendRow(rowToArchive);
        masterSheet.deleteRow(i + 1);
        archiveCount++;
      }
    }
    if (archiveCount > 0) { CacheService.getScriptCache().remove(CACHE_KEY); logActivity({ action: '事業所アーカイブ', details: `${archiveCount}件の非表示事業所をアーカイブに退避` }); }
    return archiveCount;
  } catch (e) {
    logError({ functionName: 'archiveHiddenServices', message: e.message, stack: e.stack });
    throw new Error('アーカイブ処理に失敗しました。');
  } finally { lock.releaseLock(); }
}

function archiveAllServices() {
  Auth.assertAdmin();
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error('サーバーが混み合っています。再試行してください。'); }

  try {
    const ss = SpreadsheetApp.openById(MY_SPREADSHEET_ID);
    const masterSheet = ss.getSheetByName(MY_SHEET_NAME);
    let archiveSheet = ss.getSheetByName('事業所アーカイブ');

    if (!archiveSheet) {
      archiveSheet = ss.insertSheet('事業所アーカイブ');
      const headers = masterSheet.getRange(1, 1, 1, masterSheet.getLastColumn()).getValues();
      const newHeaders = [...headers[0], 'アーカイブ日時'];
      archiveSheet.appendRow(newHeaders);
    } else {
      const lastCol = archiveSheet.getLastColumn();
      if (lastCol > 0 && archiveSheet.getRange(1, lastCol).getValue() !== 'アーカイブ日時') {
        archiveSheet.getRange(1, lastCol + 1).setValue('アーカイブ日時');
      }
    }

    const data = masterSheet.getDataRange().getValues();
    let archiveCount = 0;
    const now = new Date();

    // データ行（2行目以降）をすべてアーカイブシートへ退避
    for (let i = 1; i < data.length; i++) {
      const rowToArchive = [...data[i], now];
      archiveSheet.appendRow(rowToArchive);
      archiveCount++;
    }

    // マスタシートのデータ行を一括削除（1行目のヘッダーのみ残す）
    if (archiveCount > 0) {
      masterSheet.deleteRows(2, masterSheet.getLastRow() - 1);
      CacheService.getScriptCache().remove(CACHE_KEY);
      logActivity({ action: '年度更新リセット', details: `全事業所（${archiveCount}件）をアーカイブし、マスタをリセット` });
    }
    return archiveCount;
  } catch (e) {
    logError({ functionName: 'archiveAllServices', message: e.message, stack: e.stack });
    throw new Error('全件アーカイブ処理に失敗しました。');
  } finally { lock.releaseLock(); }
}
/**
 * スプレッドシートの自動セットアップおよび非破壊アップデートを行う関数
 */
function setupOrCreateSpreadsheetSheets() {
  const props = PropertiesService.getScriptProperties();

  // スクリプトプロパティが未設定の場合、自動的に登録します（フォールバック）
  let spreadsheetId = props.getProperty('SPREADSHEET_ID');
  if (!spreadsheetId) {
    spreadsheetId = '1nDZrvED6tu_BgdRT4QgZwBVAMXmIZBNSU-iueaNuntE';
    props.setProperty('SPREADSHEET_ID', spreadsheetId);
    console.log(`プロパティ SPREADSHEET_ID を自動登録しました: ${spreadsheetId}`);
  }

  let folderId = props.getProperty('QR_CODE_FOLDER_ID');
  if (!folderId) {
    folderId = '13WSJm5_xURSwgexVlVQ4kEQ1UI0RTe4D';
    props.setProperty('QR_CODE_FOLDER_ID', folderId);
    console.log(`プロパティ QR_CODE_FOLDER_ID を自動登録しました: ${folderId}`);
  }

  const ss = SpreadsheetApp.openById(spreadsheetId);
  const userEmail = Session.getActiveUser().getEmail().toLowerCase();

  // 各シートで必要なヘッダー定義
  const schema = {
    [MY_SHEET_NAME]: ['QRコードID', 'デイサービス名', '状態', '到着時刻', '遅延連絡', '読み仮名', '住所', '電話番号', '表示設定'],
    [USER_SHEET_NAME]: ['email', 'type', 'role'],
    [HISTORY_SHEET_NAME]: ['タイムスタンプ', 'QRコードID', '事業所名'],
    [SCHEDULE_SHEET_NAME]: ['曜日', '担当キャラクター'],
    [CHARACTER_SHEET_NAME]: ['キャラクター名', '音声フォルダID'],
    [ACTIVITY_LOG_SHEET_NAME]: ['発生日時', 'ユーザー', '操作内容', '詳細'],
    [ERROR_LOG_SHEET_NAME]: ['発生日時', 'ユーザー', 'エラー種別', '発生関数/場所', 'エラーメッセージ', 'スタックトレース'],
    [AI_CONFIG_SHEET_NAME]: ['キー', '値']
  };

  for (const sheetName in schema) {
    let sheet = ss.getSheetByName(sheetName);
    const requiredHeaders = schema[sheetName];

    if (!sheet) {
      // 1. シートが存在しない場合は新規作成してヘッダーを書き込む
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(requiredHeaders);
      sheet.getRange(1, 1, 1, requiredHeaders.length).setFontWeight("bold");
      console.log(`シート「${sheetName}」を新規作成し、初期ヘッダーを書き込みました。`);
    } else {
      // 2. 既存シートがある場合は、不足しているヘッダーのみ右端に追記（非破壊マイグレーション）
      const lastCol = sheet.getLastColumn();
      let currentHeaders = [];
      if (lastCol > 0) {
        currentHeaders = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0]
          .map(h => String(h).trim().replace(/\s/g, "").toLowerCase());
      }

      const missingHeaders = requiredHeaders.filter(req => {
        const cleanedReq = req.replace(/\s/g, "").toLowerCase();
        return !currentHeaders.includes(cleanedReq);
      });

      if (missingHeaders.length > 0) {
        const startCol = lastCol + 1;
        const targetRange = sheet.getRange(1, startCol, 1, missingHeaders.length);
        targetRange.setValues([missingHeaders]);
        targetRange.setFontWeight("bold");
        console.log(`シート「${sheetName}」に不足していた列を追加しました: ${missingHeaders.join(', ')}`);
      } else {
        console.log(`シート「${sheetName}」のヘッダーは最新です。変更はありません。`);
      }
    }

    // 3. 初期データの補正（データが0件の場合のみデフォルト値を挿入）
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) { // ヘッダー行のみ、または完全に空の場合
      if (sheetName === USER_SHEET_NAME && userEmail) {
        // 管理者本人を追加
        const headers = mapHeaders_(sheet);
        const newRow = new Array(requiredHeaders.length).fill('');
        newRow[headers['email']] = userEmail;
        newRow[headers['type']] = 'user';
        newRow[headers['role']] = 'admin';
        sheet.appendRow(newRow);
        console.log(`「ユーザー」シートに実行者を管理者(admin)として追加しました: ${userEmail}`);
      } else if (sheetName === SCHEDULE_SHEET_NAME) {
        // 週間スケジュールの初期曜日を自動設定
        const days = ['dayMonday', 'dayTuesday', 'dayWednesday', 'dayThursday', 'dayFriday', 'daySaturday', 'daySunday'];
        days.forEach(day => sheet.appendRow([day, '']));
        console.log('「週間スケジュール」シートに初期曜日データを挿入しました。');
      } else if (sheetName === AI_CONFIG_SHEET_NAME) {
        // AIプロンプトの初期設定を追加
        const defaultPrompt = `あなたはデータ分析の専門家です。
提供されたデイサービスの到着時刻データを分析し、傾向や特徴を要約してください。
特に、到着が集中する時間帯や、遅れがちな傾向があれば指摘してください。
また、業務改善に向けた簡単なアドバイスもあれば含めてください。`;
        sheet.appendRow(['プライマリープロンプト', defaultPrompt]);
        sheet.appendRow(['セカンダリープロンプト', '']);
        sheet.appendRow(['使用するAIモデル名', 'gemini-2.5-flash']);
        console.log('「AI設定」シートにデフォルトプロンプトと使用モデルを挿入しました。');
      }
    }
  }

  SpreadsheetApp.flush();
  console.log('--- スプレッドシートのセットアップがすべて完了しました！ ---');
}