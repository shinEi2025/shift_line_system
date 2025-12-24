/************************************************************
 * utils.gs
 ************************************************************/

function normalizeNameKey_(s) {
  return String(s || '').trim().replace(/[ 　\t]/g, '');
}

/**
 * 氏名から名字を抽出（スペースの前の部分）
 * @param {string} fullName - フルネーム（例：「森永 英敬」）
 * @returns {string} 名字（例：「森永」）
 */
function extractLastName_(fullName) {
  const name = String(fullName || '').trim();
  // 半角スペース、全角スペース、タブで分割
  const parts = name.split(/[ 　\t]/);
  return parts[0] || name; // スペースがない場合はそのまま返す
}

function extractNamedValues_(e) {
  const out = {};
  const named = (e && e.namedValues) ? e.namedValues : {};
  for (const k in named) {
    const arr = named[k];
    out[k] = Array.isArray(arr) ? String(arr[0] || '') : String(arr || '');
  }
  return out;
}

function getNextMonthKey_() {
  const d = new Date();
  d.setMonth(d.getMonth() + 1);
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  return `${y}-${m}`;
}

/** URLから SpreadsheetId 抽出 */
function extractSpreadsheetId_(url) {
  const m = String(url || '').match(/\/d\/([a-zA-Z0-9-_]+)/);
  return m ? m[1] : '';
}

/**
 * メールアドレスを抽出（テキストから）
 * @param {string} text - テキスト
 * @returns {string} メールアドレス（見つからない場合は空文字列）
 */
function extractEmail_(text) {
  const emailRegex = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/;
  const match = String(text || '').match(emailRegex);
  return match ? match[0] : '';
}

/**
 * メールアドレスの形式を検証
 * @param {string} email - メールアドレス
 * @returns {boolean} 有効な形式かどうか
 */
function isValidEmail_(email) {
  const emailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
  return emailRegex.test(String(email || '').trim());
}

/**
 * Teachersの氏名で講師を探す（完全一致）
 * 必須: 氏名 / lineUserId
 * 任意: teacherId / メール
 */
function findTeacherByName_(masterSs, teacherNameRaw) {
  const sh = masterSs.getSheetByName(CONFIG.SHEET_TEACHERS);
  if (!sh) throw new Error(`マスターにシート "${CONFIG.SHEET_TEACHERS}" がありません`);

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return null;

  const header = values[0];
  const idxName = header.indexOf('氏名');
  const idxLine = header.indexOf('lineUserId');
  const idxTeacherId = header.indexOf('teacherId');
  const idxEmail = header.indexOf('メール');

  if (idxName < 0) throw new Error('Teachersに「氏名」列がありません');
  if (idxLine < 0) throw new Error('Teachersに「lineUserId」列がありません');

  const key = normalizeNameKey_(teacherNameRaw);
  const hits = [];
  for (let r = 1; r < values.length; r++) {
    const nm = normalizeNameKey_(values[r][idxName]);
    if (nm && nm === key) {
      hits.push({
        row: r + 1,
        name: String(values[r][idxName]).trim(),
        lineUserId: String(values[r][idxLine] || '').trim(),
        teacherId: idxTeacherId >= 0 ? String(values[r][idxTeacherId] || '').trim() : '',
        email: idxEmail >= 0 ? String(values[r][idxEmail] || '').trim() : '',
      });
    }
  }
  if (hits.length !== 1) return null;
  return hits[0]; // row, name, lineUserId, teacherId, email を含む
}

/**
 * Teachersの氏名で1件に特定できたら lineUserId を書き込み（上書き禁止）
 * 戻り値 status:
 *  - linked               : 新規に紐付けた
 *  - already_linked_same  : すでに同じ userId が入っている（=二重返信防止用）
 *  - already_linked_other : 別の userId が入っている（事故防止）
 *  - multiple             : 同名が複数
 *  - not_found            : 見つからない
 */
function linkLineUserByName_(masterSs, nameKey, userId) {
  const sh = masterSs.getSheetByName(CONFIG.SHEET_TEACHERS);
  if (!sh) throw new Error(`マスターにシート "${CONFIG.SHEET_TEACHERS}" がありません`);

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return { status: 'not_found' };

  const header = values[0];
  const idxName = header.indexOf('氏名');
  const idxLine = header.indexOf('lineUserId');
  const idxLinkedAt = header.indexOf('lineLinkedAt');
  const idxEmail = header.indexOf('メール');

  if (idxName < 0) throw new Error('Teachersに「氏名」列がありません');
  if (idxLine < 0) throw new Error('Teachersに「lineUserId」列がありません（ヘッダー追加してください）');

  const hits = [];
  for (let r = 1; r < values.length; r++) {
    const nm = normalizeNameKey_(values[r][idxName]);
    if (nm && nm === nameKey) {
      hits.push({
        row: r + 1,
        name: String(values[r][idxName]).trim(),
        currentLineId: String(values[r][idxLine] || '').trim(),
        currentEmail: idxEmail >= 0 ? String(values[r][idxEmail] || '').trim() : '',
      });
    }
  }

  if (hits.length === 0) return { status: 'not_found' };
  if (hits.length >= 2) return { status: 'multiple', candidates: hits.map(h => h.name) };

  const target = hits[0];

  // ★すでに同じLINE userIdが入っている → 何もしない（=二重返信防止）
  if (target.currentLineId === userId) {
    return { status: 'already_linked_same', name: target.name, email: target.currentEmail };
  }

  // ★別のuserIdが入っている → 上書き禁止（事故防止）
  if (target.currentLineId && target.currentLineId !== userId) {
    return { status: 'already_linked_other', name: target.name, email: target.currentEmail };
  }

  // ★新規紐付け
  sh.getRange(target.row, idxLine + 1).setValue(userId);
  if (idxLinkedAt >= 0) sh.getRange(target.row, idxLinkedAt + 1).setValue(new Date());

  return { status: 'linked', name: target.name, email: target.currentEmail, row: target.row };
}

/**
 * Teachersシートのメールアドレスを更新
 * @param {SpreadsheetApp.Spreadsheet} masterSs - マスタースプレッドシート
 * @param {number} row - 行番号（1ベース）
 * @param {string} email - メールアドレス
 * @returns {boolean} 更新成功かどうか
 */
function updateTeacherEmail_(masterSs, row, email) {
  try {
    const sh = masterSs.getSheetByName(CONFIG.SHEET_TEACHERS);
    if (!sh) return false;

    const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const idxEmail = header.indexOf('メール');
    if (idxEmail < 0) return false;

    sh.getRange(row, idxEmail + 1).setValue(email);
    return true;
  } catch (e) {
    console.error('updateTeacherEmail_ error:', e);
    return false;
  }
}

/** Submissionsに1行追加（ヘッダー名ベース） */
function appendSubmission_(masterSs, obj) {
  const sh = masterSs.getSheetByName(CONFIG.SHEET_SUBMISSIONS);
  if (!sh) throw new Error(`マスターにシート "${CONFIG.SHEET_SUBMISSIONS}" がありません`);

  const values = sh.getDataRange().getValues();
  if (values.length < 1) throw new Error('Submissionsのヘッダー行がありません');

  const header = values[0];
  const row = header.map(h => {
    switch (h) {
      case 'timestamp': return obj.timestamp || new Date();
      case 'monthKey': return obj.monthKey || '';
      case 'teacherId': return obj.teacherId || '';
      case '氏名': return obj.name || '';
      case 'sheetUrl': return obj.sheetUrl || '';
      case 'status': return obj.status || '';
      case 'lastNotified': return obj.lastNotified || '';
      case 'submissionKey': return obj.submissionKey || '';
      case 'submittedAt': return obj.submittedAt || '';
      default: return '';
    }
  });

  sh.appendRow(row);
}

/** 講師シートに _META を書き込む */
function writeMetaToTeacherSheet_(teacherSpreadsheetId, metaSheetName, meta) {
  const ss = SpreadsheetApp.openById(teacherSpreadsheetId);
  let sh = ss.getSheetByName(metaSheetName);
  if (!sh) sh = ss.insertSheet(metaSheetName);

  const rows = Object.entries(meta).map(([k, v]) => [k, v]);
  sh.clearContents();
  if (rows.length) sh.getRange(1, 1, rows.length, 2).setValues(rows);

  sh.hideSheet();
}



/**
 * Teachersから lineUserId を取得
 * - teacherId があれば teacherId で優先検索
 * - なければ 氏名（空白ゆれ吸収）で検索
 */
function getTeacherLineUserId_(masterSs, teacherId, teacherName) {
  const sh = masterSs.getSheetByName(CONFIG.SHEET_TEACHERS);
  if (!sh) throw new Error(`マスターにシート "${CONFIG.SHEET_TEACHERS}" がありません`);

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return '';

  const header = values[0];
  const idxTeacherId = header.indexOf('teacherId');
  const idxName = header.indexOf('氏名');
  const idxLine = header.indexOf('lineUserId');

  if (idxName < 0 || idxLine < 0) throw new Error('Teachersに「氏名」「lineUserId」列が必要です');

  const keyName = normalizeNameKey_(teacherName || '');

  for (let r = 1; r < values.length; r++) {
    const rowTeacherId = idxTeacherId >= 0 ? String(values[r][idxTeacherId] || '').trim() : '';
    const rowNameKey = normalizeNameKey_(values[r][idxName]);

    const idMatch = teacherId && rowTeacherId && teacherId === rowTeacherId;
    const nameMatch = keyName && rowNameKey && keyName === rowNameKey;

    if (idMatch || (!teacherId && nameMatch) || (teacherId && !rowTeacherId && nameMatch)) {
      return String(values[r][idxLine] || '').trim();
    }
  }
  return '';
}

/** 日付が「今日」かどうか（JST基準） */
function isSameJstDate_(a, b) {
  if (!a || !b) return false;
  const tz = Session.getScriptTimeZone();
  return Utilities.formatDate(new Date(a), tz, 'yyyy-MM-dd') === Utilities.formatDate(new Date(b), tz, 'yyyy-MM-dd');
}

/**
 * Teachersから講師情報を取得（teacherIdまたは氏名で検索）
 * @param {SpreadsheetApp.Spreadsheet} masterSs - マスタースプレッドシート
 * @param {string} teacherId - 講師ID（任意）
 * @param {string} teacherName - 講師氏名（任意）
 * @returns {Object|null} 講師情報（name, email, teacherId, lineUserId）またはnull
 */
function getTeacherInfo_(masterSs, teacherId, teacherName) {
  const sh = masterSs.getSheetByName(CONFIG.SHEET_TEACHERS);
  if (!sh) return null;

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return null;

  const header = values[0];
  const idxTeacherId = header.indexOf('teacherId');
  const idxName = header.indexOf('氏名');
  const idxEmail = header.indexOf('メール');
  const idxLine = header.indexOf('lineUserId');

  if (idxName < 0) return null;

  const keyName = normalizeNameKey_(teacherName || '');

  for (let r = 1; r < values.length; r++) {
    const rowTeacherId = idxTeacherId >= 0 ? String(values[r][idxTeacherId] || '').trim() : '';
    const rowNameKey = normalizeNameKey_(values[r][idxName]);

    const idMatch = teacherId && rowTeacherId && teacherId === rowTeacherId;
    const nameMatch = keyName && rowNameKey && keyName === rowNameKey;

    if (idMatch || (!teacherId && nameMatch) || (teacherId && !rowTeacherId && nameMatch)) {
      return {
        name: String(values[r][idxName]).trim(),
        email: idxEmail >= 0 ? String(values[r][idxEmail] || '').trim() : '',
        teacherId: rowTeacherId,
        lineUserId: idxLine >= 0 ? String(values[r][idxLine] || '').trim() : '',
      };
    }
  }
  return null;
}

/**
 * 管理者向け詳細ログ出力
 * @param {Error|string} error - エラーオブジェクトまたはエラーメッセージ
 * @param {string} functionName - 関数名
 * @param {Object} context - 追加コンテキスト情報
 */
function logError_(error, functionName, context = {}) {
  const timestamp = new Date().toISOString();
  const errorMessage = error instanceof Error ? error.message : String(error);
  const stackTrace = error instanceof Error ? error.stack : '';
  
  const logDetails = [
    `[ERROR] ${timestamp}`,
    `Function: ${functionName}`,
    `Message: ${errorMessage}`,
    ...(stackTrace ? [`Stack: ${stackTrace}`] : []),
    ...(Object.keys(context).length > 0 ? [`Context: ${JSON.stringify(context, null, 2)}`] : []),
  ].join('\n');
  
  console.error(logDetails);
}

/**
 * 管理者にLINE通知を送信（例外発生時）
 * @param {Error|string} error - エラーオブジェクトまたはエラーメッセージ
 * @param {string} functionName - 関数名
 * @param {Object} context - 追加コンテキスト情報
 */
function notifyAdminOnError_(error, functionName, context = {}) {
  try {
    const adminLineUserId = PropertiesService.getScriptProperties().getProperty('ADMIN_LINE_USER_ID');
    if (!adminLineUserId) {
      console.error('ADMIN_LINE_USER_ID が未設定です（スクリプトプロパティ）。管理者通知をスキップします。');
      return;
    }
    
    const errorMessage = error instanceof Error ? error.message : String(error);
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    
    let message = `【システムエラー通知】\n`;
    message += `時刻: ${timestamp}\n`;
    message += `関数: ${functionName}\n`;
    message += `エラー: ${errorMessage}`;
    
    if (Object.keys(context).length > 0) {
      message += `\n\n詳細:\n${JSON.stringify(context, null, 2)}`;
    }
    
    // メッセージが長すぎる場合は切り詰め（LINEの上限は5000文字）
    if (message.length > 4000) {
      message = message.substring(0, 4000) + '\n\n（メッセージが長いため切り詰めました）';
    }
    
    pushLine_(adminLineUserId, message);
  } catch (notifyErr) {
    // 通知自体が失敗しても元のエラーを隠さない
    console.error('管理者通知の送信に失敗しました:', notifyErr);
  }
}

/**
 * エラーハンドリング：詳細ログ + 管理者通知
 * @param {Error|string} error - エラーオブジェクトまたはエラーメッセージ
 * @param {string} functionName - 関数名
 * @param {Object} context - 追加コンテキスト情報
 */
function handleError_(error, functionName, context = {}) {
  logError_(error, functionName, context);
  notifyAdminOnError_(error, functionName, context);
}
