/************************************************************
 * utils.gs
 *
 * システム全体で使用する定数とユーティリティ関数
 ************************************************************/

// =============================================================================
// システム設定定数
// =============================================================================

const CONFIG = {
  MASTER_SPREADSHEET_ID: '1mhBpPhuL6Aq-YRXgmCu1kMtg0g3JO7h37pHdWmM8sqE',
  SHEET_TEACHERS: 'Teachers',
  SHEET_SUBMISSIONS: 'Submissions',
  SHEET_REMINDER_SETTINGS: 'ReminderSettings',

  TEMPLATE_SPREADSHEET_ID: '1uVWYwQkr4zQ5UMwGCNL7nNkUvRwLA5v-IK2NlO1Ulyk', // フォールバック用
  TEMPLATE_FOLDER_ID: '1LRqvNMds307Hf7W2XO9M4PhcNtLiY1c5', // テンプレートフォルダID（10_Templates）
  COPIES_PARENT_FOLDER_ID: '1gxrQ_Kdh1aBGQYem__hxWP9h8HC7IKFo',

  META_SHEET_NAME: '_META',
};

const CONFIG_FORM = {
  QUESTION_TEACHER_NAME: '氏名',
  QUESTION_MONTH_KEY: '提出月', // 例：2026-01
};

// シート設定
const SHEET_CONFIG = {
  PROTECTION_DESCRIPTION: '提出後ロック',
  INPUT_SHEET_NAME: 'Input',
  META_SHEET_NAME: '_META',
  CHECK_ROW: 2,
  CHECK_COL: 3, // C2
  STATUS_CELL_A1: 'B2',
  HELP_CELL_A1: 'D2',
  TEACHER_NAME_CELL_A1: 'G3',
};

// ドライブ設定
const DRIVE_CONFIG = {
  SLEEP_AFTER_EDITOR_CHECK: 100,
  SLEEP_AFTER_ADD_EDITOR: 300,
  SLEEP_AFTER_ADD_SCRIPT_OWNER: 500,
};

// LINE設定
const LINE_CONFIG = {
  STATE_EXPIRY_MS: 24 * 60 * 60 * 1000, // 24時間
  CONFIRM_YES: ['はい', 'yes', 'y'],
  CONFIRM_NO: ['いいえ', 'no', 'n'],
};

// LINEメッセージテンプレート
const MESSAGE_SPREADSHEET_APP = `\n\nGoogleスプレッドシートアプリのインストールをお願いします。
スマートフォンからの編集には、専用アプリが必要です。

【Androidの方】
Google スプレッドシート
https://play.google.com/store/apps/details?id=com.google.android.apps.docs.editors.sheets

【iPhoneの方】
Google スプレッドシート
https://apps.apple.com/jp/app/google-%E3%82%B9%E3%83%97%E3%83%AC%E3%83%83%E3%83%89%E3%82%B7%E3%83%BC%E3%83%88/id842849113`;

// =============================================================================
// ユーティリティ関数
// =============================================================================

function normalizeNameKey_(s) {
  // 全角/半角の統一と空白除去
  let normalized = String(s || '').trim();
  // 全角英数字を半角に変換
  normalized = normalized.replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(s) {
    return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
  });
  // 空白文字を除去
  normalized = normalized.replace(/[ 　\t\n\r]/g, '');
  return normalized;
}

/**
 * テキストが氏名として扱えるかどうかを判定
 * @param {string} text - 入力テキスト
 * @returns {boolean} 氏名として扱えるかどうか
 */
function looksLikeName_(text) {
  if (!text || text.trim().length === 0) return false;
  
  const trimmed = text.trim();
  
  // 絵文字のみ、または絵文字が大部分を占める場合は会話と判断
  // 絵文字のUnicode範囲をチェック
  const emojiPattern = /[\u{1F300}-\u{1F9FF}\u{2600}-\u{26FF}\u{2700}-\u{27BF}\u{1F600}-\u{1F64F}\u{1F680}-\u{1F6FF}\u{1F1E0}-\u{1F1FF}\u{1F900}-\u{1F9FF}\u{1FA00}-\u{1FA6F}\u{1FA70}-\u{1FAFF}]/gu;
  const emojiMatches = trimmed.match(emojiPattern);
  const emojiCount = emojiMatches ? emojiMatches.length : 0;
  const textLength = trimmed.length;
  
  // 絵文字が半分以上を占める、または絵文字のみの場合は会話と判断
  if (emojiCount > 0 && (emojiCount >= textLength / 2 || emojiCount === textLength)) {
    return false;
  }
  
  // 絵文字を除いたテキストで判定
  const textWithoutEmoji = trimmed.replace(emojiPattern, '').trim();
  if (textWithoutEmoji.length === 0) {
    return false; // 絵文字のみ
  }
  
  // 長すぎる場合は会話と判断（20文字以上、絵文字を除く）
  if (textWithoutEmoji.length > 20) return false;
  
  // 一般的な会話フレーズが含まれている場合は会話と判断
  const conversationPhrases = [
    '承知しました',
    '了解しました',
    'わかりました',
    'お疲れさま',
    'お疲れ様',
    'ありがとう',
    'よろしく',
    'お願い',
    'お休み',
    '授業',
    '振替',
    '明日',
    '今日',
    'から',
    'また',
    'になります',
    'になります',
    'お願いいたします',
    'お願いします',
    '改善が必要',
    '必要'
  ];
  
  for (const phrase of conversationPhrases) {
    if (textWithoutEmoji.includes(phrase)) {
      return false; // 会話と判断
    }
  }
  
  // 句読点や長い文章の場合は会話と判断
  if (textWithoutEmoji.includes('。') || textWithoutEmoji.includes('、') || textWithoutEmoji.match(/[！？]/)) {
    // ただし、短い名前の可能性もあるので、文字数で判断
    if (textWithoutEmoji.length > 10) return false;
  }
  
  // 漢字、ひらがな、カタカナ、アルファベットが含まれている場合は名前の可能性
  if (textWithoutEmoji.match(/[\u4e00-\u9faf\u3040-\u309f\u30a0-\u30ffa-zA-Z]/)) {
    return true;
  }
  
  return false;
}

/**
 * テキストから氏名を抽出（メールアドレスや挨拶文を除去）
 * @param {string} text - 入力テキスト
 * @returns {string} 抽出された氏名（見つからない場合は空文字列）
 */
function extractNameFromText_(text) {
  if (!text) return '';
  
  // メールアドレスを除去
  let cleaned = text.replace(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g, '').trim();
  
  // ふりがな（括弧内の読み方）を除去
  // 例：「奥園凌(おくぞのりょう)」→「奥園凌」
  cleaned = cleaned.replace(/[（(][^）)]*[）)]/g, '').trim();
  
  // 一般的な挨拶文やフレーズを除去
  const greetings = [
    'よろしくお願いいたします',
    'よろしくお願いします',
    'よろしく',
    'お願いします',
    'お願いいたします',
    'ありがとうございます',
    'ありがとう',
    'お世話になります',
    'お世話になっております',
    '初めまして',
    'はじめまして',
    '講師登録',
    '登録',
    '先生',
    'さん'
  ];
  
  for (const greeting of greetings) {
    cleaned = cleaned.replace(new RegExp(greeting, 'gi'), '').trim();
  }
  
  // 記号や数字のみの行を除去
  cleaned = cleaned.replace(/^[0-9\s\-_.,!?。、！？]+$/g, '').trim();
  
  // 複数行の場合は最初の行を優先
  const lines = cleaned.split(/[\n\r]/);
  for (const line of lines) {
    const trimmed = line.trim();
    if (trimmed && trimmed.length >= 2) {
      // 2文字以上で、メールアドレスでない場合
      if (!trimmed.match(/^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/)) {
        return trimmed;
      }
    }
  }
  
  return cleaned;
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

/**
 * 月キーを正規化（YYYY-MM形式の文字列に変換）
 * Dateオブジェクトやその他の形式からYYYY-MM形式に変換
 * @param {string|Date} monthKey - 月キー（Dateオブジェクトまたは文字列）
 * @returns {string} YYYY-MM形式の文字列
 */
function normalizeMonthKey_(monthKey) {
  if (!monthKey) return '';
  
  // 既にYYYY-MM形式の文字列の場合
  if (typeof monthKey === 'string') {
    const match = monthKey.match(/^(\d{4})-(\d{2})/);
    if (match) {
      return `${match[1]}-${match[2]}`;
    }
  }
  
  // Dateオブジェクトの場合
  if (monthKey instanceof Date) {
    const y = monthKey.getFullYear();
    const m = String(monthKey.getMonth() + 1).padStart(2, '0');
    return `${y}-${m}`;
  }
  
  // その他の場合、文字列に変換してから試行
  const str = String(monthKey).trim();
  const match = str.match(/^(\d{4})-(\d{2})/);
  if (match) {
    return `${match[1]}-${match[2]}`;
  }
  
  // 日付文字列から抽出を試行
  const dateMatch = str.match(/(\d{4})[\/\-年](\d{1,2})/);
  if (dateMatch) {
    const y = dateMatch[1];
    const m = String(parseInt(dateMatch[2], 10)).padStart(2, '0');
    return `${y}-${m}`;
  }
  
  // 変換できない場合は空文字列を返す
  return '';
}

/**
 * 月の表記をYYYY-MM形式に変換
 * 例：「2月」→「2026-02」（現在の年を使用）
 * 例：「12月」→「2026-12」
 * @param {string} monthText - 月の表記（例：「2月」「12月」）
 * @returns {string} YYYY-MM形式の文字列（変換できない場合は空文字列）
 */
function parseMonthText_(monthText) {
  if (!monthText) return '';
  
  const str = String(monthText).trim();
  // 「2月」「12月」などの形式を検出
  const match = str.match(/^(\d{1,2})月$/);
  if (match) {
    const month = parseInt(match[1], 10);
    if (month >= 1 && month <= 12) {
      const now = new Date();
      const year = now.getFullYear();
      const monthStr = String(month).padStart(2, '0');
      return `${year}-${monthStr}`;
    }
  }
  
  return '';
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
 * Gmailアドレスかどうかを判定
 * @param {string} email - メールアドレス
 * @returns {boolean} Gmailアドレスかどうか
 */
function isGmailAddress_(email) {
  const emailLower = String(email || '').trim().toLowerCase();
  return emailLower.endsWith('@gmail.com');
}

/**
 * 次のTeacher IDを取得（自動採番）
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} masterSs - マスタースプレッドシート
 * @returns {string} 次のTeacher ID（例：T027）
 */
function getNextTeacherId_(masterSs) {
  const sh = masterSs.getSheetByName(CONFIG.SHEET_TEACHERS);
  if (!sh) throw new Error(`マスターにシート "${CONFIG.SHEET_TEACHERS}" がありません`);

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return 'T001'; // データがない場合はT001から開始

  const header = values[0];
  const idxTeacherId = header.indexOf('teacherId');
  if (idxTeacherId < 0) return 'T001'; // teacherId列がない場合

  let maxNum = 0;
  for (let r = 1; r < values.length; r++) {
    const teacherId = String(values[r][idxTeacherId] || '').trim();
    // T001, T002などの形式から数字を抽出
    const match = teacherId.match(/^T(\d+)$/i);
    if (match) {
      const num = parseInt(match[1], 10);
      if (num > maxNum) maxNum = num;
    }
  }

  const nextNum = maxNum + 1;
  return `T${String(nextNum).padStart(3, '0')}`; // T027形式
}

/**
 * Teachersシートに新規講師を手動追加（管理者用ヘルパー関数）
 * LINE User IDがなくても登録可能
 * 
 * 使用例：
 * addTeacherManually('谷口知子', 'satorara0510@gmail.com')
 * 
 * @param {string} teacherName - 講師氏名
 * @param {string} email - メールアドレス（任意）
 * @returns {Object} 追加された講師情報（teacherId, name, row）
 */
function addTeacherManually(teacherName, email = '') {
  const master = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
  const sh = master.getSheetByName(CONFIG.SHEET_TEACHERS);
  if (!sh) throw new Error(`マスターにシート "${CONFIG.SHEET_TEACHERS}" がありません`);

  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idxTeacherId = header.indexOf('teacherId');
  const idxName = header.indexOf('氏名');
  const idxLine = header.indexOf('lineUserId');
  const idxLinkedAt = header.indexOf('lineLinkedAt');
  const idxEmail = header.indexOf('メール');

  if (idxName < 0) throw new Error('Teachersに「氏名」列がありません');

  // 既に存在するかチェック
  const existing = findTeacherByName_(master, teacherName);
  if (existing) {
    console.log(`既に登録済み: ${teacherName} (teacherId: ${existing.teacherId})`);
    return {
      teacherId: existing.teacherId,
      name: existing.name,
      row: existing.row,
      email: existing.email || email
    };
  }

  // 次のTeacher IDを取得
  const teacherId = getNextTeacherId_(master);

  // 新しい行を準備
  const newRow = [];
  for (let i = 0; i < header.length; i++) {
    if (i === idxTeacherId) {
      newRow[i] = teacherId;
    } else if (i === idxName) {
      newRow[i] = teacherName;
    } else if (i === idxLine) {
      newRow[i] = ''; // LINE User IDは空（後でLINE登録時に更新）
    } else if (i === idxLinkedAt) {
      newRow[i] = ''; // LINE登録時に設定
    } else if (i === idxEmail && email) {
      newRow[i] = email;
    } else {
      newRow[i] = '';
    }
  }

  // 行を追加
  sh.appendRow(newRow);
  const newRowNum = sh.getLastRow();

  console.log(`講師を追加しました: ${teacherName} (teacherId: ${teacherId}, row: ${newRowNum})`);
  
  return {
    teacherId,
    name: teacherName,
    row: newRowNum,
    email: email || ''
  };
}

/**
 * Teachersシートに新規講師を追加
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} masterSs - マスタースプレッドシート
 * @param {string} teacherName - 講師氏名
 * @param {string} lineUserId - LINE User ID
 * @param {string} email - メールアドレス（任意）
 * @returns {Object} 追加された講師情報（teacherId, name, row）
 */
function addNewTeacher_(masterSs, teacherName, lineUserId, email = '') {
  const sh = masterSs.getSheetByName(CONFIG.SHEET_TEACHERS);
  if (!sh) throw new Error(`マスターにシート "${CONFIG.SHEET_TEACHERS}" がありません`);

  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idxTeacherId = header.indexOf('teacherId');
  const idxName = header.indexOf('氏名');
  const idxLine = header.indexOf('lineUserId');
  const idxLinkedAt = header.indexOf('lineLinkedAt');
  const idxEmail = header.indexOf('メール');

  if (idxName < 0) throw new Error('Teachersに「氏名」列がありません');
  if (idxLine < 0) throw new Error('Teachersに「lineUserId」列がありません');

  // 次のTeacher IDを取得
  const teacherId = getNextTeacherId_(masterSs);

  // 新しい行を準備
  const newRow = [];
  for (let i = 0; i < header.length; i++) {
    if (i === idxTeacherId) {
      newRow[i] = teacherId;
    } else if (i === idxName) {
      newRow[i] = teacherName;
    } else if (i === idxLine) {
      newRow[i] = lineUserId;
    } else if (i === idxLinkedAt) {
      newRow[i] = new Date();
    } else if (i === idxEmail && email) {
      newRow[i] = email;
    } else {
      newRow[i] = '';
    }
  }

  // 行を追加
  sh.appendRow(newRow);
  const newRowNum = sh.getLastRow();

  return {
    teacherId,
    name: teacherName,
    row: newRowNum,
    email: email || ''
  };
}

/**
 * Teachersの氏名で講師を探す（柔軟なマッチング）
 * 必須: 氏名 / lineUserId
 * 任意: teacherId / メール
 * 
 * マッチング方法：
 * 1. 完全一致（正規化後）
 * 2. 部分一致（正規化後、入力が名簿の一部、または名簿が入力の一部）
 * 3. 類似度マッチング（文字列の類似度が高い場合）
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

  const inputKey = normalizeNameKey_(teacherNameRaw);
  if (!inputKey || inputKey.length < 1) return null;

  const exactMatches = [];
  const partialMatches = [];
  
  for (let r = 1; r < values.length; r++) {
    const dbName = String(values[r][idxName] || '').trim();
    const dbKey = normalizeNameKey_(dbName);
    
    if (!dbKey || dbKey.length < 1) continue;
    
    // 完全一致
    if (dbKey === inputKey) {
      exactMatches.push({
        row: r + 1,
        name: dbName,
        lineUserId: String(values[r][idxLine] || '').trim(),
        teacherId: idxTeacherId >= 0 ? String(values[r][idxTeacherId] || '').trim() : '',
        email: idxEmail >= 0 ? String(values[r][idxEmail] || '').trim() : '',
        matchType: 'exact'
      });
    }
    // 部分一致（入力が名簿の一部、または名簿が入力の一部）
    else if (dbKey.includes(inputKey) || inputKey.includes(dbKey)) {
      // 短い方の長さが長い方の70%以上の場合のみ部分一致として扱う
      const shorter = Math.min(dbKey.length, inputKey.length);
      const longer = Math.max(dbKey.length, inputKey.length);
      if (shorter >= longer * 0.7) {
        partialMatches.push({
          row: r + 1,
          name: dbName,
          lineUserId: String(values[r][idxLine] || '').trim(),
          teacherId: idxTeacherId >= 0 ? String(values[r][idxTeacherId] || '').trim() : '',
          email: idxEmail >= 0 ? String(values[r][idxEmail] || '').trim() : '',
          matchType: 'partial'
        });
      }
    }
  }
  
  // 完全一致を優先
  if (exactMatches.length === 1) {
    return exactMatches[0];
  }
  if (exactMatches.length > 1) {
    return null; // 複数一致
  }
  
  // 部分一致を確認
  if (partialMatches.length === 1) {
    return partialMatches[0];
  }
  if (partialMatches.length > 1) {
    return null; // 複数一致
  }
  
  return null;
}

/**
 * LINE User IDで講師を検索
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} masterSs - マスタースプレッドシート
 * @param {string} lineUserId - LINE User ID
 * @returns {Object|null} 講師情報（name, email, teacherId, lineUserId, row）またはnull
 */
function findTeacherByLineUserId_(masterSs, lineUserId) {
  if (!lineUserId || lineUserId.trim().length === 0) return null;
  
  const sh = masterSs.getSheetByName(CONFIG.SHEET_TEACHERS);
  if (!sh) return null;

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return null;

  const header = values[0];
  const idxName = header.indexOf('氏名');
  const idxLine = header.indexOf('lineUserId');
  const idxTeacherId = header.indexOf('teacherId');
  const idxEmail = header.indexOf('メール');

  if (idxName < 0 || idxLine < 0) return null;

  for (let r = 1; r < values.length; r++) {
    const rowLineUserId = String(values[r][idxLine] || '').trim();
    if (rowLineUserId === lineUserId) {
      return {
        row: r + 1,
        name: String(values[r][idxName]).trim(),
        lineUserId: rowLineUserId,
        teacherId: idxTeacherId >= 0 ? String(values[r][idxTeacherId] || '').trim() : '',
        email: idxEmail >= 0 ? String(values[r][idxEmail] || '').trim() : '',
      };
    }
  }
  
  return null;
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

  const exactHits = [];
  const partialHits = [];
  
  for (let r = 1; r < values.length; r++) {
    const dbName = String(values[r][idxName] || '').trim();
    const dbKey = normalizeNameKey_(dbName);
    
    if (!dbKey || dbKey.length < 1) continue;
    
    // 完全一致
    if (dbKey === nameKey) {
      exactHits.push({
        row: r + 1,
        name: dbName,
        currentLineId: String(values[r][idxLine] || '').trim(),
        currentEmail: idxEmail >= 0 ? String(values[r][idxEmail] || '').trim() : '',
      });
    }
    // 部分一致（入力が名簿の一部、または名簿が入力の一部）
    else if (dbKey.includes(nameKey) || nameKey.includes(dbKey)) {
      // 短い方の長さが長い方の70%以上の場合のみ部分一致として扱う
      const shorter = Math.min(dbKey.length, nameKey.length);
      const longer = Math.max(dbKey.length, nameKey.length);
      if (shorter >= longer * 0.7) {
        partialHits.push({
          row: r + 1,
          name: dbName,
          currentLineId: String(values[r][idxLine] || '').trim(),
          currentEmail: idxEmail >= 0 ? String(values[r][idxEmail] || '').trim() : '',
        });
      }
    }
  }
  
  // 完全一致を優先
  const hits = exactHits.length > 0 ? exactHits : partialHits;

  if (hits.length === 0) return { status: 'not_found' };
  if (hits.length >= 2) return {
    status: 'multiple',
    candidates: hits.map(h => h.name),
    candidatesWithInfo: hits.map(h => ({ name: h.name, email: h.currentEmail, row: h.row }))
  };

  const target = hits[0];

  // ★すでに同じLINE userIdが入っている → 完全一致
  if (target.currentLineId === userId) {
    return { status: 'already_linked_same', name: target.name, email: target.currentEmail, row: target.row };
  }

  // ★別のuserIdが入っている → LINE IDが変更された可能性
  if (target.currentLineId && target.currentLineId !== userId) {
    return { status: 'already_linked_other', name: target.name, email: target.currentEmail, row: target.row, oldLineId: target.currentLineId };
  }

  // ★新規紐付け
  sh.getRange(target.row, idxLine + 1).setValue(userId);
  if (idxLinkedAt >= 0) sh.getRange(target.row, idxLinkedAt + 1).setValue(new Date());

  return { status: 'linked', name: target.name, email: target.currentEmail, row: target.row };
}

/**
 * Teachersシートのメールアドレスを更新
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} masterSs - マスタースプレッドシート
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

/** Submissionsに既存エントリがあるかチェック（submissionKeyで検索） */
function findSubmissionByKey_(masterSs, submissionKey) {
  const sh = masterSs.getSheetByName(CONFIG.SHEET_SUBMISSIONS);
  if (!sh) return null;

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return null;

  const header = values[0];
  const idxSubmissionKey = header.indexOf('submissionKey');
  const idxSheetUrl = header.indexOf('sheetUrl');
  const idxStatus = header.indexOf('status');
  if (idxSubmissionKey < 0) return null;

  for (let r = 1; r < values.length; r++) {
    const key = String(values[r][idxSubmissionKey] || '').trim();
    if (key === submissionKey) {
      return {
        row: r + 1,
        header: header,
        sheetUrl: idxSheetUrl >= 0 ? String(values[r][idxSheetUrl] || '').trim() : '',
        status: idxStatus >= 0 ? String(values[r][idxStatus] || '').trim() : ''
      };
    }
  }
  return null;
}

/**
 * Submissionsに既存エントリがあるかチェック（monthKeyとteacherId/氏名で検索）
 * submissionKeyの生成方式が異なる場合でも、同じ講師・同じ月のエントリを見つけられる
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} masterSs - マスタースプレッドシート
 * @param {string} monthKey - 月キー（YYYY-MM形式）
 * @param {string} teacherId - 講師ID（任意）
 * @param {string} teacherName - 講師氏名（任意）
 * @returns {Object|null} 見つかった場合は{row, header, sheetUrl, status, submissionKey}、なければnull
 */
function findSubmissionByMonthAndTeacher_(masterSs, monthKey, teacherId, teacherName) {
  const sh = masterSs.getSheetByName(CONFIG.SHEET_SUBMISSIONS);
  if (!sh) return null;

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return null;

  const header = values[0];
  const idxMonthKey = header.indexOf('monthKey');
  const idxTeacherId = header.indexOf('teacherId');
  const idxName = header.indexOf('氏名');
  const idxSheetUrl = header.indexOf('sheetUrl');
  const idxStatus = header.indexOf('status');
  const idxSubmissionKey = header.indexOf('submissionKey');

  if (idxMonthKey < 0) return null;

  const normalizedMonthKey = normalizeMonthKey_(monthKey);
  const teacherNameKey = normalizeNameKey_(teacherName || '');

  for (let r = 1; r < values.length; r++) {
    const rowMonthKey = normalizeMonthKey_(values[r][idxMonthKey]);
    if (rowMonthKey !== normalizedMonthKey) continue;

    const rowTeacherId = idxTeacherId >= 0 ? String(values[r][idxTeacherId] || '').trim() : '';
    const rowName = idxName >= 0 ? String(values[r][idxName] || '').trim() : '';
    const rowNameKey = normalizeNameKey_(rowName);

    // teacherIdで一致、または氏名で一致
    const idMatch = teacherId && rowTeacherId && teacherId === rowTeacherId;
    const nameMatch = teacherNameKey && rowNameKey && teacherNameKey === rowNameKey;

    if (idMatch || nameMatch) {
      return {
        row: r + 1,
        header: header,
        sheetUrl: idxSheetUrl >= 0 ? String(values[r][idxSheetUrl] || '').trim() : '',
        status: idxStatus >= 0 ? String(values[r][idxStatus] || '').trim() : '',
        submissionKey: idxSubmissionKey >= 0 ? String(values[r][idxSubmissionKey] || '').trim() : ''
      };
    }
  }
  return null;
}

/** 有効なstatus値 */
const VALID_STATUSES = ['created', 'submitted', 'teacher_not_found', 'template_not_found', ''];

/**
 * statusの値を検証
 * @param {string} status - ステータス値
 * @returns {boolean} 有効なステータスかどうか
 */
function isValidStatus_(status) {
  return VALID_STATUSES.includes(status);
}

/** Submissionsの既存エントリを更新（ヘッダー名ベース） */
function updateSubmission_(masterSs, row, header, obj) {
  const sh = masterSs.getSheetByName(CONFIG.SHEET_SUBMISSIONS);
  if (!sh) throw new Error(`マスターにシート "${CONFIG.SHEET_SUBMISSIONS}" がありません`);

  // statusのバリデーション
  if (obj.status !== undefined && !isValidStatus_(obj.status)) {
    console.error(`[updateSubmission_] Invalid status value: "${obj.status}" for row ${row}. Valid values: ${VALID_STATUSES.join(', ')}`);
    throw new Error(`Invalid status value: "${obj.status}"`);
  }

  for (let i = 0; i < header.length; i++) {
    const colName = header[i];
    let value = '';
    switch (colName) {
      case 'timestamp': value = obj.timestamp !== undefined ? obj.timestamp : null; break;
      case 'monthKey': value = obj.monthKey !== undefined ? obj.monthKey : null; break;
      case 'teacherId': value = obj.teacherId !== undefined ? obj.teacherId : null; break;
      case '氏名': value = obj.name !== undefined ? obj.name : null; break;
      case 'sheetUrl': value = obj.sheetUrl !== undefined ? obj.sheetUrl : null; break;
      case 'status': value = obj.status !== undefined ? obj.status : null; break;
      case 'lastNotified': value = obj.lastNotified !== undefined ? obj.lastNotified : null; break;
      case 'submissionKey': value = obj.submissionKey !== undefined ? obj.submissionKey : null; break;
      case 'submittedAt': value = obj.submittedAt !== undefined ? obj.submittedAt : null; break;
      case 'ackNotifiedAt': value = obj.ackNotifiedAt !== undefined ? obj.ackNotifiedAt : null; break;
      default: continue;
    }
    if (value !== null) {
      sh.getRange(row, i + 1).setValue(value);
    }
  }
}

/** Submissionsに1行追加（ヘッダー名ベース） */
function appendSubmission_(masterSs, obj) {
  const sh = masterSs.getSheetByName(CONFIG.SHEET_SUBMISSIONS);
  if (!sh) throw new Error(`マスターにシート "${CONFIG.SHEET_SUBMISSIONS}" がありません`);

  // statusのバリデーション
  if (obj.status && !isValidStatus_(obj.status)) {
    console.error(`[appendSubmission_] Invalid status value: "${obj.status}". Valid values: ${VALID_STATUSES.join(', ')}`);
    throw new Error(`Invalid status value: "${obj.status}"`);
  }

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
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} masterSs - マスタースプレッドシート
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

// =============================================================================
// 共通ヘルパー関数（リファクタリング用）
// =============================================================================

/**
 * 状態の有効期限をチェック
 * @param {number|Date} timestamp - 状態の作成時刻
 * @param {number} expiryMs - 有効期限（ミリ秒）、デフォルトは24時間
 * @returns {boolean} 期限切れの場合true
 */
function isStateExpired_(timestamp, expiryMs = LINE_CONFIG.STATE_EXPIRY_MS) {
  if (!timestamp) return true;
  const created = typeof timestamp === 'number' ? timestamp : new Date(timestamp).getTime();
  return Date.now() - created > expiryMs;
}

/**
 * 確認応答（はい/いいえ）を解析
 * @param {string} text - ユーザーの応答テキスト
 * @returns {'yes'|'no'|null} 'yes', 'no', またはnull（判定不能）
 */
function parseConfirmationResponse_(text) {
  if (!text) return null;
  const normalized = String(text).trim().toLowerCase();
  if (LINE_CONFIG.CONFIRM_YES.includes(normalized)) return 'yes';
  if (LINE_CONFIG.CONFIRM_NO.includes(normalized)) return 'no';
  return null;
}

/**
 * チェックボックスやフラグの値をブール値として評価
 * TRUE, true, 1, ○, チェック済み などを true として扱う
 * @param {any} value - チェック対象の値
 * @returns {boolean} チェックされている場合 true
 */
function isChecked_(value) {
  if (value === true || value === 1) return true;
  if (typeof value === 'string') {
    const normalized = value.trim().toLowerCase();
    return ['true', '1', '○', 'yes', 'チェック済み', 'on'].includes(normalized);
  }
  return false;
}

/**
 * 休職中の講師名リストを取得
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} master - マスタースプレッドシート
 * @returns {Set<string>} 休職中の講師名のSet
 */
function getOnLeaveTeachers_(master) {
  const result = new Set();
  try {
    const sh = master.getSheetByName(CONFIG.SHEET_TEACHERS);
    if (!sh) return result;

    const values = sh.getDataRange().getValues();
    if (values.length < 2) return result;

    const header = values[0];
    const idxName = header.indexOf('氏名');
    const idxOnLeave = header.indexOf('休職中');

    if (idxName < 0 || idxOnLeave < 0) return result;

    for (let r = 1; r < values.length; r++) {
      const teacherName = String(values[r][idxName] || '').trim();
      const onLeave = isChecked_(values[r][idxOnLeave]);
      if (teacherName && onLeave) {
        result.add(teacherName);
      }
    }
  } catch (err) {
    console.error('getOnLeaveTeachers_ error:', err);
  }
  return result;
}

/**
 * Submissionsシートのヘッダー列インデックスを取得
 * @param {Array<string>} header - ヘッダー行
 * @returns {Object} 列名→インデックスのマップ
 */
function getSubmissionIndices_(header) {
  return {
    monthKey: header.indexOf('monthKey'),
    status: header.indexOf('status'),
    name: header.indexOf('氏名'),
    teacherId: header.indexOf('teacherId'),
    url: header.indexOf('sheetUrl'),
    submittedAt: header.indexOf('submittedAt'),
    ackNotifiedAt: header.indexOf('ackNotifiedAt'),
    lockedAt: header.indexOf('lockedAt'),
    submissionKey: header.indexOf('submissionKey'),
    // 管理者リマインド用
    managerReminder2Weeks: header.indexOf('managerReminder2Weeks'),
    managerReminder10Days: header.indexOf('managerReminder10Days'),
    managerReminder1Week: header.indexOf('managerReminder1Week'),
    managerReminder3Days: header.indexOf('managerReminder3Days'),
    managerReminder1Day: header.indexOf('managerReminder1Day'),
    managerReminder1st: header.indexOf('managerReminder1st'),
  };
}

/**
 * Submissionsから未提出/提出済みリストを構築
 * @param {Array<Array>} values - シートデータ（ヘッダー含む）
 * @param {Object} indices - getSubmissionIndices_の戻り値
 * @param {string} targetMonth - 対象月（YYYY-MM形式）
 * @returns {Object} { unapplied: Array, submitted: Array }
 */
function buildSubmissionLists_(values, indices, targetMonth) {
  const unapplied = [];
  const submitted = [];

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const mk = normalizeMonthKey_(row[indices.monthKey]);
    if (mk !== targetMonth) continue;

    const status = String(row[indices.status] || '').trim();
    const teacherName = String(row[indices.name] || '').trim();
    if (!teacherName) continue;

    const entry = {
      row: r + 1,
      name: teacherName,
      teacherId: indices.teacherId >= 0 ? String(row[indices.teacherId] || '').trim() : '',
      sheetUrl: indices.url >= 0 ? String(row[indices.url] || '').trim() : '',
    };

    if (status === 'submitted') {
      submitted.push(entry);
    } else if (['created', '', 'teacher_not_found', 'template_not_found'].includes(status)) {
      unapplied.push(entry);
    }
  }

  return { unapplied, submitted };
}

/**
 * リマインド日の種類を判定
 * @param {Date} today - 今日の日付
 * @param {Date} monthStart - 対象月の1日
 * @returns {string|null} リマインド種別（'2weeks', '10days', '1week', '3days', '1day', 'firstDay'）またはnull
 */
function getReminderType_(today, monthStart) {
  const reminderDays = [
    { days: 14, type: '2weeks' },
    { days: 10, type: '10days' },
    { days: 7, type: '1week' },
    { days: 3, type: '3days' },
    { days: 1, type: '1day' },
    { days: 0, type: 'firstDay' },
  ];

  for (const reminder of reminderDays) {
    const targetDate = new Date(monthStart);
    targetDate.setDate(targetDate.getDate() - reminder.days);
    if (isSameJstDate_(today, targetDate)) {
      return reminder.type;
    }
  }
  return null;
}

/**
 * リマインド種別に対応する列インデックスを取得
 * @param {string} reminderType - リマインド種別
 * @param {Object} indices - getSubmissionIndices_の戻り値
 * @returns {number} 列インデックス（-1の場合は該当列なし）
 */
function getReminderColumnIndex_(reminderType, indices) {
  const mapping = {
    '2weeks': indices.managerReminder2Weeks,
    '10days': indices.managerReminder10Days,
    '1week': indices.managerReminder1Week,
    '3days': indices.managerReminder3Days,
    '1day': indices.managerReminder1Day,
    'firstDay': indices.managerReminder1st,
  };
  return mapping[reminderType] ?? -1;
}

/**
 * 提出状況のメッセージをフォーマット
 * @param {Array} unappliedList - 未提出者リスト
 * @param {Array} submittedList - 提出済みリスト
 * @param {string} monthKey - 対象月
 * @returns {string} フォーマットされたメッセージ
 */
function formatSubmissionListMessage_(unappliedList, submittedList, monthKey) {
  let message = `【${monthKey} シフト提出状況】\n\n`;

  if (unappliedList.length > 0) {
    message += '【未提出者】\n';
    const withUrl = unappliedList.filter(item => item.sheetUrl?.trim());
    const withoutUrl = unappliedList.filter(item => !item.sheetUrl?.trim());

    withUrl.forEach((item, index) => {
      message += `${index + 1}. ${item.name}（シート作成済み）\n`;
    });

    withoutUrl.forEach((item, index) => {
      const startNum = withUrl.length + index + 1;
      message += `${startNum}. ${item.name}（シート未作成）\n`;
    });

    message += '\n';
  }

  if (submittedList.length > 0) {
    message += '【提出済み】\n';
    submittedList.forEach((item, index) => {
      const name = typeof item === 'string' ? item : item.name;
      message += `${index + 1}. ${name}\n`;
    });
  }

  if (unappliedList.length === 0 && submittedList.length === 0) {
    message += '（データなし）';
  }  return message;
}

// =============================================================================
// リマインダー設定関連
// =============================================================================

/**
 * ReminderSettingsシートからリマインダー設定を読み込む
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} masterSs - マスタースプレッドシート
 * @returns {Array} リマインダー設定の配列
 */
function getReminderSettings_(masterSs) {
  const sh = masterSs.getSheetByName(CONFIG.SHEET_REMINDER_SETTINGS);
  if (!sh) {
    console.log('ReminderSettingsシートが見つかりません。デフォルト設定を使用します。');
    return getDefaultReminderSettings_();
  }

  const values = sh.getDataRange().getValues();
  if (values.length < 2) {
    return getDefaultReminderSettings_();
  }

  const header = values[0];
  const idxId = header.indexOf('id');
  const idxType = header.indexOf('type');
  const idxDays = header.indexOf('daysBeforeDeadline');
  const idxTarget = header.indexOf('targetAudience');
  const idxEnabled = header.indexOf('enabled');
  const idxMessage = header.indexOf('messageTemplate');

  if (idxId < 0 || idxDays < 0 || idxEnabled < 0) {
    console.log('ReminderSettingsシートに必要な列がありません。デフォルト設定を使用します。');
    return getDefaultReminderSettings_();
  }

  const settings = [];
  for (let r = 1; r < values.length; r++) {
    const enabled = values[r][idxEnabled];
    if (enabled !== true && enabled !== 'TRUE' && enabled !== 'true') continue;

    settings.push({
      id: String(values[r][idxId] || '').trim(),
      type: idxType >= 0 ? String(values[r][idxType] || '').trim() : '',
      daysBeforeDeadline: Number(values[r][idxDays]) || 0,
      targetAudience: idxTarget >= 0 ? String(values[r][idxTarget] || 'manager').trim() : 'manager',
      enabled: true,
      messageTemplate: idxMessage >= 0 ? String(values[r][idxMessage] || '').trim() : '',
    });
  }

  return settings.length > 0 ? settings : getDefaultReminderSettings_();
}

/**
 * デフォルトのリマインダー設定を取得
 * @returns {Array} デフォルトリマインダー設定の配列
 */
function getDefaultReminderSettings_() {
  return [
    { id: 'initial_request', type: '初回申請依頼', daysBeforeDeadline: 21, targetAudience: 'teacher', enabled: true, messageTemplate: '【シフト申請のお願い】\n{name}先生、{monthKey}のシフト申請をお願いします。' },
    { id: 'reminder_2weeks', type: '2週間前リマインド', daysBeforeDeadline: 14, targetAudience: 'manager', enabled: true, messageTemplate: '【シフト未提出リマインド（{monthKey}）】\n{monthKey}のシフト提出期限まで2週間です。\n\n{submissionList}' },
    { id: 'reminder_10days', type: '10日前リマインド', daysBeforeDeadline: 10, targetAudience: 'manager', enabled: true, messageTemplate: '【シフト未提出リマインド（{monthKey}）】\n{monthKey}のシフト提出期限まで10日です。\n\n{submissionList}' },
    { id: 'reminder_1week_teacher', type: '1週間前リマインド', daysBeforeDeadline: 7, targetAudience: 'teacher', enabled: true, messageTemplate: '【シフト未提出リマインド（{monthKey}）】\n{name}先生、{monthKey}のシフト提出期限まで1週間です。' },
    { id: 'reminder_1week_manager', type: '1週間前リマインド', daysBeforeDeadline: 7, targetAudience: 'manager', enabled: true, messageTemplate: '【シフト未提出リマインド（{monthKey}）】\n{monthKey}のシフト未提出者をお知らせします。\n\n{submissionList}' },
    { id: 'reminder_3days_teacher', type: '3日前リマインド', daysBeforeDeadline: 3, targetAudience: 'teacher', enabled: true, messageTemplate: '【シフト未提出リマインド（{monthKey}）】\n{name}先生、{monthKey}のシフト提出期限まで3日です。至急、提出をお願いします！' },
    { id: 'reminder_3days_manager', type: '3日前リマインド', daysBeforeDeadline: 3, targetAudience: 'manager', enabled: true, messageTemplate: '【シフト未提出リマインド（{monthKey}）】\n{monthKey}のシフト未提出者をお知らせします。\n\n{submissionList}' },
    { id: 'reminder_1day_teacher', type: '1日前リマインド', daysBeforeDeadline: 1, targetAudience: 'teacher', enabled: true, messageTemplate: '【シフト未提出リマインド（{monthKey}）】\n{name}先生、{monthKey}のシフト提出期限は明日です。至急、提出をお願いします！' },
    { id: 'reminder_1day_manager', type: '1日前リマインド', daysBeforeDeadline: 1, targetAudience: 'manager', enabled: true, messageTemplate: '【シフト未提出リマインド（{monthKey}）】\n{monthKey}のシフト未提出者をお知らせします。\n\n{submissionList}' },
    { id: 'deadline_day_teacher', type: '締切日', daysBeforeDeadline: 0, targetAudience: 'teacher', enabled: true, messageTemplate: '【シフト提出状況（{monthKey}）】\n{name}先生、本日が{monthKey}のシフト提出期限です。至急、提出をお願いします！' },
    { id: 'deadline_day_manager', type: '締切日', daysBeforeDeadline: 0, targetAudience: 'manager', enabled: true, messageTemplate: '【シフト提出状況（{monthKey}）】\n{monthKey}のシフト未提出者をお知らせします。\n\n{submissionList}' },
  ];
}

/**
 * ReminderSettingsシートを初期化（デフォルト値で作成）
 * Google Apps Scriptエディタで手動実行
 */
function initializeReminderSettingsSheet() {
  const master = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
  let sh = master.getSheetByName(CONFIG.SHEET_REMINDER_SETTINGS);

  if (sh) {
    console.log('ReminderSettingsシートは既に存在します。');
    return 'シートは既に存在します';
  }

  sh = master.insertSheet(CONFIG.SHEET_REMINDER_SETTINGS);

  // ヘッダー行
  const headers = ['id', 'type', 'daysBeforeDeadline', 'targetAudience', 'enabled', 'messageTemplate', 'description'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4a86e8').setFontColor('white');

  // デフォルトデータ
  const defaultData = [
    ['initial_request', '初回申請依頼', 21, 'teacher', true, '【シフト申請のお願い】\n{name}先生、{monthKey}のシフト申請をお願いします。', '来月のシフト申請を依頼（3週間前）'],
    ['reminder_2weeks', '2週間前リマインド', 14, 'manager', true, '【シフト未提出リマインド（{monthKey}）】\n{monthKey}のシフト提出期限まで2週間です。\n\n{submissionList}', '管理者へ未提出者リスト通知'],
    ['reminder_10days', '10日前リマインド', 10, 'manager', true, '【シフト未提出リマインド（{monthKey}）】\n{monthKey}のシフト提出期限まで10日です。\n\n{submissionList}', '管理者へ未提出者リスト通知'],
    ['reminder_1week_teacher', '1週間前リマインド', 7, 'teacher', true, '【シフト未提出リマインド（{monthKey}）】\n{name}先生、{monthKey}のシフト提出期限まで1週間です。', '講師へリマインド'],
    ['reminder_1week_manager', '1週間前リマインド', 7, 'manager', true, '【シフト未提出リマインド（{monthKey}）】\n{monthKey}のシフト未提出者をお知らせします。\n\n{submissionList}', '管理者へ未提出者リスト通知'],
    ['reminder_3days_teacher', '3日前リマインド', 3, 'teacher', true, '【シフト未提出リマインド（{monthKey}）】\n{name}先生、{monthKey}のシフト提出期限まで3日です。至急、提出をお願いします！', '講師へリマインド'],
    ['reminder_3days_manager', '3日前リマインド', 3, 'manager', true, '【シフト未提出リマインド（{monthKey}）】\n{monthKey}のシフト未提出者をお知らせします。\n\n{submissionList}', '管理者へ未提出者リスト通知'],
    ['reminder_1day_teacher', '1日前リマインド', 1, 'teacher', true, '【シフト未提出リマインド（{monthKey}）】\n{name}先生、{monthKey}のシフト提出期限は明日です。至急、提出をお願いします！', '講師へリマインド'],
    ['reminder_1day_manager', '1日前リマインド', 1, 'manager', true, '【シフト未提出リマインド（{monthKey}）】\n{monthKey}のシフト未提出者をお知らせします。\n\n{submissionList}', '管理者へ未提出者リスト通知'],
    ['deadline_day_teacher', '締切日', 0, 'teacher', true, '【シフト提出状況（{monthKey}）】\n{name}先生、本日が{monthKey}のシフト提出期限です。至急、提出をお願いします！', '講師へ締切当日通知'],
    ['deadline_day_manager', '締切日', 0, 'manager', true, '【シフト提出状況（{monthKey}）】\n{monthKey}のシフト未提出者をお知らせします。\n\n{submissionList}', '管理者へ締切当日通知'],
  ];

  sh.getRange(2, 1, defaultData.length, defaultData[0].length).setValues(defaultData);

  // 列幅調整
  sh.setColumnWidth(1, 150);  // id
  sh.setColumnWidth(2, 150);  // type
  sh.setColumnWidth(3, 150);  // daysBeforeDeadline
  sh.setColumnWidth(4, 120);  // targetAudience
  sh.setColumnWidth(5, 80);   // enabled
  sh.setColumnWidth(6, 400);  // messageTemplate
  sh.setColumnWidth(7, 250);  // description

  // チェックボックス設定（enabled列）
  sh.getRange(2, 5, defaultData.length, 1).insertCheckboxes();

  // データ入力規則（targetAudience列）
  const targetRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['manager', 'teacher'], true)
    .build();
  sh.getRange(2, 4, defaultData.length, 1).setDataValidation(targetRule);

  console.log('ReminderSettingsシートを作成しました');
  return 'シートを作成しました';
}

/**
 * 今日がリマインダー送信日かどうかを判定
 * @param {Date} today - 今日の日付
 * @param {Date} monthStart - 対象月の1日
 * @param {number} daysBeforeDeadline - 締切何日前か
 * @returns {boolean} リマインダー送信日ならtrue
 */
function isReminderDay_(today, monthStart, daysBeforeDeadline) {
  const targetDate = new Date(monthStart);
  targetDate.setDate(targetDate.getDate() - daysBeforeDeadline);
  return isSameJstDate_(today, targetDate);
}

/**
 * メッセージテンプレートの変数を置換
 * @param {string} template - メッセージテンプレート
 * @param {Object} vars - 置換用変数 {name, monthKey, submissionList}
 * @returns {string} 置換後のメッセージ
 */
function replaceMessageVariables_(template, vars) {
  let message = template;
  if (vars.name) message = message.replace(/\{name\}/g, vars.name);
  if (vars.monthKey) message = message.replace(/\{monthKey\}/g, vars.monthKey);
  if (vars.submissionList) message = message.replace(/\{submissionList\}/g, vars.submissionList);
  return message;
}
