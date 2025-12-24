/************************************************************
 * Code.gs
 * - LINE webhook: doPost
 * - Form trigger: onFormSubmit
 * 
 * 【設定が必要なスクリプトプロパティ】
 * - LINE_CHANNEL_ACCESS_TOKEN: LINE Botのチャネルアクセストークン
 * - ADMIN_LINE_USER_ID: 管理者のLINE User ID（例外通知用、任意）
 ************************************************************/

const CONFIG = {
  MASTER_SPREADSHEET_ID: '1mhBpPhuL6Aq-YRXgmCu1kMtg0g3JO7h37pHdWmM8sqE',
  SHEET_TEACHERS: 'Teachers',
  SHEET_SUBMISSIONS: 'Submissions',

  TEMPLATE_SPREADSHEET_ID: '1uVWYwQkr4zQ5UMwGCNL7nNkUvRwLA5v-IK2NlO1Ulyk',
  COPIES_PARENT_FOLDER_ID: '1gxrQ_Kdh1aBGQYem__hxWP9h8HC7IKFo',

  META_SHEET_NAME: '_META',
};

const CONFIG_FORM = {
  QUESTION_TEACHER_NAME: '氏名',
  QUESTION_MONTH_KEY: '提出月', // 例：2026-01
};

function doGet() {
  return ContentService.createTextOutput('OK');
}

/**
 * LINE Webhook：講師が氏名を送信 → TeachersにlineUserId紐付け（上書き禁止）
 */
function doPost(e) {
  try {
    const body = (e && e.postData && e.postData.contents) ? e.postData.contents : '';

    // 到達ログ（念のため）
    const props = PropertiesService.getScriptProperties();
    const n = Number(props.getProperty('HIT_COUNT') || '0') + 1;
    props.setProperty('HIT_COUNT', String(n));
    props.setProperty('LAST_BODY', body.slice(0, 3000));

    if (!body) return ContentService.createTextOutput('OK');

    const payload = JSON.parse(body);
    const events = payload.events || [];

    const master = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
    const adminLineUserId = PropertiesService.getScriptProperties().getProperty('ADMIN_LINE_USER_ID') || '';

    for (const ev of events) {
      if (ev.type !== 'message') continue;
      if (!ev.message || ev.message.type !== 'text') continue;

      const userId = ev.source && ev.source.userId ? ev.source.userId : '';
      const textRaw = String(ev.message.text || '').trim();
      const replyToken = ev.replyToken || '';
      if (!userId || !replyToken) continue;

      // 管理者からのロック解除コマンドを処理
      if (adminLineUserId && userId === adminLineUserId) {
        // コマンド形式: "変更依頼: 講師名 月" または "変更依頼: 講師名" または "変更依頼:講師名"
        const unlockResult = handleAdminUnlockCommand_(master, textRaw);
        if (unlockResult.handled) {
          replyLine_(replyToken, unlockResult.message);
          continue;
        }
        // コマンドとして認識されなかった場合は通常の名前検索に進む（管理者も登録可能にするため）
      }

      const nameKey = normalizeNameKey_(textRaw);
      const result = linkLineUserByName_(master, nameKey, userId);

      if (result.status === 'linked') {
        replyLine_(replyToken, `登録OK：${result.name} さん\n今後はこのLINEでシフト連絡します。`);
      } else if (result.status === 'already_linked_same') {
        // ★二重返信を止める：何も返さない（静かに無視）
        // replyLine_(replyToken, 'すでに登録済みです。'); ←返したいならこれ
      } else if (result.status === 'already_linked_other') {
        replyLine_(replyToken, `この氏名は別のLINEと紐付いています：${result.name}\n教室まで連絡してください。`);
      } else if (result.status === 'multiple') {
        replyLine_(
          replyToken,
          `同じ氏名が複数います（候補：${result.candidates.join(' / ')}）\n` +
          `フルネームをそのまま送ってください（空白は気にしなくてOK）。`
        );
      } else {
        replyLine_(
          replyToken,
          `名簿（Teachers）に一致する氏名が見つかりませんでした：\n「${textRaw}」\n` +
          `Teachersの氏名表記と一致するように送ってください（空白は気にしなくてOK）。`
        );
      }

    }

    return ContentService.createTextOutput('OK');
  } catch (err) {
    handleError_(err, 'doPost', { 
      body: (e && e.postData && e.postData.contents) ? String(e.postData.contents).slice(0, 500) : 'no body' 
    });
    return ContentService.createTextOutput('OK');
  }
}

/**
 * Googleフォーム送信（インストール型トリガー推奨）
 * - テンプレから講師用シートをコピー作成
 * - 編集権限付与（Teachersのメール）
 * - Submissions追記
 * - _META書き込み
 * - LINEにURLをPush
 */
function onFormSubmit(e) {
  let teacherNameRaw = 'unknown';
  let monthKey = 'unknown';
  try {
    const nv = extractNamedValues_(e);

    const master = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);

    teacherNameRaw = (nv[CONFIG_FORM.QUESTION_TEACHER_NAME] || '').trim();
    if (!teacherNameRaw) throw new Error(`フォームに「${CONFIG_FORM.QUESTION_TEACHER_NAME}」がありません（または空です）`);

    monthKey = (nv[CONFIG_FORM.QUESTION_MONTH_KEY] || '').trim() || getNextMonthKey_();

    const teacher = findTeacherByName_(master, teacherNameRaw);
    if (!teacher) {
      appendSubmission_(master, {
        timestamp: new Date(),
        monthKey,
        teacherId: '',
        name: teacherNameRaw,
        sheetUrl: '',
        status: 'teacher_not_found',
        lastNotified: '',
        submissionKey: `${monthKey}|${normalizeNameKey_(teacherNameRaw)}`,
        submittedAt: '',
      });
      return;
    }

    const teacherId = teacher.teacherId || '';
    const teacherName = teacher.name;
    const teacherEmail = teacher.email || '';
    const lineUserId = teacher.lineUserId || '';

    // 月フォルダ確保 → テンプレコピー
    const monthFolderId = ensureMonthFolder_(monthKey, CONFIG.COPIES_PARENT_FOLDER_ID);
    const fileName = `${monthKey}_${teacherName}_シフト提出`;
    const newSpreadsheetId = copyTemplateSpreadsheet_(monthFolderId, CONFIG.TEMPLATE_SPREADSHEET_ID, fileName);
    const sheetUrl = `https://docs.google.com/spreadsheets/d/${newSpreadsheetId}/edit`;

    // 編集権限付与
    if (teacherEmail) {
      grantEditPermission_(newSpreadsheetId, teacherEmail);
    }

    // Submissionsに記録
    const submissionKey = `${monthKey}|${teacherId || normalizeNameKey_(teacherName)}`;
    appendSubmission_(master, {
      timestamp: new Date(),
      monthKey,
      teacherId,
      name: teacherName,
      sheetUrl,
      status: 'created',
      lastNotified: '',
      submissionKey,
      submittedAt: '',
    });

    // _META書き込み（テンプレ側表示や回収のため）
    writeMetaToTeacherSheet_(newSpreadsheetId, CONFIG.META_SHEET_NAME, {
      MASTER_SPREADSHEET_ID: CONFIG.MASTER_SPREADSHEET_ID,
      SUBMISSIONS_SHEET_NAME: CONFIG.SHEET_SUBMISSIONS,
      SUBMISSION_KEY: submissionKey,
      MONTH_KEY: monthKey,
      TEACHER_ID: teacherId,
      TEACHER_NAME: teacherName,
    });

    // 講師名をG3に直接設定（シートを開いた時にすぐ表示されるように）
    try {
      const teacherSs = SpreadsheetApp.openById(newSpreadsheetId);
      const inputSheet = teacherSs.getSheetByName('Input');
      if (inputSheet) {
        inputSheet.getRange('G3').setValue(teacherName);
      }
    } catch (e) {
      // G3の設定に失敗しても続行（onOpen()で後から設定される）
      console.error('Failed to set teacher name in G3:', e);
    }

    // LINEにURL送信（登録済みのみ）
    if (lineUserId) {
      pushLine_(lineUserId,
        `【シフト提出URL（${monthKey}）】\n${sheetUrl}\n\n入力後、☑（提出）を入れてください。`
      );
    }

  } catch (err) {
    handleError_(err, 'onFormSubmit', {
      teacherNameRaw,
      monthKey
    });
  }
}


/**
 * 管理者からの「henshin」単独コマンド：最新の提出済みシートをロック解除
 * @param {SpreadsheetApp.Spreadsheet} masterSs - マスタースプレッドシート
 * @returns {Object} {handled: boolean, message: string}
 */
function handleAdminUnlockLatest_(masterSs) {
  try {
    const sh = masterSs.getSheetByName(CONFIG.SHEET_SUBMISSIONS);
    if (!sh) {
      return { handled: true, message: 'Submissionsシートが見つかりません' };
    }

    const values = sh.getDataRange().getValues();
    if (values.length < 2) {
      return { handled: true, message: '提出データが見つかりません' };
    }

    const header = values[0];
    const idxUrl = header.indexOf('sheetUrl');
    const idxStatus = header.indexOf('status');
    const idxName = header.indexOf('氏名');
    const idxMonthKey = header.indexOf('monthKey');
    const idxLockedAt = header.indexOf('lockedAt');
    const idxTeacherId = header.indexOf('teacherId');
    const idxSubmittedAt = header.indexOf('submittedAt');

    if (idxUrl < 0 || idxStatus < 0 || idxName < 0) {
      return { handled: true, message: 'Submissionsに必要列がありません' };
    }

    // 最新の提出済みシートを検索（submittedAtが最新のもの）
    let latestRow = -1;
    let latestSubmittedAt = null;

    for (let r = 1; r < values.length; r++) {
      const status = String(values[r][idxStatus] || '').trim();
      if (status !== 'submitted') continue;

      const submittedAt = values[r][idxSubmittedAt];
      if (submittedAt && (!latestSubmittedAt || new Date(submittedAt) > new Date(latestSubmittedAt))) {
        latestRow = r + 1;
        latestSubmittedAt = submittedAt;
      }
    }

    if (latestRow < 0) {
      return { handled: true, message: '提出済みのデータが見つかりません' };
    }

    const targetUrl = String(values[latestRow - 1][idxUrl] || '').trim();
    const targetTeacherId = idxTeacherId >= 0 ? String(values[latestRow - 1][idxTeacherId] || '').trim() : '';
    const targetTeacherName = String(values[latestRow - 1][idxName] || '').trim();
    const targetMonthKey = idxMonthKey >= 0 ? String(values[latestRow - 1][idxMonthKey] || '').trim() : '';

    if (!targetUrl) {
      return { handled: true, message: 'シートURLが見つかりません' };
    }

    // ロック解除
    const spreadsheetId = extractSpreadsheetId_(targetUrl);
    if (!spreadsheetId) {
      return { handled: true, message: 'シートIDの取得に失敗しました' };
    }

    // 講師情報を取得
    const teacher = getTeacherInfo_(masterSs, targetTeacherId, targetTeacherName);
    const teacherEmail = teacher ? teacher.email || '' : '';
    const lineUserId = teacher ? teacher.lineUserId || '' : '';

    // ロック解除
    const unlocked = unlockTeacherSheet_(spreadsheetId, teacherEmail);
    if (!unlocked) {
      return { handled: true, message: 'ロック解除に失敗しました' };
    }

    // SubmissionsのlockedAtをクリア
    if (idxLockedAt >= 0) {
      sh.getRange(latestRow, idxLockedAt + 1).setValue('');
    }

    // 講師にLINE通知
    if (lineUserId) {
      pushLine_(lineUserId,
        `【シフト変更依頼】\n${targetTeacherName}さん（${targetMonthKey}）のシフトを変更していただくようお願いします。\nシートの編集が可能になりました。\n${targetUrl}`
      );
    }

    return { handled: true, message: `ロック解除しました：${targetTeacherName}さん（${targetMonthKey}）` };

  } catch (err) {
    handleError_(err, 'handleAdminUnlockLatest_');
    return { handled: true, message: 'エラーが発生しました：' + (err.message || String(err)) };
  }
}

/**
 * 管理者からのロック解除コマンドを処理
 * コマンド形式: "変更依頼: 講師名 月" または "変更依頼: 講師名"
 * @param {SpreadsheetApp.Spreadsheet} masterSs - マスタースプレッドシート
 * @param {string} command - コマンド文字列
 * @returns {Object} {handled: boolean, message: string}
 */
function handleAdminUnlockCommand_(masterSs, command) {
  try {
    // コマンド形式: 
    // - "変更依頼: 講師名 月" または "変更依頼: 講師名" または "変更依頼:講師名"
    // - "変更依頼 講師名 月" または "変更依頼 講師名"（コロンなし、スペースのみ）
    // コロンの有無、スペースの有無を柔軟に対応
    const trimmedCommand = command.trim();
    
    // パターン1: コロンあり（全角/半角）
    let match = trimmedCommand.match(/^変更依頼[：:]\s*(.+?)(?:\s+(\d{4}-\d{2}))?\s*$/);
    
    // パターン2: コロンなし、スペースで始まる
    if (!match) {
      match = trimmedCommand.match(/^変更依頼\s+(.+?)(?:\s+(\d{4}-\d{2}))?\s*$/);
    }
    
    if (!match || !match[1]) {
      return { handled: false, message: '' };
    }

    let teacherNameRaw = match[1].trim();
    // 月の形式（YYYY-MM）が講師名に含まれている場合は除外
    if (teacherNameRaw.match(/^\d{4}-\d{2}$/)) {
      return { handled: false, message: '' };
    }
    
    if (!teacherNameRaw) {
      return { handled: false, message: '' };
    }
    const monthKey = match[2] ? match[2].trim() : '';

    // Submissionsから該当する提出を検索
    const sh = masterSs.getSheetByName(CONFIG.SHEET_SUBMISSIONS);
    if (!sh) {
      return { handled: true, message: 'Submissionsシートが見つかりません' };
    }

    const values = sh.getDataRange().getValues();
    if (values.length < 2) {
      return { handled: true, message: '提出データが見つかりません' };
    }

    const header = values[0];
    const idxUrl = header.indexOf('sheetUrl');
    const idxStatus = header.indexOf('status');
    const idxName = header.indexOf('氏名');
    const idxMonthKey = header.indexOf('monthKey');
    const idxLockedAt = header.indexOf('lockedAt');
    const idxTeacherId = header.indexOf('teacherId');

    if (idxUrl < 0 || idxStatus < 0 || idxName < 0) {
      return { handled: true, message: 'Submissionsに必要列がありません' };
    }

    // 該当する提出を検索
    let targetRow = -1;
    let targetUrl = '';
    let targetTeacherId = '';
    let targetTeacherName = '';
    let targetMonthKey = '';

    for (let r = 1; r < values.length; r++) {
      const name = String(values[r][idxName] || '').trim();
      const mk = idxMonthKey >= 0 ? String(values[r][idxMonthKey] || '').trim() : '';
      const status = String(values[r][idxStatus] || '').trim();

      const nameMatch = normalizeNameKey_(name) === normalizeNameKey_(teacherNameRaw);
      const monthMatch = !monthKey || mk === monthKey;

      if (nameMatch && monthMatch && status === 'submitted') {
        targetRow = r + 1;
        targetUrl = String(values[r][idxUrl] || '').trim();
        targetTeacherId = idxTeacherId >= 0 ? String(values[r][idxTeacherId] || '').trim() : '';
        targetTeacherName = name;
        targetMonthKey = mk;
        break;
      }
    }

    if (targetRow < 0) {
      return { handled: true, message: `提出済みのデータが見つかりません：${teacherNameRaw}${monthKey ? ' ' + monthKey : ''}` };
    }

    if (!targetUrl) {
      return { handled: true, message: 'シートURLが見つかりません' };
    }

    // ロック解除
    const spreadsheetId = extractSpreadsheetId_(targetUrl);
    if (!spreadsheetId) {
      return { handled: true, message: 'シートIDの取得に失敗しました' };
    }

    // 講師情報を取得
    const teacher = getTeacherInfo_(masterSs, targetTeacherId, targetTeacherName);
    const teacherEmail = teacher ? teacher.email || '' : '';
    const lineUserId = teacher ? teacher.lineUserId || '' : '';

    // ロック解除
    const unlocked = unlockTeacherSheet_(spreadsheetId, teacherEmail);
    if (!unlocked) {
      return { handled: true, message: 'ロック解除に失敗しました' };
    }

    // SubmissionsのlockedAtをクリア
    if (idxLockedAt >= 0) {
      sh.getRange(targetRow, idxLockedAt + 1).setValue('');
    }

    // 講師にLINE通知
    if (lineUserId) {
      pushLine_(lineUserId,
        `【シフト変更依頼】\n${targetTeacherName}さん（${targetMonthKey}）のシフトを変更していただくようお願いします。\nシートの編集が可能になりました。\n${targetUrl}`
      );
    }

    return { handled: true, message: `ロック解除しました：${targetTeacherName}さん（${targetMonthKey}）` };

  } catch (err) {
    handleError_(err, 'handleAdminUnlockCommand_', { command });
    return { handled: true, message: 'エラーが発生しました：' + (err.message || String(err)) };
  }
}

/************************************************************
 * テンプレ側（講師用シート）の関数
 * これらの関数は講師用シート（テンプレートからコピーされたシート）で実行されます
 ************************************************************/

const SUBMIT_CONFIG = {
  INPUT_SHEET_NAME: 'Input',
  META_SHEET_NAME: '_META',

  CHECK_ROW: 2,
  CHECK_COL: 3, // C2

  STATUS_CELL_A1: 'B2',
  HELP_CELL_A1: 'D2',
  TEACHER_NAME_CELL_A1: 'G3',
};

/**
 * 講師用シートが開かれた時に実行
 * - パネルを確保
 * - 講師名を設定
 * - ロック解除状態をチェック
 */
function onOpen() {
  ensurePanel_();
  setTeacherNameFromMeta_();
  // ロック解除状態をチェック（マスター側から解除された場合の復元）
  checkAndUnlockIfNeeded_();
}

/**
 * 講師用シートの編集時に実行
 * - Input!C2がTRUEになったら提出済みにする
 * - シートをロック（編集不可にする）
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (sh.getName() !== SUBMIT_CONFIG.INPUT_SHEET_NAME) return;

    if (e.range.getRow() !== SUBMIT_CONFIG.CHECK_ROW || e.range.getColumn() !== SUBMIT_CONFIG.CHECK_COL) return;

    const val = e.range.getValue();
    if (val !== true) return;

    // 提出済みに設定
    sh.getRange(SUBMIT_CONFIG.STATUS_CELL_A1).setValue('提出済');

    // シートをロック（編集不可にする）
    lockSheetAfterSubmission_();

    SpreadsheetApp.getActiveSpreadsheet().toast('提出済にしました。シートは編集不可になりました。', 'シフト提出', 5);
  } catch (err) {
    console.error('onEdit error:', err);
  }
}

/**
 * 提出後にシートをロック（編集不可にする）
 * - すべてのシートを保護
 * - 現在のユーザー（講師）の編集権限を削除
 * - 管理者（スクリプト実行者）は常に編集可能
 */
function lockSheetAfterSubmission_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    const currentUser = Session.getActiveUser().getEmail();
    const scriptOwner = Session.getEffectiveUser().getEmail();

    // 既存の保護を削除（重複防止）
    sheets.forEach(sheet => {
      const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      protections.forEach(prot => {
        if (prot.getDescription() === '提出後ロック') {
          prot.remove();
        }
      });
    });

    // 各シートを保護
    sheets.forEach(sheet => {
      const protection = sheet.protect().setDescription('提出後ロック');

      // 管理者（スクリプト実行者）は常に編集可能
      try {
        protection.addEditor(scriptOwner);
      } catch (e) {
        // 既にエディターリストに含まれている場合は無視
      }

      // 現在のユーザー（講師）の編集権限を削除
      if (currentUser && currentUser !== scriptOwner) {
        try {
          protection.removeEditor(currentUser);
        } catch (e) {
          // エディターリストに含まれていない場合は無視
        }
      }
    });

    // ファイルレベルでも講師の編集権限を削除（閲覧のみに変更）
    try {
      const file = DriveApp.getFileById(ss.getId());
      if (currentUser && currentUser !== scriptOwner) {
        file.removeEditor(currentUser);
        file.addViewer(currentUser);
      }
    } catch (e) {
      // ファイルレベルの権限変更に失敗しても続行
      console.error('File-level permission change failed:', e);
    }

  } catch (err) {
    console.error('lockSheetAfterSubmission_ error:', err);
    // ロックに失敗しても提出処理は続行
  }
}

/**
 * シートのロックを解除（管理者からの変更依頼時）
 * この関数は管理者が手動で実行するか、マスター側から呼び出される
 * onOpen時に自動的にロック解除状態をチェックして復元する
 */
function unlockSheetForRevision_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    const currentUser = Session.getActiveUser().getEmail();

    // すべての保護を解除（説明が「提出後ロック」のもの、または空の保護も）
    let protectionRemoved = false;
    sheets.forEach(sheet => {
      const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      protections.forEach(prot => {
        const desc = prot.getDescription();
        if (desc === '提出後ロック' || desc === '' || !desc) {
          try {
            prot.remove();
            protectionRemoved = true;
          } catch (e) {
            console.error('Failed to remove protection:', e);
          }
        }
      });
    });

    // ファイルレベルで編集権限を再付与
    try {
      const file = DriveApp.getFileById(ss.getId());
      const viewers = file.getViewers().map(u => u.getEmail());
      
      if (currentUser && viewers.includes(currentUser)) {
        file.removeViewer(currentUser);
        file.addEditor(currentUser);
      }
    } catch (e) {
      console.error('File-level permission restoration failed:', e);
    }

    if (protectionRemoved) {
      SpreadsheetApp.getActiveSpreadsheet().toast('ロックを解除しました。編集可能になりました。', 'ロック解除', 5);
    }
    return true;
  } catch (err) {
    console.error('unlockSheetForRevision_ error:', err);
    return false;
  }
}

/**
 * onOpen時にロック状態をチェックして、必要に応じてロック解除
 * マスター側からロック解除された場合、講師がシートを開いた時に自動的に編集可能になる
 */
function checkAndUnlockIfNeeded_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const file = DriveApp.getFileById(ss.getId());
    const currentUser = Session.getActiveUser().getEmail();
    
    // 現在のユーザーがビューアーの場合、ロック解除を試みる
    const viewers = file.getViewers().map(u => u.getEmail());
    if (currentUser && viewers.includes(currentUser)) {
      // マスター側でロック解除されている可能性があるので、保護を解除
      unlockSheetForRevision_();
    }
  } catch (e) {
    // エラーは無視
  }
}

/**
 * パネルを確保（Inputシートの状態表示）
 */
function ensurePanel_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SUBMIT_CONFIG.INPUT_SHEET_NAME);
  if (!sh) return;

  if (!String(sh.getRange(SUBMIT_CONFIG.STATUS_CELL_A1).getValue() || '').trim()) {
    sh.getRange(SUBMIT_CONFIG.STATUS_CELL_A1).setValue('未提出');
  }

  const help = sh.getRange(SUBMIT_CONFIG.HELP_CELL_A1);
  if (!String(help.getValue() || '').trim()) help.setValue('入力後、☑で提出完了');
}

/**
 * _METAシートから講師名を取得して表示
 * シートを開いた時に必ず講師名がG3に表示されるようにする
 */
function setTeacherNameFromMeta_() {
  try {
    const meta = getMeta_();
    const name = String(meta.TEACHER_NAME || '').trim();
    if (!name) return;

    const sh = SpreadsheetApp.getActive().getSheetByName(SUBMIT_CONFIG.INPUT_SHEET_NAME);
    if (!sh) return;

    const cell = sh.getRange(SUBMIT_CONFIG.TEACHER_NAME_CELL_A1);
    const currentValue = String(cell.getValue() || '').trim();
    // 現在の値が空、または異なる場合は更新
    if (!currentValue || currentValue !== name) {
      cell.setValue(name);
    }
  } catch (e) {
    console.error('setTeacherNameFromMeta_ error:', e);
  }
}

/**
 * _METAシートから情報を取得
 */
function getMeta_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SUBMIT_CONFIG.META_SHEET_NAME);
  if (!sh) return {};

  const data = sh.getDataRange().getValues();
  const obj = {};
  for (const [k, v] of data) if (k) obj[String(k).trim()] = String(v || '').trim();
  return obj;
}

// synced from vscode
