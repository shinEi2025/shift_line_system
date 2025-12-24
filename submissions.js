/************************************************************
 * submissions.gs
 * 1) pollSubmissionsAndUpdate:
 *    - 講師シート Input!C2 が TRUE になったら提出済みにする
 *    - その瞬間に「提出受理しました」をLINEで1回だけ送る（ackNotifiedAtで抑止）
 *
 * 2) remindUnsubmitted:
 *    - 未提出（status != submitted）の講師だけにリマインド（1日1回）
 *    - reminderNotifiedAtで抑止
 ************************************************************/

function pollSubmissionsAndUpdate() {
  try {
    const master = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
    const sh = master.getSheetByName(CONFIG.SHEET_SUBMISSIONS);
    if (!sh) throw new Error('Submissionsが見つかりません');

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return;

  const header = values[0];
  const idxUrl = header.indexOf('sheetUrl');
  const idxStatus = header.indexOf('status');
  const idxSubmittedAt = header.indexOf('submittedAt');
  const idxTeacherId = header.indexOf('teacherId');
  const idxName = header.indexOf('氏名');
  const idxMonthKey = header.indexOf('monthKey');
  const idxAck = header.indexOf('ackNotifiedAt');
  const idxLockedAt = header.indexOf('lockedAt');

  if (idxUrl < 0 || idxStatus < 0 || idxSubmittedAt < 0) {
    throw new Error('Submissionsに必要列がありません（sheetUrl/status/submittedAt）');
  }
  if (idxAck < 0) {
    throw new Error('Submissionsに ackNotifiedAt 列を追加してください（提出受理LINEの一回送信制御に必要）');
  }

  for (let r = 1; r < values.length; r++) {
    const status = String(values[r][idxStatus] || '').trim();
    if (status === 'submitted') continue;

    const url = String(values[r][idxUrl] || '').trim();
    if (!url) continue;

    const id = extractSpreadsheetId_(url);
    if (!id) continue;

    // 講師シートの提出フラグ
    const submitted = readTeacherSubmittedFlag_(id);
    if (!submitted) continue;

    // ① Submissions 更新
    sh.getRange(r + 1, idxStatus + 1).setValue('submitted');
    sh.getRange(r + 1, idxSubmittedAt + 1).setValue(new Date());

    const teacherId = idxTeacherId >= 0 ? String(values[r][idxTeacherId] || '').trim() : '';
    const teacherName = idxName >= 0 ? String(values[r][idxName] || '').trim() : '';
    // monthKeyを適切にフォーマット（Dateオブジェクトの場合は文字列に変換）
    let monthKey = '';
    if (idxMonthKey >= 0) {
      const monthKeyValue = values[r][idxMonthKey];
      if (monthKeyValue instanceof Date) {
        // Dateオブジェクトの場合は YYYY-MM 形式に変換
        const year = monthKeyValue.getFullYear();
        const month = String(monthKeyValue.getMonth() + 1).padStart(2, '0');
        monthKey = `${year}-${month}`;
      } else {
        monthKey = String(monthKeyValue || '').trim();
      }
    }

    // ② 講師シートの表示も更新（任意：B2を提出済に）
    let teacherEmail = '';
    try {
      const ss = SpreadsheetApp.openById(id);
      const input = ss.getSheetByName('Input');
      if (input) input.getRange('B2').setValue('提出済');
      
      // 講師のメールアドレスを取得（ロック用）
      const teacher = getTeacherInfo_(master, teacherId, teacherName);
      teacherEmail = teacher ? teacher.email || '' : '';
    } catch (e) {}

    // ③ シートをロック（提出後は編集不可）
    const alreadyLocked = idxLockedAt >= 0 && values[r][idxLockedAt];
    if (!alreadyLocked && teacherEmail) {
      const locked = lockTeacherSheet_(id, teacherEmail);
      if (locked && idxLockedAt >= 0) {
        sh.getRange(r + 1, idxLockedAt + 1).setValue(new Date());
      }
    }

    // ④ 提出受理LINE（1回だけ）
    const ackAlready = values[r][idxAck];
    if (ackAlready) continue;

    const lineUserId = getTeacherLineUserId_(master, teacherId, teacherName);
    if (lineUserId) {
      const lastName = extractLastName_(teacherName);
      pushLine_(lineUserId, `【提出受理】\n${lastName}先生（${monthKey}）のシフト提出を受け付けました。ありがとうございます！`);
      sh.getRange(r + 1, idxAck + 1).setValue(new Date());
    }
  }
  } catch (err) {
    handleError_(err, 'pollSubmissionsAndUpdate');
  }
}

/** 講師用シートの Input!C2 が TRUE なら提出済み */
function readTeacherSubmittedFlag_(teacherSpreadsheetId) {
  const ss = SpreadsheetApp.openById(teacherSpreadsheetId);
  const sh = ss.getSheetByName('Input');
  if (!sh) return false;
  return sh.getRange('C2').getValue() === true;
}

/**
 * 未提出リマインド（未提出だけ・1日1回）
 * - まずは「最新の monthKey（YYYY-MM）」を対象にする
 * - reminderNotifiedAt が今日なら送らない
 */
function remindUnsubmitted() {
  try {
    const master = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
    const sh = master.getSheetByName(CONFIG.SHEET_SUBMISSIONS);
    if (!sh) throw new Error('Submissionsが見つかりません');

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return;

  const header = values[0];
  const idxMonthKey = header.indexOf('monthKey');
  const idxStatus = header.indexOf('status');
  const idxUrl = header.indexOf('sheetUrl');
  const idxTeacherId = header.indexOf('teacherId');
  const idxName = header.indexOf('氏名');
  const idxReminder = header.indexOf('reminderNotifiedAt');

  if (idxMonthKey < 0 || idxStatus < 0 || idxUrl < 0) {
    throw new Error('Submissionsに必要列がありません（monthKey/status/sheetUrl）');
  }
  if (idxReminder < 0) {
    throw new Error('Submissionsに reminderNotifiedAt 列を追加してください（リマインドの一回/日制御に必要）');
  }

  // 対象月：存在するmonthKeyのうち最大（YYYY-MMなら文字列maxでOK）
  let targetMonth = '';
  for (let r = 1; r < values.length; r++) {
    const mk = String(values[r][idxMonthKey] || '').trim();
    if (mk && mk > targetMonth) targetMonth = mk;
  }
  if (!targetMonth) return;

  const today = new Date();

  for (let r = 1; r < values.length; r++) {
    const mk = String(values[r][idxMonthKey] || '').trim();
    if (mk !== targetMonth) continue;

    const status = String(values[r][idxStatus] || '').trim();
    if (status === 'submitted') continue;

    const url = String(values[r][idxUrl] || '').trim();
    if (!url) continue;

    // すでに今日送ってたらスキップ
    const last = values[r][idxReminder];
    if (last && isSameJstDate_(last, today)) continue;

    const teacherId = idxTeacherId >= 0 ? String(values[r][idxTeacherId] || '').trim() : '';
    const teacherName = idxName >= 0 ? String(values[r][idxName] || '').trim() : '';

    const lineUserId = getTeacherLineUserId_(master, teacherId, teacherName);
    if (!lineUserId) continue;

    const lastName = extractLastName_(teacherName);
    pushLine_(lineUserId,
      `【シフト未提出リマインド】\n${lastName}先生（${targetMonth}）の提出がまだのようです。\nこちらから入力・提出（☑）をお願いします。\n${url}`
    );

    sh.getRange(r + 1, idxReminder + 1).setValue(new Date());
  }
  } catch (err) {
    handleError_(err, 'remindUnsubmitted');
  }
}
