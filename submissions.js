/************************************************************
 * submissions.gs
 * 1) pollSubmissionsAndUpdate:
 *    - 講師シート Input!C2 が TRUE になったら提出済みにする
 *    - その瞬間に「提出受理しました」をLINEで1回だけ送る（ackNotifiedAtで抑止）
 *
 * 2) remindUnsubmitted:
 *    - 未提出（status != submitted）の講師だけにリマインド（1日1回）
 *    - reminderNotifiedAtで抑止
 *
 * 3) remindUnappliedToManager:
 *    - 管理者に未提出者リストをLINE通知（今月と来月を対象）
 *    - 3週間前（来月のみ、午後3時）：テンプレート確認し、準備状況に応じて管理者に通知
 *    - 対象月の2週間前、10日前、1週間前、3日前、1日前、1日に通知
 *    - 毎月1日には前月の未提出者リストを管理者に通知し、未提出者にもLINE通知を送信
 *    - 1日には提出済みリストも表示（SHIFT SYNC）
 *    - managerReminder3Weeks / managerReminder2Weeks / managerReminder10Days / managerReminder1Week / managerReminder3Days / managerReminder1Day / managerReminder1stで抑止
 * 
 * 4) setupRemindUnappliedToManagerTrigger:
 *    - remindUnappliedToManagerのトリガーを設定（手動実行用）
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

    // パフォーマンス最適化：処理対象のレコードを先にフィルタリング
    // status='submitted' かつ ackNotifiedAt が設定済みのレコードはスキップ
    const recordsToCheck = [];
    for (let r = 1; r < values.length; r++) {
      const status = String(values[r][idxStatus] || '').trim();
      const url = String(values[r][idxUrl] || '').trim();
      const ackAlready = values[r][idxAck];

      // 既に提出済みでLINE通知も完了している場合はスキップ（最重要の最適化）
      if (status === 'submitted' && ackAlready) continue;

      // URLがない場合はスキップ
      if (!url) continue;

      const id = extractSpreadsheetId_(url);
      if (!id) continue;

      recordsToCheck.push({
        rowIndex: r,
        id: id,
        url: url,
        status: status,
        ackAlready: ackAlready,
        values: values[r]
      });
    }

    console.log(`[pollSubmissionsAndUpdate] 処理対象レコード数: ${recordsToCheck.length}/${values.length - 1}`);

    // 処理対象がない場合は早期リターン
    if (recordsToCheck.length === 0) return;

    for (const record of recordsToCheck) {
      const r = record.rowIndex;
      const id = record.id;

      // 各レコードの処理を個別にtry-catchで囲む（1つが失敗しても他の処理を続行）
      try {
        // 既に提出済み（status='submitted'）の場合は、LINE通知の処理のみ行う
        if (record.status === 'submitted') {
          // ackNotifiedAtが未設定の場合のみLINE通知を送信
          if (!record.ackAlready) {
            const teacherId = idxTeacherId >= 0 ? String(record.values[idxTeacherId] || '').trim() : '';
            const teacherName = idxName >= 0 ? String(record.values[idxName] || '').trim() : '';
            const monthKey = formatMonthKey_(record.values[idxMonthKey]);

            try {
              const lineUserId = getTeacherLineUserId_(master, teacherId, teacherName);
              if (lineUserId) {
                const lastName = extractLastName_(teacherName);
                pushLine_(lineUserId, `【提出受理】\n${lastName}先生（${monthKey}）のシフト提出を受け付けました。ありがとうございます！`);
                sh.getRange(r + 1, idxAck + 1).setValue(new Date());
              }
            } catch (e) {
              console.error(`[pollSubmissionsAndUpdate] Failed to send LINE notification for ${teacherName}:`, e);
            }
          }
          continue;
        }

        // 講師シートの提出フラグをチェック（エラーハンドリング付き）
        let submitted = false;
        try {
          submitted = readTeacherSubmittedFlag_(id);
        } catch (e) {
          console.error(`[pollSubmissionsAndUpdate] Failed to read submission flag for ${id}:`, e);
          continue; // このスプレッドシートの処理をスキップ
        }

        if (!submitted) continue;

        // ① Submissions 更新
        const now = new Date();
        try {
          sh.getRange(r + 1, idxStatus + 1).setValue('submitted');
          sh.getRange(r + 1, idxSubmittedAt + 1).setValue(now);
        } catch (e) {
          console.error(`[pollSubmissionsAndUpdate] Failed to update Submissions row ${r + 1}:`, e);
          continue; // 更新に失敗した場合は次のレコードへ
        }

        const teacherId = idxTeacherId >= 0 ? String(record.values[idxTeacherId] || '').trim() : '';
        const teacherName = idxName >= 0 ? String(record.values[idxName] || '').trim() : '';
        const monthKey = formatMonthKey_(record.values[idxMonthKey]);

        // ② 講師シートの表示も更新（任意：B2を提出済に）
        let teacherEmail = '';
        try {
          const ss = SpreadsheetApp.openById(id);
          const input = ss.getSheetByName('Input');
          if (input) input.getRange('B2').setValue('提出済');

          // 講師のメールアドレスを取得（ロック用）
          const teacher = getTeacherInfo_(master, teacherId, teacherName);
          teacherEmail = teacher ? teacher.email || '' : '';
        } catch (e) {
          console.error(`[pollSubmissionsAndUpdate] Failed to update teacher sheet ${id}:`, e);
          // シートの更新に失敗しても続行
        }

        // ③ シートをロック（提出後は編集不可）
        const alreadyLocked = idxLockedAt >= 0 && record.values[idxLockedAt];
        if (!alreadyLocked && teacherEmail) {
          try {
            const locked = lockTeacherSheet_(id, teacherEmail);
            if (locked && idxLockedAt >= 0) {
              sh.getRange(r + 1, idxLockedAt + 1).setValue(now);
            }
          } catch (e) {
            console.error(`[pollSubmissionsAndUpdate] Failed to lock sheet ${id}:`, e);
            // ロックに失敗しても続行（通知は送る）
          }
        }

        // ④ 提出受理LINE通知
        try {
          const lineUserId = getTeacherLineUserId_(master, teacherId, teacherName);
          if (lineUserId) {
            const lastName = extractLastName_(teacherName);
            pushLine_(lineUserId, `【提出受理】\n${lastName}先生（${monthKey}）のシフト提出を受け付けました。ありがとうございます！`);
            sh.getRange(r + 1, idxAck + 1).setValue(now);
          }
        } catch (e) {
          console.error(`[pollSubmissionsAndUpdate] Failed to send LINE notification for ${teacherName}:`, e);
          // LINE通知に失敗しても続行
        }
      } catch (err) {
        // 個々のレコード処理で予期しないエラーが発生した場合
        console.error(`[pollSubmissionsAndUpdate] Unexpected error processing row ${r + 1}:`, err);
        // 次のレコードの処理を続行
        continue;
      }
    }
  } catch (err) {
    handleError_(err, 'pollSubmissionsAndUpdate');
  }
}

/**
 * monthKeyをYYYY-MM形式にフォーマット
 * @param {any} monthKeyValue - monthKeyの値（Dateまたは文字列）
 * @returns {string} YYYY-MM形式の文字列
 */
function formatMonthKey_(monthKeyValue) {
  if (!monthKeyValue) return '';
  if (monthKeyValue instanceof Date) {
    const year = monthKeyValue.getFullYear();
    const month = String(monthKeyValue.getMonth() + 1).padStart(2, '0');
    return `${year}-${month}`;
  }
  return String(monthKeyValue || '').trim();
}

/** 講師用シートの Input!C2 が TRUE なら提出済み */
function readTeacherSubmittedFlag_(teacherSpreadsheetId) {
  const ss = SpreadsheetApp.openById(teacherSpreadsheetId);
  const sh = ss.getSheetByName('Input');
  if (!sh) return false;
  return sh.getRange('C2').getValue() === true;
}

/**
 * 管理者向け未提出リマインド
 * - 来月のシフト提出について、来月の1週間前、3日前、1日前、1日に管理者にLINE通知
 * - 来月1日には未提出者リストと提出済みリストの両方を通知
 * - 毎月1日には前月の未提出者リストを管理者に通知し、未提出者にもLINE通知を送信
 */
function remindUnappliedToManager() {
  try {
    const master = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
    const sh = master.getSheetByName(CONFIG.SHEET_SUBMISSIONS);
    if (!sh) throw new Error('Submissionsが見つかりません');

    const values = sh.getDataRange().getValues();
    if (values.length < 2) return;

    const header = values[0];
    const idxMonthKey = header.indexOf('monthKey');
    const idxStatus = header.indexOf('status');
    const idxName = header.indexOf('氏名');
    const idxTeacherId = header.indexOf('teacherId');
    const idxUrl = header.indexOf('sheetUrl');

    if (idxMonthKey < 0 || idxStatus < 0 || idxName < 0) {
      throw new Error('Submissionsに必要列がありません（monthKey/status/氏名）');
    }

    const adminLineUserId = PropertiesService.getScriptProperties().getProperty('ADMIN_LINE_USER_ID');
    if (!adminLineUserId) {
      console.log('ADMIN_LINE_USER_IDが未設定のため、管理者リマインドをスキップします');
      return;
    }

    // リマインダー設定をスプレッドシートから読み込む
    const reminderSettings = getReminderSettings_(master);
    console.log(`リマインダー設定を読み込みました: ${reminderSettings.length}件`);

    // 休職中の講師リストを取得
    const onLeaveTeachers = getOnLeaveTeachers_(master);

    const today = new Date();
    const currentMonthKey = getCurrentMonthKey_();
    const nextMonthKey = getNextMonthKey_();
    const prevMonthKey = getPreviousMonthKey_();

    // 毎月1日の処理：前月の未提出者を通知
    const isFirstDayOfMonth = today.getDate() === 1;
    if (isFirstDayOfMonth && prevMonthKey) {
      handlePreviousMonthUnapplied_(master, sh, values, header, prevMonthKey, idxMonthKey, idxStatus, idxName, idxTeacherId, idxUrl, adminLineUserId, onLeaveTeachers);
    }

    // 初回申請依頼処理（来月分、設定された日数前）
    if (nextMonthKey) {
      const initialRequestSetting = reminderSettings.find(s => s.id === 'initial_request');
      if (initialRequestSetting) {
        processInitialShiftRequest_(master, sh, values, header, nextMonthKey, idxMonthKey, adminLineUserId, initialRequestSetting, today);
      }
    }

    // 対象月を決定（今月と来月）
    const targetMonths = [];
    if (currentMonthKey) targetMonths.push(currentMonthKey);
    if (nextMonthKey && nextMonthKey !== currentMonthKey) targetMonths.push(nextMonthKey);

    for (const targetMonth of targetMonths) {
      const monthStart = getMonthStartDate_(targetMonth);
      if (!monthStart) continue;

      // 今日該当するリマインダー設定を検索（初回申請依頼以外）
      const matchedReminder = reminderSettings.find(setting => {
        if (setting.id === 'initial_request') return false; // 初回申請依頼は別処理
        return isReminderDay_(today, monthStart, setting.daysBeforeDeadline);
      });

      if (!matchedReminder) continue; // 今日はリマインド日ではない

      console.log(`リマインダー該当: ${matchedReminder.id} (${targetMonth})`);

      // 対象月の未提出者と提出済み者を取得
      const { unappliedList, appliedList } = getSubmissionLists_(values, idxMonthKey, idxStatus, idxName, idxTeacherId, idxUrl, targetMonth, onLeaveTeachers);

      // 未提出者リストのテキストを生成
      const submissionListText = formatSubmissionListForReminder_(unappliedList, appliedList, matchedReminder.daysBeforeDeadline === 0);

      // メッセージテンプレートを置換
      const message = replaceMessageVariables_(matchedReminder.messageTemplate, {
        monthKey: targetMonth,
        submissionList: submissionListText
      });

      // 管理者に通知（manager の場合）
      if (matchedReminder.targetAudience === 'manager' &&
          (unappliedList.length > 0 || (matchedReminder.daysBeforeDeadline === 0 && appliedList.length > 0))) {
        pushLine_(adminLineUserId, message.trim());
      }

      // 講師に通知（teacher の場合）
      if (matchedReminder.targetAudience === 'teacher' && unappliedList.length > 0) {
        sendUnappliedReminderToTeachers_(master, unappliedList, targetMonth);
      }
    }
  } catch (err) {
    handleError_(err, 'remindUnappliedToManager');
  }
}

/**
 * 初回シフト申請依頼を処理
 */
function processInitialShiftRequest_(master, sh, values, header, nextMonthKey, idxMonthKey, adminLineUserId, setting, today) {
  const monthStart = getMonthStartDate_(nextMonthKey);
  if (!monthStart) return;

  // 今日が初回申請依頼日かチェック
  if (!isReminderDay_(today, monthStart, setting.daysBeforeDeadline)) return;

  // テンプレート確認
  checkTemplateThreeWeeksBefore_(master, sh, values, header, nextMonthKey, idxMonthKey, adminLineUserId);
}

/**
 * 提出状況リストを取得
 */
function getSubmissionLists_(values, idxMonthKey, idxStatus, idxName, idxTeacherId, idxUrl, targetMonth, onLeaveTeachers) {
  const unappliedList = [];
  const appliedList = [];

  for (let r = 1; r < values.length; r++) {
    const mk = normalizeMonthKey_(values[r][idxMonthKey]);
    if (mk !== targetMonth) continue;

    const status = String(values[r][idxStatus] || '').trim();
    const teacherName = String(values[r][idxName] || '').trim();

    if (!teacherName) continue;

    // 休職中の講師はスキップ
    if (onLeaveTeachers.has(teacherName)) continue;

    if (status === 'submitted') {
      appliedList.push(teacherName);
    } else if (status === 'created' || status === '' || status === 'teacher_not_found' || status === 'template_not_found') {
      const sheetUrl = idxUrl >= 0 ? String(values[r][idxUrl] || '').trim() : '';
      unappliedList.push({
        row: r + 1,
        name: teacherName,
        teacherId: idxTeacherId >= 0 ? String(values[r][idxTeacherId] || '').trim() : '',
        sheetUrl: sheetUrl
      });
    }
  }

  return { unappliedList, appliedList };
}

/**
 * リマインダー用の提出状況リストをフォーマット
 */
function formatSubmissionListForReminder_(unappliedList, appliedList, includeApplied) {
  let text = '';

  if (unappliedList.length > 0) {
    text += '【未提出者】\n';
    const withSheetUrl = unappliedList.filter(item => item.sheetUrl && item.sheetUrl.trim());
    const withoutSheetUrl = unappliedList.filter(item => !item.sheetUrl || !item.sheetUrl.trim());

    withSheetUrl.forEach((item, index) => {
      text += `${index + 1}. ${item.name}（シート作成済み）\n`;
    });

    if (withoutSheetUrl.length > 0) {
      const startNum = withSheetUrl.length + 1;
      withoutSheetUrl.forEach((item, index) => {
        text += `${startNum + index}. ${item.name}（シート未作成）\n`;
      });
    }
  }

  if (includeApplied && appliedList.length > 0) {
    if (text) text += '\n';
    text += '【提出済み】\n';
    appliedList.forEach((item, index) => {
      text += `${index + 1}. ${item}\n`;
    });
  }

  return text || '（該当者なし）';
}

/**
 * 現在の月キーを取得（YYYY-MM形式）
 * @returns {string} 現在の月キー
 */
function getCurrentMonthKey_() {
  const d = new Date();
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  return `${y}-${m}`;
}

/**
 * 前月の月キーを取得（YYYY-MM形式）
 * @returns {string} 前月の月キー
 */
function getPreviousMonthKey_() {
  const d = new Date();
  d.setMonth(d.getMonth() - 1);
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  return `${y}-${m}`;
}

/**
 * 月キーからその月の1日のDateオブジェクトを取得
 * @param {string} monthKey - 月キー（YYYY-MM形式）
 * @returns {Date|null} その月の1日のDateオブジェクト、変換できない場合はnull
 */
function getMonthStartDate_(monthKey) {
  const match = String(monthKey || '').trim().match(/^(\d{4})-(\d{2})/);
  if (!match) return null;

  const year = parseInt(match[1], 10);
  const month = parseInt(match[2], 10) - 1; // JavaScriptの月は0ベース
  if (month < 0 || month > 11) return null;

  return new Date(year, month, 1);
}

/**
 * 前月の未提出者を処理（毎月1日に実行）
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} master - マスタースプレッドシート
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh - Submissionsシート
 * @param {Array} values - Submissionsシートの全データ
 * @param {Array} header - ヘッダー行
 * @param {string} prevMonthKey - 前月の月キー
 * @param {number} idxMonthKey - monthKey列のインデックス
 * @param {number} idxStatus - status列のインデックス
 * @param {number} idxName - 氏名列のインデックス
 * @param {number} idxTeacherId - teacherId列のインデックス
 * @param {number} idxUrl - sheetUrl列のインデックス
 * @param {string} adminLineUserId - 管理者のLINE User ID
 * @param {Set<string>} onLeaveTeachers - 休職中の講師名のSet
 */
function handlePreviousMonthUnapplied_(master, sh, values, header, prevMonthKey, idxMonthKey, idxStatus, idxName, idxTeacherId, idxUrl, adminLineUserId, onLeaveTeachers) {
  try {
    const idxManagerReminder1st = header.indexOf('managerReminder1st');

    // 前月の未提出者と提出済み者を取得
    const unappliedList = [];
    const appliedList = [];

    for (let r = 1; r < values.length; r++) {
      const mk = normalizeMonthKey_(values[r][idxMonthKey]);
      if (mk !== prevMonthKey) continue;

      const status = String(values[r][idxStatus] || '').trim();
      const teacherName = String(values[r][idxName] || '').trim();

      if (!teacherName) continue;

      // 休職中の講師はスキップ
      if (onLeaveTeachers && onLeaveTeachers.has(teacherName)) continue;

      if (status === 'submitted') {
        appliedList.push(teacherName);
      } else if (status === 'created' || status === '' || status === 'teacher_not_found' || status === 'template_not_found') {
        // すでに通知済みかチェック
        let alreadyNotified = false;
        if (idxManagerReminder1st >= 0 && values[r][idxManagerReminder1st]) {
          const lastNotifyDate = values[r][idxManagerReminder1st];
          const today = new Date();
          if (lastNotifyDate && isSameJstDate_(lastNotifyDate, today)) {
            alreadyNotified = true;
          }
        }
        
        if (!alreadyNotified) {
          const sheetUrl = idxUrl >= 0 ? String(values[r][idxUrl] || '').trim() : '';
          unappliedList.push({
            row: r + 1,
            name: teacherName,
            teacherId: idxTeacherId >= 0 ? String(values[r][idxTeacherId] || '').trim() : '',
            sheetUrl: sheetUrl
          });
        }
      }
    }

    // 管理者に通知
    if (unappliedList.length > 0 || appliedList.length > 0) {
      let message = `【SHIFT SYNC - ${prevMonthKey}】\n`;
      message += `${prevMonthKey}のシフト提出期限です。\n\n`;

      if (unappliedList.length > 0) {
        message += `【未提出者】\n`;
        
        // sheetUrlがある場合とない場合で分けて表示
        const withSheetUrl = unappliedList.filter(item => item.sheetUrl && item.sheetUrl.trim());
        const withoutSheetUrl = unappliedList.filter(item => !item.sheetUrl || !item.sheetUrl.trim());
        
        if (withSheetUrl.length > 0) {
          withSheetUrl.forEach((item, index) => {
            message += `${index + 1}. ${item.name}（シート作成済み）\n`;
          });
        }
        
        if (withoutSheetUrl.length > 0) {
          const startNum = withSheetUrl.length > 0 ? withSheetUrl.length + 1 : 1;
          withoutSheetUrl.forEach((item, index) => {
            message += `${startNum + index}. ${item.name}（シート未作成・フォーム送信待ち）\n`;
          });
        }
        
        message += `\n`;
      }

      if (appliedList.length > 0) {
        message += `【提出済み】\n`;
        appliedList.forEach((item, index) => {
          message += `${index + 1}. ${item}\n`;
        });
      }

      pushLine_(adminLineUserId, message.trim());

      // 未提出者にLINE通知を送信
      if (unappliedList.length > 0) {
        sendUnappliedReminderToTeachers_(master, unappliedList, prevMonthKey);

        // 通知済みフラグを更新
        if (idxManagerReminder1st >= 0) {
          unappliedList.forEach(item => {
            sh.getRange(item.row, idxManagerReminder1st + 1).setValue(new Date());
          });
        }
      }
    }
  } catch (err) {
    handleError_(err, 'handlePreviousMonthUnapplied_');
  }
}

/**
 * 未提出者にLINEリマインド通知を送信
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} master - マスタースプレッドシート
 * @param {Array} unappliedList - 未提出者リスト
 * @param {string} monthKey - 月キー
 */
function sendUnappliedReminderToTeachers_(master, unappliedList, monthKey) {
  try {
    for (const item of unappliedList) {
      const lineUserId = getTeacherLineUserId_(master, item.teacherId, item.name);
      if (lineUserId) {
        const lastName = extractLastName_(item.name);
        let message = `【シフト未提出リマインド】\n${lastName}先生（${monthKey}）の提出がまだのようです。`;
        
        if (item.sheetUrl && item.sheetUrl.trim()) {
          // sheetUrlがある場合：シートが作成済みなので、URLを送信
          message += `\nこちらから入力・提出（☑）をお願いします。\n${item.sheetUrl}`;
        } else {
          // sheetUrlがない場合：シートがまだ作成されていない
          message += `\nシフト申請用紙の準備がまだのようです。管理者に連絡するか、フォーム送信をお待ちください。`;
        }
        
        pushLine_(lineUserId, message);
      }
    }
  } catch (err) {
    handleError_(err, 'sendUnappliedReminderToTeachers_');
  }
}

/**
 * 全講師に来月のシフト申請依頼をLINE通知
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} master - マスタースプレッドシート
 * @param {string} nextMonthKey - 来月の月キー
 * @returns {Object} 通知結果（notifiedCount, notifiedTeachers）
 */
function notifyAllTeachersAboutNextMonthShift_(master, nextMonthKey) {
  const result = {
    notifiedCount: 0,
    notifiedTeachers: []
  };
  
  try {
    const sh = master.getSheetByName(CONFIG.SHEET_TEACHERS);
    if (!sh) {
      console.error('Teachersシートが見つかりません');
      return result;
    }

    const values = sh.getDataRange().getValues();
    if (values.length < 2) return result;

    const header = values[0];
    const idxName = header.indexOf('氏名');
    const idxLine = header.indexOf('lineUserId');
    const idxOnLeave = header.indexOf('休職中');

    if (idxName < 0 || idxLine < 0) {
      console.error('Teachersに「氏名」または「lineUserId」列がありません');
      return result;
    }

    // LINE User IDが登録されている全講師に通知（休職中を除く）
    for (let r = 1; r < values.length; r++) {
      const lineUserId = String(values[r][idxLine] || '').trim();
      const teacherName = String(values[r][idxName] || '').trim();

      // 休職中フラグをチェック（TRUE, true, 1, ○ などを休職中とみなす）
      const onLeave = idxOnLeave >= 0 && isChecked_(values[r][idxOnLeave]);
      if (onLeave) continue;

      if (lineUserId && teacherName) {
        const lastName = extractLastName_(teacherName);
        const message = `【シフト申請のお願い】\n${lastName}先生、${nextMonthKey}のシフト申請をお願いします。`;
        pushLine_(lineUserId, message);
        
        // 通知結果を記録
        result.notifiedCount++;
        result.notifiedTeachers.push(teacherName);
      }
    }
    
    return result;
  } catch (err) {
    handleError_(err, 'notifyAllTeachersAboutNextMonthShift_');
    return result;
  }
}

/**
 * 全講師のSubmissionsエントリを作成（テンプレート確認時に実行）
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} master - マスタースプレッドシート
 * @param {string} monthKey - 月キー
 */
function createSubmissionsForAllTeachers_(master, monthKey) {
  try {
    const teachersSh = master.getSheetByName(CONFIG.SHEET_TEACHERS);
    if (!teachersSh) {
      console.error('Teachersシートが見つかりません');
      return;
    }

    const teachersValues = teachersSh.getDataRange().getValues();
    if (teachersValues.length < 2) return;

    const teachersHeader = teachersValues[0];
    const idxName = teachersHeader.indexOf('氏名');
    const idxTeacherId = teachersHeader.indexOf('teacherId');
    const idxEmail = teachersHeader.indexOf('メール');
    const idxLine = teachersHeader.indexOf('lineUserId');
    const idxOnLeave = teachersHeader.indexOf('休職中');

    if (idxName < 0) {
      console.error('Teachersに「氏名」列がありません');
      return;
    }

    let createdCount = 0;
    let skippedCount = 0;

    // 全講師をループ（休職中を除く）
    for (let r = 1; r < teachersValues.length; r++) {
      const teacherName = String(teachersValues[r][idxName] || '').trim();
      if (!teacherName) continue;

      // 休職中フラグをチェック
      const onLeave = idxOnLeave >= 0 && isChecked_(teachersValues[r][idxOnLeave]);
      if (onLeave) continue;

      const teacherId = idxTeacherId >= 0 ? String(teachersValues[r][idxTeacherId] || '').trim() : '';
      const email = idxEmail >= 0 ? String(teachersValues[r][idxEmail] || '').trim() : '';
      const lineUserId = idxLine >= 0 ? String(teachersValues[r][idxLine] || '').trim() : '';

      // submissionKeyを生成
      const submissionKey = `${monthKey}-${teacherId || normalizeNameKey_(teacherName)}`;

      // 既存エントリをチェック（submissionKeyと、monthKey+teacherId/氏名の両方で検索）
      let existing = findSubmissionByKey_(master, submissionKey);
      if (!existing) {
        // submissionKeyで見つからない場合、monthKeyとteacherId/氏名の組み合わせで検索
        existing = findSubmissionByMonthAndTeacher_(master, monthKey, teacherId, teacherName);
      }
      if (existing) {
        // teacherIdが確定しているのにsubmissionKeyが名前ベース（不正）の場合は更新
        if (teacherId && existing.submissionKey !== submissionKey) {
          updateSubmission_(master, existing.row, existing.header, {
            teacherId: teacherId,
            submissionKey: submissionKey,
          });
          console.log(`Submission修正: ${teacherName} (${existing.submissionKey} → ${submissionKey})`);
        }
        skippedCount++;
        continue;
      }

      // 新規エントリを作成
      appendSubmission_(master, {
        timestamp: new Date(),
        monthKey: monthKey,
        teacherId: teacherId,
        name: teacherName,
        sheetUrl: '', // シートはまだ作成されていない
        status: 'created', // 初期状態
        lastNotified: '',
        submissionKey: submissionKey,
        submittedAt: '',
      });

      createdCount++;
    }

    console.log(`Submissionsエントリ作成完了: ${createdCount}件作成、${skippedCount}件スキップ（既存）`);
  } catch (err) {
    handleError_(err, 'createSubmissionsForAllTeachers_');
  }
}

/**
 * 3週間前のテンプレート確認処理（来月のみ、午後3時）
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} master - マスタースプレッドシート
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh - Submissionsシート
 * @param {Array} values - Submissionsシートの全データ
 * @param {Array} header - ヘッダー行
 * @param {string} nextMonthKey - 来月の月キー
 * @param {number} idxMonthKey - monthKey列のインデックス
 * @param {string} adminLineUserId - 管理者のLINE User ID
 */
function checkTemplateThreeWeeksBefore_(master, sh, values, header, nextMonthKey, idxMonthKey, adminLineUserId) {
  try {
    const today = new Date();
    const currentHour = today.getHours();
    
    // 午後3時（15時）でない場合はスキップ
    if (currentHour !== 15) {
      return;
    }
    
    const props = PropertiesService.getScriptProperties();

    const monthStart = getMonthStartDate_(nextMonthKey);
    if (!monthStart) return;

    // 3週間前（21日前）の日付を計算
    const threeWeeksBefore = new Date(monthStart);
    threeWeeksBefore.setDate(threeWeeksBefore.getDate() - 21);

    // 3週間前の日付以降かチェック（3週間前以降、毎日チェック）
    // 日付のみで比較（時刻は無視）
    const todayDateOnly = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    const threeWeeksBeforeDateOnly = new Date(threeWeeksBefore.getFullYear(), threeWeeksBefore.getMonth(), threeWeeksBefore.getDate());
    
    if (todayDateOnly < threeWeeksBeforeDateOnly) {
      return; // まだ3週間前になっていない
    }

    // 通知済みフラグをチェック（ScriptPropertiesで管理、同じ日に複数回通知しないように）
    const lastNotifyKey = `TEMPLATE_CHECK_${nextMonthKey}`;
    const lastNotifyDateStr = props.getProperty(lastNotifyKey);
    
    if (lastNotifyDateStr) {
      const lastNotifyDate = new Date(lastNotifyDateStr);
      if (isSameJstDate_(lastNotifyDate, today)) {
        return; // 既に今日通知済み
      }
    }

    // テンプレートの存在をチェック
    const templateFolderId = CONFIG.TEMPLATE_FOLDER_ID;
    if (!templateFolderId) {
      console.log('TEMPLATE_FOLDER_IDが未設定のため、テンプレート確認をスキップします');
      return;
    }

    const templateSpreadsheetId = findTemplateByMonth_(nextMonthKey, templateFolderId);

    if (!templateSpreadsheetId) {
      // テンプレートが存在しない場合：管理者にのみ通知（講師には送らない）
      const message = `【シフト申請用紙作成依頼】\n${nextMonthKey}のシフト申請用紙のテンプレートが準備されていません。\nシフト申請用紙を作成してください。`;
      pushLine_(adminLineUserId, message);
      
      // 通知済みフラグを更新（ScriptProperties）
      props.setProperty(lastNotifyKey, today.toISOString());
    } else {
      // テンプレートが存在する場合のみ講師に通知（初回のみ）
      const templateReadyKey = `TEMPLATE_READY_${nextMonthKey}`;
      const alreadyNotifiedReady = props.getProperty(templateReadyKey);
      
      if (!alreadyNotifiedReady) {
        // 1. 全講師にLINE通知を送信（LINE User IDが登録されている講師のみ）
        const notificationResult = notifyAllTeachersAboutNextMonthShift_(master, nextMonthKey);
        
        // 2. 全講師のSubmissionsエントリを作成（まだ存在しない場合のみ）
        createSubmissionsForAllTeachers_(master, nextMonthKey);
        
        // 3. 管理者に結果を報告（誰に通知を送ったか）
        let adminMessage = `【シフト申請依頼の送信結果】\n${nextMonthKey}のシフト申請用紙のテンプレートは準備されています。\n\n`;
        
        if (notificationResult.notifiedCount > 0) {
          adminMessage += `以下の講師（${notificationResult.notifiedCount}名）にシフト申請依頼を送信しました：\n`;
          notificationResult.notifiedTeachers.forEach((name, index) => {
            adminMessage += `${index + 1}. ${name}\n`;
          });
        } else {
          adminMessage += `LINE User IDが登録されている講師がいませんでした。`;
        }
        
        adminMessage += `\n全コースの申請フォームは既存のシステム（フォーム送信）でコピー・作成してください。`;
        pushLine_(adminLineUserId, adminMessage);
        
        // テンプレート準備済みフラグを更新
        props.setProperty(templateReadyKey, 'true');
      }
      
      // 通知済みフラグを更新（ScriptProperties）
      props.setProperty(lastNotifyKey, today.toISOString());
    }
  } catch (err) {
    handleError_(err, 'checkTemplateThreeWeeksBefore_');
  }
}

/**
 * 管理者向けリマインドのトリガーを設定（手動実行用）
 * この関数を実行すると、remindUnappliedToManagerの時間ベーストリガーが設定されます
 * テンプレート確認のため、午後3時に実行することを推奨
 */
function setupRemindUnappliedToManagerTrigger() {
  try {
    // 既存のトリガーを削除
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'remindUnappliedToManager') {
        ScriptApp.deleteTrigger(trigger);
      }
    });

    // 新しいトリガーを作成（毎日午後3時に実行）
    ScriptApp.newTrigger('remindUnappliedToManager')
      .timeBased()
      .everyDays(1)
      .atHour(15)
      .create();

    console.log('remindUnappliedToManagerのトリガーを設定しました（毎日15時）');
    return 'トリガーを設定しました：remindUnappliedToManager（毎日15時）';
  } catch (err) {
    handleError_(err, 'setupRemindUnappliedToManagerTrigger');
    return 'エラーが発生しました：' + (err.message || String(err));
  }
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

/**
 * 【手動実行用】来月のシフト申請依頼を全講師に送信
 * GASエディタから直接実行可能
 * デプロイ忘れ等で初回通知が送れなかった場合に使用
 */
function sendInitialShiftRequestManually() {
  try {
    const master = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
    const nextMonthKey = getNextMonthKey_();

    // テンプレートの存在確認
    const templateFolderId = CONFIG.TEMPLATE_FOLDER_ID;
    if (!templateFolderId) {
      console.log('TEMPLATE_FOLDER_IDが未設定です');
      return 'エラー: TEMPLATE_FOLDER_IDが未設定です';
    }

    const templateSpreadsheetId = findTemplateByMonth_(nextMonthKey, templateFolderId);
    if (!templateSpreadsheetId) {
      console.log(`${nextMonthKey}のテンプレートが見つかりません`);
      return `エラー: ${nextMonthKey}のテンプレートが見つかりません`;
    }

    // 全講師にLINE通知を送信
    const result = notifyAllTeachersAboutNextMonthShift_(master, nextMonthKey);

    // 全講師のSubmissionsエントリを作成
    createSubmissionsForAllTeachers_(master, nextMonthKey);

    // 管理者に結果を報告
    const adminLineUserId = PropertiesService.getScriptProperties().getProperty('ADMIN_LINE_USER_ID');
    if (adminLineUserId && result.notifiedCount > 0) {
      let adminMessage = `【シフト申請依頼の送信結果（手動実行）】\n${nextMonthKey}のシフト申請依頼を送信しました。\n\n`;
      adminMessage += `通知した講師（${result.notifiedCount}名）：\n`;
      result.notifiedTeachers.forEach((name, index) => {
        adminMessage += `${index + 1}. ${name}\n`;
      });
      pushLine_(adminLineUserId, adminMessage);
    }

    console.log(`初回通知送信完了: ${result.notifiedCount}名`);
    console.log(`通知した講師: ${result.notifiedTeachers.join(', ')}`);

    return `送信完了: ${result.notifiedCount}名に通知しました`;
  } catch (err) {
    handleError_(err, 'sendInitialShiftRequestManually');
    return 'エラーが発生しました: ' + (err.message || String(err));
  }
}

/**
 * 権限承認テスト用関数
 * UrlFetchAppの権限を承認するためのダミー関数
 * GASエディタから手動で実行してください
 */
function testUrlFetchPermission() {
  // UrlFetchAppの権限を確認するための最小限のリクエスト
  const response = UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true });
  console.log('UrlFetchApp permission test: Status ' + response.getResponseCode());
  SpreadsheetApp.getActiveSpreadsheet().toast('UrlFetchApp権限が正常に機能しています', '権限テスト', 3);
}

/**
 * 【手動実行用】SubmissionsのteacherIdとsubmissionKeyを一括修復
 * teacherIdが空またはsubmissionKeyが名前ベースの行をTeachersシートの情報で修正する
 */
function repairSubmissionKeys() {
  const master = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);

  // Teachersシートから講師名→teacherIdのマップを作成
  const teachersSh = master.getSheetByName(CONFIG.SHEET_TEACHERS);
  if (!teachersSh) throw new Error('Teachersシートが見つかりません');

  const teachersValues = teachersSh.getDataRange().getValues();
  const teachersHeader = teachersValues[0];
  const tIdxName = teachersHeader.indexOf('氏名');
  const tIdxTeacherId = teachersHeader.indexOf('teacherId');

  if (tIdxName < 0 || tIdxTeacherId < 0) throw new Error('Teachersに「氏名」または「teacherId」列がありません');

  const teacherMap = {};
  for (let r = 1; r < teachersValues.length; r++) {
    const name = String(teachersValues[r][tIdxName] || '').trim();
    const id = String(teachersValues[r][tIdxTeacherId] || '').trim();
    if (name && id) {
      teacherMap[normalizeNameKey_(name)] = { name, teacherId: id };
    }
  }

  // Submissionsシートを処理
  const subSh = master.getSheetByName(CONFIG.SHEET_SUBMISSIONS);
  if (!subSh) throw new Error('Submissionsシートが見つかりません');

  const values = subSh.getDataRange().getValues();
  const header = values[0];
  const idxName = header.indexOf('氏名');
  const idxTeacherId = header.indexOf('teacherId');
  const idxMonthKey = header.indexOf('monthKey');
  const idxSubmissionKey = header.indexOf('submissionKey');

  if (idxName < 0 || idxTeacherId < 0 || idxMonthKey < 0 || idxSubmissionKey < 0) {
    throw new Error('Submissionsに必要列がありません（氏名/teacherId/monthKey/submissionKey）');
  }

  let updatedCount = 0;

  for (let r = 1; r < values.length; r++) {
    const name = String(values[r][idxName] || '').trim();
    const currentTeacherId = String(values[r][idxTeacherId] || '').trim();
    const monthKey = normalizeMonthKey_(values[r][idxMonthKey]);
    const currentSubmissionKey = String(values[r][idxSubmissionKey] || '').trim();

    if (!name || !monthKey) continue;

    const teacher = teacherMap[normalizeNameKey_(name)];
    if (!teacher) continue;

    const correctSubmissionKey = `${monthKey}-${teacher.teacherId}`;

    if (!currentTeacherId || currentSubmissionKey !== correctSubmissionKey) {
      updateSubmission_(master, r + 1, header, {
        teacherId: teacher.teacherId,
        submissionKey: correctSubmissionKey,
      });
      console.log(`修復: 行${r + 1} ${name} (${currentSubmissionKey} → ${correctSubmissionKey})`);
      updatedCount++;
    }
  }

  console.log(`修復完了: ${updatedCount}件更新`);
  return `修復完了: ${updatedCount}件更新`;
}
