/************************************************************
 * テンプレート側（講師用シート）のコード
 * 
 * このファイルは、テンプレートファイルまたはコピーされた講師用シートに
 * 配置する必要があるコードです。
 * 
 * 【使用方法】
 * 1. テンプレートファイルのApps Scriptエディタを開く
 * 2. このファイルの内容をすべてコピーして、テンプレートファイルのスクリプトに貼り付ける
 * 3. 既存のコードをすべて削除して、このコードに置き換える
 * 
 * 【注意事項】
 * - このコードは、コピーされた各講師用シートに含まれる必要があります
 * - マスター側のコード（Code.js）と同期させる必要があります
 * - コードを更新した場合は、テンプレートファイルにも反映してください
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

    // ファイルレベルの権限変更は通知メールが送信されるため、ここでは行わない
    // シートの保護だけで編集を制限する
    // マスター側（drive.js）の lockTeacherSheet_ で通知なしの権限変更を行う

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

    // ファイルレベルの権限変更は通知メールが送信されるため、ここでは行わない
    // マスター側（drive.js）の unlockTeacherSheet_ で通知なしの権限変更を行う

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

    // シートレベルの保護をチェック
    const sheets = ss.getSheets();
    let hasProtection = false;

    sheets.forEach(sheet => {
      const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      protections.forEach(prot => {
        const desc = prot.getDescription();
        if (desc === '提出後ロック') {
          hasProtection = true;
        }
      });
    });

    // 保護がある場合、unlockSheetForRevision_を呼んでシートレベルの保護を解除
    // （マスター側で既にファイルレベルの権限が変更されている場合）
    if (hasProtection) {
      // ファイルの編集権限があるかチェック
      try {
        const file = DriveApp.getFileById(ss.getId());
        const currentUser = Session.getActiveUser().getEmail();
        const editors = file.getEditors().map(u => u.getEmail());

        if (currentUser && editors.includes(currentUser)) {
          // 編集権限があるのに保護がある = マスター側でロック解除済み
          unlockSheetForRevision_();
        }
      } catch (e) {
        // ファイル権限取得に失敗した場合は何もしない
      }
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



