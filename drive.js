/************************************************************
 * drive.gs
 ************************************************************/

/** 月フォルダ（例：2026-01）を親フォルダ内に作成・取得 */
function ensureMonthFolder_(monthKey, parentFolderId) {
  const parent = DriveApp.getFolderById(parentFolderId);
  const it = parent.getFoldersByName(monthKey);
  if (it.hasNext()) return it.next().getId();
  return parent.createFolder(monthKey).getId();
}

/** テンプレのスプレッドシートをフォルダにコピーして、新しいSpreadsheetIdを返す */
function copyTemplateSpreadsheet_(folderId, templateSpreadsheetId, fileName) {
  const folder = DriveApp.getFolderById(folderId);
  const templateFile = DriveApp.getFileById(templateSpreadsheetId);
  const copied = templateFile.makeCopy(fileName, folder);
  return copied.getId();
}

/** 編集権限付与（Teachersのメール宛て） */
function grantEditPermission_(spreadsheetId, email) {
  try {
    const file = DriveApp.getFileById(spreadsheetId);
    const editors = file.getEditors().map(u => u.getEmail());
    if (editors.includes(email)) return;
    file.addEditor(email);
  } catch (err) {
    logError_(err, 'grantEditPermission_', { spreadsheetId, email });
    // 編集権限付与の失敗は致命的ではないため、管理者通知は送らない
  }
}

/**
 * 講師シートをロック（提出後、編集不可にする）
 * - すべてのシートを保護し、講師の編集権限を削除
 * - 管理者（スクリプト実行者）は常に編集可能
 */
function lockTeacherSheet_(spreadsheetId, teacherEmail) {
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheets = ss.getSheets();
    
    // 既存の保護を削除（重複防止）
    const existingProtections = [];
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
      
      // 講師の編集権限を削除（閲覧のみ）
      // 保護のデフォルトで所有者とスクリプト実行者は編集可能
      if (teacherEmail) {
        try {
          protection.removeEditor(teacherEmail);
        } catch (e) {
          // エディターリストに含まれていない場合は無視
        }
      }
    });
    
    // ファイルレベルで講師の編集権限を削除（閲覧のみに変更）
    if (teacherEmail) {
      const file = DriveApp.getFileById(spreadsheetId);
      file.removeEditor(teacherEmail);
      file.addViewer(teacherEmail);
    }
    
    return true;
  } catch (err) {
    logError_(err, 'lockTeacherSheet_', { spreadsheetId, teacherEmail });
    return false;
  }
}

/**
 * 講師シートのロックを解除（管理者からの変更依頼時）
 * - すべての保護を解除
 * - 講師に編集権限を再付与
 * - テンプレート側でロックされている場合も確実に解除
 */
function unlockTeacherSheet_(spreadsheetId, teacherEmail) {
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheets = ss.getSheets();
    const scriptOwner = Session.getEffectiveUser().getEmail();
    
    // すべての保護を解除（説明が「提出後ロック」のもの、または空の保護も）
    let protectionRemoved = false;
    sheets.forEach(sheet => {
      const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      protections.forEach(prot => {
        const desc = prot.getDescription();
        // 説明が「提出後ロック」または空の保護も解除（テンプレート側で作成された可能性）
        if (desc === '提出後ロック' || desc === '' || !desc) {
          try {
            prot.remove();
            protectionRemoved = true;
          } catch (e) {
            // 保護の削除に失敗しても続行
            console.error('Failed to remove protection:', e);
          }
        }
      });
    });
    
    // ファイルレベルで講師に編集権限を再付与
    const file = DriveApp.getFileById(spreadsheetId);
    
    // スクリプト実行者（管理者）がエディターであることを確認
    try {
      const editors = file.getEditors().map(u => u.getEmail());
      if (!editors.includes(scriptOwner)) {
        file.addEditor(scriptOwner);
      }
    } catch (e) {
      console.error('Failed to ensure script owner is editor:', e);
    }
    
    // 講師に編集権限を再付与
    if (teacherEmail) {
      try {
        // まずビューアーから削除
        const viewers = file.getViewers().map(u => u.getEmail());
        if (viewers.includes(teacherEmail)) {
          file.removeViewer(teacherEmail);
        }
      } catch (e) {
        // ビューアーでない場合は無視
      }
      
      try {
        // エディターとして追加
        const editors = file.getEditors().map(u => u.getEmail());
        if (!editors.includes(teacherEmail)) {
          file.addEditor(teacherEmail);
        }
      } catch (e) {
        // エディター追加に失敗した場合はログ
        console.error('Failed to add editor:', e);
      }
    }
    
    return true;
  } catch (err) {
    // より詳細なエラーログを出力
    const errorDetails = {
      spreadsheetId: spreadsheetId,
      teacherEmail: teacherEmail,
      errorMessage: err.message || String(err),
      errorStack: err.stack || ''
    };
    logError_(err, 'unlockTeacherSheet_', errorDetails);
    console.error('unlockTeacherSheet_ failed:', errorDetails);
    return false;
  }
}
