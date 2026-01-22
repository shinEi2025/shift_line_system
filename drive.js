/************************************************************
 * drive.gs
 *
 * 定数は utils.js で定義:
 *   CONFIG, SHEET_CONFIG, DRIVE_CONFIG
 ************************************************************/

/**
 * ファイルのエディターリストを取得
 * @param {GoogleAppsScript.Drive.File} file - Driveファイルオブジェクト
 * @returns {Array<string>} エディターのメールアドレス配列
 */
function getFileEditors_(file) {
  try {
    return file.getEditors().map(u => u.getEmail());
  } catch (e) {
    console.error('Failed to get file editors:', e);
    return [];
  }
}

/**
 * 指定されたメールアドレスがエディター権限を持っているか確認
 * @param {GoogleAppsScript.Drive.File} file - Driveファイルオブジェクト
 * @param {string} email - 確認するメールアドレス
 * @returns {boolean} エディター権限がある場合true
 */
function hasEditorAccess_(file, email) {
  const editors = getFileEditors_(file);
  return editors.includes(email);
}

/**
 * ファイルへのエディター権限を確保（既にエディターの場合はスキップ）
 * @param {GoogleAppsScript.Drive.File} file - Driveファイルオブジェクト
 * @param {string} email - 追加するメールアドレス
 * @returns {boolean} エディター権限が確保された場合true
 */
function ensureEditor_(file, email) {
  try {
    // 既にエディターの場合は何もしない（メール通知を防ぐ）
    if (hasEditorAccess_(file, email)) {
      return true;
    }

    // メール通知なしでエディターとして追加（Drive Advanced Service使用）
    const fileId = file.getId();
    const result = addPermissionWithoutNotification_(fileId, email, 'writer');

    if (result) {
      Utilities.sleep(DRIVE_CONFIG.SLEEP_AFTER_ADD_EDITOR);
      return true;
    }

    // Drive Advanced Serviceが失敗した場合、エラーをログに記録
    console.error('ensureEditor_: Drive Advanced Serviceでの権限追加に失敗しました。通知なしでの追加ができません。');
    return false;
  } catch (err) {
    // エディターリストを再度確認して、追加されていれば成功とみなす
    const editors = getFileEditors_(file);
    return editors.includes(email);
  }
}

/**
 * ファイルへのビューアー権限を確保（メール通知なし）
 * @param {GoogleAppsScript.Drive.File} file - Driveファイルオブジェクト
 * @param {string} email - 追加するメールアドレス
 * @returns {boolean} ビューアー権限が確保された場合true
 */
function ensureViewer_(file, email) {
  try {
    const fileId = file.getId();
    const result = addPermissionWithoutNotification_(fileId, email, 'reader');

    if (result) {
      return true;
    }

    // Drive Advanced Serviceが失敗した場合、エラーをログに記録
    console.error('ensureViewer_: Drive Advanced Serviceでの権限追加に失敗しました。通知なしでの追加ができません。');
    return false;
  } catch (err) {
    console.error('ensureViewer_ error:', err);
    return false;
  }
}

/**
 * 「リンクを知っているすべての人」を編集者に設定
 * @param {string} fileId - ファイルID
 * @returns {boolean} 設定に成功した場合true
 */
function setAnyoneWithLinkCanEdit_(fileId) {
  try {
    Drive.Permissions.create(
      {
        type: 'anyone',
        role: 'writer'
      },
      fileId
    );
    console.log(`setAnyoneWithLinkCanEdit_: Successfully set anyone with link as editor for file: ${fileId}`);
    return true;
  } catch (err) {
    console.error(`setAnyoneWithLinkCanEdit_ error: ${err.message || err}`);
    return false;
  }
}

/**
 * 「リンクを知っているすべての人」を閲覧者に設定（編集権限を削除）
 * @param {string} fileId - ファイルID
 * @returns {boolean} 設定に成功した場合true
 */
function setAnyoneWithLinkCanView_(fileId) {
  try {
    // まず既存の「anyone」権限を取得
    const permissions = Drive.Permissions.list(fileId).permissions || [];
    const anyonePermission = permissions.find(p => p.type === 'anyone');

    if (anyonePermission) {
      // 既存の権限を更新（writerからreaderに変更）
      Drive.Permissions.update(
        { role: 'reader' },
        fileId,
        anyonePermission.id
      );
      console.log(`setAnyoneWithLinkCanView_: Successfully updated anyone permission to reader for file: ${fileId}`);
    } else {
      // 権限が存在しない場合は新規作成
      Drive.Permissions.create(
        {
          type: 'anyone',
          role: 'reader'
        },
        fileId
      );
      console.log(`setAnyoneWithLinkCanView_: Successfully created anyone reader permission for file: ${fileId}`);
    }
    return true;
  } catch (err) {
    console.error(`setAnyoneWithLinkCanView_ error: ${err.message || err}`);
    return false;
  }
}

/**
 * Drive Advanced Serviceを使用して、メール通知なしで権限を追加
 * @param {string} fileId - ファイルID
 * @param {string} email - 追加するメールアドレス
 * @param {string} role - 'writer' または 'reader'
 * @returns {boolean} 追加に成功した場合true
 */
function addPermissionWithoutNotification_(fileId, email, role) {
  try {
    // Drive Advanced Service (v3) を使用
    Drive.Permissions.create(
      {
        type: 'user',
        role: role,
        emailAddress: email
      },
      fileId,
      {
        sendNotificationEmail: false
      }
    );
    return true;
  } catch (err) {
    // 権限が既に存在する場合などのエラーは許容
    console.log(`addPermissionWithoutNotification_ (${role}): ${err.message || err}`);
    return false;
  }
}

/**
 * 実行ユーザー（Effective User）のメールアドレスを取得
 * @param {string} spreadsheetId - スプレッドシートID（フォールバック用）
 * @returns {string} 実行ユーザーのメールアドレス
 */
function getEffectiveUserEmail_(spreadsheetId) {
  try {
    return Session.getEffectiveUser().getEmail();
  } catch (e) {
    // 権限がない場合、ファイル所有者を使用
    try {
      const file = DriveApp.getFileById(spreadsheetId);
      return file.getOwner().getEmail();
    } catch (e2) {
      // それでも取得できない場合、アクティブユーザーを使用
      return Session.getActiveUser().getEmail();
    }
  }
}

/** 月フォルダ（例：2026-01）を親フォルダ内に作成・取得 */
function ensureMonthFolder_(monthKey, parentFolderId) {
  const parent = DriveApp.getFolderById(parentFolderId);
  const it = parent.getFoldersByName(monthKey);
  if (it.hasNext()) return it.next().getId();
  return parent.createFolder(monthKey).getId();
}

/**
 * 月ごとのテンプレートを検索
 * @param {string} monthKey - 月キー（YYYY-MM形式、例：2026-01）
 * @param {string} templateFolderId - テンプレートフォルダID
 * @returns {string|null} テンプレートスプレッドシートID（見つからない場合はnull）
 */
function findTemplateByMonth_(monthKey, templateFolderId) {
  try {
    const templateFolder = DriveApp.getFolderById(templateFolderId);
    const files = templateFolder.getFiles();
    
    // 月キーに基づいてテンプレートファイル名を検索
    // 例：2026-01 → "2026-01_シフト提出テンプレ"
    const templateNamePattern = `${monthKey}_シフト提出テンプレ`;
    
    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      
      // ファイル名が月キーで始まり、テンプレート名を含むかチェック
      if (fileName.startsWith(monthKey) && fileName.includes('シフト提出テンプレ')) {
        // スプレッドシートかどうかを確認
        if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
          return file.getId();
        }
      }
    }
    
    return null;
  } catch (err) {
    logError_(err, 'findTemplateByMonth_', { monthKey, templateFolderId });
    return null;
  }
}

/** テンプレのスプレッドシートをフォルダにコピーして、新しいSpreadsheetIdを返す */
function copyTemplateSpreadsheet_(folderId, templateSpreadsheetId, fileName) {
  const folder = DriveApp.getFolderById(folderId);
  const templateFile = DriveApp.getFileById(templateSpreadsheetId);
  const copied = templateFile.makeCopy(fileName, folder);
  const newSpreadsheetId = copied.getId();
  
  // コピー後にマスター側のコードを展開（Apps Script APIを使用）
  // 注意: この機能を使用するには、以下の前提条件が必要です：
  // 1. OAuth2認証が完了している（getAppsScriptAPIAccessToken()が動作すること）
  // 2. スクリプトプロジェクトIDが取得できること（getScriptIdForSpreadsheet_()が動作すること）
  // 
  // コピーされた新しいシートのスクリプトプロジェクトIDを自動的に取得して保存する
  try {
    // まず、スクリプトプロジェクトIDを取得（キャッシュまたはAPI経由）
    let scriptId = getScriptIdForSpreadsheet_(newSpreadsheetId);
    
    // 取得できない場合、Apps Script APIを使用して取得を試みる
    if (!scriptId) {
      scriptId = findScriptIdBySpreadsheetId_(newSpreadsheetId);
      if (scriptId) {
        // 見つかった場合は保存
        saveScriptIdForSpreadsheet_(newSpreadsheetId, scriptId);
      }
    }
    
    // スクリプトプロジェクトIDが取得できた場合、コードを展開
    if (scriptId) {
      const deployResult = deployTemplateSideCodeToSpreadsheet_(newSpreadsheetId, scriptId);
      if (deployResult) {
        console.log(`[copyTemplateSpreadsheet_] Successfully deployed template code to: ${newSpreadsheetId}`);
      } else {
        console.log(`[copyTemplateSpreadsheet_] Template code deployment failed (check logs for details)`);
      }
    } else {
      console.log(`[copyTemplateSpreadsheet_] Script ID not found for new spreadsheet: ${newSpreadsheetId}`);
      console.log(`[copyTemplateSpreadsheet_] Using template file code instead (makeCopy includes script code)`);
      console.log(`[copyTemplateSpreadsheet_] Note: If template has no code, you need to manually deploy code or save script ID`);
    }
  } catch (err) {
    console.error('Failed to deploy template-side code:', err);
    logError_(err, 'deployTemplateSideCodeToSpreadsheet_', { spreadsheetId: newSpreadsheetId });
    // エラーが発生しても処理を続行（テンプレートのコードが使用される）
  }
  
  return newSpreadsheetId;
}

/**
 * Apps Script APIを使用して、スプレッドシートIDに紐づくスクリプトプロジェクトIDを検索
 * @param {string} spreadsheetId - スプレッドシートID
 * @return {string|null} スクリプトプロジェクトID（見つからない場合はnull）
 */
function findScriptIdBySpreadsheetId_(spreadsheetId) {
  try {
    const accessToken = getAppsScriptAPIAccessToken();
    if (!accessToken) {
      return null;
    }
    
    // Apps Script APIを使用して、すべてのプロジェクトを取得
    const url = 'https://script.googleapis.com/v1/projects';
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        'Authorization': 'Bearer ' + accessToken
      }
    });
    
    const result = JSON.parse(response.getContentText());
    
    if (result.error) {
      console.error(`[findScriptIdBySpreadsheetId_] API error: ${JSON.stringify(result.error)}`);
      return null;
    }
    
    const projects = result.projects || [];
    
    // 各プロジェクトのメタデータを確認して、スプレッドシートIDと一致するものを探す
    for (const project of projects) {
      try {
        // プロジェクトのメタデータを取得
        const projectUrl = `https://script.googleapis.com/v1/projects/${project.scriptId}`;
        const projectResponse = UrlFetchApp.fetch(projectUrl, {
          method: 'get',
          headers: {
            'Authorization': 'Bearer ' + accessToken
          }
        });
        
        const projectData = JSON.parse(projectResponse.getContentText());
        
        // プロジェクトのparentIdがスプレッドシートIDと一致するか確認
        // 注意: この方法は確実ではない可能性があります（APIの仕様により）
        if (projectData.parentId === spreadsheetId) {
          return project.scriptId;
        }
      } catch (e) {
        // 個別のプロジェクトの取得に失敗しても続行
        console.error(`[findScriptIdBySpreadsheetId_] Failed to get project details for ${project.scriptId}:`, e);
      }
    }
    
    return null;
  } catch (err) {
    console.error('[findScriptIdBySpreadsheetId_] Error:', err);
    return null;
  }
}

/**
 * 講師シートをロック（提出後、編集不可にする）
 * - 「リンクを知っているすべての人」を閲覧のみに変更
 * - すべてのシートを保護
 * - 管理者（スクリプト実行者）は常に編集可能
 */
function lockTeacherSheet_(spreadsheetId, teacherEmail) {
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheets = ss.getSheets();

    // 「リンクを知っているすべての人」を閲覧のみに変更
    setAnyoneWithLinkCanView_(spreadsheetId);

    // 既存の保護を削除（重複防止）
    sheets.forEach(sheet => {
      const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      protections.forEach(prot => {
        if (prot.getDescription() === SHEET_CONFIG.PROTECTION_DESCRIPTION) {
          prot.remove();
        }
      });
    });

    // 各シートを保護
    sheets.forEach(sheet => {
      const protection = sheet.protect().setDescription(SHEET_CONFIG.PROTECTION_DESCRIPTION);
      // 保護のデフォルトで所有者とスクリプト実行者は編集可能
    });

    return true;
  } catch (err) {
    logError_(err, 'lockTeacherSheet_', { spreadsheetId, teacherEmail });
    return false;
  }
}

/**
 * スクリプト実行者がファイルを編集できることを確保
 * @param {GoogleAppsScript.Drive.File} file - Driveファイルオブジェクト
 * @param {string} scriptOwner - スクリプト実行者のメールアドレス
 * @param {string} owner - ファイル所有者のメールアドレス
 * @returns {Object} {success: boolean, isEditor: boolean, error: string|null}
 */
function ensureScriptOwnerCanEdit_(file, scriptOwner, owner) {
  const editors = getFileEditors_(file);
  let isEditor = (scriptOwner === owner) || editors.includes(scriptOwner);

  if (isEditor) {
    return { success: true, isEditor: true, error: null };
  }

  // 所有者と異なる場合のみ、エディターとして追加を試みる
  if (!ensureEditor_(file, scriptOwner)) {
    return {
      success: false,
      isEditor: false,
      error: `スクリプト実行者（${scriptOwner}）をエディターに追加できませんでした。所有者（${owner}）がスクリプト実行者の編集権限を手動で付与してください。`
    };
  }

  Utilities.sleep(DRIVE_CONFIG.SLEEP_AFTER_ADD_SCRIPT_OWNER - DRIVE_CONFIG.SLEEP_AFTER_ADD_EDITOR);
  const updatedEditors = getFileEditors_(file);
  isEditor = updatedEditors.includes(scriptOwner);

  return { success: true, isEditor: isEditor, error: null };
}

/**
 * スプレッドシートのすべてのシート保護を削除
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - スプレッドシートオブジェクト
 * @param {string} scriptOwner - スクリプト実行者のメールアドレス
 * @param {string} owner - ファイル所有者のメールアドレス
 * @returns {Object} {success: boolean, errors: Array<string>}
 */
function removeAllSheetProtections_(ss, scriptOwner, owner) {
  const sheets = ss.getSheets();
  let protectionRemoved = false;
  const errors = [];

  sheets.forEach((sheet) => {
    try {
      const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      if (protections.length === 0) {
        protectionRemoved = true;
        return;
      }

      protections.forEach((prot, pIdx) => {
        try {
          // スクリプト実行者が所有者と異なる場合、保護のエディターに追加を試みる
          if (scriptOwner !== owner) {
            try {
              prot.addEditor(scriptOwner);
            } catch (e) {
              // 既にエディターの場合や追加できない場合は無視
            }
          }

          prot.remove();
          protectionRemoved = true;
        } catch (e) {
          errors.push(`シート「${sheet.getName()}」の保護${pIdx + 1}削除失敗: ${e.message || String(e)}`);
        }
      });
    } catch (e) {
      errors.push(`シート「${sheet.getName()}」の保護取得失敗: ${e.message || String(e)}`);
    }
  });

  const success = protectionRemoved || errors.length === 0;
  return { success, errors };
}

/**
 * 講師のファイル編集権限を復元
 * @param {GoogleAppsScript.Drive.File} file - Driveファイルオブジェクト
 * @param {string} teacherEmail - 講師のメールアドレス
 * @param {Array<string>} viewers - 現在のビューアーリスト
 * @returns {Object} {success: boolean, errors: Array<string>}
 */
function restoreTeacherEditAccess_(file, teacherEmail, viewers) {
  const errors = [];

  if (!teacherEmail) {
    return { success: false, errors: ['講師のメールアドレスが指定されていません。'] };
  }

  // ビューアーから削除
  if (viewers.includes(teacherEmail)) {
    try {
      file.removeViewer(teacherEmail);
    } catch (e) {
      errors.push(`ビューアーから削除できませんでした。エラー: ${e.message || String(e)}`);
    }
  }

  // エディターとして追加
  if (!ensureEditor_(file, teacherEmail)) {
    errors.push(`講師（${teacherEmail}）をエディターに追加できませんでした。所有者が手動で編集権限を付与してください。`);
  }

  return { success: errors.length === 0, errors };
}

/**
 * 講師シートのロックを解除（管理者からの変更依頼時）
 * - 「リンクを知っているすべての人」を編集者に変更
 * - すべてのシート保護を削除
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} teacherEmail - 講師のメールアドレス（互換性のため保持）
 * @returns {Object} {success: boolean, errorMessage: string|null}
 */
function unlockTeacherSheet_(spreadsheetId, teacherEmail) {
  const errors = [];
  const scriptOwner = getEffectiveUserEmail_(spreadsheetId);

  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const file = DriveApp.getFileById(spreadsheetId);
    const owner = file.getOwner().getEmail();

    // ステップ1: スクリプト実行者の編集権限を確保
    const step1 = ensureScriptOwnerCanEdit_(file, scriptOwner, owner);
    if (step1.error) {
      errors.push(`ステップ1失敗: ${step1.error}`);
    }

    // ステップ2: すべてのシート保護を削除
    const step2 = removeAllSheetProtections_(ss, scriptOwner, owner);
    if (!step2.success && step2.errors.length > 0) {
      errors.push(`ステップ2失敗: 保護の削除に失敗しました。\n${step2.errors.join('\n')}`);
    }

    // ステップ3: 「リンクを知っているすべての人」を編集者に変更
    if (!setAnyoneWithLinkCanEdit_(spreadsheetId)) {
      errors.push(`ステップ3失敗: リンク共有権限を編集者に変更できませんでした。`);
    }

    // エラーがある場合は詳細を返す
    if (errors.length > 0) {
      const errorMessage = `ロック解除が部分的に失敗しました：\n\n${errors.join('\n\n')}\n\n` +
        `詳細情報:\n` +
        `- シートID: ${spreadsheetId}\n` +
        `- スクリプト実行者: ${scriptOwner}\n` +
        `- シート所有者: ${owner}\n` +
        `- スクリプト実行者がエディター: ${step1.isEditor ? 'はい' : 'いいえ'}`;

      logError_(new Error(errorMessage), 'unlockTeacherSheet_', {
        spreadsheetId, teacherEmail, scriptOwner, owner, errors
      });

      return { success: false, errorMessage };
    }

    return { success: true, errorMessage: null };
  } catch (err) {
    let owner = '(取得失敗)';
    try {
      const file = DriveApp.getFileById(spreadsheetId);
      owner = file.getOwner().getEmail();
    } catch (e) {}

    const errorMessage = `ロック解除中に予期しないエラーが発生しました：\n\n${err.message || String(err)}\n\n` +
      `詳細情報:\n` +
      `- シートID: ${spreadsheetId}\n` +
      `- スクリプト実行者: ${scriptOwner}\n` +
      `- シート所有者: ${owner}\n` +
      `- 講師メール: ${teacherEmail || '(未指定)'}\n` +
      (errors.length > 0 ? `\n部分的なエラー:\n${errors.join('\n')}` : '');

    logError_(err, 'unlockTeacherSheet_', {
      spreadsheetId, teacherEmail, scriptOwner,
      errorMessage: err.message || String(err),
      errorStack: err.stack || '', errors
    });

    return { success: false, errorMessage };
  }
}

/**
 * テンプレート側のコードを文字列として取得
 * template-side-code.jsファイルの内容を返す
 *
 * 注意: Google Apps Scriptではローカルファイルを直接読み込めないため、
 * template-side-code.jsファイルの内容を文字列として返します。
 * ファイルを更新した場合は、この関数も更新してください。
 *
 * @returns {string} テンプレート側のコード
 */
function getTemplateSideCode_() {
  // template-side-code.jsファイルの内容を文字列として返す
  // ファイルを更新した場合は、この文字列も更新してください
  return `/************************************************************
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
}`;
}

/**
 * スプレッドシートIDからスクリプトプロジェクトIDを取得
 * 
 * 注意: Google Apps Scriptの制約により、スプレッドシートIDから直接スクリプトプロジェクトIDを
 * 取得することはできません。そのため、以下の方法を使用します：
 * 
 * 1. スクリプトプロパティにマッピングが保存されている場合は、それを使用
 * 2. 保存されていない場合は、Apps Script APIを使用して取得を試みる
 * 3. 取得できない場合は、nullを返す
 * 
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {string|null} スクリプトプロジェクトID（取得できない場合はnull）
 */
function getScriptIdForSpreadsheet_(spreadsheetId) {
  try {
    // まず、スクリプトプロパティから取得を試みる
    const props = PropertiesService.getScriptProperties();
    const scriptIdKey = `SCRIPT_ID_${spreadsheetId}`;
    const cachedScriptId = props.getProperty(scriptIdKey);
    
    if (cachedScriptId) {
      return cachedScriptId;
    }
    
    // スクリプトプロパティに保存されていない場合、Apps Script APIを使用して取得を試みる
    // 注意: この方法は、スプレッドシートに紐づいたスクリプトプロジェクトが
    // 既に存在する場合にのみ動作します
    try {
      const accessToken = getAppsScriptAPIAccessToken();
      if (!accessToken) {
        console.log(`[getScriptIdForSpreadsheet_] Access token not available for spreadsheet: ${spreadsheetId}`);
        return null;
      }
      
      // Apps Script APIを使用して、プロジェクトリストを取得
      // 注意: この方法は、すべてのプロジェクトをリストアップするため、効率が悪い可能性があります
      // より効率的な方法は、スクリプトプロパティにマッピングを保存することです
      const url = 'https://script.googleapis.com/v1/projects';
      const response = UrlFetchApp.fetch(url, {
        method: 'get',
        headers: {
          'Authorization': 'Bearer ' + accessToken
        }
      });
      
      const result = JSON.parse(response.getContentText());
      
      if (result.error) {
        console.error(`[getScriptIdForSpreadsheet_] API error: ${JSON.stringify(result.error)}`);
        return null;
      }
      
      // プロジェクトリストから、スプレッドシートIDに一致するものを探す
      // 注意: この方法は、プロジェクト名やメタデータからスプレッドシートIDを推測する必要があります
      // 実際には、この方法は信頼性が低いため、スクリプトプロパティにマッピングを保存することを推奨します
      
      // 現時点では、スクリプトプロパティにマッピングを保存する方法を推奨します
      return null;
      
    } catch (apiErr) {
      console.error(`[getScriptIdForSpreadsheet_] Failed to get script ID via API: ${apiErr.message}`);
      return null;
    }
    
  } catch (err) {
    console.error('getScriptIdForSpreadsheet_ error:', err);
    logError_(err, 'getScriptIdForSpreadsheet_', { spreadsheetId });
    return null;
  }
}

/**
 * スクリプトプロジェクトIDをスクリプトプロパティに保存
 * 
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} scriptId - スクリプトプロジェクトID
 */
function saveScriptIdForSpreadsheet_(spreadsheetId, scriptId) {
  try {
    const props = PropertiesService.getScriptProperties();
    const scriptIdKey = `SCRIPT_ID_${spreadsheetId}`;
    props.setProperty(scriptIdKey, scriptId);
    console.log(`[saveScriptIdForSpreadsheet_] Saved script ID for spreadsheet: ${spreadsheetId} -> ${scriptId}`);
  } catch (err) {
    console.error('saveScriptIdForSpreadsheet_ error:', err);
    logError_(err, 'saveScriptIdForSpreadsheet_', { spreadsheetId, scriptId });
  }
}

/**
 * Apps Script APIを使用して、スプレッドシートにテンプレート側のコードを展開
 * 
 * 【前提条件】
 * 1. OAuth2認証が完了している（getAppsScriptAPIAccessToken()が動作すること）
 * 2. スクリプトプロジェクトIDが取得できること（getScriptIdForSpreadsheet_()が動作すること）
 * 
 * 【使用方法】
 * 1. 初回実行時: スクリプトプロジェクトIDを手動で取得して、saveScriptIdForSpreadsheet_()で保存
 * 2. 2回目以降: 自動的にスクリプトプロパティから取得して使用
 * 
 * 【スクリプトプロジェクトIDの取得方法】
 * 1. スプレッドシートを開く
 * 2. 「拡張機能」→「Apps Script」を選択
 * 3. ブラウザのURLからスクリプトIDを取得
 *    URLの形式: https://script.google.com/home/projects/{SCRIPT_ID}/edit
 *    {SCRIPT_ID}の部分がスクリプトプロジェクトIDです
 * 4. saveScriptIdForSpreadsheet_(spreadsheetId, scriptId) を実行して保存
 * 
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} scriptId - スクリプトプロジェクトID（省略可能、省略時は自動取得を試みる）
 * @returns {boolean} 成功した場合true
 */
function deployTemplateSideCodeToSpreadsheet_(spreadsheetId, scriptId = null) {
  try {
    // テンプレート側のコードを取得
    const templateCode = getTemplateSideCode_();
    
    // スクリプトプロジェクトIDを取得
    let targetScriptId = scriptId;
    if (!targetScriptId) {
      targetScriptId = getScriptIdForSpreadsheet_(spreadsheetId);
    }
    
    if (!targetScriptId) {
      console.log(`[deployTemplateSideCodeToSpreadsheet_] Script ID not found for spreadsheet: ${spreadsheetId}`);
      console.log(`[deployTemplateSideCodeToSpreadsheet_] Hint: Use saveScriptIdForSpreadsheet_() to save the script ID first`);
      // スクリプトプロジェクトIDが取得できない場合でも、エラーとしない
      // （テンプレートファイルにコードが含まれているため、コピー時に自動的にコードもコピーされる）
      return false;
    }
    
    // OAuth2認証を使用してApps Script APIにアクセス
    let accessToken;
    try {
      accessToken = getAppsScriptAPIAccessToken();
    } catch (authErr) {
      console.error(`[deployTemplateSideCodeToSpreadsheet_] Authentication failed: ${authErr.message}`);
      console.log(`[deployTemplateSideCodeToSpreadsheet_] Hint: Run getOAuthAuthorizationUrl() to authenticate first`);
      return false;
    }
    
    if (!accessToken) {
      console.error(`[deployTemplateSideCodeToSpreadsheet_] Access token not available`);
      return false;
    }
    
    // 既存のコードを取得
    const getUrl = `https://script.googleapis.com/v1/projects/${targetScriptId}/content`;
    const getResponse = UrlFetchApp.fetch(getUrl, {
      method: 'get',
      headers: {
        'Authorization': 'Bearer ' + accessToken
      }
    });
    
    const getResult = JSON.parse(getResponse.getContentText());
    
    if (getResult.error) {
      console.error(`[deployTemplateSideCodeToSpreadsheet_] Failed to get existing code: ${JSON.stringify(getResult.error)}`);
      return false;
    }
    
    // 既存のファイルを更新または追加
    const existingFiles = getResult.files || [];
    const codeFileName = 'Code'; // デフォルトのファイル名
    
    // 既存のCodeファイルを探す
    let codeFileExists = false;
    const updatedFiles = existingFiles.map(file => {
      if (file.name === codeFileName) {
        codeFileExists = true;
        return {
          name: file.name,
          type: file.type || 'SERVER_JS',
          source: templateCode
        };
      }
      // 他のファイルはそのまま保持
      return {
        name: file.name,
        type: file.type,
        source: file.source || ''
      };
    });
    
    // Codeファイルが存在しない場合は追加
    if (!codeFileExists) {
      updatedFiles.push({
        name: codeFileName,
        type: 'SERVER_JS',
        source: templateCode
      });
    }
    
    // 新しいコードで更新
    const updateUrl = `https://script.googleapis.com/v1/projects/${targetScriptId}/content`;
    const updateResponse = UrlFetchApp.fetch(updateUrl, {
      method: 'put',
      headers: {
        'Authorization': 'Bearer ' + accessToken,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        files: updatedFiles
      })
    });
    
    const updateResult = JSON.parse(updateResponse.getContentText());
    
    if (updateResult.error) {
      console.error(`[deployTemplateSideCodeToSpreadsheet_] Failed to update code: ${JSON.stringify(updateResult.error)}`);
      return false;
    }
    
    console.log(`[deployTemplateSideCodeToSpreadsheet_] Successfully deployed template code to spreadsheet: ${spreadsheetId}`);
    return true;
    
  } catch (err) {
    console.error('deployTemplateSideCodeToSpreadsheet_ error:', err);
    logError_(err, 'deployTemplateSideCodeToSpreadsheet_', { spreadsheetId, scriptId });
    // エラーが発生しても処理を続行（テンプレートのコードが使用される）
    return false;
  }
}
