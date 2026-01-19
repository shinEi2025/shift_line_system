/************************************************************
 * 自動コード展開機能のテスト関数
 * 
 * このファイルには、OAuth2認証とスクリプトプロジェクトID管理が
 * 正常に動作しているか確認するためのテスト関数が含まれています。
 * 
 * 【使用方法】
 * 1. このファイルをマスター側のApps Scriptプロジェクトに追加
 * 2. 各テスト関数を実行して、結果を確認
 ************************************************************/

/**
 * OAuth2認証の状態を確認するテスト関数
 * @return {Object} 認証状態の詳細情報
 */
function testOAuth2Auth() {
  try {
    console.log('=== OAuth2認証テスト開始 ===');
    
    // アクセストークンの取得を試みる
    try {
      const accessToken = getAppsScriptAPIAccessToken();
      
      if (accessToken) {
        console.log('✓ OAuth2認証: 成功');
        console.log('  アクセストークン: ' + accessToken.substring(0, 20) + '...');
        return {
          success: true,
          message: 'OAuth2認証は正常に動作しています',
          hasAccessToken: true
        };
      } else {
        console.log('✗ OAuth2認証: アクセストークンが取得できませんでした');
        return {
          success: false,
          message: 'アクセストークンが取得できませんでした。getOAuthAuthorizationUrl()を実行して認証してください。',
          hasAccessToken: false
        };
      }
    } catch (authErr) {
      console.log('✗ OAuth2認証: エラー - ' + authErr.message);
      return {
        success: false,
        message: 'OAuth2認証エラー: ' + authErr.message,
        error: authErr.message,
        hasAccessToken: false
      };
    }
  } catch (err) {
    console.error('テスト実行エラー:', err);
    return {
      success: false,
      message: 'テスト実行中にエラーが発生しました: ' + err.message,
      error: err.message
    };
  }
}

/**
 * スクリプトプロジェクトID管理の状態を確認するテスト関数
 * @param {string} spreadsheetId - テスト対象のスプレッドシートID（任意）
 * @return {Object} スクリプトプロジェクトID管理の状態
 */
function testScriptIdManagement(spreadsheetId = null) {
  try {
    console.log('=== スクリプトプロジェクトID管理テスト開始 ===');
    
    // テスト用のスプレッドシートIDを使用（指定がない場合）
    if (!spreadsheetId) {
      spreadsheetId = CONFIG.TEMPLATE_SPREADSHEET_ID;
      console.log('テンプレートスプレッドシートIDを使用: ' + spreadsheetId);
    }
    
    // スクリプトプロジェクトIDの取得を試みる
    const scriptId = getScriptIdForSpreadsheet_(spreadsheetId);
    
    if (scriptId) {
      console.log('✓ スクリプトプロジェクトID取得: 成功');
      console.log('  スクリプトID: ' + scriptId);
      return {
        success: true,
        message: 'スクリプトプロジェクトIDは正常に取得できました',
        spreadsheetId: spreadsheetId,
        scriptId: scriptId
      };
    } else {
      console.log('✗ スクリプトプロジェクトID取得: 失敗');
      console.log('  ヒント: saveScriptIdForSpreadsheet_()を使用してスクリプトプロジェクトIDを保存してください');
      return {
        success: false,
        message: 'スクリプトプロジェクトIDが取得できませんでした。saveScriptIdForSpreadsheet_(spreadsheetId, scriptId)を実行して保存してください。',
        spreadsheetId: spreadsheetId
      };
    }
  } catch (err) {
    console.error('テスト実行エラー:', err);
    return {
      success: false,
      message: 'テスト実行中にエラーが発生しました: ' + err.message,
      error: err.message
    };
  }
}

/**
 * 自動コード展開機能の完全なテスト関数
 * @param {string} spreadsheetId - テスト対象のスプレッドシートID（任意）
 * @return {Object} テスト結果の詳細情報
 */
function testCodeDeployment(spreadsheetId = null) {
  try {
    console.log('=== 自動コード展開機能テスト開始 ===');
    
    const results = {
      oauth2: null,
      scriptId: null,
      deployment: null,
      overall: { success: false, message: '' }
    };
    
    // 1. OAuth2認証のテスト
    console.log('\n[1/3] OAuth2認証のテスト...');
    results.oauth2 = testOAuth2Auth();
    
    if (!results.oauth2.success) {
      results.overall = {
        success: false,
        message: 'OAuth2認証が失敗しています。先にOAuth2認証を完了してください。'
      };
      return results;
    }
    
    // 2. スクリプトプロジェクトID管理のテスト
    console.log('\n[2/3] スクリプトプロジェクトID管理のテスト...');
    results.scriptId = testScriptIdManagement(spreadsheetId);
    
    if (!results.scriptId.success) {
      results.overall = {
        success: false,
        message: 'スクリプトプロジェクトIDが取得できません。先にスクリプトプロジェクトIDを保存してください。'
      };
      return results;
    }
    
    // 3. 実際のコード展開のテスト（オプション）
    // spreadsheetIdが指定されていない場合は、テンプレートスプレッドシートIDを使用
    const targetSpreadsheetId = spreadsheetId || CONFIG.TEMPLATE_SPREADSHEET_ID;
    
    if (targetSpreadsheetId && results.scriptId.scriptId) {
      console.log('\n[3/3] コード展開のテスト...');
      console.log('  対象スプレッドシートID: ' + targetSpreadsheetId);
      try {
        const deployResult = deployTemplateSideCodeToSpreadsheet_(targetSpreadsheetId, results.scriptId.scriptId);
        
        if (deployResult) {
          console.log('✓ コード展開: 成功');
          results.deployment = {
            success: true,
            message: 'コード展開が正常に実行されました'
          };
          results.overall = {
            success: true,
            message: 'すべてのテストが成功しました！自動コード展開機能は正常に動作しています。'
          };
        } else {
          console.log('✗ コード展開: 失敗');
          results.deployment = {
            success: false,
            message: 'コード展開が失敗しました。実行ログを確認してください。'
          };
          results.overall = {
            success: false,
            message: 'コード展開が失敗しました。実行ログを確認してください。'
          };
        }
      } catch (deployErr) {
        console.log('✗ コード展開: エラー - ' + deployErr.message);
        results.deployment = {
          success: false,
          message: 'コード展開中にエラーが発生しました: ' + deployErr.message,
          error: deployErr.message
        };
        results.overall = {
          success: false,
          message: 'コード展開中にエラーが発生しました: ' + deployErr.message
        };
      }
    } else {
      console.log('[3/3] コード展開のテスト: スキップ（スクリプトIDが取得できませんでした）');
      results.deployment = {
        success: null,
        message: 'テストがスキップされました'
      };
      results.overall = {
        success: true,
        message: 'OAuth2認証とスクリプトプロジェクトID管理は正常に動作しています。コード展開のテストを実行するには、スプレッドシートIDを指定してください。'
      };
    }
    
    console.log('\n=== テスト完了 ===');
    console.log('全体的な結果: ' + (results.overall.success ? '成功' : '失敗'));
    console.log(results.overall.message);
    
    return results;
    
  } catch (err) {
    console.error('テスト実行エラー:', err);
    return {
      oauth2: null,
      scriptId: null,
      deployment: null,
      overall: {
        success: false,
        message: 'テスト実行中にエラーが発生しました: ' + err.message,
        error: err.message
      }
    };
  }
}

/**
 * すべてのテストを実行して結果を表示する関数（簡単な確認用）
 */
function runAllTests() {
  console.log('=== 自動コード展開機能の完全テスト ===\n');
  
  const results = testCodeDeployment();
  
  console.log('\n=== テスト結果サマリー ===');
  console.log('OAuth2認証: ' + (results.oauth2?.success ? '✓ 成功' : '✗ 失敗'));
  console.log('スクリプトID管理: ' + (results.scriptId?.success ? '✓ 成功' : '✗ 失敗'));
  console.log('コード展開: ' + (results.deployment?.success === null ? '- スキップ' : (results.deployment?.success ? '✓ 成功' : '✗ 失敗')));
  console.log('\n全体結果: ' + (results.overall.success ? '✓ すべて正常' : '✗ 問題あり'));
  console.log(results.overall.message);
  
  return results;
}

