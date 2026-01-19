// OAuth2ライブラリを使わずに実装（UrlFetchAppを直接使用）

/**
 * Apps Script APIのアクセストークンを取得します
 * @return {string} アクセストークン
 */
function getAppsScriptAPIAccessToken() {
  var props = PropertiesService.getUserProperties();
  var accessToken = props.getProperty('oauth_access_token');
  var expiresAt = props.getProperty('oauth_expires_at');
  
  // トークンが存在し、有効期限内の場合
  if (accessToken && expiresAt && new Date().getTime() < parseInt(expiresAt)) {
    return accessToken;
  }
  
  // リフレッシュトークンがある場合、トークンを更新
  var refreshToken = props.getProperty('oauth_refresh_token');
  if (refreshToken) {
    return refreshAccessToken(refreshToken);
  }
  
  // 認証が必要な場合
  throw new Error('認証が必要です。getOAuthAuthorizationUrl() を実行して認証URLを取得してください。');
}

/**
 * OAuth2認証URLを取得します（初回認証用）
 * @return {string} 認証URL
 */
function getOAuthAuthorizationUrl() {
  try {
    var CLIENT_ID = PropertiesService.getScriptProperties().getProperty('OAUTH_CLIENT_ID');
    
    if (!CLIENT_ID) {
      throw new Error('OAUTH_CLIENT_ID がスクリプトプロパティに設定されていません。');
    }
    
    // リダイレクトURIを取得（Webアプリとしてデプロイされている場合）
    var redirectUri;
    try {
      redirectUri = ScriptApp.getService().getUrl();
    } catch (e) {
      // Webアプリとしてデプロイされていない場合、手動で設定する必要がある
      // スクリプトプロパティから取得を試みる
      redirectUri = PropertiesService.getScriptProperties().getProperty('OAUTH_REDIRECT_URI');
      
      if (!redirectUri) {
        throw new Error('リダイレクトURIが取得できません。Webアプリとしてデプロイするか、スクリプトプロパティに OAUTH_REDIRECT_URI を設定してください。');
      }
    }
    
    var authUrl = 'https://accounts.google.com/o/oauth2/auth?' +
      'client_id=' + encodeURIComponent(CLIENT_ID) +
      '&redirect_uri=' + encodeURIComponent(redirectUri) +
      '&response_type=code' +
      '&scope=' + encodeURIComponent('https://www.googleapis.com/auth/script.projects') +
      '&access_type=offline' +
      '&prompt=consent';
    
    Logger.log('認証URL: ' + authUrl);
    Logger.log('リダイレクトURI: ' + redirectUri);
    
    return authUrl;
  } catch (error) {
    Logger.log('エラー: ' + error.toString());
    throw error;
  }
}

/**
 * OAuth2認証コードからアクセストークンを取得します
 * @param {string} code 認証コード
 * @return {string} アクセストークン
 */
function exchangeAuthorizationCode(code) {
  var CLIENT_ID = PropertiesService.getScriptProperties().getProperty('OAUTH_CLIENT_ID');
  var CLIENT_SECRET = PropertiesService.getScriptProperties().getProperty('OAUTH_CLIENT_SECRET');
  
  // リダイレクトURIを取得
  var redirectUri;
  try {
    redirectUri = ScriptApp.getService().getUrl();
  } catch (e) {
    redirectUri = PropertiesService.getScriptProperties().getProperty('OAUTH_REDIRECT_URI');
    if (!redirectUri) {
      throw new Error('リダイレクトURIが取得できません。');
    }
  }
  
  var payload = {
    'code': code,
    'client_id': CLIENT_ID,
    'client_secret': CLIENT_SECRET,
    'redirect_uri': redirectUri,
    'grant_type': 'authorization_code'
  };
  
  var options = {
    'method': 'post',
    'contentType': 'application/x-www-form-urlencoded',
    'payload': Object.keys(payload).map(function(key) {
      return encodeURIComponent(key) + '=' + encodeURIComponent(payload[key]);
    }).join('&')
  };
  
  var response = UrlFetchApp.fetch('https://accounts.google.com/o/oauth2/token', options);
  var result = JSON.parse(response.getContentText());
  
  if (result.error) {
    throw new Error('認証エラー: ' + result.error);
  }
  
  // トークンを保存
  var props = PropertiesService.getUserProperties();
  props.setProperty('oauth_access_token', result.access_token);
  props.setProperty('oauth_refresh_token', result.refresh_token);
  
  var expiresAt = new Date().getTime() + (result.expires_in * 1000);
  props.setProperty('oauth_expires_at', expiresAt.toString());
  
  return result.access_token;
}

/**
 * リフレッシュトークンを使ってアクセストークンを更新します
 * @param {string} refreshToken リフレッシュトークン
 * @return {string} 新しいアクセストークン
 */
function refreshAccessToken(refreshToken) {
  var CLIENT_ID = PropertiesService.getScriptProperties().getProperty('OAUTH_CLIENT_ID');
  var CLIENT_SECRET = PropertiesService.getScriptProperties().getProperty('OAUTH_CLIENT_SECRET');
  
  var payload = {
    'refresh_token': refreshToken,
    'client_id': CLIENT_ID,
    'client_secret': CLIENT_SECRET,
    'grant_type': 'refresh_token'
  };
  
  var options = {
    'method': 'post',
    'contentType': 'application/x-www-form-urlencoded',
    'payload': Object.keys(payload).map(function(key) {
      return encodeURIComponent(key) + '=' + encodeURIComponent(payload[key]);
    }).join('&')
  };
  
  var response = UrlFetchApp.fetch('https://accounts.google.com/o/oauth2/token', options);
  var result = JSON.parse(response.getContentText());
  
  if (result.error) {
    throw new Error('トークン更新エラー: ' + result.error);
  }
  
  // トークンを保存
  var props = PropertiesService.getUserProperties();
  props.setProperty('oauth_access_token', result.access_token);
  
  if (result.expires_in) {
    var expiresAt = new Date().getTime() + (result.expires_in * 1000);
    props.setProperty('oauth_expires_at', expiresAt.toString());
  }
  
  return result.access_token;
}

/**
 * OAuth2認証コールバックのハンドラー（Code.jsのdoGet()から呼び出される）
 * @param {Object} e リクエストパラメータ
 * @return {HtmlOutput} HTML出力
 */
function handleOAuthCallback_(e) {
  try {
    var accessToken = exchangeAuthorizationCode(e.parameter.code);
    return HtmlService.createHtmlOutput(
      '<html><body style="font-family: Arial; padding: 20px;">' +
      '<h2>認証が成功しました！</h2>' +
      '<p>このウィンドウを閉じてください。</p>' +
      '</body></html>'
    );
  } catch (error) {
    return HtmlService.createHtmlOutput(
      '<html><body style="font-family: Arial; padding: 20px;">' +
      '<h2>認証エラー</h2>' +
      '<p>' + error.toString() + '</p>' +
      '</body></html>'
    );
  }
}

/**
 * OAuth認証ページを表示（Code.jsのdoGet()から呼び出される）
 * @return {HtmlOutput} HTML出力
 */
function showOAuthPage_() {
  try {
    var authUrl = getOAuthAuthorizationUrl();
    return HtmlService.createHtmlOutput(
      '<html><body style="font-family: Arial; padding: 20px;">' +
      '<h2>OAuth認証</h2>' +
      '<p><a href="' + authUrl + '" target="_blank" style="background-color: #4285f4; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px;">認証を開始</a></p>' +
      '<p style="font-size: 12px; color: #666;">または、以下のURLをコピーしてブラウザで開いてください：</p>' +
      '<p style="font-size: 12px; word-break: break-all;">' + authUrl + '</p>' +
      '</body></html>'
    );
  } catch (urlError) {
    return HtmlService.createHtmlOutput(
      '<html><body style="font-family: Arial; padding: 20px;">' +
      '<h2>エラー: 認証URLの取得に失敗しました</h2>' +
      '<p style="color: red;">' + urlError.toString() + '</p>' +
      '<h3>確認事項:</h3>' +
      '<ul>' +
      '<li>スクリプトプロパティに <code>OAUTH_CLIENT_ID</code> が設定されているか確認してください</li>' +
      '<li>Webアプリとしてデプロイされているか確認してください</li>' +
      '<li>または、スクリプトプロパティに <code>OAUTH_REDIRECT_URI</code> を設定してください</li>' +
      '</ul>' +
      '<p>実行ログを確認して、詳細なエラーメッセージを確認してください。</p>' +
      '</body></html>'
    );
  }
}
