# テンプレートファイルのGASコード更新手順

## 概要

このシステムでは、マスター側（Code.js）に全てのコードが統合されていますが、テンプレートファイルにも**講師用シートで実行されるコード**を含める必要があります。

## なぜテンプレートにコードが必要なのか？

1. **コピー時にコードもコピーされる**: `copyTemplateSpreadsheet_()` が `makeCopy()` を使用するため、テンプレートに含まれるGASコードもコピーされます
2. **各シートで動作する必要がある**: `onOpen()`, `onEdit()` などの関数は、各講師用シート（コピーされたシート）に紐づいたスクリプトとして存在する必要があります
3. **マスター側のコードだけでは動作しない**: マスター側のコードはマスターシートにのみ紐づいているため、コピーされたシートでは動作しません

## 自動コード展開機能（実験的）

`copyTemplateSpreadsheet_()` 関数を拡張して、コピー後に自動的にマスター側のコードを展開する機能を追加しました。

**注意**: 現在の実装では、Google Apps Scriptの制約により、実際のコード展開は実装されていません。代わりに、テンプレートファイルにコードが含まれているため、コピー時に自動的にコードもコピーされます。

将来的に、Apps Script APIを使用してコードを展開する機能を実装する場合は、以下の手順が必要です：

1. **Apps Script APIの有効化**: Google Cloud ConsoleでApps Script APIを有効にする
2. **OAuth2認証の設定**: Apps Script APIを使用するためのOAuth2認証を設定する
3. **スクリプトプロジェクトIDの管理**: スプレッドシートに紐づいたスクリプトプロジェクトのIDを管理する仕組みを構築する

### Apps Script APIの有効化手順

Google Cloud ConsoleでApps Script APIを有効化するには、以下の手順を実行します：

1. **Google Cloud Consoleにアクセス**
   - [Google Cloud Console](https://console.cloud.google.com/) にアクセス

2. **プロジェクトの選択**
   - 画面上部のプロジェクト選択ドロップダウンから、Apps Scriptプロジェクトに関連付けられたGCPプロジェクトを選択
   - プロジェクトが存在しない場合は、新しいプロジェクトを作成
     - 「プロジェクトを選択」→「新しいプロジェクト」をクリック
     - プロジェクト名を入力して「作成」をクリック

3. **Apps Script APIを有効化**
   - 左側のメニューから「APIとサービス」→「ライブラリ」を選択
   - 検索ボックスに「Apps Script API」と入力
   - 「Google Apps Script API」をクリック
   - 「有効にする」ボタンをクリック

4. **確認**
   - 「APIとサービス」→「有効なAPI」に移動
   - 「Google Apps Script API」が一覧に表示されていることを確認

**注意事項**:
- Apps Script APIを使用するには、該当のGCPプロジェクトとApps Scriptプロジェクトが同じGoogleアカウントに関連付けられている必要があります
- デフォルトのGCPプロジェクトを使用している場合でも、Apps Script APIは明示的に有効化する必要があります

### OAuth2認証の設定手順

Apps Script APIを使用するには、OAuth2認証を設定する必要があります。以下の手順を実行します：

1. **OAuth同意画面の設定**
   - Google Cloud Consoleで、左側のメニューから「APIとサービス」→「OAuth同意画面」を選択
   - ユーザータイプを選択（通常は「内部」を選択、G Suite以外の場合は「外部」）
   - 「作成」をクリック
   - 必須項目を入力：
     - アプリ名：適切な名前を入力（例：「シフト提出管理システム」）
     - ユーザーサポートメール：自分のメールアドレスを選択
     - デベロッパーの連絡先情報：自分のメールアドレスを入力
   - 「保存して次へ」をクリック
   - スコープの追加（必要な場合）：
     - 「スコープを追加または削除」をクリック
     - `https://www.googleapis.com/auth/script.projects` を追加（Apps Script APIに必要）
     - 「更新」→「保存して次へ」をクリック
   - テストユーザーの追加（外部アプリの場合、または内部アプリでもテストモードの場合）：
     - **重要**: OAuth同意画面が「テスト」モードの場合、認証に使用するGoogleアカウントをテストユーザーとして追加する必要があります
     - 「テストユーザー」セクションで「+ ユーザーを追加」をクリック
     - 認証に使用するGoogleアカウントのメールアドレスを入力（例: `schooliegakuen@gmail.com`）
     - 「追加」をクリック
     - 複数のアカウントを使用する場合は、それぞれ追加してください
     - **注意**: テストユーザーを追加しないと、「エラー 403: access_denied」が発生します
   - 「ダッシュボードに戻る」をクリック

   **OAuth同意画面の設定が完了したら:**
   - 「OAuth の概要」画面が表示されることを確認
   - 次のステップ（OAuth 2.0 クライアントIDの作成）に進みます

   **「エラー 403: access_denied」が表示される場合:**
   - OAuth同意画面が「テスト」モードになっている可能性があります
   - Google Cloud Consoleで「APIとサービス」→「OAuth同意画面」を開く
   - 「テストユーザー」セクションで、認証に使用するGoogleアカウントを追加してください
   - または、OAuth同意画面を「公開」に変更することもできます（審査が必要な場合があります）

2. **OAuth 2.0 クライアントIDの作成**

   OAuth同意画面の設定が完了したら、次にOAuthクライアントを作成します。

   **方法A: OAuth の概要画面から作成（推奨）**
   - 現在表示されている「OAuth の概要」画面で、「OAuth クライアントを作成」ボタンをクリック
   - または、左側のメニューから「クライアント」を選択

   **方法B: 従来の方法**
   - 「APIとサービス」→「認証情報」を選択
   - 画面上部の「+ 認証情報を作成」→「OAuth 2.0 クライアント ID」を選択

   **クライアントの設定:**
   - アプリケーションの種類で「ウェブアプリケーション」を選択
   - 名前を入力（例：「Apps Script API クライアント」）
   - **承認済みのリダイレクト URI**:
     - **OAuth2ライブラリを使用する場合（方法A）**: 通常は不要（空のままでOK）
     - **ライブラリを使わずに実装する場合（方法B）**: 
       - 一時的に空欄のまま作成しても構いません（後で編集可能）
       - または、後でWebアプリとしてデプロイした際に表示されるURLを追加します
       - 方法B-1（手動認証）を使用する場合は、リダイレクトURIは不要です
       - 方法B-2（Webアプリ）を使用する場合は、WebアプリのURLをリダイレクトURIとして設定します
   - 「作成」をクリック
   - 表示された「クライアント ID」と「クライアント シークレット」をコピーして保存（後で使用）
     - **重要**: この2つの値は後でApps Scriptのスクリプトプロパティに設定する必要があります

3. **スクリプトプロパティの設定**

   OAuth 2.0 クライアントIDとシークレットをApps Scriptのスクリプトプロパティに保存します：

   - Apps Scriptエディタを開く（https://script.google.com/ または スプレッドシートから「拡張機能」→「Apps Script」）
   - 左側のメニューから「プロジェクトの設定」（⚙️ アイコン）をクリック
   - 下にスクロールして「スクリプト プロパティ」セクションを表示
   - 「スクリプト プロパティを追加」をクリック
   - 以下の2つのプロパティを追加：
     - **プロパティ**: `OAUTH_CLIENT_ID` → **値**: ステップ2で取得したクライアント ID
     - 「スクリプト プロパティを追加」を再度クリック
     - **プロパティ**: `OAUTH_CLIENT_SECRET` → **値**: ステップ2で取得したクライアント シークレット
   - 「保存」をクリック

4. **OAuth2ライブラリの追加**

   Apps ScriptでOAuth2認証を使用するために、OAuth2ライブラリを追加します。このライブラリは、OAuth2認証を簡単に実装するためのヘルパーライブラリです。

   **詳細な手順:**

   1. **Apps Scriptエディタを開く**
      - Googleドライブでスプレッドシートを開き、「拡張機能」→「Apps Script」を選択
      - または、https://script.google.com/ にアクセスして、プロジェクトを開く

   2. **ライブラリメニューを開く**
      - 左側のメニュー（ファイル一覧が表示されている場所）の上部にある「ライブラリ」（📚 アイコン、または「ライブラリ」という文字）をクリック
      - 左側のメニューにライブラリのセクションが表示されます

   3. **ライブラリを追加**
      - 「ライブラリを追加」ボタンをクリック（または「+」アイコン）
      - ダイアログボックスが表示されます

   4. **スクリプトIDを入力**
      - 「スクリプトIDを追加」または「スクリプトID」という入力欄に、以下のIDをコピー＆ペーストします：
        ```
        1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuBYs9t7he0AF_gdUPRcIhL
        ```
      - これは「Apps Script OAuth2」ライブラリのIDです

   5. **検索して追加**
      - 「検索」ボタンをクリック（またはEnterキーを押す）
      - ライブラリの情報が表示されます（名前: "OAuth2" など）

   6. **バージョンを選択**
      - バージョン選択ドロップダウンが表示されます
      - 通常は最新版（最も大きな番号）を選択します
      - 例: "24" や "最新" と表示される場合

   7. **識別子を設定（オプション）**
      - 「識別子」フィールドには、コード内で使用する名前を入力します
      - 通常は `OAuth2` のまま（変更不要）でOKです
      - コード内で `OAuth2.createService()` のように使用する際の名前になります

   8. **保存**
      - 「追加」ボタン（または「保存」ボタン）をクリック

   9. **追加の確認**
      - 左側のメニューに「ライブラリ」セクションが表示され、「OAuth2」というライブラリが追加されていることを確認します
      - ライブラリ名の横にバージョン番号が表示されます

   **もし「スクリプトID」の入力欄が見つからない場合:**
   - ダイアログボックスの上部にタブがある場合、「スクリプトID」タブを選択してください
   - または、「URLから追加」ではなく「スクリプトIDから追加」を選択してください

   **エラーが表示される場合:**
   - スクリプトIDが正しくコピーされているか確認してください（スペースが入っていないか）
   - インターネット接続を確認してください
   - しばらく待ってから再度試してください

   **「ライブラリを検索できませんでした」というエラーが表示される場合:**

   このエラーは、ライブラリIDが古いか、アクセス権限の問題で表示されることがあります。以下の代替方法を試してください：

   **方法A: 別のOAuth2ライブラリIDを試す**
   - 以下の別のライブラリIDを試してください：
     ```
     MswhXl8fVhTFUH_Q3UOJbXvxhLh3-eEQr7BHTl0UW_9yOLBNcEPDzgdW
     ```
   - これは同じOAuth2ライブラリの別のバージョンです

   **方法B: OAuth2ライブラリを使わずに実装する（推奨）**
   
   OAuth2ライブラリが使えない場合は、UrlFetchAppを直接使ってOAuth2認証を実装できます。この方法はライブラリに依存しないため、より確実に動作します。
   
   ステップ5で提供されるコードの代わりに、以下の代替実装を使用してください（ステップ5の「代替実装（ライブラリ不要）」セクションを参照）。

5. **Apps ScriptでのOAuth2認証の実装**

   OAuth2ライブラリが使用できる場合は「方法A」、使用できない場合は「方法B」を選択してください。

   **方法A: OAuth2ライブラリを使用する場合**

   Apps Scriptプロジェクトに、OAuth2認証用のコードを追加します。以下のコードを適切なファイル（例：`Code.js` または `oauth.js`）に追加します：

   ```javascript
   // OAuth2ライブラリを使用（ライブラリID: 1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuBYs9t7he0AF_gdUPRcIhL)
   
   /**
    * Apps Script APIのアクセストークンを取得します
    * @return {string} アクセストークン
    */
   function getAppsScriptAPIAccessToken() {
     var service = getOAuthService();
     if (service.hasAccess()) {
       return service.getAccessToken();
     } else {
       var authorizationUrl = service.getAuthorizationUrl();
       Logger.log('次のURLにアクセスして認証してください: ' + authorizationUrl);
       throw new Error('認証が必要です。上記のURLにアクセスしてください。');
     }
   }
   
   /**
    * OAuth2サービスを取得します
    * @return {OAuth2Service} OAuth2サービスオブジェクト
    */
   function getOAuthService() {
     // スクリプトプロパティから取得
     var CLIENT_ID = PropertiesService.getScriptProperties().getProperty('OAUTH_CLIENT_ID');
     var CLIENT_SECRET = PropertiesService.getScriptProperties().getProperty('OAUTH_CLIENT_SECRET');
     
     return OAuth2.createService('apps-script-api')
       .setAuthorizationBaseUrl('https://accounts.google.com/o/oauth2/auth')
       .setTokenUrl('https://accounts.google.com/o/oauth2/token')
       .setClientId(CLIENT_ID)
       .setClientSecret(CLIENT_SECRET)
       .setScope('https://www.googleapis.com/auth/script.projects')
       .setCallbackFunction('authCallback')
       .setPropertyStore(PropertiesService.getUserProperties());
   }
   
   /**
    * OAuth2認証のコールバック関数
    * @param {Object} request リクエストオブジェクト
    * @return {HtmlOutput} HTML出力
    */
   function authCallback(request) {
     var service = getOAuthService();
     var authorized = service.handleCallback(request);
     if (authorized) {
       return HtmlService.createHtmlOutput('認証が成功しました。このウィンドウを閉じてください。');
     } else {
       return HtmlService.createHtmlOutput('認証が拒否されました。');
     }
   }
   ```

   **方法B: OAuth2ライブラリを使わずに実装する場合（推奨）**

   OAuth2ライブラリが使用できない場合は、UrlFetchAppを使って直接OAuth2認証を実装します。この方法はライブラリに依存しないため、より確実に動作します。

   Apps Scriptプロジェクトに、以下のコードを適切なファイル（例：`Code.js` または `oauth.js`）に追加します：

   ```javascript
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
     var redirectUri = ScriptApp.getService().getUrl();
     
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
    * OAuth2認証のコールバック関数（Webアプリとしてデプロイした場合）
    * @param {Object} e リクエストパラメータ
    * @return {HtmlOutput} HTML出力
    */
   function doGet(e) {
     try {
       if (e.parameter.code) {
         // 認証コードが返ってきた場合
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
       } else {
         // 認証URLを取得して表示
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
           // 認証URLの取得に失敗した場合
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
     } catch (error) {
       return HtmlService.createHtmlOutput(
         '<html><body style="font-family: Arial; padding: 20px;">' +
         '<h2>エラーが発生しました</h2>' +
         '<p>' + error.toString() + '</p>' +
         '<p>スクリプトプロパティに OAUTH_CLIENT_ID が設定されているか確認してください。</p>' +
         '</body></html>'
       );
     }
   }
   ```

   **注意**: 方法Bを使用する場合、Webアプリとしてデプロイするか、`getOAuthAuthorizationUrl()` 関数を実行して表示されたURLに手動でアクセスする必要があります。

### 方法Bの実装手順（詳細ガイド）

方法B（OAuth2ライブラリを使わない実装）を実装する場合の、完全なステップバイステップガイドです。

**概要**: この方法では、OAuth2ライブラリを使わずに、UrlFetchAppを使って直接OAuth2認証を実装します。ライブラリに依存しないため、より確実に動作します。

**実装の流れ**:
1. OAuth2認証用のコードを追加（5つの関数）
2. Webアプリとしてデプロイ（推奨）または手動認証
3. リダイレクトURIをGoogle Cloud Consoleで設定
4. 認証を実行
5. アクセストークンを取得してApps Script APIを使用

#### 前提条件

- ✅ OAuth同意画面の設定が完了している（ステップ1）
- ✅ OAuth 2.0 クライアントIDとシークレットを取得済み（ステップ2）
- ✅ Apps Scriptのスクリプトプロパティに `OAUTH_CLIENT_ID` と `OAUTH_CLIENT_SECRET` を設定済み（ステップ3）

#### 実装手順

**ステップ1: コードファイルの作成**

1. Apps Scriptエディタを開く
2. 左側のメニューから「+」をクリックして新しいファイルを作成
3. ファイル名を `oauth.js` に変更（任意、わかりやすい名前なら何でもOK）
4. または、既存の `Code.js` ファイルにコードを追加しても構いません

**ステップ2: OAuth2認証コードの追加**

1. ドキュメントの「方法B: OAuth2ライブラリを使わずに実装する場合（推奨）」セクションを探す
2. ```javascript から ``` までのコード全体をコピー（最初の行から最後の行まで）
3. 作成したファイル（`oauth.js` など）に貼り付け
4. **必ず「保存」をクリック**（Ctrl+S / Cmd+S、または「保存」アイコン）

**重要**: 
- コードを貼り付けたら、必ず「保存」をクリックしてください
- 保存しないと、関数が認識されません
- 保存後、関数一覧（エディタ上部のドロップダウン）に `getAppsScriptAPIAccessToken` が表示されることを確認してください

**ステップ3: コードの確認**

1. **関数が認識されているか確認**
   - エディタ上部の関数選択ドロップダウンを開く
   - 以下の関数が表示されることを確認：
     - `getAppsScriptAPIAccessToken`
     - `getOAuthAuthorizationUrl`
     - `exchangeAuthorizationCode`
     - `refreshAccessToken`
     - `doGet`
   - 表示されない場合は、ファイルを保存し直してください

2. **コードの内容を確認**
   - ファイル内で `getAppsScriptAPIAccessToken` を検索（Ctrl+F / Cmd+F）
   - 関数が定義されていることを確認

3. **構文エラーがないか確認**
   - エディタ下部にエラーメッセージが表示されていないか確認
   - 赤い下線が表示されている場合は、構文エラーがある可能性があります

**ステップ4: 初回認証の実行（方法B-2: Webアプリとしてデプロイ - 推奨）**

この方法が最も簡単で確実です。

1. **Webアプリとしてデプロイ**
   - Apps Scriptエディタの右上にある「デプロイ」ボタンをクリック
   - 「新しいデプロイ」を選択
   - 「種類の選択」の右側にある「設定」アイコン（歯車）をクリック
   - 「ウェブアプリ」を選択
   - 以下の設定を行います：
     - **説明**: 「OAuth認証用」など、わかりやすい名前を入力
     - **次のユーザーとして実行**: 「自分」を選択
     - **アクセスできるユーザー**: 「自分」を選択
   - 「デプロイ」ボタンをクリック
   - 初回デプロイの場合、認証が求められる場合があります。承認してください

2. **リダイレクトURIの設定**
   - デプロイが完了すると、WebアプリのURLが表示されます
   - このURLをコピーします（例: `https://script.google.com/macros/s/ABC123.../exec`）
   - Google Cloud Consoleに戻ります
   - 「APIとサービス」→「認証情報」を選択
   - 作成したOAuth 2.0 クライアントIDをクリックして編集
   - 「承認済みのリダイレクト URI」セクションで「+ URIを追加」をクリック
   - コピーしたWebアプリのURLを貼り付けます
   - 「保存」をクリック

3. **認証の実行**
   - デプロイされたWebアプリのURLにブラウザでアクセス
   - 「認証を開始」というリンクが表示されるはずです
   - **もし「OK」だけが表示される場合**: 以下のトラブルシューティングを参照してください

   **トラブルシューティング: 「OK」だけが表示される場合**

   この問題は、`doGet()` 関数が正しく動作していない可能性があります。以下の手順で確認・修正してください：

   **方法1: `doGet()` 関数を修正する**

   `doGet()` 関数を以下のように修正してください：

   ```javascript
   /**
    * OAuth2認証のコールバック関数（Webアプリとしてデプロイした場合）
    * @param {Object} e リクエストパラメータ
    * @return {HtmlOutput} HTML出力
    */
   function doGet(e) {
     try {
       if (e.parameter.code) {
         // 認証コードが返ってきた場合
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
       } else {
         // 認証URLを取得して表示
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
           // 認証URLの取得に失敗した場合
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
     } catch (error) {
       return HtmlService.createHtmlOutput(
         '<html><body style="font-family: Arial; padding: 20px;">' +
         '<h2>エラーが発生しました</h2>' +
         '<p>' + error.toString() + '</p>' +
         '<p>スクリプトプロパティに OAUTH_CLIENT_ID が設定されているか確認してください。</p>' +
         '</body></html>'
       );
     }
   }
   ```

   **方法2: 手動で認証URLを取得する（代替方法）**

   Webアプリが正しく動作しない場合は、以下の手順で手動認証を行ってください：

   1. Apps Scriptエディタで、`getOAuthAuthorizationUrl()` 関数を実行
   2. 実行ログに表示された認証URLをコピー
   3. そのURLをブラウザで開く
   4. 認証を完了する
   5. リダイレクト後のURLから `code=` の後の値をコピー
   6. Apps Scriptエディタで、`exchangeAuthorizationCode('コピーした認証コード')` を実行

   **方法3: デプロイを再実行する**

   1. Apps Scriptエディタで「デプロイ」→「デプロイを管理」を選択
   2. 既存のデプロイを削除
   3. 再度「新しいデプロイ」を作成
   4. 設定を確認してデプロイ

   - リンクをクリック
   - Googleアカウントの認証画面が表示されるので、承認を行います
   - 認証が完了すると、「認証が成功しました。このウィンドウを閉じてください。」というメッセージが表示されます
   - ウィンドウを閉じてください

4. **動作確認**
   - Apps Scriptエディタに戻ります
   - `getAppsScriptAPIAccessToken()` 関数を選択して実行します
   - 実行ログにアクセストークン（長い文字列）が表示されれば成功です

**ステップ5: 認証情報の確認**

認証情報は自動的に保存されています。確認するには：

1. Apps Scriptエディタで、「プロジェクトの設定」（⚙️ アイコン）をクリック
2. 下にスクロールして「ユーザー プロパティ」を表示
3. 以下のプロパティが保存されていることを確認：
   - `oauth_access_token` - アクセストークン
   - `oauth_refresh_token` - リフレッシュトークン
   - `oauth_expires_at` - トークンの有効期限

**ステップ6: Apps Script APIを使用する**

認証が完了し、アクセストークンを取得できるようになったら、Apps Script APIを使用してスクリプトプロジェクトのコードを更新できます。

**基本的な使用方法:**

```javascript
/**
 * Apps Script APIを使用してスクリプトプロジェクトのコードを取得
 * @param {string} scriptId - スクリプトプロジェクトID
 * @return {Object} スクリプトプロジェクトの内容
 */
function getScriptContent(scriptId) {
  var accessToken = getAppsScriptAPIAccessToken();
  
  var url = 'https://script.googleapis.com/v1/projects/' + scriptId + '/content';
  var options = {
    'method': 'get',
    'headers': {
      'Authorization': 'Bearer ' + accessToken
    }
  };
  
  var response = UrlFetchApp.fetch(url, options);
  var result = JSON.parse(response.getContentText());
  
  if (result.error) {
    throw new Error('Apps Script API エラー: ' + JSON.stringify(result.error));
  }
  
  return result;
}

/**
 * Apps Script APIを使用してスクリプトプロジェクトのコードを更新
 * @param {string} scriptId - スクリプトプロジェクトID
 * @param {string} fileName - ファイル名（例: 'Code'）
 * @param {string} code - 更新するコード
 * @return {Object} 更新結果
 */
function updateScriptContent(scriptId, fileName, code) {
  var accessToken = getAppsScriptAPIAccessToken();
  
  // まず既存のコードを取得
  var currentContent = getScriptContent(scriptId);
  
  // ファイルを更新
  var files = currentContent.files.map(function(file) {
    if (file.name === fileName) {
      return {
        'name': file.name,
        'type': file.type,
        'source': code
      };
    }
    return file;
  });
  
  // 新しいファイルが存在しない場合は追加
  var fileExists = files.some(function(file) {
    return file.name === fileName;
  });
  
  if (!fileExists) {
    files.push({
      'name': fileName,
      'type': 'SERVER_JS',
      'source': code
    });
  }
  
  var url = 'https://script.googleapis.com/v1/projects/' + scriptId + '/content';
  var options = {
    'method': 'put',
    'headers': {
      'Authorization': 'Bearer ' + accessToken,
      'Content-Type': 'application/json'
    },
    'payload': JSON.stringify({
      'files': files
    })
  };
  
  var response = UrlFetchApp.fetch(url, options);
  var result = JSON.parse(response.getContentText());
  
  if (result.error) {
    throw new Error('Apps Script API エラー: ' + JSON.stringify(result.error));
  }
  
  return result;
}
```

**スクリプトプロジェクトIDの取得:**

スプレッドシートに紐づいたスクリプトプロジェクトのIDを取得する方法：

1. **手動で取得する方法:**
   - スプレッドシートを開く
   - 「拡張機能」→「Apps Script」を選択
   - ブラウザのURLからスクリプトIDを取得
   - URLの形式: `https://script.google.com/home/projects/{SCRIPT_ID}/edit`
   - `{SCRIPT_ID}` の部分がスクリプトプロジェクトIDです

2. **スクリプトプロパティに保存する方法:**
   - スクリプトプロパティに `SPREADSHEET_SCRIPT_ID` として保存
   - または、スプレッドシートIDとスクリプトIDのマッピングを管理

**使用例:**

```javascript
// スクリプトプロジェクトIDを取得（例）
var scriptId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_SCRIPT_ID');

// テンプレート側のコードを取得
var templateCode = getTemplateSideCode_();

// スクリプトプロジェクトのコードを更新
try {
  var result = updateScriptContent(scriptId, 'Code', templateCode);
  Logger.log('コードの更新が完了しました');
} catch (error) {
  Logger.log('エラー: ' + error.toString());
}
```

**注意事項:**
- スクリプトプロジェクトIDは、スプレッドシートごとに異なります
- コードを更新すると、既存のコードが上書きされます
- バックアップを取るか、既存のコードを取得してから更新することを推奨します

#### 代替方法: 方法B-1（手動認証 - Webアプリデプロイ不要）

Webアプリとしてデプロイしたくない場合の手動認証方法です。この方法は少し複雑ですが、デプロイが不要です。

**注意**: この方法では、コード内の `getOAuthAuthorizationUrl()` 関数を少し修正する必要があります。

**方法B-1の実装手順:**

1. **コードの修正**

   `getOAuthAuthorizationUrl()` 関数と `exchangeAuthorizationCode()` 関数を以下のように修正します：

   ```javascript
   /**
    * OAuth2認証URLを取得します（初回認証用 - 手動認証版）
    * @return {string} 認証URL
    */
   function getOAuthAuthorizationUrl() {
     var CLIENT_ID = PropertiesService.getScriptProperties().getProperty('OAUTH_CLIENT_ID');
     // 手動認証用のリダイレクトURI（Google Cloud Consoleで設定する必要がある）
     var redirectUri = 'http://localhost:8080/callback'; // または 'urn:ietf:wg:oauth:2.0:oob'
     
     var authUrl = 'https://accounts.google.com/o/oauth2/auth?' +
       'client_id=' + encodeURIComponent(CLIENT_ID) +
       '&redirect_uri=' + encodeURIComponent(redirectUri) +
       '&response_type=code' +
       '&scope=' + encodeURIComponent('https://www.googleapis.com/auth/script.projects') +
       '&access_type=offline' +
       '&prompt=consent';
     
     Logger.log('以下のURLにアクセスして認証してください: ' + authUrl);
     Logger.log('認証後、リダイレクト先のURLから「code=」の後の値をコピーしてください');
     return authUrl;
   }
   
   /**
    * OAuth2認証コードからアクセストークンを取得します（手動認証版）
    * @param {string} code 認証コード
    * @return {string} アクセストークン
    */
   function exchangeAuthorizationCode(code) {
     var CLIENT_ID = PropertiesService.getScriptProperties().getProperty('OAUTH_CLIENT_ID');
     var CLIENT_SECRET = PropertiesService.getScriptProperties().getProperty('OAUTH_CLIENT_SECRET');
     // getOAuthAuthorizationUrl() で使用したものと同じリダイレクトURIを使用
     var redirectUri = 'http://localhost:8080/callback'; // getOAuthAuthorizationUrl() と同じ値を指定
     
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
   ```

2. **Google Cloud ConsoleでのリダイレクトURI設定**

   - Google Cloud Consoleで、「APIとサービス」→「認証情報」を選択
   - OAuth 2.0 クライアントIDを編集
   - 「承認済みのリダイレクト URI」に以下のいずれかを追加：
     - `http://localhost:8080/callback` （推奨）
     - または `urn:ietf:wg:oauth:2.0:oob` （非推奨、廃止予定）

3. **認証の実行**

   - Apps Scriptエディタで、`getOAuthAuthorizationUrl()` 関数を実行
   - ログに表示されたURLをコピー
   - ブラウザでURLを開く
   - 認証を完了する
   - リダイレクト後、ブラウザのアドレスバーから `code=` の後の値をコピー
   - Apps Scriptエディタで、以下のように実行：
     ```javascript
     exchangeAuthorizationCode('コピーした認証コード');
     ```
   - これで認証が完了します

**注意**: 方法B-1は少し複雑なので、可能であれば方法B-2（Webアプリデプロイ）を推奨します。

#### トラブルシューティング

**問題1: 「認証が必要です」というエラーが表示される**

- ユーザープロパティにトークンが保存されているか確認してください
- 認証を再度実行してください

**問題2: 「リダイレクトURI不一致」というエラー（エラー 400: redirect_uri_mismatch）**

このエラーは、Google Cloud Consoleで設定したリダイレクトURIと、認証リクエストで使用しているリダイレクトURIが一致していない場合に発生します。

**解決方法:**

1. **使用されているリダイレクトURIを確認**
   - Apps Scriptエディタで `getOAuthAuthorizationUrl()` 関数を実行
   - 実行ログに表示された「リダイレクトURI: ...」をコピー
   - 例: `https://script.google.com/macros/s/AKfycbz_tkdZo6sCiZPSUPALMr1vizf73-rVm2zDpNl4dzo/dev`

2. **Google Cloud ConsoleでリダイレクトURIを追加**
   - Google Cloud Consoleにアクセス
   - 「APIとサービス」→「認証情報」を選択
   - OAuth 2.0 クライアントIDをクリックして編集
   - 「承認済みのリダイレクト URI」セクションで「+ URIを追加」をクリック
   - ステップ1でコピーしたリダイレクトURIを貼り付け
   - **重要**: URLは完全に一致している必要があります（末尾の `/dev` や `/exec` も含む）
   - 「保存」をクリック

3. **開発版と本番版の両方を追加（推奨）**
   - 開発版: `https://script.google.com/macros/s/{SCRIPT_ID}/dev`
   - 本番版: `https://script.google.com/macros/s/{SCRIPT_ID}/exec`
   - 両方を追加しておくと、どちらでも動作します

4. **設定後の確認**
   - 数秒待ってから、再度認証URLにアクセス
   - エラーが解消されていることを確認

**注意事項:**
- Webアプリを再デプロイした場合、URLが変わる可能性があるので、リダイレクトURIも更新してください
- URLは完全に一致している必要があります（大文字小文字、スラッシュの有無など）

**問題3: トークンの有効期限切れ**

- トークンは自動的にリフレッシュされますが、問題が発生する場合は、再度認証を実行してください

**問題3-2: 「エラー 403: access_denied」- アプリがGoogleの審査プロセスを完了していません**

このエラーは、OAuth同意画面が「テスト」モードになっていて、認証に使用するGoogleアカウントがテストユーザーとして追加されていない場合に発生します。

**解決方法:**

1. **Google Cloud ConsoleでOAuth同意画面を開く**
   - Google Cloud Consoleにアクセス
   - 「APIとサービス」→「OAuth同意画面」を選択

2. **テストユーザーを追加**
   - 「テストユーザー」セクションを表示
   - 「+ ユーザーを追加」をクリック
   - 認証に使用するGoogleアカウントのメールアドレスを入力（例: `schooliegakuen@gmail.com`）
   - 「追加」をクリック
   - 複数のアカウントを使用する場合は、それぞれ追加してください

3. **設定の確認**
   - テストユーザーが正しく追加されているか確認
   - 数秒待ってから、再度認証URLにアクセス

4. **代替方法: OAuth同意画面を公開する（オプション）**
   - OAuth同意画面を「公開」に変更することもできます
   - ただし、Googleの審査が必要な場合があります
   - テスト環境では、テストユーザーを追加する方法を推奨します

**注意事項:**
- テストユーザーを追加しないと、認証時に「エラー 403: access_denied」が発生します
- テストユーザーは、OAuth同意画面の「テストユーザー」セクションで管理できます

**問題4: 「Script function not found: getAppsScriptAPIAccessToken」エラー**

このエラーは、方法BのコードがApps Scriptプロジェクトに追加されていないことを意味します。

**解決方法:**

1. **コードが追加されているか確認**
   - Apps Scriptエディタで、左側のファイル一覧を確認
   - `oauth.js` または `Code.js` などのファイルを開く
   - `getAppsScriptAPIAccessToken` 関数が存在するか確認（Ctrl+F / Cmd+F で検索）

2. **コードを追加する**
   - ドキュメントの「方法B: OAuth2ライブラリを使わずに実装する場合（推奨）」セクションから、コード全体をコピー
   - Apps Scriptエディタで新しいファイルを作成（左側の「+」をクリック）
   - ファイル名を `oauth.js` に変更（任意）
   - コードを貼り付けて保存（Ctrl+S / Cmd+S）

3. **必要な関数がすべて含まれているか確認**
   以下の5つの関数がすべて含まれていることを確認：
   - `getAppsScriptAPIAccessToken()`
   - `getOAuthAuthorizationUrl()`
   - `exchangeAuthorizationCode(code)`
   - `refreshAccessToken(refreshToken)`
   - `doGet(e)`

4. **保存を確認**
   - コードを保存した後、関数一覧（ドロップダウン）に `getAppsScriptAPIAccessToken` が表示されることを確認
   - 表示されない場合は、ファイルを保存し直してください

**問題5: Webアプリにアクセスしたら「OK」だけが表示される**

`doGet()` 関数が正しく実装されている場合でも、この問題が発生することがあります。以下の手順でデバッグしてください：

1. **実行ログを確認**
   - Apps Scriptエディタで「実行ログ」を開く（表示 → 実行ログ、または Ctrl+Enter / Cmd+Enter）
   - WebアプリのURLにアクセスした直後にログを確認
   - `getOAuthAuthorizationUrl()` 関数が呼び出されたか、エラーが発生していないか確認
   
   **よくあるエラーメッセージと対処法:**
   
   - **「OAUTH_CLIENT_ID がスクリプトプロパティに設定されていません。」**
     → スクリプトプロパティに `OAUTH_CLIENT_ID` を設定してください
   
   - **「リダイレクトURIが取得できません。Webアプリとしてデプロイするか、スクリプトプロパティに OAUTH_REDIRECT_URI を設定してください。」**
     → 以下のいずれかを実行：
       - Webアプリとしてデプロイを完了する
       - または、スクリプトプロパティに `OAUTH_REDIRECT_URI` を追加（値はWebアプリのURL）
   
   - **「認証URL: ...」と「リダイレクトURI: ...」が表示されている場合**
     → 関数は正常に動作しています。ブラウザに表示されない場合は、`doGet()` 関数のHTML出力部分を確認してください

2. **スクリプトプロパティの確認**
   - 「プロジェクトの設定」→「スクリプト プロパティ」で以下を確認：
     - `OAUTH_CLIENT_ID` が設定されているか
     - 値が正しいか（コピー&ペーストの際にスペースが入っていないか）

3. **Webアプリのデプロイ状態を確認**
   - `getOAuthAuthorizationUrl()` 関数内の `ScriptApp.getService().getUrl()` が正しく動作しているか確認
   - デプロイが正しく完了しているか確認
   - WebアプリのURLが正しく取得できているか確認

4. **代替方法: リダイレクトURIを手動設定**
   - スクリプトプロパティに `OAUTH_REDIRECT_URI` を追加
   - 値は、Webアプリとしてデプロイした際に表示されたURL（例: `https://script.google.com/macros/s/.../exec`）
   - これにより、`ScriptApp.getService().getUrl()` が失敗しても、手動設定したURIが使用されます

5. **手動認証方法を使用**
   - 上記の方法で解決しない場合は、手動認証方法（方法B-1）を使用してください
   - または、`getOAuthAuthorizationUrl()` 関数を直接実行して、ログに表示されたURLをブラウザで開いてください

6. **初回認証の実行**

   OAuth2認証を初めて使用する場合は、認証を実行する必要があります：

   **方法Aを使用した場合（OAuth2ライブラリ）:**

   - Apps Scriptエディタで、`getAppsScriptAPIAccessToken()` 関数を選択
   - 関数を実行（▶️ 実行ボタンをクリック）
   - 初回実行時は、認証URLがログ（「実行ログ」）に出力されます
   - ログに表示されたURLをコピーして、新しいタブで開きます
   - Googleアカウントの認証画面が表示されるので、承認を行います
   - 認証が完了すると、「認証が成功しました」というメッセージが表示されます
   - そのウィンドウを閉じて、Apps Scriptエディタに戻ります
   - 再度 `getAppsScriptAPIAccessToken()` 関数を実行すると、アクセストークンが返されます

   **方法Bを使用した場合（ライブラリなし）:**

   以下の2つの方法があります：

   **方法B-1: 手動で認証URLにアクセスする方法**
   - Apps Scriptエディタで、`getOAuthAuthorizationUrl()` 関数を実行
   - ログ（「実行ログ」）に認証URLが表示されます
   - そのURLをコピーして、ブラウザで開きます
   - Googleアカウントの認証画面が表示されるので、承認を行います
   - 認証が完了すると、リダイレクトされますが、エラーが表示される場合があります（これは正常です）
   - ブラウザのアドレスバーに表示されるURLから、`code=` の後の値（認証コード）をコピーします
   - Apps Scriptエディタで、以下のように実行します：
     ```javascript
     exchangeAuthorizationCode('コピーした認証コード');
     ```
   - これで認証が完了し、アクセストークンが保存されます

   **方法B-2: Webアプリとしてデプロイする方法（推奨）**
   - Apps Scriptエディタで、「デプロイ」→「新しいデプロイ」を選択
   - 種類の選択で「ウェブアプリ」を選択
   - 説明を入力（例：「OAuth認証用」）
   - 「次のユーザーとして実行」で「自分」を選択
   - 「アクセスできるユーザー」で「自分」を選択
   - 「デプロイ」をクリック
   - 表示されたWebアプリのURLにアクセス
   - 「認証を開始」リンクをクリックして認証を完了

   **共通の注意事項:**
   - 認証情報はユーザープロパティに保存されるため、一度認証すれば次回以降は自動的に使用されます
   - アクセストークンの有効期限が切れた場合は、リフレッシュトークンが自動的に使用されて更新されます（方法Bの場合）

**注意事項**:
- OAuth2クライアントIDとシークレットは機密情報です。スクリプトプロパティに保存し、コードに直接記述しないでください
- テスト環境では、OAuth同意画面の「テストユーザー」として自分のアカウントを追加する必要があります
- 本番環境で使用する場合は、OAuth同意画面を公開する必要があります（審査が必要な場合があります）

## テンプレートファイルのコード更新手順（推奨）

現在の推奨方法は、テンプレートファイルに手動でコードを配置する方法です。

### 1. テンプレート側のコードを取得

`template-side-code.js` ファイルを開いて、内容を確認します。

### 2. テンプレートファイルのApps Scriptエディタを開く

1. Google Driveでテンプレートファイルを開く
2. 「拡張機能」→「Apps Script」を選択
3. 既存のコードをすべて削除

### 3. 新しいコードを貼り付ける

1. `template-side-code.js` の内容をすべてコピー
2. テンプレートファイルのApps Scriptエディタに貼り付け
3. 「保存」をクリック（Ctrl+S / Cmd+S）

### 4. 確認

- コードが正しく保存されていることを確認
- エラーがないことを確認

## コードの同期について

**重要な注意事項**:

- マスター側のコード（`Code.js` の984行目以降）を更新した場合は、必ず `template-side-code.js` も更新してください
- `template-side-code.js` を更新した後は、テンプレートファイルにも反映してください
- これにより、新しくコピーされる講師用シートには常に最新のコードが含まれます

## 実装されている関数

### `getTemplateSideCode_()`

テンプレート側のコードを文字列として取得します。`drive.js` に実装されています。

### `deployTemplateSideCodeToSpreadsheet_(spreadsheetId)`

スプレッドシートにテンプレート側のコードを展開する関数です。現在はプレースホルダーとして実装されており、実際の展開処理は実装されていません。

将来的に、Apps Script APIを使用してコードを展開する機能を実装する場合は、この関数を拡張します。

## 既存のコピー済みシートについて

既にコピーされた講師用シートには、コピー時のコードが含まれています。
これらのシートのコードを更新するには：

1. 各シートを個別に開いて、Apps Scriptエディタでコードを更新する
2. または、新しいテンプレートから再コピーする（推奨されない）

## 参考資料

- [Apps Script API](https://developers.google.com/apps-script/api)
- [OAuth2認証](https://developers.google.com/identity/protocols/oauth2)
- [Google Apps Script API ガイド](https://developers.google.com/apps-script/api/guides/v1/scripts/update)
