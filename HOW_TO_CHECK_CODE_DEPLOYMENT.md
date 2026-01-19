# コード展開機能の動作確認方法

## なぜ実行ログに `deployTemplateSideCodeToSpreadsheet_` が表示されないのか？

`deployTemplateSideCodeToSpreadsheet_` 関数は、`copyTemplateSpreadsheet_` 関数の**内部**で呼び出されるため、実行ログの一覧には直接表示されません。

実行ログには、**直接実行された関数**（例：`onFormSubmit`、`onOpen`）のみが表示されます。

## 確認方法

### 方法1: `onFormSubmit` の実行ログの詳細を確認

1. **実行ログ画面で `onFormSubmit` をクリック**
   - 実行ログの一覧で、`onFormSubmit` の行をクリック
   - 例：`2025/12/28 15:54:44` の `onFormSubmit` をクリック

2. **「ログ」タブを確認**
   - 詳細画面が開いたら、「ログ」タブをクリック
   - 以下のようなメッセージが表示されていれば、コード展開機能が動作しています：
     ```
     [copyTemplateSpreadsheet_] Successfully deployed template code to: 1tWdPgLsZjxaP8e3OVIRVOOV-pbz7wDQbFuO3hBhWgWw
     ```
   - または、以下のようなメッセージが表示されている場合、コード展開はスキップされています：
     ```
     [copyTemplateSpreadsheet_] Template code deployment skipped (script ID not found or auth not available)
     [copyTemplateSpreadsheet_] Using template file code instead (makeCopy includes script code)
     ```

### 方法2: 実行ログの「ログ」タブで直接確認

1. **実行ログ画面の「ログ」タブをクリック**
   - 実行ログ画面の上部に「ログ」タブがある場合、それをクリック

2. **ログメッセージを検索**
   - `[copyTemplateSpreadsheet_]` という文字列で検索
   - または、`deployTemplateSideCodeToSpreadsheet_` という文字列で検索

### 方法3: Apps Scriptエディタの実行ログで確認

1. **Apps Scriptエディタを開く**
   - マスタースプレッドシートを開く
   - 「拡張機能」→「Apps Script」を選択

2. **実行ログを確認**
   - エディタ下部の「実行ログ」タブをクリック
   - 最新の実行ログを確認
   - `[copyTemplateSpreadsheet_]` で始まるメッセージを探す

## コード展開機能が動作していない場合

### 確認事項1: スクリプトプロジェクトIDが保存されているか

1. **スクリプトプロパティを確認**
   - Apps Scriptエディタの左側メニューから「プロジェクトの設定」（⚙️ アイコン）をクリック
   - 下にスクロールして「スクリプト プロパティ」セクションを表示
   - `SCRIPT_ID_1uVWYwQkr4zQ5UMwGCNL7nNkUvRwLA5v-IK2NlO1Ulyk` というプロパティが存在するか確認

2. **存在しない場合**
   - スクリプトプロジェクトIDを保存する手順を再度実行してください

### 確認事項2: OAuth2認証が完了しているか

1. **認証状態を確認**
   - Apps Scriptエディタで、以下の関数を実行：
     ```javascript
     function testAuth() {
       try {
         const token = getAppsScriptAPIAccessToken();
         Logger.log('認証成功: アクセストークンが取得できました');
         return true;
       } catch (e) {
         Logger.log('認証エラー: ' + e.toString());
         Logger.log('認証が必要です。getOAuthAuthorizationUrl() を実行して認証URLを取得してください。');
         return false;
       }
     }
     ```

2. **認証が必要な場合**
   - `getOAuthAuthorizationUrl()` を実行して認証URLを取得
   - 認証URLにアクセスして認証を完了

### 確認事項3: コピーされたシートのスクリプトプロジェクトIDが保存されているか

**重要**: テンプレートファイルのスクリプトプロジェクトIDを保存しましたが、**コピーされたシート**のスクリプトプロジェクトIDは、コピーごとに異なります。

現在の実装では、コピーされたシートのスクリプトプロジェクトIDを自動的に取得することはできません。そのため、以下のいずれかの方法を使用する必要があります：

#### 方法A: テンプレートファイルのスクリプトプロジェクトIDを使用（推奨）

テンプレートファイルのスクリプトプロジェクトIDを、コピーされたシートにも適用するようにコードを修正します。

#### 方法B: コピーされたシートのスクリプトプロジェクトIDを手動で保存

1. **コピーされたシートを開く**
   - 例：`2026-02_森永 英敬_シフト提出`

2. **スクリプトプロジェクトIDを取得**
   - 「拡張機能」→「Apps Script」を選択
   - URLからスクリプトプロジェクトIDを取得

3. **保存**
   - マスター側のApps Scriptエディタで以下を実行：
     ```javascript
     saveScriptIdForSpreadsheet_(
       'コピーされたシートのスプレッドシートID',
       'コピーされたシートのスクリプトプロジェクトID'
     );
     ```

## 現在の動作状況

実行ログを見ると、`onFormSubmit` は正常に実行されていますが、`deployTemplateSideCodeToSpreadsheet_` の実行ログが表示されていない可能性があります。

これは、以下のいずれかの理由によるものです：

1. **スクリプトプロジェクトIDが取得できていない**
   - コピーされたシートのスクリプトプロジェクトIDが保存されていない
   - `getScriptIdForSpreadsheet_()` が `null` を返している

2. **OAuth2認証が完了していない**
   - `getAppsScriptAPIAccessToken()` がエラーを返している

3. **コード展開がスキップされている**
   - 上記の理由により、コード展開がスキップされ、テンプレートファイルのコードが使用されている

## 確認手順（まとめ）

1. **実行ログの詳細を確認**
   - `onFormSubmit` の実行ログをクリック
   - 「ログ」タブで `[copyTemplateSpreadsheet_]` のメッセージを確認

2. **スクリプトプロパティを確認**
   - テンプレートファイルのスクリプトプロジェクトIDが保存されているか確認

3. **OAuth2認証を確認**
   - `getAppsScriptAPIAccessToken()` が正常に動作するか確認

4. **コピーされたシートのスクリプトプロジェクトIDを確認**
   - コピーされたシートのスクリプトプロジェクトIDが保存されているか確認

## 補足: コード展開機能が動作しなくても問題ない理由

コード展開機能が動作しなくても、システムは正常に動作します。理由：

1. **テンプレートファイルにコードが含まれている**
   - テンプレートファイルをコピーする際、コードも自動的にコピーされます

2. **`makeCopy()` の動作**
   - `makeCopy()` を使用すると、スプレッドシートと一緒にApps Scriptコードもコピーされます

3. **フォールバック機能**
   - コード展開が失敗した場合でも、テンプレートファイルのコードが使用されます

コード展開機能は、**テンプレートファイルを更新しなくても、最新のコードを展開できる**という利便性を提供する機能です。必須ではありません。


