# スクリプトプロジェクトID保存の詳細手順

## ステップ3: スクリプトプロパティに保存（超詳細版）

このステップでは、取得したスクリプトプロジェクトIDを保存します。

---

## 手順1: マスター側のApps Scriptエディタを開く

### 1-1. マスタースプレッドシートを開く

1. **Google Driveを開く**
   - ブラウザで https://drive.google.com にアクセス
   - Googleアカウントでログイン

2. **マスタースプレッドシートを探す**
   - マスタースプレッドシートのID: `1mhBpPhuL6Aq-YRXgmCu1kMtg0g3JO7h37pHdWmM8sqE`
   - 検索ボックスにこのIDを入力して検索
   - または、直接URLを開く: `https://docs.google.com/spreadsheets/d/1mhBpPhuL6Aq-YRXgmCu1kMtg0g3JO7h37pHdWmM8sqE/edit`

3. **スプレッドシートを開く**
   - スプレッドシートが開いたら、上部にメニューバーが表示されます

### 1-2. Apps Scriptエディタを開く

1. **メニューバーを確認**
   - スプレッドシートの上部に以下のメニューが表示されます：
     ```
     ファイル  編集  表示  挿入  書式  データ  ツール  拡張機能  ヘルプ
     ```

2. **「拡張機能」メニューをクリック**
   - メニューバーの「拡張機能」をクリック
   - ドロップダウンメニューが表示されます

3. **「Apps Script」を選択**
   - ドロップダウンメニューから「Apps Script」をクリック
   - 新しいタブでApps Scriptエディタが開きます

---

## 手順2: Apps Scriptエディタの画面を確認

Apps Scriptエディタが開いたら、以下のような画面が表示されます：

```
┌─────────────────────────────────────────────────────────┐
│  [保存] [実行] [デプロイ]  [関数を選択 ▼]  [ログ]      │
├─────────────────────────────────────────────────────────┤
│                                                          │
│  左側: ファイル一覧                                      │
│  ├─ Code.gs                                             │
│  ├─ drive.js                                            │
│  ├─ oauth.js                                            │
│  └─ ...                                                 │
│                                                          │
│  中央: コードエディタ                                    │
│  （ここにコードが表示されます）                          │
│                                                          │
│  下部: 実行ログ                                         │
│  （実行結果が表示されます）                              │
│                                                          │
└─────────────────────────────────────────────────────────┘
```

### 画面の説明

- **上部ツールバー**: 保存、実行、デプロイなどのボタン
- **左側パネル**: ファイル一覧（Code.gs、drive.js、oauth.jsなど）
- **中央パネル**: コードエディタ（コードを編集する場所）
- **下部パネル**: 実行ログ（実行結果が表示される場所）

---

## 手順3: コードを追加する

### 3-1. ファイルを選択

1. **左側のファイル一覧を確認**
   - `drive.js` というファイルがあるはずです
   - このファイルをクリックして開きます

2. **ファイルが開かない場合**
   - 左側の「+」ボタンをクリックして新しいファイルを作成
   - ファイル名を `test.js` などに変更

### 3-2. コードを追加

**重要**: 以下の2つの方法のうち、どちらか一方を選択してください。

#### 方法A: drive.jsファイルに追加（推奨）

1. **drive.jsファイルを開く**
   - 左側のファイル一覧から `drive.js` をクリックして開く
   - ファイルが表示されない場合は、`drive.gs` という名前かもしれません

2. **コードエディタの下部にスクロール**
   - ファイルの最後（最後の `}` の後）に移動

3. **以下のコードをコピー＆ペースト**
   ```javascript
   /**
    * スクリプトプロジェクトIDを保存する関数（一時的な実行用）
    */
   function saveScriptIdForTemplate() {
     // テンプレートファイルのスプレッドシートIDとスクリプトプロジェクトIDを指定
     // ステップ2で取得したスクリプトプロジェクトIDをここに貼り付けます
     const templateSpreadsheetId = '1uVWYwQkr4zQ5UMwGCNL7nNkUvRwLA5v-IK2NlO1Ulyk'; // テンプレートファイルのスプレッドシートID
     const scriptProjectId = 'ここにステップ2で取得したスクリプトプロジェクトIDを貼り付け'; // 例: '1a2b3c4d5e6f7g8h9i0j1k2l3m4n5o6p'
     
     // 保存を実行
     saveScriptIdForSpreadsheet_(templateSpreadsheetId, scriptProjectId);
     
     // 結果をログに出力
     Logger.log('スクリプトプロジェクトIDを保存しました:');
     Logger.log('スプレッドシートID: ' + templateSpreadsheetId);
     Logger.log('スクリプトプロジェクトID: ' + scriptProjectId);
   }
   ```

4. **スクリプトプロジェクトIDを貼り付け**
   - ステップ2で取得したスクリプトプロジェクトIDをコピー
   - コード内の `'ここにステップ2で取得したスクリプトプロジェクトIDを貼り付け'` の部分を選択
   - 実際のスクリプトプロジェクトIDに置き換えます
   - 例: `const scriptProjectId = '1a2b3c4d5e6f7g8h9i0j1k2l3m4n5o6p';`

#### 方法B: 新しいファイルを作成（エラーが発生する場合）

もし `saveScriptIdForSpreadsheet_ is not defined` というエラーが発生する場合は、以下の方法を試してください：

1. **新しいファイルを作成**
   - 左側の「+」ボタンをクリック
   - ファイル名を `saveScriptId.gs` に変更

2. **以下のコード全体をコピー＆ペースト**
   ```javascript
   /**
    * スクリプトプロジェクトIDをスクリプトプロパティに保存
    */
   function saveScriptIdForSpreadsheet_(spreadsheetId, scriptId) {
     try {
       const props = PropertiesService.getScriptProperties();
       const scriptIdKey = 'SCRIPT_ID_' + spreadsheetId;
       props.setProperty(scriptIdKey, scriptId);
       Logger.log('[saveScriptIdForSpreadsheet_] Saved script ID for spreadsheet: ' + spreadsheetId + ' -> ' + scriptId);
     } catch (err) {
       Logger.log('saveScriptIdForSpreadsheet_ error: ' + err.toString());
     }
   }
   
   /**
    * スクリプトプロジェクトIDを保存する関数（一時的な実行用）
    */
   function saveScriptIdForTemplate() {
     // テンプレートファイルのスプレッドシートIDとスクリプトプロジェクトIDを指定
     // ステップ2で取得したスクリプトプロジェクトIDをここに貼り付けます
     const templateSpreadsheetId = '1uVWYwQkr4zQ5UMwGCNL7nNkUvRwLA5v-IK2NlO1Ulyk'; // テンプレートファイルのスプレッドシートID
     const scriptProjectId = 'ここにステップ2で取得したスクリプトプロジェクトIDを貼り付け'; // 例: '1a2b3c4d5e6f7g8h9i0j1k2l3m4n5o6p'
     
     // 保存を実行
     saveScriptIdForSpreadsheet_(templateSpreadsheetId, scriptProjectId);
     
     // 結果をログに出力
     Logger.log('スクリプトプロジェクトIDを保存しました:');
     Logger.log('スプレッドシートID: ' + templateSpreadsheetId);
     Logger.log('スクリプトプロジェクトID: ' + scriptProjectId);
   }
   ```

3. **スクリプトプロジェクトIDを貼り付け**
   - ステップ2で取得したスクリプトプロジェクトIDをコピー
   - コード内の `'ここにステップ2で取得したスクリプトプロジェクトIDを貼り付け'` の部分を選択
   - 実際のスクリプトプロジェクトIDに置き換えます
   - 例: `const scriptProjectId = '1a2b3c4d5e6f7g8h9i0j1k2l3m4n5o6p';`

### 3-3. 保存

1. **保存ボタンをクリック**
   - エディタ上部の「保存」ボタン（💾 アイコン）をクリック
   - または、キーボードショートカット: `Ctrl+S`（Windows）または `Cmd+S`（Mac）

2. **保存確認**
   - 保存が完了すると、エディタ上部に「保存済み」と表示されます

---

## 手順4: 関数を実行する

### 4-1. 関数を選択

1. **「関数を選択」ドロップダウンを確認**
   - エディタ上部の「関数を選択」というドロップダウンをクリック
   - ドロップダウンが開き、関数の一覧が表示されます

2. **関数を選択**
   - 一覧から `saveScriptIdForTemplate` を選択
   - 選択すると、関数がハイライト表示されます

### 4-2. 実行ボタンをクリック

1. **「実行」ボタンを確認**
   - エディタ上部の「実行」ボタン（▶️ アイコン）をクリック
   - または、キーボードショートカット: `Ctrl+Enter`（Windows）または `Cmd+Enter`（Mac）

2. **認証（初回のみ）**
   - 初回実行時は、認証が求められる場合があります
   - 画面中央に「承認が必要です」というダイアログが表示されます

### 4-3. 認証手順（初回のみ）

1. **「権限を確認」をクリック**
   - ダイアログ内の「権限を確認」ボタンをクリック

2. **Googleアカウントを選択**
   - Googleアカウントの選択画面が表示されます
   - 使用するGoogleアカウントを選択

3. **警告画面が表示される場合**
   - 「Googleはこのアプリを確認していません」という警告が表示される場合があります
   - 「詳細」をクリック
   - 「（プロジェクト名）に移動（安全ではないページ）」をクリック

4. **権限を許可**
   - 「許可」ボタンをクリック
   - 認証が完了すると、Apps Scriptエディタに戻ります

---

## 手順5: 実行結果を確認する

### 5-1. 実行ログを確認

1. **「実行ログ」タブをクリック**
   - エディタ下部の「実行ログ」タブをクリック
   - 実行結果が表示されます

2. **成功メッセージを確認**
   - 以下のようなメッセージが表示されれば成功：
     ```
     スクリプトプロジェクトIDを保存しました:
     スプレッドシートID: 1uVWYwQkr4zQ5UMwGCNL7nNkUvRwLA5v-IK2NlO1Ulyk
     スクリプトプロジェクトID: 1a2b3c4d5e6f7g8h9i0j1k2l3m4n5o6p
     ```

3. **エラーメッセージが表示される場合**
   - エラーメッセージが表示された場合は、以下を確認：
     - スクリプトプロジェクトIDが正しく貼り付けられているか
     - コードの構文エラーがないか（引用符が正しく閉じられているかなど）

### 5-2. スクリプトプロパティを確認（オプション）

1. **「プロジェクトの設定」を開く**
   - エディタ左側のメニューから「プロジェクトの設定」（⚙️ アイコン）をクリック

2. **「スクリプト プロパティ」を確認**
   - 下にスクロールして「スクリプト プロパティ」セクションを表示
   - `SCRIPT_ID_1uVWYwQkr4zQ5UMwGCNL7nNkUvRwLA5v-IK2NlO1Ulyk` というプロパティが表示されていれば成功
   - 値には、保存したスクリプトプロジェクトIDが表示されます

---

## 手順6: 一時的な関数を削除（オプション）

保存が完了したら、一時的に追加した `saveScriptIdForTemplate` 関数は削除しても構いません。

1. **関数を選択**
   - コードエディタで `saveScriptIdForTemplate` 関数全体を選択

2. **削除**
   - `Delete` キーまたは `Backspace` キーを押して削除

3. **保存**
   - `Ctrl+S`（Windows）または `Cmd+S`（Mac）で保存

---

## トラブルシューティング

### 問題1: 「ReferenceError: saveScriptIdForSpreadsheet_ is not defined」エラー

**原因:**
- `saveScriptIdForSpreadsheet_` 関数が `drive.js` ファイルに定義されていない
- または、`drive.js` ファイルがプロジェクトに含まれていない

**解決方法:**
1. **方法Bを使用する**（上記の「方法B: 新しいファイルを作成」を参照）
   - 新しいファイル `saveScriptId.gs` を作成
   - `saveScriptIdForSpreadsheet_` 関数と `saveScriptIdForTemplate` 関数の両方を追加

2. **または、drive.jsファイルを確認**
   - 左側のファイル一覧に `drive.js` または `drive.gs` があるか確認
   - ファイルを開いて、`saveScriptIdForSpreadsheet_` 関数が定義されているか確認
   - 定義されていない場合は、`drive.js` ファイルの最後に以下のコードを追加：
     ```javascript
     function saveScriptIdForSpreadsheet_(spreadsheetId, scriptId) {
       try {
         const props = PropertiesService.getScriptProperties();
         const scriptIdKey = 'SCRIPT_ID_' + spreadsheetId;
         props.setProperty(scriptIdKey, scriptId);
         Logger.log('[saveScriptIdForSpreadsheet_] Saved script ID for spreadsheet: ' + spreadsheetId + ' -> ' + scriptId);
       } catch (err) {
         Logger.log('saveScriptIdForSpreadsheet_ error: ' + err.toString());
       }
     }
     ```

### 問題2: 「関数を選択」ドロップダウンに `saveScriptIdForTemplate` が表示されない

**解決方法:**
- コードが正しく保存されているか確認
- 関数名のスペルミスがないか確認
- エディタをリロード（ページを再読み込み）

### 問題3: 実行ボタンをクリックしても何も起こらない

**解決方法:**
- 関数が正しく選択されているか確認
- コードの構文エラーがないか確認（エディタ下部にエラーメッセージが表示される場合があります）
- ブラウザのコンソールを確認（F12キーで開発者ツールを開く）

### 問題3: 認証エラーが発生する

**解決方法:**
- Googleアカウントが正しく選択されているか確認
- ブラウザのポップアップブロッカーが有効になっていないか確認
- 別のブラウザで試す

### 問題4: 実行ログにエラーメッセージが表示される

**解決方法:**
- エラーメッセージの内容を確認
- スクリプトプロジェクトIDが正しく貼り付けられているか確認
- スプレッドシートIDが正しいか確認

---

## 完了

これで、スクリプトプロジェクトIDの保存が完了しました。

次回から、テンプレートファイルをコピーする際に、自動的にコードが展開されるようになります。

