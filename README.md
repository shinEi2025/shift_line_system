# シフト提出管理システム

Google Apps Script (GAS) と LINE Bot を使用した講師のシフト提出管理システムです。

## 📋 概要

このシステムは以下の機能を提供します：

- **LINE Bot連携**: 講師がLINE Botに氏名を送信して登録
- **自動シート作成**: Googleフォーム送信時に講師用シフト提出シートを自動作成
- **提出管理**: 講師がシフトを入力・提出すると自動的にロック
- **リマインド機能**: 未提出の講師に自動リマインド
- **管理者機能**: 管理者が提出済みシートのロックを解除して再編集可能にする

## 📁 ファイル構成

```
my-gas-project/
├── Code.js              # メインロジック（LINE webhook、フォーム送信処理、テンプレート側関数）
├── drive.js             # Drive操作（フォルダ作成、シートコピー、権限管理、ロック/アンロック）
├── line.js              # LINE API操作（返信、プッシュ通知）
├── submissions.js       # 提出状況の監視とリマインド
├── utils.js             # ユーティリティ関数（名前正規化、講師検索、エラーハンドリング）
├── appsscript.json      # GASプロジェクト設定
├── FLOW.md              # 詳細なフロー図（必読）
└── README.md            # このファイル
```

## 🚀 セットアップ

### 1. 必要なスクリプトプロパティ

Google Apps Scriptのスクリプトプロパティに以下を設定：

- `LINE_CHANNEL_ACCESS_TOKEN`: LINE Botのチャネルアクセストークン
- `ADMIN_LINE_USER_ID`: 管理者のLINE User ID（任意、例外通知用）

### 2. マスタースプレッドシートの準備

以下のシートが必要：

- **Teachersシート**: 講師情報
  - 必須列: `氏名`, `lineUserId`
  - 推奨列: `teacherId`, `メール`, `lineLinkedAt`

- **Submissionsシート**: 提出状況
  - 必須列: `timestamp`, `monthKey`, `teacherId`, `氏名`, `sheetUrl`, `status`, `submittedAt`
  - 推奨列: `lockedAt`, `ackNotifiedAt`, `reminderNotifiedAt`, `submissionKey`

### 3. テンプレートスプレッドシートの準備

講師用シフト提出シートのテンプレートを用意：

- **Inputシート**: シフト入力用
  - `B2`: 提出状態表示（未提出/提出済）
  - `C2`: 提出チェックボックス（TRUEで提出）
  - `D2`: ヘルプテキスト
  - `G3`: 講師名表示

- **_METAシート**: メタ情報（自動生成、非表示）

### 4. トリガーの設定

#### 必須トリガー
- **onFormSubmit**: Googleフォーム送信時（インストール型トリガー推奨）

#### 推奨トリガー（定期実行）
- **pollSubmissionsAndUpdate**: 提出状況の監視（例：5分ごと）
- **remindUnsubmitted**: 未提出リマインド（例：毎日9時）

### 5. LINE Webhookの設定

LINE Developers ConsoleでWebhook URLを設定：
```
https://script.google.com/macros/s/{SCRIPT_ID}/exec
```

## 📖 使用方法

### 講師の登録

1. 講師がLINE Botに氏名を送信
2. システムがTeachersシートで検索
3. 一致すればLINE User IDを登録
4. 「登録OK」メッセージを返信

### シフト提出依頼

1. 管理者がGoogleフォームを送信
   - 氏名を入力
   - 提出月を入力（省略可、次月が自動設定）
2. システムが自動的に：
   - 講師用シートを作成
   - 講師に編集権限を付与
   - LINE通知を送信（登録済みの場合）

### 講師による提出

1. 講師がシートを開いてシフトを入力
2. `Input!C2`にチェック（☑）を入れる
3. シートが自動的にロック（編集不可）
4. 提出受理通知がLINEに送信

### 管理者によるロック解除

1. 管理者がLINE Botにコマンドを送信：
   ```
   変更依頼: {講師名} {月}
   ```
   または
   ```
   変更依頼: {講師名}
   ```
2. システムが自動的に：
   - シートのロックを解除
   - 講師に編集権限を再付与
   - 講師にLINE通知を送信

## 📊 詳細なフロー

詳細なフロー図は [FLOW.md](./FLOW.md) を参照してください。

## 🔧 開発

### clasp の使用

```bash
# ファイルをプッシュ
clasp push

# 変更を監視して自動プッシュ
clasp push --watch

# クラウドから最新を取得
clasp pull

# 関数を直接実行
clasp run 関数名

# ログを確認
clasp logs
clasp logs --watch
```

## ⚠️ 注意事項

- LINE登録は上書き防止機能あり（既に別のLINE IDが登録されている場合、上書き不可）
- 提出後のシートは自動的にロックされ、講師は編集不可
- 管理者（スクリプト実行者）は常に編集可能
- エラー発生時は自動的に管理者にLINE通知（ADMIN_LINE_USER_IDが設定されている場合）

## 📝 ライセンス

このプロジェクトは内部使用を目的としています。

