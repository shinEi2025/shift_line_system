# シフト提出システム - フロー図

## システム概要

このシステムは、Google Apps Script (GAS) と LINE Bot を使用して、講師のシフト提出を管理するシステムです。

## 主要コンポーネント

- **マスタースプレッドシート**: 講師情報と提出状況を管理
- **テンプレートスプレッドシート**: 講師用シフト提出シートのテンプレート
- **LINE Bot**: 講師との連絡と通知
- **Googleフォーム**: シフト提出依頼のトリガー

---

## フロー図

### 1. LINE登録フロー（初回のみ）

```
講師がLINE Botに氏名を送信
    ↓
doPost() が受信
    ↓
Teachersシートで氏名を検索
    ↓
┌─────────────────────────┐
│ 検索結果による分岐        │
└─────────────────────────┘
    │
    ├─ 見つからない
    │   → 「名簿に一致する氏名が見つかりませんでした」
    │
    ├─ 複数見つかる
    │   → 「同じ氏名が複数います（候補：...）」
    │
    ├─ 既に別のLINE IDが登録済み
    │   → 「この氏名は別のLINEと紐付いています」
    │
    ├─ 既に同じLINE IDが登録済み
    │   → 何も返さない（二重返信防止）
    │
    └─ 新規登録
        → lineUserId を Teachers に記録
        → 「登録OK：{氏名} さん\n今後はこのLINEでシフト連絡します。」
```

**関数**: `doPost()` → `linkLineUserByName_()`

---

### 2. フォーム送信フロー（シート作成）

```
管理者がGoogleフォームを送信
（氏名、提出月を入力）
    ↓
onFormSubmit() がトリガー
    ↓
Teachersシートで講師情報を検索
    ↓
┌─────────────────────────┐
│ 講師が見つからない場合    │
└─────────────────────────┘
    → Submissionsに status='teacher_not_found' で記録
    → 処理終了
    ↓
┌─────────────────────────┐
│ 講師が見つかった場合      │
└─────────────────────────┘
    ↓
1. 月フォルダを作成/取得
   （例：2026-01）
    ↓
2. テンプレートから講師用シートをコピー
   ファイル名: {月}_{氏名}_シフト提出
    ↓
3. 講師に編集権限を付与
   （Teachersシートのメールアドレス）
    ↓
4. Submissionsに記録
   - timestamp
   - monthKey
   - teacherId
   - 氏名
   - sheetUrl
   - status='created'
   - submissionKey
    ↓
5. 講師用シートの _META シートに情報を書き込み
   - MASTER_SPREADSHEET_ID
   - SUBMISSIONS_SHEET_NAME
   - SUBMISSION_KEY
   - MONTH_KEY
   - TEACHER_ID
   - TEACHER_NAME
    ↓
6. LINE通知（lineUserIdが登録済みの場合）
   「【シフト提出URL（{月}）】\n{URL}\n\n入力後、☑（提出）を入れてください。」
```

**関数**: `onFormSubmit()` → `findTeacherByName_()`, `ensureMonthFolder_()`, `copyTemplateSpreadsheet_()`, `grantEditPermission_()`, `appendSubmission_()`, `writeMetaToTeacherSheet_()`, `pushLine_()`

---

### 3. 講師による提出フロー

```
講師がシートを開く
    ↓
onOpen() が実行
    ↓
1. パネルを確保（未提出/提出済の表示）
2. _METAから講師名を取得して表示
3. ロック解除状態をチェック
   （管理者がロック解除した場合、自動的に編集可能に）
    ↓
講師がシフトを入力
    ↓
講師が Input!C2 にチェック（☑）を入れる
    ↓
onEdit() がトリガー
    ↓
1. Input!B2 を「提出済」に変更
    ↓
2. lockSheetAfterSubmission_() を実行
   - すべてのシートを保護
   - 講師の編集権限を削除（閲覧のみに変更）
   - 管理者（スクリプト実行者）は編集可能のまま
    ↓
3. トースト通知
   「提出済にしました。シートは編集不可になりました。」
```

**関数**: `onOpen()`, `onEdit()` → `lockSheetAfterSubmission_()`

---

### 4. 提出状況の監視フロー（定期実行）

```
pollSubmissionsAndUpdate() が定期実行
（時間ベーストリガーで設定）
    ↓
Submissionsシートを確認
    ↓
status != 'submitted' の行をチェック
    ↓
各講師シートの Input!C2 を確認
    ↓
┌─────────────────────────┐
│ C2 が TRUE の場合        │
└─────────────────────────┘
    ↓
1. Submissionsを更新
   - status = 'submitted'
   - submittedAt = 現在時刻
    ↓
2. 講師シートの Input!B2 を「提出済」に更新
    ↓
3. シートをロック
   - lockTeacherSheet_() を実行
   - lockedAt を記録
    ↓
4. 提出受理LINE通知（1回だけ）
   - ackNotifiedAt で制御
   - 「【提出受理】\n{氏名}さん（{月}）のシフト提出を受け付けました。ありがとうございます！」
```

**関数**: `pollSubmissionsAndUpdate()` → `readTeacherSubmittedFlag_()`, `lockTeacherSheet_()`, `pushLine_()`

---

### 5. 未提出リマインドフロー（定期実行）

```
remindUnsubmitted() が定期実行
（1日1回、時間ベーストリガーで設定）
    ↓
Submissionsシートを確認
    ↓
最新の monthKey を取得
    ↓
status != 'submitted' の行をチェック
    ↓
reminderNotifiedAt が今日でない場合
    ↓
LINE通知を送信
「【シフト未提出リマインド】\n{氏名}さん（{月}）の提出がまだのようです。\nこちらから入力・提出（☑）をお願いします。\n{URL}」
    ↓
reminderNotifiedAt を更新
```

**関数**: `remindUnsubmitted()` → `pushLine_()`

---

### 6. 管理者によるロック解除フロー

```
管理者がLINE Botにコマンドを送信
「変更依頼: {講師名} {月}」
または
「変更依頼: {講師名}」
    ↓
doPost() が受信
    ↓
管理者のLINE User IDを確認
    ↓
handleAdminUnlockCommand_() を実行
    ↓
1. Submissionsシートで該当する提出を検索
   - 氏名が一致
   - 月が一致（指定されている場合）
   - status = 'submitted'
    ↓
2. 講師情報を取得
   - teacherId または 氏名から
   - メールアドレスとLINE User IDを取得
    ↓
3. unlockTeacherSheet_() を実行
   - すべての保護を解除
   - 講師に編集権限を再付与
   - スクリプト実行者（管理者）がエディターであることを確認
    ↓
4. Submissionsの lockedAt をクリア
    ↓
5. 講師にLINE通知
   「【シフト変更依頼】\n{氏名}さん（{月}）のシフトを変更していただくようお願いします。\nシートの編集が可能になりました。\n{URL}」
    ↓
6. 管理者に返信
   「ロック解除しました：{氏名}さん（{月}）」
```

**関数**: `doPost()` → `handleAdminUnlockCommand_()` → `unlockTeacherSheet_()` → `pushLine_()`

**補足**: 講師がシートを開いた時（`onOpen()`）に、`checkAndUnlockIfNeeded_()` が実行され、ビューアーになっている場合は自動的にロック解除されます。

---

## データ構造

### Teachersシート
| 列名 | 説明 |
|------|------|
| 氏名 | 講師の氏名（空白を除いて正規化して検索） |
| lineUserId | LINE BotのUser ID（上書き禁止） |
| lineLinkedAt | LINE登録日時 |
| teacherId | 講師ID（任意） |
| メール | メールアドレス（編集権限付与用） |

### Submissionsシート
| 列名 | 説明 |
|------|------|
| timestamp | フォーム送信日時 |
| monthKey | 提出月（YYYY-MM形式） |
| teacherId | 講師ID |
| 氏名 | 講師の氏名 |
| sheetUrl | 講師用シートのURL |
| status | 状態（created, submitted, teacher_not_found） |
| submittedAt | 提出日時 |
| lockedAt | ロック日時 |
| ackNotifiedAt | 提出受理通知送信日時 |
| reminderNotifiedAt | リマインド通知送信日時 |
| submissionKey | 提出キー（{monthKey}\|{teacherId}） |

### 講師用シートの _META シート
| キー | 値 |
|------|-----|
| MASTER_SPREADSHEET_ID | マスタースプレッドシートID |
| SUBMISSIONS_SHEET_NAME | Submissionsシート名 |
| SUBMISSION_KEY | 提出キー |
| MONTH_KEY | 月キー |
| TEACHER_ID | 講師ID |
| TEACHER_NAME | 講師名 |

---

## 設定が必要なスクリプトプロパティ

1. **LINE_CHANNEL_ACCESS_TOKEN**: LINE Botのチャネルアクセストークン
2. **ADMIN_LINE_USER_ID**: 管理者のLINE User ID（例外通知用、任意）

---

## トリガー設定

### 必須トリガー
- **onFormSubmit**: Googleフォーム送信時（インストール型トリガー推奨）

### 推奨トリガー（定期実行）
- **pollSubmissionsAndUpdate**: 提出状況の監視（例：5分ごと）
- **remindUnsubmitted**: 未提出リマインド（例：毎日9時）

---

## エラーハンドリング

- すべての主要関数で `try-catch` を使用
- エラー発生時は `handleError_()` でログ出力と管理者通知
- LINE APIの失敗は再試行可能なため、管理者通知は送らない

---

## セキュリティ機能

1. **LINE登録の上書き防止**: 既に別のLINE IDが登録されている場合、上書きを禁止
2. **提出後のロック**: 提出後は講師が編集不可（閲覧のみ）
3. **管理者によるロック解除**: 管理者のみがロック解除可能
4. **スクリプト実行者は常に編集可能**: 管理者は常にシートを編集可能

