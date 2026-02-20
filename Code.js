/************************************************************
 * Code.gs
 * - LINE webhook: doPost
 * - Form trigger: onFormSubmit
 *
 * 定数は utils.js で定義:
 *   CONFIG, CONFIG_FORM, SHEET_CONFIG, DRIVE_CONFIG, LINE_CONFIG,
 *   MESSAGE_SPREADSHEET_APP
 *
 * 【設定が必要なスクリプトプロパティ】
 * - LINE_CHANNEL_ACCESS_TOKEN: LINE Botのチャネルアクセストークン
 * - ADMIN_LINE_USER_ID: 管理者のLINE User ID（例外通知用、任意）
 ************************************************************/

/**
 * GETリクエストのハンドラー
 * - OAuth認証コールバック（codeパラメータあり）
 * - 通常のヘルスチェック（パラメータなし）
 */
function doGet(e) {
  // OAuth認証コールバックの場合
  if (e && e.parameter && e.parameter.code) {
    return handleOAuthCallback_(e);
  }

  // OAuth認証ページの表示（authパラメータあり）
  if (e && e.parameter && e.parameter.auth) {
    return showOAuthPage_();
  }

  // 通常のヘルスチェック
  return ContentService.createTextOutput('OK');
}

/**
 * LINE Webhook：講師が氏名を送信 → TeachersにlineUserId紐付け（上書き禁止）
 */
function doPost(e) {
  try {
    const body = (e && e.postData && e.postData.contents) ? e.postData.contents : '';

    // 到達ログ（念のため）
    const props = PropertiesService.getScriptProperties();
    const n = Number(props.getProperty('HIT_COUNT') || '0') + 1;
    props.setProperty('HIT_COUNT', String(n));
    props.setProperty('LAST_BODY', body.slice(0, 3000));

    if (!body) return ContentService.createTextOutput('OK');

    const payload = JSON.parse(body);
    const events = payload.events || [];

    const master = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
    const adminLineUserId = PropertiesService.getScriptProperties().getProperty('ADMIN_LINE_USER_ID') || '';

    for (const ev of events) {
      if (ev.type !== 'message') continue;
      if (!ev.message || ev.message.type !== 'text') continue;

      const userId = ev.source && ev.source.userId ? ev.source.userId : '';
      const textRaw = String(ev.message.text || '').trim();
      const replyToken = ev.replyToken || '';
      if (!userId || !replyToken) continue;

      // 管理者からのロック解除コマンドを処理
      if (adminLineUserId && userId === adminLineUserId) {
        // 月の選択待ち状態を処理
        const monthSelectKey = `MONTH_SELECT_${userId}`;
        const monthSelectData = props.getProperty(monthSelectKey);
        if (monthSelectData) {
          try {
            const selectInfo = JSON.parse(monthSelectData);
            const selectedMonth = textRaw.trim();
            
            // 24時間以内の選択のみ有効
            if (isStateExpired_(selectInfo.timestamp)) {
              props.deleteProperty(monthSelectKey);
            } else if (selectInfo.availableMonths.includes(selectedMonth)) {
              // 選択された月でロック解除を実行
              const unlockResult = handleAdminUnlockCommand_(master, `変更依頼 ${selectInfo.teacherName} ${selectedMonth}`);
              if (unlockResult.handled) {
                replyLine_(replyToken, unlockResult.message);
              } else {
                replyLine_(replyToken, `エラー：ロック解除に失敗しました。`);
              }
              props.deleteProperty(monthSelectKey);
              continue;
            } else {
              // 無効な選択の場合、再度選択肢を提示
              const monthList = selectInfo.availableMonths.map((m, i) => `${i + 1}. ${m}`).join('\n');
              replyLine_(replyToken, `無効な選択です。以下の月から選択してください：\n${monthList}`);
              continue;
            }
          } catch (e) {
            console.error('Month selection error:', e);
            props.deleteProperty(monthSelectKey);
          }
        }
        
        // コマンド形式: "変更依頼: 講師名 月" または "変更依頼: 講師名" または "変更依頼:講師名"
        const unlockResult = handleAdminUnlockCommand_(master, textRaw);
        if (unlockResult.handled) {
          replyLine_(replyToken, unlockResult.message);
          continue;
        }
        // コマンドとして認識されなかった場合は通常の名前検索に進む（管理者も登録可能にするため）
      }

      // メールアドレス更新の確認応答を処理
      const emailUpdateKey = `EMAIL_UPDATE_${userId}`;
      const emailUpdateData = props.getProperty(emailUpdateKey);
      if (emailUpdateData) {
        try {
          const updateInfo = JSON.parse(emailUpdateData);

          // 24時間以内の確認のみ有効
          if (isStateExpired_(updateInfo.timestamp)) {
            props.deleteProperty(emailUpdateKey);
          } else {
            const confirmation = parseConfirmationResponse_(textRaw);
            if (confirmation === 'yes') {
              // メールアドレスを更新
              const teacher = findTeacherByName_(master, updateInfo.name);
              if (teacher) {
                updateTeacherEmail_(master, teacher.row, updateInfo.newEmail);
                replyLine_(replyToken, `メールアドレスを変更しました：${updateInfo.newEmail}`);
              } else {
                replyLine_(replyToken, `エラー：講師情報が見つかりませんでした。`);
              }
              props.deleteProperty(emailUpdateKey);
              continue;
            } else if (confirmation === 'no') {
              replyLine_(replyToken, `メールアドレスの変更をキャンセルしました。`);
              props.deleteProperty(emailUpdateKey);
              continue;
            } else {
              // 確認待ち状態を維持（「はい」「いいえ」以外のメッセージ）
              replyLine_(replyToken, `メールアドレスの変更を続けますか？\n「はい」または「いいえ」と送信してください。`);
              continue;
            }
          }
        } catch (e) {
          console.error('Email update confirmation error:', e);
          props.deleteProperty(emailUpdateKey);
        }
      }

      // 初回登録時のメールアドレス待ち状態を処理
      const emailRequestKey = `EMAIL_REQUEST_${userId}`;
      const emailRequestData = props.getProperty(emailRequestKey);
      if (emailRequestData) {
        try {
          const requestInfo = JSON.parse(emailRequestData);
          const extractedEmail = extractEmail_(textRaw);
          
          // 24時間以内のリクエストのみ有効
          if (isStateExpired_(requestInfo.timestamp)) {
            props.deleteProperty(emailRequestKey);
          } else if (extractedEmail && isValidEmail_(extractedEmail)) {
            // メールアドレスが送信された場合
            // Gmailアドレスかどうかをチェック
            if (isGmailAddress_(extractedEmail)) {
              // Gmailアドレスの場合、自動登録
              const teacher = findTeacherByName_(master, requestInfo.name);
              if (teacher && teacher.row) {
                updateTeacherEmail_(master, teacher.row, extractedEmail);
                const lastName = extractLastName_(requestInfo.name);
                replyLine_(replyToken, `メールアドレスを登録しました：${extractedEmail}\n\n${lastName}先生、登録が完了しました。`);
              } else {
                replyLine_(replyToken, `エラー：講師情報が見つかりませんでした。`);
              }
              props.deleteProperty(emailRequestKey);
            } else {
              // Gmailアドレスでない場合、再度促す
              const lastName = extractLastName_(requestInfo.name);
              replyLine_(replyToken, `${lastName}先生、Gmailアドレスを登録してください。\nGmailアドレスを送信してください。\n例：example@gmail.com`);
            }
            continue;
          } else {
            // メールアドレスが含まれていない場合、再度促す
            const lastName = extractLastName_(requestInfo.name);
            replyLine_(replyToken, `${lastName}先生、Gmailアドレスを登録してください。\nGmailアドレスを送信してください。\n例：example@gmail.com`);
            continue;
          }
        } catch (e) {
          console.error('Email request processing error:', e);
          props.deleteProperty(emailRequestKey);
        }
      }

      // 「講師登録」プレフィックスのチェック
      const registrationPrefix = '講師登録';
      const hasRegistrationPrefix = textRaw.trim().startsWith(registrationPrefix);
      
      // メールアドレスが含まれているかチェック
      const extractedEmail = extractEmail_(textRaw);
      const hasEmail = extractedEmail && extractedEmail.trim().length > 0;
      
      // まず、このLINE User IDで既に登録されている講師を検索
      const existingTeacherByLineId = findTeacherByLineUserId_(master, userId);
      
      // 既に登録済みで、メールアドレスとLINE IDの両方が登録されている場合
      if (existingTeacherByLineId) {
        const hasEmailRegistered = existingTeacherByLineId.email && existingTeacherByLineId.email.trim().length > 0;
        const hasLineIdRegistered = existingTeacherByLineId.lineUserId && existingTeacherByLineId.lineUserId.trim().length > 0;
        
        // メールアドレスとLINE IDの両方が登録済みの場合
        if (hasEmailRegistered && hasLineIdRegistered) {
          // メールアドレスが含まれていないメッセージは通常の会話として無視
          if (!hasEmail && !hasRegistrationPrefix) {
            continue; // 無視（エラーメッセージを送らない）
          }
          // メールアドレスが含まれている場合は、メールアドレス更新処理に進む
        }
        // どちらかが未登録の場合は、登録処理を続行
      }
      
      // メールアドレスが含まれていない場合、既に完全登録済みでない限り登録処理を試みる
      // メールアドレスが含まれている場合は必ず登録処理を実行
      
      // 氏名を抽出（メールアドレスや挨拶文を除去）
      let extractedName = extractNameFromText_(textRaw);
      
      // 「講師登録」プレフィックスを除去
      if (hasRegistrationPrefix) {
        extractedName = extractedName.replace(new RegExp(`^${registrationPrefix}\\s*`), '').trim();
      }
      
      // 名前が抽出できない場合、メールアドレスを除いたテキスト全体を使用
      if (!extractedName || extractedName.length < 1) {
        let textWithoutEmail = textRaw.replace(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g, '').trim();
        if (hasRegistrationPrefix) {
          textWithoutEmail = textWithoutEmail.replace(new RegExp(`^${registrationPrefix}\\s*`), '').trim();
        }
        extractedName = textWithoutEmail;
      }
      
      // 氏名を正規化して検索
      const nameKey = normalizeNameKey_(extractedName || textRaw);
      console.log(`[DEBUG] extractedName: "${extractedName}", nameKey: "${nameKey}", extractedEmail: "${extractedEmail}"`);
      const result = linkLineUserByName_(master, nameKey, userId);
      console.log(`[DEBUG] linkLineUserByName_ result: ${JSON.stringify(result)}`);

      if (result.status === 'not_found') {
        // メールアドレスが含まれている場合のみ登録処理を実行
        if (hasEmail || hasRegistrationPrefix) {
          // 名前とメールアドレスの両方がある場合、または「講師登録」プレフィックスがある場合は自動登録
          const teacherName = extractedName || textRaw.replace(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g, '').trim();

          // 名前が抽出できた場合は自動登録
          if (teacherName && teacherName.length >= 2) {
            const newTeacher = addNewTeacher_(master, teacherName, userId, extractedEmail || '');
            const lastName = extractLastName_(newTeacher.name);

            let message = `登録OK：${lastName}先生\n今後はこのLINEでシフト連絡します。${MESSAGE_SPREADSHEET_APP}`;

            if (extractedEmail && isValidEmail_(extractedEmail)) {
              // Gmailアドレスかどうかをチェック
              if (isGmailAddress_(extractedEmail)) {
                message += `\n\nメールアドレスを登録しました：${extractedEmail}`;
              } else {
                // Gmailアドレスでない場合、Gmailアドレスを要求
                const props = PropertiesService.getScriptProperties();
                props.setProperty(`EMAIL_REQUEST_${userId}`, JSON.stringify({
                  name: newTeacher.name,
                  timestamp: new Date().getTime()
                }));
                message += `\n\nGmailアドレスを登録してください。\nGmailアドレスを送信してください。\n例：example@gmail.com`;
              }
            } else {
              // メールアドレスが登録されていない場合、自動で要求
              const props = PropertiesService.getScriptProperties();
              props.setProperty(`EMAIL_REQUEST_${userId}`, JSON.stringify({
                name: newTeacher.name,
                timestamp: new Date().getTime()
              }));
              message += `\n\nGmailアドレスを登録してください。\nGmailアドレスを送信してください。\n例：example@gmail.com`;
            }

            replyLine_(replyToken, message);
            continue;
          } else {
            // 名前が抽出できない場合はエラーメッセージ
            replyLine_(replyToken, `お名前（フルネーム）とGmailアドレスの両方を送ってください。\n例：山田太郎 taro@gmail.com`);
            continue;
          }
        } else {
          // メールアドレスが含まれていない場合
          // 名前が2文字以上なら、名前を保存してメールアドレスを要求する
          const teacherName = extractedName || textRaw.trim();
          if (teacherName && teacherName.length >= 2 && looksLikeName_(teacherName)) {
            // 名前だけ送信された場合、名前を一時保存してメールアドレスを要求
            const newTeacher = addNewTeacher_(master, teacherName, userId, '');
            const lastName = extractLastName_(newTeacher.name);
            const props = PropertiesService.getScriptProperties();
            props.setProperty(`EMAIL_REQUEST_${userId}`, JSON.stringify({
              name: newTeacher.name,
              timestamp: new Date().getTime()
            }));
            replyLine_(replyToken, `${lastName}先生、Gmailアドレスを登録してください。\nGmailアドレスを送信してください。\n例：example@gmail.com`);
            continue;
          }
          // 名前として認識できない場合は無視
          continue;
        }
      }

      if (result.status === 'multiple') {
        replyLine_(
          replyToken,
          `同じ氏名が複数います（候補：${result.candidates.join(' / ')}）\n` +
          `フルネームをそのまま送ってください（空白は気にしなくてOK）。`
        );
        continue;
      }

      if (result.status === 'already_linked_other') {
        // 別のLINE IDが登録されている場合
        const currentEmail = result.email || '';
        const lastName = extractLastName_(result.name);
        
        // メールアドレスが含まれている場合、メールアドレスで確認
        if (extractedEmail && isValidEmail_(extractedEmail)) {
          if (currentEmail && currentEmail === extractedEmail) {
            // メールアドレスが一致 → LINE IDを更新（アカウント変更の可能性）
            const sh = master.getSheetByName(CONFIG.SHEET_TEACHERS);
            const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
            const idxLine = header.indexOf('lineUserId');
            const idxLinkedAt = header.indexOf('lineLinkedAt');
            
            if (idxLine >= 0 && result.row) {
              sh.getRange(result.row, idxLine + 1).setValue(userId);
              if (idxLinkedAt >= 0) sh.getRange(result.row, idxLinkedAt + 1).setValue(new Date());
              
              replyLine_(
                replyToken,
                `新しいシステムが導入されました。\nアカウントを変更しましたか？再登録しました：${lastName}先生`
              );
              continue;
            }
          }
        }
        
        // メールアドレスが一致しない、または含まれていない場合
        replyLine_(replyToken, `この氏名は別のLINEと紐付いています：${result.name}\n教室まで連絡してください。`);
        continue;
      }

      if (result.status === 'already_linked_same') {
        // 既に登録済みの場合、完全一致チェック
        const currentEmail = result.email || '';
        console.log(`[DEBUG] already_linked_same: currentEmail="${currentEmail}", extractedEmail="${extractedEmail}", isValidEmail=${isValidEmail_(extractedEmail)}`);

        // メールアドレスが含まれている場合
        if (extractedEmail && isValidEmail_(extractedEmail)) {
          if (currentEmail && currentEmail === extractedEmail) {
            // LINE IDもメールアドレスも完全一致 → 既に登録済み
            const lastName = extractLastName_(result.name);
            replyLine_(replyToken, `既に登録済みです：${lastName}先生`);
            continue;
          } else if (currentEmail && currentEmail !== extractedEmail) {
            // メールアドレスが違う場合、確認を求める
            const props = PropertiesService.getScriptProperties();
            props.setProperty(`EMAIL_UPDATE_${userId}`, JSON.stringify({
              name: result.name,
              oldEmail: currentEmail,
              newEmail: extractedEmail,
              timestamp: new Date().getTime()
            }));
            replyLine_(
              replyToken,
              `登録されているメールアドレスと違います。\n現在: ${currentEmail}\n送信: ${extractedEmail}\n\n変更しますか？「はい」または「いいえ」と送信してください。`
            );
            continue;
          } else if (!currentEmail) {
            // メールアドレスが登録されていない場合、登録
            const lastName = extractLastName_(result.name);
            if (result.row) {
              updateTeacherEmail_(master, result.row, extractedEmail);
              replyLine_(replyToken, `登録OK：${lastName}先生\n今後はこのLINEでシフト連絡します。${MESSAGE_SPREADSHEET_APP}\n\nメールアドレスを登録しました：${extractedEmail}`);
            } else {
              // result.rowがない場合でも返信を送信
              replyLine_(replyToken, `登録OK：${lastName}先生\n今後はこのLINEでシフト連絡します。${MESSAGE_SPREADSHEET_APP}`);
            }
            continue;
          }
          // フォールバック：条件に一致しない場合（通常は到達しない）
          const lastName = extractLastName_(result.name);
          replyLine_(replyToken, `既に登録済みです：${lastName}先生`);
          continue;
        } else {
          // メールアドレスが含まれていない場合、既に登録済みと返す
          const lastName = extractLastName_(result.name);
          replyLine_(replyToken, `既に登録済みです：${lastName}先生`);
          continue;
        }
      }

      if (result.status === 'linked') {
        // 新規登録の場合
        const lastName = extractLastName_(result.name);
        const currentEmail = result.email || '';
        
        // メールアドレスが含まれている場合
        if (extractedEmail && isValidEmail_(extractedEmail)) {
          if (currentEmail && currentEmail === extractedEmail) {
            // メールアドレスが既に登録済み → 既に登録済みと返す
            replyLine_(replyToken, `既に登録済みです：${lastName}先生`);
            continue;
          } else if (currentEmail && currentEmail !== extractedEmail) {
            // メールアドレスが違う場合、確認を求める
            const props = PropertiesService.getScriptProperties();
            props.setProperty(`EMAIL_UPDATE_${userId}`, JSON.stringify({
              name: result.name,
              oldEmail: currentEmail,
              newEmail: extractedEmail,
              timestamp: new Date().getTime()
            }));
            replyLine_(
              replyToken,
              `登録OK：${lastName}先生\n今後はこのLINEでシフト連絡します。${MESSAGE_SPREADSHEET_APP}\n\n登録されているメールアドレスと違います。\n現在: ${currentEmail}\n送信: ${extractedEmail}\n\n変更しますか？「はい」または「いいえ」と送信してください。`
            );
            continue;
          } else if (!currentEmail) {
            // メールアドレスが登録されていない場合
            // Gmailアドレスかどうかをチェック
            if (isGmailAddress_(extractedEmail)) {
              // Gmailアドレスの場合、登録
              if (result.row) {
                updateTeacherEmail_(master, result.row, extractedEmail);
                replyLine_(replyToken, `登録OK：${lastName}先生\n今後はこのLINEでシフト連絡します。${MESSAGE_SPREADSHEET_APP}\n\nメールアドレスを登録しました：${extractedEmail}`);
              } else {
                replyLine_(replyToken, `登録OK：${lastName}先生\n今後はこのLINEでシフト連絡します。${MESSAGE_SPREADSHEET_APP}`);
              }
            } else {
              // Gmailアドレスでない場合、Gmailアドレスを要求
              const props = PropertiesService.getScriptProperties();
              props.setProperty(`EMAIL_REQUEST_${userId}`, JSON.stringify({
                name: result.name,
                timestamp: new Date().getTime()
              }));
              replyLine_(replyToken, `登録OK：${lastName}先生\n今後はこのLINEでシフト連絡します。${MESSAGE_SPREADSHEET_APP}\n\nGmailアドレスを登録してください。\nGmailアドレスを送信してください。\n例：example@gmail.com`);
            }
            continue;
          }
        }
        
        // メールアドレスが含まれていない場合
        let message = `登録OK：${lastName}先生\n今後はこのLINEでシフト連絡します。${MESSAGE_SPREADSHEET_APP}`;
        if (!currentEmail) {
          // メールアドレスが登録されていない場合、自動で要求
          const props = PropertiesService.getScriptProperties();
          props.setProperty(`EMAIL_REQUEST_${userId}`, JSON.stringify({
            name: result.name,
            timestamp: new Date().getTime()
          }));
          message += `\n\n${lastName}先生、Gmailアドレスを登録してください。\nGmailアドレスを送信してください。\n例：example@gmail.com`;
        }
        replyLine_(replyToken, message);
        continue;
      }

    }

    return ContentService.createTextOutput('OK');
  } catch (err) {
    handleError_(err, 'doPost', { 
      body: (e && e.postData && e.postData.contents) ? String(e.postData.contents).slice(0, 500) : 'no body' 
    });
    return ContentService.createTextOutput('OK');
  }
}

/**
 * Googleフォーム送信（インストール型トリガー推奨）
 * - テンプレから講師用シートをコピー作成
 * - 編集権限付与（Teachersのメール）
 * - Submissions追記
 * - _META書き込み
 * - LINEにURLをPush
 */
function onFormSubmit(e) {
  let teacherNameRaw = 'unknown';
  let monthKey = 'unknown';
  try {
    const nv = extractNamedValues_(e);

    const master = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);

    teacherNameRaw = (nv[CONFIG_FORM.QUESTION_TEACHER_NAME] || '').trim();
    if (!teacherNameRaw) throw new Error(`フォームに「${CONFIG_FORM.QUESTION_TEACHER_NAME}」がありません（または空です）`);

    monthKey = (nv[CONFIG_FORM.QUESTION_MONTH_KEY] || '').trim() || getNextMonthKey_();

    const teacher = findTeacherByName_(master, teacherNameRaw);
    if (!teacher) {
      // 講師が見つからない場合でも、既存のsubmissionがあれば更新する（重複防止）
      const submissionKeyNotFound = `${monthKey}-${normalizeNameKey_(teacherNameRaw)}`;
      let existingNotFound = findSubmissionByKey_(master, submissionKeyNotFound);
      if (!existingNotFound) {
        existingNotFound = findSubmissionByMonthAndTeacher_(master, monthKey, '', teacherNameRaw);
      }

      if (existingNotFound) {
        // 既存エントリを更新
        updateSubmission_(master, existingNotFound.row, existingNotFound.header, {
          timestamp: new Date(),
          status: 'teacher_not_found',
        });
      } else {
        // 新規エントリを作成
        appendSubmission_(master, {
          timestamp: new Date(),
          monthKey,
          teacherId: '',
          name: teacherNameRaw,
          sheetUrl: '',
          status: 'teacher_not_found',
          lastNotified: '',
          submissionKey: submissionKeyNotFound,
          submittedAt: '',
        });
      }
      return;
    }

    const teacherId = teacher.teacherId || '';
    const teacherName = teacher.name;
    const teacherEmail = teacher.email || '';
    const lineUserId = teacher.lineUserId || '';

    // Submissionsに記録（既存エントリがあれば更新、なければ新規作成）
    const submissionKey = `${monthKey}-${teacherId || normalizeNameKey_(teacherName)}`;

    // 既存のsubmissionを検索（submissionKeyで検索、見つからなければmonthKeyとteacherId/氏名で検索）
    // これにより、submissionKeyの生成方式が異なる場合でも重複を防げる
    let existingSubmission = findSubmissionByKey_(master, submissionKey);
    if (!existingSubmission) {
      // submissionKeyで見つからない場合、monthKeyとteacherId/氏名の組み合わせで検索
      existingSubmission = findSubmissionByMonthAndTeacher_(master, monthKey, teacherId, teacherName);
    }

    // 既存のsubmissionがあり、かつsheetUrlが存在する場合は、新しいシートを作成せず既存のシートを使用
    let newSpreadsheetId = null;
    let sheetUrl = '';
    
    if (existingSubmission && existingSubmission.sheetUrl && existingSubmission.sheetUrl.trim()) {
      // 既存のシートを使用
      sheetUrl = existingSubmission.sheetUrl.trim();
      newSpreadsheetId = extractSpreadsheetId_(sheetUrl);
      
      // 既存エントリを更新（再送信の場合はackNotifiedAtをクリアして再提出受理通知を送れるようにする）
      updateSubmission_(master, existingSubmission.row, existingSubmission.header, {
        timestamp: new Date(), // フォーム送信日時で更新
        teacherId: teacherId, // teacherIdを正しい値に更新
        name: teacherName,
        sheetUrl: sheetUrl,
        status: 'created', // シート作成済みに更新
        submissionKey: submissionKey, // submissionKeyを正しい値に更新
        ackNotifiedAt: '', // 再送信時はackNotifiedAtをクリア（再提出受理通知を送るため）
      });
    } else {
      // 新しいシートを作成
      // 月ごとのテンプレートを検索
      let templateSpreadsheetId = null;
      if (CONFIG.TEMPLATE_FOLDER_ID) {
        templateSpreadsheetId = findTemplateByMonth_(monthKey, CONFIG.TEMPLATE_FOLDER_ID);
      }
      
      // 月ごとのテンプレートが見つからない場合、エラーとして処理
      if (!templateSpreadsheetId) {
        // エラーログ（管理者向け、詳細は最小限）
        const errorMsg = `${monthKey}のテンプレートが見つかりません`;
        console.error(`[${new Date().toISOString()}] onFormSubmit: ${errorMsg}`, { monthKey, teacherName });
        
        // 講師にLINE通知（シンプルなメッセージ）
        if (lineUserId) {
          pushLine_(lineUserId, `${monthKey}のシフト申請用紙はありません。管理者に連絡してください。`);
        }
        
        // Submissionsにエラー状態で記録
        if (existingSubmission) {
          updateSubmission_(master, existingSubmission.row, existingSubmission.header, {
            timestamp: new Date(),
            teacherId: teacherId, // teacherIdを正しい値に更新
            name: teacherName,
            status: 'template_not_found',
            submissionKey: submissionKey, // submissionKeyを正しい値に更新
          });
        } else {
          appendSubmission_(master, {
            timestamp: new Date(),
            monthKey,
            teacherId: teacherId || '',
            name: teacherName,
            sheetUrl: '',
            status: 'template_not_found',
            lastNotified: '',
            submissionKey,
            submittedAt: '',
          });
        }
        return;
      }
      
      // 月フォルダ確保 → テンプレコピー
      const monthFolderId = ensureMonthFolder_(monthKey, CONFIG.COPIES_PARENT_FOLDER_ID);
      const fileName = `${monthKey}_${teacherName}_シフト提出`;
      newSpreadsheetId = copyTemplateSpreadsheet_(monthFolderId, templateSpreadsheetId, fileName);
      sheetUrl = `https://docs.google.com/spreadsheets/d/${newSpreadsheetId}/edit`;

      // 「リンクを知っているすべての人」を編集者に設定（新しいシートの場合のみ）
      setAnyoneWithLinkCanEdit_(newSpreadsheetId);

      // Submissionsに記録
      if (existingSubmission) {
        // 既存エントリを更新
        updateSubmission_(master, existingSubmission.row, existingSubmission.header, {
          timestamp: new Date(), // フォーム送信日時で更新
          teacherId: teacherId, // teacherIdを正しい値に更新
          name: teacherName,
          sheetUrl: sheetUrl,
          status: 'created', // シート作成済みに更新
          submissionKey: submissionKey, // submissionKeyを正しい値に更新
          ackNotifiedAt: '', // 再送信時はackNotifiedAtをクリア（再提出受理通知を送るため）
        });
      } else {
        // 新規エントリを作成
        appendSubmission_(master, {
          timestamp: new Date(),
          monthKey,
          teacherId,
          name: teacherName,
          sheetUrl,
          status: 'created',
          lastNotified: '',
          submissionKey,
          submittedAt: '',
        });
      }
    }

    // _META書き込み（テンプレ側表示や回収のため）
    // 既存のシートの場合でも、_METAとG3は更新する（念のため）
    if (newSpreadsheetId) {
      writeMetaToTeacherSheet_(newSpreadsheetId, CONFIG.META_SHEET_NAME, {
        MASTER_SPREADSHEET_ID: CONFIG.MASTER_SPREADSHEET_ID,
        SUBMISSIONS_SHEET_NAME: CONFIG.SHEET_SUBMISSIONS,
        SUBMISSION_KEY: submissionKey,
        MONTH_KEY: monthKey,
        TEACHER_ID: teacherId,
        TEACHER_NAME: teacherName,
      });

      // 講師名をG3に直接設定（シートを開いた時にすぐ表示されるように）
      try {
        const teacherSs = SpreadsheetApp.openById(newSpreadsheetId);
        const inputSheet = teacherSs.getSheetByName('Input');
        if (inputSheet) {
          inputSheet.getRange('G3').setValue(teacherName);
        }
      } catch (e) {
        // G3の設定に失敗しても続行（onOpen()で後から設定される）
        console.error('Failed to set teacher name in G3:', e);
      }
    }

    // LINEにURL送信（新しいシートの場合のみ、登録済みのみ）
    // 既存のsubmissionがあり、既存のシートを使用する場合は、LINE通知を送信しない（重複通知を防ぐ）
    if (lineUserId && (!existingSubmission || !existingSubmission.sheetUrl || !existingSubmission.sheetUrl.trim())) {
      pushLine_(lineUserId,
        `【シフト提出URL（${monthKey}）】\n${sheetUrl}\n\n※編集するには登録したGmailでGoogleにログインしてください。\n入力後、☑（提出）を入れてください。`
      );
    }

  } catch (err) {
    handleError_(err, 'onFormSubmit', {
      teacherNameRaw,
      monthKey
    });
  }
}


/**
 * 管理者からの「henshin」単独コマンド：最新の提出済みシートをロック解除
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} masterSs - マスタースプレッドシート
 * @returns {Object} {handled: boolean, message: string}
 */
function handleAdminUnlockLatest_(masterSs) {
  try {
    const sh = masterSs.getSheetByName(CONFIG.SHEET_SUBMISSIONS);
    if (!sh) {
      return { handled: true, message: 'Submissionsシートが見つかりません' };
    }

    const values = sh.getDataRange().getValues();
    if (values.length < 2) {
      return { handled: true, message: '提出データが見つかりません' };
    }

    const header = values[0];
    const idxUrl = header.indexOf('sheetUrl');
    const idxStatus = header.indexOf('status');
    const idxName = header.indexOf('氏名');
    const idxMonthKey = header.indexOf('monthKey');
    const idxLockedAt = header.indexOf('lockedAt');
    const idxTeacherId = header.indexOf('teacherId');
    const idxSubmittedAt = header.indexOf('submittedAt');

    if (idxUrl < 0 || idxStatus < 0 || idxName < 0) {
      return { handled: true, message: 'Submissionsに必要列がありません' };
    }

    // 最新の提出済みシートを検索（submittedAtが最新のもの）
    let latestRow = -1;
    let latestSubmittedAt = null;

    for (let r = 1; r < values.length; r++) {
      const status = String(values[r][idxStatus] || '').trim();
      if (status !== 'submitted') continue;

      const submittedAt = values[r][idxSubmittedAt];
      if (submittedAt && (!latestSubmittedAt || new Date(submittedAt) > new Date(latestSubmittedAt))) {
        latestRow = r + 1;
        latestSubmittedAt = submittedAt;
      }
    }

    if (latestRow < 0) {
      return { handled: true, message: '提出済みのデータが見つかりません' };
    }

    const targetUrl = String(values[latestRow - 1][idxUrl] || '').trim();
    const targetTeacherId = idxTeacherId >= 0 ? String(values[latestRow - 1][idxTeacherId] || '').trim() : '';
    const targetTeacherName = String(values[latestRow - 1][idxName] || '').trim();
    const targetMonthKey = idxMonthKey >= 0 ? normalizeMonthKey_(values[latestRow - 1][idxMonthKey]) : '';

    if (!targetUrl) {
      return { handled: true, message: 'シートURLが見つかりません' };
    }

    // ロック解除
    const spreadsheetId = extractSpreadsheetId_(targetUrl);
    if (!spreadsheetId) {
      return { handled: true, message: 'シートIDの取得に失敗しました' };
    }

    // 講師情報を取得
    const teacher = getTeacherInfo_(masterSs, targetTeacherId, targetTeacherName);
    const teacherEmail = teacher ? teacher.email || '' : '';
    const lineUserId = teacher ? teacher.lineUserId || '' : '';

    // メールアドレスがない場合は警告
    if (!teacherEmail) {
      const lastName = extractLastName_(targetTeacherName);
      return { handled: true, message: `ロック解除に失敗しました：${lastName}先生のメールアドレスが登録されていません。Teachersシートにメールアドレスを追加してください。` };
    }

    // ロック解除
    const unlockResult = unlockTeacherSheet_(spreadsheetId, teacherEmail);
    if (!unlockResult.success) {
      const lastName = extractLastName_(targetTeacherName);
      const errorMsg = unlockResult.errorMessage || '原因不明のエラー';
      return { handled: true, message: `ロック解除に失敗しました：${lastName}先生（${targetMonthKey}）\n\n${errorMsg}` };
    }

    // Submissionsの状態をリセット（再提出可能にする）
    if (idxLockedAt >= 0) {
      sh.getRange(latestRow, idxLockedAt + 1).setValue('');
    }
    // statusを'created'に戻す（再提出可能にする）
    sh.getRange(latestRow, idxStatus + 1).setValue('created');
    // submittedAtをクリア
    if (idxSubmittedAt >= 0) {
      sh.getRange(latestRow, idxSubmittedAt + 1).setValue('');
    }
    // ackNotifiedAtをクリア（再提出受理通知を送るため）
    const idxAckNotifiedAt = header.indexOf('ackNotifiedAt');
    if (idxAckNotifiedAt >= 0) {
      sh.getRange(latestRow, idxAckNotifiedAt + 1).setValue('');
    }

    // 講師シートのチェックボックスとステータスをリセット
    try {
      const teacherSs = SpreadsheetApp.openById(spreadsheetId);
      const inputSheet = teacherSs.getSheetByName('Input');
      if (inputSheet) {
        // チェックボックス（C2）をFALSEにリセット
        inputSheet.getRange('C2').setValue(false);
        // ステータス（B2）を「未提出」に戻す
        inputSheet.getRange('B2').setValue('未提出');
      }
    } catch (e) {
      console.error('Failed to reset checkbox and status:', e);
      // エラーが発生しても続行
    }

    // 講師にLINE通知
    if (lineUserId) {
      pushLine_(lineUserId,
        `【シフト変更依頼】\n${extractLastName_(targetTeacherName)}先生（${targetMonthKey}）のシフトを変更していただくようお願いします。\nシートの編集が可能になりました。\n\n※編集するには登録したGmailでGoogleにログインしてください。\n変更後、☑（提出）を入れてください。\n${targetUrl}`
      );
    }

    const lastName = extractLastName_(targetTeacherName);
    return { handled: true, message: `ロック解除しました：${lastName}先生（${targetMonthKey}）` };

  } catch (err) {
    handleError_(err, 'handleAdminUnlockLatest_');
    return { handled: true, message: 'エラーが発生しました：' + (err.message || String(err)) };
  }
}

/**
 * ロック解除コマンドをパース
 * @param {string} command - コマンド文字列
 * @returns {Object|null} {teacherName: string, monthKey: string} またはnull
 */
function parseUnlockCommand_(command) {
  const trimmedCommand = command.trim();
  let monthKey = '';
  let teacherName = '';

  // パターン0: 月が先頭（例：「2月変更依頼 森永英敬」）
  let match = trimmedCommand.match(/^(\d{1,2})月変更依頼\s+(.+)$/);
  if (match) {
    monthKey = parseMonthText_(`${match[1]}月`);
    teacherName = match[2].trim();
    return { teacherName, monthKey };
  }

  // パターン1: コロンあり（全角/半角）
  match = trimmedCommand.match(/^変更依頼[：:]\s*(.+?)(?:\s+(\d{4}-\d{2}))?\s*$/);

  // パターン2: コロンなし、スペースで始まる
  if (!match) {
    match = trimmedCommand.match(/^変更依頼\s+(.+?)(?:\s+(\d{4}-\d{2}))?\s*$/);
  }

  if (!match || !match[1]) return null;

  teacherName = match[1].trim();
  // 月の形式（YYYY-MM）が講師名に含まれている場合は除外
  if (teacherName.match(/^\d{4}-\d{2}$/)) return null;
  if (!teacherName) return null;

  monthKey = match[2] ? match[2].trim() : '';
  return { teacherName, monthKey };
}

/**
 * 再提出のためSubmissionsをリセット
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh - Submissionsシート
 * @param {number} row - 行番号
 * @param {Object} indices - 列インデックス
 */
function resetSubmissionForResubmit_(sh, row, indices) {
  if (indices.lockedAt >= 0) {
    sh.getRange(row, indices.lockedAt + 1).setValue('');
  }
  sh.getRange(row, indices.status + 1).setValue('created');
  if (indices.submittedAt >= 0) {
    sh.getRange(row, indices.submittedAt + 1).setValue('');
  }
  if (indices.ackNotifiedAt >= 0) {
    sh.getRange(row, indices.ackNotifiedAt + 1).setValue('');
  }
}

/**
 * 再提出のため講師シートをリセット
 * @param {string} spreadsheetId - スプレッドシートID
 */
function resetTeacherSheetForResubmit_(spreadsheetId) {
  try {
    const teacherSs = SpreadsheetApp.openById(spreadsheetId);
    const inputSheet = teacherSs.getSheetByName('Input');
    if (inputSheet) {
      inputSheet.getRange('C2').setValue(false);
      inputSheet.getRange('B2').setValue('未提出');
    }
  } catch (e) {
    console.error('Failed to reset teacher sheet:', e);
  }
}

/**
 * 管理者からのロック解除コマンドを処理
 * コマンド形式: "変更依頼: 講師名 月" または "変更依頼: 講師名"
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} masterSs - マスタースプレッドシート
 * @param {string} command - コマンド文字列
 * @returns {Object} {handled: boolean, message: string}
 */
function handleAdminUnlockCommand_(masterSs, command) {
  try {
    // コマンドをパース
    const parsed = parseUnlockCommand_(command);
    if (!parsed) {
      return { handled: false, message: '' };
    }
    const teacherNameRaw = parsed.teacherName;
    let monthKey = parsed.monthKey;

    // Submissionsから該当する提出を検索
    const sh = masterSs.getSheetByName(CONFIG.SHEET_SUBMISSIONS);
    if (!sh) {
      return { handled: true, message: 'Submissionsシートが見つかりません' };
    }

    const values = sh.getDataRange().getValues();
    if (values.length < 2) {
      return { handled: true, message: '提出データが見つかりません' };
    }

    const header = values[0];
    const idxUrl = header.indexOf('sheetUrl');
    const idxStatus = header.indexOf('status');
    const idxName = header.indexOf('氏名');
    const idxMonthKey = header.indexOf('monthKey');
    const idxLockedAt = header.indexOf('lockedAt');
    const idxTeacherId = header.indexOf('teacherId');
    const idxSubmittedAt = header.indexOf('submittedAt');
    const idxAckNotifiedAt = header.indexOf('ackNotifiedAt');

    if (idxUrl < 0 || idxStatus < 0 || idxName < 0) {
      return { handled: true, message: 'Submissionsに必要列がありません' };
    }

    // 月が指定されていない場合、提出済みの月のリストを取得
    if (!monthKey) {
      const availableMonths = [];
      const monthMap = new Map(); // 月ごとの情報を保持
      
      for (let r = 1; r < values.length; r++) {
        const name = String(values[r][idxName] || '').trim();
        const mk = idxMonthKey >= 0 ? normalizeMonthKey_(values[r][idxMonthKey]) : '';
        const status = String(values[r][idxStatus] || '').trim();
        
        const nameMatch = normalizeNameKey_(name) === normalizeNameKey_(teacherNameRaw);
        
        if (nameMatch && status === 'submitted' && mk) {
          if (!availableMonths.includes(mk)) {
            availableMonths.push(mk);
            monthMap.set(mk, {
              row: r + 1,
              url: String(values[r][idxUrl] || '').trim(),
              teacherId: idxTeacherId >= 0 ? String(values[r][idxTeacherId] || '').trim() : '',
              teacherName: name
            });
          }
        }
      }
      
      if (availableMonths.length === 0) {
        return { handled: true, message: `提出済みのデータが見つかりません：${teacherNameRaw}\n提出済み（status='submitted'）のデータが必要です。` };
      } else if (availableMonths.length === 1) {
        // 1つだけの場合は自動的に選択
        monthKey = availableMonths[0];
      } else {
        // 複数ある場合は選択肢を提示
        const props = PropertiesService.getScriptProperties();
        const adminLineUserId = PropertiesService.getScriptProperties().getProperty('ADMIN_LINE_USER_ID') || '';
        if (adminLineUserId) {
          props.setProperty(`MONTH_SELECT_${adminLineUserId}`, JSON.stringify({
            teacherName: teacherNameRaw,
            availableMonths: availableMonths,
            timestamp: new Date().getTime()
          }));
          
          const monthList = availableMonths.map((m, i) => `${i + 1}. ${m}`).join('\n');
          return { handled: true, message: `${teacherNameRaw}先生の提出済みシフトが複数あります。\n変更したい月を選択してください：\n${monthList}\n\n月を送信してください（例：2026-01）` };
        } else {
          return { handled: true, message: `${teacherNameRaw}先生の提出済みシフトが複数あります（${availableMonths.join('、')}）。\n月を指定してください：変更依頼 ${teacherNameRaw} 2026-01` };
        }
      }
    }

    // 該当する提出を検索
    let targetRow = -1;
    let targetUrl = '';
    let targetTeacherId = '';
    let targetTeacherName = '';
    let targetMonthKey = '';

    for (let r = 1; r < values.length; r++) {
      const name = String(values[r][idxName] || '').trim();
      const mk = idxMonthKey >= 0 ? normalizeMonthKey_(values[r][idxMonthKey]) : '';
      const status = String(values[r][idxStatus] || '').trim();

      const nameMatch = normalizeNameKey_(name) === normalizeNameKey_(teacherNameRaw);
      const monthMatch = !monthKey || mk === monthKey;

      if (nameMatch && monthMatch && status === 'submitted') {
        targetRow = r + 1;
        targetUrl = String(values[r][idxUrl] || '').trim();
        targetTeacherId = idxTeacherId >= 0 ? String(values[r][idxTeacherId] || '').trim() : '';
        targetTeacherName = name;
        targetMonthKey = mk;
        break;
      }
    }

    if (targetRow < 0) {
      // より詳細なエラーメッセージ：提出済みでない可能性も考慮
      let statusInfo = '';
      for (let r = 1; r < values.length; r++) {
        const name = String(values[r][idxName] || '').trim();
        const mk = idxMonthKey >= 0 ? normalizeMonthKey_(values[r][idxMonthKey]) : '';
        const status = String(values[r][idxStatus] || '').trim();
        const nameMatch = normalizeNameKey_(name) === normalizeNameKey_(teacherNameRaw);
        const monthMatch = !monthKey || mk === monthKey;
        if (nameMatch && monthMatch) {
          statusInfo = `（状態: ${status}）`;
          break;
        }
      }
      return { handled: true, message: `提出済みのデータが見つかりません：${teacherNameRaw}${monthKey ? ' ' + monthKey : ''}${statusInfo}\n提出済み（status='submitted'）のデータが必要です。` };
    }

    if (!targetUrl) {
      return { handled: true, message: 'シートURLが見つかりません' };
    }

    // ロック解除
    const spreadsheetId = extractSpreadsheetId_(targetUrl);
    if (!spreadsheetId) {
      return { handled: true, message: 'シートIDの取得に失敗しました' };
    }

    // 講師情報を取得
    const teacher = getTeacherInfo_(masterSs, targetTeacherId, targetTeacherName);
    const teacherEmail = teacher ? teacher.email || '' : '';
    const lineUserId = teacher ? teacher.lineUserId || '' : '';

    // メールアドレスがない場合は警告
    if (!teacherEmail) {
      const lastName = extractLastName_(targetTeacherName);
      return { handled: true, message: `ロック解除に失敗しました：${lastName}先生のメールアドレスが登録されていません。Teachersシートにメールアドレスを追加してください。` };
    }

    // ロック解除
    const unlockResult = unlockTeacherSheet_(spreadsheetId, teacherEmail);
    if (!unlockResult.success) {
      const lastName = extractLastName_(targetTeacherName);
      const errorMsg = unlockResult.errorMessage || '原因不明のエラー';
      return { handled: true, message: `ロック解除に失敗しました：${lastName}先生（${targetMonthKey}）\n\n${errorMsg}` };
    }

    // Submissionsの状態をリセット（再提出可能にする）
    resetSubmissionForResubmit_(sh, targetRow, {
      lockedAt: idxLockedAt,
      status: idxStatus,
      submittedAt: idxSubmittedAt,
      ackNotifiedAt: idxAckNotifiedAt
    });

    // 講師シートのチェックボックスとステータスをリセット
    resetTeacherSheetForResubmit_(spreadsheetId);

    // 講師にLINE通知
    if (lineUserId) {
      pushLine_(lineUserId,
        `【シフト変更依頼】\n${extractLastName_(targetTeacherName)}先生（${targetMonthKey}）のシフトを変更していただくようお願いします。\nシートの編集が可能になりました。\n\n※編集するには登録したGmailでGoogleにログインしてください。\n変更後、☑（提出）を入れてください。\n${targetUrl}`
      );
    }

    const lastName = extractLastName_(targetTeacherName);
    return { handled: true, message: `ロック解除しました：${lastName}先生（${targetMonthKey}）` };

  } catch (err) {
    handleError_(err, 'handleAdminUnlockCommand_', { command });
    return { handled: true, message: 'エラーが発生しました：' + (err.message || String(err)) };
  }
}

/**
 * 谷口知子先生のメールアドレスを追加する関数（管理者用）
 * Google Apps Scriptエディタで直接実行可能
 */
function registerTaniguchi() {
  try {
    const master = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
    const sh = master.getSheetByName(CONFIG.SHEET_TEACHERS);
    if (!sh) {
      throw new Error('Teachersシートが見つかりません');
    }
    
    // 谷口知子先生を検索（柔軟なマッチング）
    const existing = findTeacherByName_(master, '谷口知子');
    if (!existing) {
      // 見つからない場合は新規登録
      const result = addTeacherManually('谷口知子', 'satorara0510@gmail.com');
      console.log(`新規登録完了:`);
      console.log(`- 氏名: ${result.name}`);
      console.log(`- teacherId: ${result.teacherId}`);
      console.log(`- メール: ${result.email}`);
      return `新規登録完了: ${result.name} (teacherId: ${result.teacherId})`;
    }
    
    console.log(`既に登録済みです:`);
    console.log(`- 氏名: ${existing.name}`);
    console.log(`- teacherId: ${existing.teacherId}`);
    console.log(`- メール: ${existing.email || '(未登録)'}`);
    console.log(`- LINE User ID: ${existing.lineUserId || '(未登録)'}`);
    console.log(`- 行番号: ${existing.row}`);
    
    // メールアドレスを直接更新
    const email = 'satorara0510@gmail.com';
    const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const idxEmail = header.indexOf('メール');
    
    if (idxEmail < 0) {
      throw new Error('メール列が見つかりません');
    }
    
    // メールアドレスを更新
    sh.getRange(existing.row, idxEmail + 1).setValue(email);
    
    console.log(`メールアドレスを追加/更新しました: ${email}`);
    
    // 更新後の情報を確認
    const updated = findTeacherByName_(master, '谷口知子');
    console.log(`更新後の情報:`);
    console.log(`- 氏名: ${updated.name}`);
    console.log(`- teacherId: ${updated.teacherId}`);
    console.log(`- メール: ${updated.email}`);
    
    return `メールアドレスを追加しました: ${email}`;
  } catch (err) {
    console.error('登録エラー:', err);
    throw err;
  }
}

/**
 * 講師を手動登録してLINE通知を送る汎用関数（管理者用）
 * @param {string} teacherName - 講師名
 * @param {string} email - メールアドレス
 * @param {string} lineUserId - LINE User ID
 */
function registerTeacherManually_(teacherName, email, lineUserId) {
  const master = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
  const sh = master.getSheetByName(CONFIG.SHEET_TEACHERS);
  if (!sh) {
    throw new Error('Teachersシートが見つかりません');
  }

  // 既存の講師を検索
  const existing = findTeacherByName_(master, teacherName);

  if (existing) {
    console.log(`既に登録済み: ${existing.name}`);
    // LINE User IDとメールアドレスを更新
    const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const idxEmail = header.indexOf('メール');
    const idxLine = header.indexOf('lineUserId');
    const idxLinkedAt = header.indexOf('lineLinkedAt');

    if (idxEmail >= 0) sh.getRange(existing.row, idxEmail + 1).setValue(email);
    if (idxLine >= 0) sh.getRange(existing.row, idxLine + 1).setValue(lineUserId);
    if (idxLinkedAt >= 0) sh.getRange(existing.row, idxLinkedAt + 1).setValue(new Date());
  } else {
    // 新規登録
    const newTeacher = addNewTeacher_(master, teacherName, lineUserId, email);
    console.log(`新規登録完了: ${newTeacher.name} (teacherId: ${newTeacher.teacherId})`);
  }

  // LINE通知を送信
  const lastName = extractLastName_(teacherName);
  const message = `登録OK：${lastName}先生\n今後はこのLINEでシフト連絡します。${MESSAGE_SPREADSHEET_APP}\n\nメールアドレスを登録しました：${email}`;
  pushLine_(lineUserId, message);

  console.log('LINE通知を送信しました');
  return `登録完了: ${teacherName} (LINE通知送信済み)`;
}

/**
 * 吉本先生のシフト提出シートを修復する関数（管理者用）
 * status=createdなのにsheetUrlが空の場合に使用
 * Google Apps Scriptエディタで直接実行
 */
function repairYoshimotoSubmission() {
  const teacherName = '吉本偉大';
  const monthKey = '2026-02';

  const master = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);

  // 講師情報を取得
  const teacher = findTeacherByName_(master, teacherName);
  if (!teacher) {
    console.error(`講師が見つかりません: ${teacherName}`);
    return `講師が見つかりません: ${teacherName}`;
  }

  console.log(`講師情報: ${JSON.stringify(teacher)}`);

  const teacherId = teacher.teacherId || '';
  const lineUserId = teacher.lineUserId || '';

  // 月ごとのテンプレートを検索
  let templateSpreadsheetId = null;
  if (CONFIG.TEMPLATE_FOLDER_ID) {
    templateSpreadsheetId = findTemplateByMonth_(monthKey, CONFIG.TEMPLATE_FOLDER_ID);
  }

  if (!templateSpreadsheetId) {
    console.error(`テンプレートが見つかりません: ${monthKey}`);
    return `テンプレートが見つかりません: ${monthKey}`;
  }

  console.log(`テンプレートID: ${templateSpreadsheetId}`);

  // 月フォルダ確保 → テンプレコピー
  const monthFolderId = ensureMonthFolder_(monthKey, CONFIG.COPIES_PARENT_FOLDER_ID);
  const fileName = `${monthKey}_${teacher.name}_シフト提出`;
  const newSpreadsheetId = copyTemplateSpreadsheet_(monthFolderId, templateSpreadsheetId, fileName);
  const sheetUrl = `https://docs.google.com/spreadsheets/d/${newSpreadsheetId}/edit`;

  console.log(`新しいシートを作成しました: ${sheetUrl}`);

  // 「リンクを知っているすべての人」を編集者に設定
  setAnyoneWithLinkCanEdit_(newSpreadsheetId);

  // _META書き込み
  const submissionKey = `${monthKey}-${teacherId || normalizeNameKey_(teacher.name)}`;
  writeMetaToTeacherSheet_(newSpreadsheetId, CONFIG.META_SHEET_NAME, {
    MASTER_SPREADSHEET_ID: CONFIG.MASTER_SPREADSHEET_ID,
    SUBMISSIONS_SHEET_NAME: CONFIG.SHEET_SUBMISSIONS,
    SUBMISSION_KEY: submissionKey,
    MONTH_KEY: monthKey,
    TEACHER_ID: teacherId,
    TEACHER_NAME: teacher.name,
  });

  // 講師名をG3に直接設定
  try {
    const teacherSs = SpreadsheetApp.openById(newSpreadsheetId);
    const inputSheet = teacherSs.getSheetByName('Input');
    if (inputSheet) {
      inputSheet.getRange('G3').setValue(teacher.name);
    }
  } catch (e) {
    console.error('Failed to set teacher name in G3:', e);
  }

  // Submissionsを更新
  const existingSubmission = findSubmissionByKey_(master, submissionKey);
  if (existingSubmission) {
    updateSubmission_(master, existingSubmission.row, existingSubmission.header, {
      sheetUrl: sheetUrl,
      status: 'created',
    });
    console.log(`Submissionsを更新しました: row ${existingSubmission.row}`);
  } else {
    console.error('既存のSubmissionが見つかりません');
    return '既存のSubmissionが見つかりません';
  }

  // LINEにURL送信
  if (lineUserId) {
    pushLine_(lineUserId,
      `【シフト提出URL（${monthKey}）】\n${sheetUrl}\n\n※編集するには登録したGmailでGoogleにログインしてください。\n入力後、☑（提出）を入れてください。`
    );
    console.log('LINE通知を送信しました');
  } else {
    console.log('LINE User IDが登録されていないため、LINE通知は送信しませんでした');
  }

  return `修復完了: ${teacher.name} (${monthKey})\nシートURL: ${sheetUrl}`;
}

/**
 * 大久保先生に登録完了メッセージを送る関数（管理者用）
 * Google Apps Scriptエディタで直接実行
 */
function sendMessageToOkubo() {
  const master = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
  const teacher = findTeacherByName_(master, '大久保拓海');

  if (!teacher) {
    console.error('大久保先生が見つかりません');
    return '大久保先生が見つかりません';
  }

  console.log(`大久保先生の情報:`);
  console.log(`- 氏名: ${teacher.name}`);
  console.log(`- teacherId: ${teacher.teacherId}`);
  console.log(`- メール: ${teacher.email || '(未登録)'}`);
  console.log(`- LINE User ID: ${teacher.lineUserId || '(未登録)'}`);

  if (!teacher.lineUserId) {
    console.error('大久保先生のLINE User IDが登録されていません');
    return '大久保先生のLINE User IDが登録されていません';
  }

  const lastName = extractLastName_(teacher.name);
  let message = `登録OK：${lastName}先生\n今後はこのLINEでシフト連絡します。${MESSAGE_SPREADSHEET_APP}`;

  if (teacher.email) {
    message += `\n\nメールアドレスを登録しました：${teacher.email}`;
  }

  pushLine_(teacher.lineUserId, message);
  console.log('LINE通知を送信しました');
  return `${lastName}先生にメッセージを送信しました`;
}

/**
 * 落合将生先生を登録してLINE通知を送る関数（管理者用）
 */
function registerOchiai() {
  const lineUserId = 'U6889bd1984ab7697850c7be19e1f5f74';
  return registerTeacherManually_('落合将生', 'masaki.180228@gmail.com', lineUserId);
}

/**
 * 末河翔真先生に登録完了メッセージを送る関数（管理者用）
 * Google Apps Scriptエディタで直接実行
 */
function sendMessageToSuekawa() {
  const lineUserId = 'U8ce9ba286178abdf94096e08051942b3';
  const email = 'shoma180808@gmail.com';
  const message = `登録OK：末河先生\n今後はこのLINEでシフト連絡します。${MESSAGE_SPREADSHEET_APP}\n\nメールアドレスを登録しました：${email}`;
  pushLine_(lineUserId, message);
  console.log('LINE通知を送信しました');
  return '末河先生にメッセージを送信しました';
}

/**
 * 落合先生と末河先生に初回メッセージを送る関数（管理者用）
 * Google Apps Scriptエディタで直接実行
 */
function sendInitialMessageToNewTeachers() {
  const teachers = [
    { name: '落合', lineUserId: 'U6889bd1984ab7697850c7be19e1f5f74', email: 'masaki.180228@gmail.com' },
    { name: '末河', lineUserId: 'U8ce9ba286178abdf94096e08051942b3', email: 'shoma180808@gmail.com' }
  ];

  const results = [];
  for (const teacher of teachers) {
    const message = `登録OK：${teacher.name}先生\n今後はこのLINEでシフト連絡します。${MESSAGE_SPREADSHEET_APP}\n\nメールアドレスを登録しました：${teacher.email}`;
    pushLine_(teacher.lineUserId, message);
    console.log(`${teacher.name}先生にメッセージを送信しました`);
    results.push(`${teacher.name}先生: 送信完了`);
  }

  return results.join('\n');
}

/**
 * 落合先生と末河先生にシフト申請依頼メッセージを送る関数（管理者用）
 * Google Apps Scriptエディタで直接実行
 */
function sendShiftRequestToNewTeachers() {
  const monthKey = '2026-02';
  const teachers = [
    { name: '落合', lineUserId: 'U6889bd1984ab7697850c7be19e1f5f74' },
    { name: '末河', lineUserId: 'U8ce9ba286178abdf94096e08051942b3' }
  ];

  const results = [];
  for (const teacher of teachers) {
    const message = `【シフト申請のお願い】\n${teacher.name}先生、${monthKey}のシフト申請をお願いします。`;
    pushLine_(teacher.lineUserId, message);
    console.log(`${teacher.name}先生にシフト申請依頼を送信しました`);
    results.push(`${teacher.name}先生: 送信完了`);
  }

  return results.join('\n');
}

/**
 * 吉本先生に登録完了メッセージを送る関数（管理者用）
 * Google Apps Scriptエディタで直接実行
 */
function sendMessageToYoshimoto() {
  const lineUserId = 'U6a3995aafcc36c0a275b9621f70ca7a2';
  const email = 'jibenxingda8@gmail.com';
  const message = `登録OK：吉本先生\n今後はこのLINEでシフト連絡します。${MESSAGE_SPREADSHEET_APP}\n\nメールアドレスを登録しました：${email}`;
  pushLine_(lineUserId, message);
  console.log('吉本先生にメッセージを送信しました');
  return '吉本先生にメッセージを送信しました';
}

/**
 * 最後のWebhookリクエストを確認する関数（デバッグ用）
 */
function checkLastBody() {
  const props = PropertiesService.getScriptProperties();
  const lastBody = props.getProperty('LAST_BODY') || '';
  const hitCount = props.getProperty('HIT_COUNT') || '0';

  console.log('HIT_COUNT:', hitCount);
  console.log('LAST_BODY:', lastBody);

  try {
    const payload = JSON.parse(lastBody);
    console.log('Parsed payload:', JSON.stringify(payload, null, 2));

    if (payload.events && payload.events.length > 0) {
      const event = payload.events[0];
      console.log('User ID:', event.source?.userId);
      console.log('Message:', event.message?.text);
    }
  } catch (e) {
    console.log('Parse error:', e);
  }

  return { hitCount, lastBody };
}
