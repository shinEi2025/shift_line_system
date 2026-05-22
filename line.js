/************************************************************
 * line.gs
 ************************************************************/

function getLineToken_() {
  const token = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  if (!token) throw new Error('LINE_CHANNEL_ACCESS_TOKEN が未設定です（スクリプトプロパティ）。');
  return token;
}

/** LINE Reply（返信）- 成功/失敗を返す */
function replyLine_(replyToken, text) {
  const token = getLineToken_();
  const url = 'https://api.line.me/v2/bot/message/reply';
  const payload = { replyToken, messages: [{ type: 'text', text }] };

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: `Bearer ${token}` },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  const code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    const errorMsg = `LINE返信に失敗しました: HTTP ${code} / ${res.getContentText()}`;
    logError_(errorMsg, 'replyLine_', { replyToken: replyToken.slice(0, 20) + '...', httpCode: code });
    return false; // 失敗
  }
  return true; // 成功
}

/**
 * LINE Reply（返信）with Push fallback
 * replyLine_が失敗した場合、pushLine_でフォールバックする
 * @param {string} replyToken - LINEのreplyToken
 * @param {string} userId - LINEのユーザーID
 * @param {string} text - 送信するメッセージ
 */
function replyLineWithFallback_(replyToken, userId, text) {
  const success = replyLine_(replyToken, text);
  if (!success && userId) {
    console.log(`[replyLineWithFallback_] Reply failed, falling back to push for userId: ${userId.slice(0, 10)}...`);
    pushLine_(userId, text);
  }
}

/** LINE Push */
function pushLine_(toUserId, text) {
  const token = getLineToken_();
  const url = 'https://api.line.me/v2/bot/message/push';
  const payload = { to: toUserId, messages: [{ type: 'text', text }] };

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: `Bearer ${token}` },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  const code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    const errorMsg = `LINEプッシュに失敗しました: HTTP ${code} / ${res.getContentText()}`;
    logError_(errorMsg, 'pushLine_', { toUserId: toUserId.slice(0, 20) + '...', httpCode: code });
    // LINE APIの失敗は通知対象外（再試行可能なため）
  }
}
