/************************************************************
 * line.gs
 ************************************************************/

function getLineToken_() {
  const token = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  if (!token) throw new Error('LINE_CHANNEL_ACCESS_TOKEN が未設定です（スクリプトプロパティ）。');
  return token;
}

/** LINE Reply（返信） */
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
    // LINE APIの失敗は通知対象外（再試行可能なため）
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
