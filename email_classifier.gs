// =====================================================================
// email_classifier.gs
// 「要処理」ラベルの未読メールを Claude API で分類し、
// スプレッドシートへ記録 + Slack 通知を行う
// =====================================================================

var LABEL_PENDING   = '要処理';
var LABEL_DONE      = '処理済み';
var LOG_SHEET_NAME  = 'メールログ';
var ERR_SHEET_NAME  = 'エラーログ';
var CLAUDE_MODEL    = 'claude-haiku-4-5-20251001';
var CLAUDE_ENDPOINT = 'https://api.anthropic.com/v1/messages';

// ---------------------------------------------------------------------
// メイン処理：「要処理」ラベルの未読メールを一件ずつ処理する
// ---------------------------------------------------------------------
function classifyEmails() {
  var props       = PropertiesService.getScriptProperties();
  var claudeKey   = props.getProperty('CLAUDE_API_KEY');
  var slackUrl    = props.getProperty('SLACK_WEBHOOK_URL');

  if (!claudeKey || !slackUrl) {
    logError(null, 'スクリプトプロパティが未設定です（CLAUDE_API_KEY / SLACK_WEBHOOK_URL）');
    return;
  }

  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet    = getOrCreateSheet(ss, LOG_SHEET_NAME,  ['受信日時', '送信者', '件名', '分類', '要約']);
  var errSheet    = getOrCreateSheet(ss, ERR_SHEET_NAME,  ['発生日時', 'メッセージID', 'エラー内容']);

  var pendingLabel = GmailApp.getUserLabelByName(LABEL_PENDING);
  if (!pendingLabel) {
    logError(errSheet, null, '「' + LABEL_PENDING + '」ラベルが Gmail に存在しません');
    return;
  }

  // 「処理済み」ラベルが存在しない場合は作成する
  var doneLabel = GmailApp.getUserLabelByName(LABEL_DONE) || GmailApp.createLabel(LABEL_DONE);

  // 「要処理」ラベルのスレッドを最大 50 件取得（GAS の制限に配慮）
  var threads = pendingLabel.getThreads(0, 50);

  threads.forEach(function(thread) {
    // スレッド内の未読メッセージだけを処理対象にする
    var messages = thread.getMessages().filter(function(m) {
      return m.isUnread();
    });

    messages.forEach(function(message) {
      try {
        processMessage(message, claudeKey, slackUrl, logSheet);
      } catch (e) {
        logError(errSheet, message.getId(), e.message);
      }
    });

    // スレッド単位でラベルを付け替える
    thread.addLabel(doneLabel);
    thread.removeLabel(pendingLabel);
    thread.markRead();
  });

  Logger.log('メール分類処理が完了しました。処理スレッド数: ' + threads.length);
}

// ---------------------------------------------------------------------
// 一通のメールを Claude API で分類・要約し、ログ記録と Slack 通知を行う
// ---------------------------------------------------------------------
function processMessage(message, claudeKey, slackUrl, logSheet) {
  var receivedAt = message.getDate();
  var sender     = message.getFrom();
  var subject    = message.getSubject();
  var body       = message.getPlainBody();

  // 本文が長すぎる場合は先頭 3000 文字に切り詰めてトークンを節約する
  var trimmedBody = body.length > 3000 ? body.substring(0, 3000) + '…（以下省略）' : body;

  // Claude API で分類・要約を取得する
  var result = callClaudeApi(claudeKey, subject, trimmedBody);

  // スプレッドシートに記録する
  logSheet.appendRow([
    Utilities.formatDate(receivedAt, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm'),
    sender,
    subject,
    result.category,
    result.summary
  ]);

  // Slack に通知する
  notifySlack(slackUrl, sender, subject, result.category, result.summary);
}

// ---------------------------------------------------------------------
// Claude API を呼び出し、メールの分類と要約を返す
// 戻り値: { category: string, summary: string }
// ---------------------------------------------------------------------
function callClaudeApi(apiKey, subject, body) {
  var prompt = [
    'あなたはメール分類アシスタントです。',
    '以下のメールを「クレーム」「質問」「注文」「その他」のいずれか一つに分類し、内容を日本語で100字以内に要約してください。',
    '必ず次のJSON形式のみで回答してください。説明文は不要です。',
    '{"category":"分類","summary":"要約"}',
    '',
    '【件名】',
    subject,
    '',
    '【本文】',
    body
  ].join('\n');

  var payload = {
    model: CLAUDE_MODEL,
    max_tokens: 256,
    messages: [
      { role: 'user', content: prompt }
    ]
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(CLAUDE_ENDPOINT, options);
  var code     = response.getResponseCode();

  if (code !== 200) {
    throw new Error('Claude API エラー (HTTP ' + code + '): ' + response.getContentText());
  }

  var responseJson = JSON.parse(response.getContentText());
  var rawText      = responseJson.content[0].text.trim();

  // レスポンスからJSONを抽出する（前後に余分なテキストが混入しても対応）
  var jsonMatch = rawText.match(/\{[\s\S]*\}/);
  if (!jsonMatch) {
    throw new Error('Claude API のレスポンスを JSON として解析できませんでした: ' + rawText);
  }

  var parsed = JSON.parse(jsonMatch[0]);

  // 想定外の分類値が返ってきた場合は「その他」にフォールバックする
  var validCategories = ['クレーム', '質問', '注文', 'その他'];
  var category = validCategories.indexOf(parsed.category) !== -1 ? parsed.category : 'その他';

  return {
    category: category,
    summary:  parsed.summary || ''
  };
}

// ---------------------------------------------------------------------
// Slack Incoming Webhook で担当者に通知する
// ---------------------------------------------------------------------
function notifySlack(webhookUrl, sender, subject, category, summary) {
  // 分類ごとに絵文字を変えて視認性を上げる
  var emoji = { 'クレーム': '🚨', '質問': '❓', '注文': '📦', 'その他': '📩' };
  var icon  = emoji[category] || '📩';

  var text = [
    icon + ' *新着メール分類通知*',
    '*送信者:* ' + sender,
    '*件名:* '   + subject,
    '*分類:* '   + category,
    '*要約:* '   + summary
  ].join('\n');

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ text: text }),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(webhookUrl, options);
  if (response.getResponseCode() !== 200) {
    throw new Error('Slack 通知エラー (HTTP ' + response.getResponseCode() + '): ' + response.getContentText());
  }
}

// ---------------------------------------------------------------------
// エラー情報を「エラーログ」シートに追記する
// ---------------------------------------------------------------------
function logError(errSheet, messageId, errorText) {
  Logger.log('[ERROR] ' + errorText);

  // errSheet が null の場合（シート取得前のエラー）はログ出力のみ
  if (!errSheet) return;

  var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
  errSheet.appendRow([timestamp, messageId || '', errorText]);
}

// ---------------------------------------------------------------------
// 指定した名前のシートを取得し、存在しなければヘッダー付きで新規作成する
// ---------------------------------------------------------------------
function getOrCreateSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sheet;
}

// ---------------------------------------------------------------------
// 5分ごとの時間駆動トリガーを登録する
// ※ GASエディタでこの関数を一度だけ手動実行してください
//    重複登録を防ぐため、同名の既存トリガーは削除してから再登録します
// ---------------------------------------------------------------------
function setupEvery5MinTrigger() {
  var targetFunctionName = 'classifyEmails';

  // 同名の既存トリガーをすべて削除して重複を防ぐ
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === targetFunctionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 5分ごとに実行されるトリガーを登録
  ScriptApp.newTrigger(targetFunctionName)
    .timeBased()
    .everyMinutes(5)
    .create();

  Logger.log('5分ごとのトリガーを登録しました。');
}
