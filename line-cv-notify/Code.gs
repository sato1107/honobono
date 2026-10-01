/**
 * 「腸活ダイエット」シートで、企画列に CV と書かれた日に LINE へ通知する。
 *
 * 通知例：本日はCVの投稿予定日です かめさん_CV21
 *
 * 事前にスクリプトプロパティへ次の2つを登録する（README.md 参照）。
 *   LINE_CHANNEL_ACCESS_TOKEN … LINE公式アカウント（Messaging API）のチャネルアクセストークン
 *   LINE_TO                   … 通知先のユーザーID（U…）またはグループID（C…）
 */

const ACCOUNTS = [
  { name: 'かめさん', spreadsheetId: '1CyZas3MhqxKDNdKnh8UHV4sU1XzIL1rlN3oRoEaslc0' },
  { name: 'ころさん', spreadsheetId: '1Jc9C0GUtwvafTgb9hMxiSzXb2HG47J1_IACMznmbyZY' },
];

const SHEET_NAME = '腸活ダイエット';
const FIRST_DATA_ROW = 4; // 1〜3行目は見出し
const DATE_COL = 3; // C列：投稿日付
const PLAN_COL = 5; // E列：企画

const NOTIFY_HOUR = 8; // 毎朝この時台に通知

/** 毎日トリガーから呼ばれる本体 */
function notifyTodayCv() {
  const today = new Date();
  const lines = findCvLines(today);
  if (lines.length === 0) return;
  pushLine(lines.join('\n'));
}

/** 指定日に CV 投稿がある行を探し、通知文の配列を返す */
function findCvLines(date) {
  const lines = [];
  ACCOUNTS.forEach(function (account) {
    const sheet = SpreadsheetApp.openById(account.spreadsheetId).getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(account.name + 'のスプレッドシートに「' + SHEET_NAME + '」シートがありません');

    const lastRow = sheet.getLastRow();
    if (lastRow < FIRST_DATA_ROW) return;
    const numRows = lastRow - FIRST_DATA_ROW + 1;
    const dates = sheet.getRange(FIRST_DATA_ROW, DATE_COL, numRows, 1).getValues();
    const plans = sheet.getRange(FIRST_DATA_ROW, PLAN_COL, numRows, 1).getDisplayValues();

    for (let i = 0; i < numRows; i++) {
      const plan = String(plans[i][0]).trim();
      if (!/CV/i.test(plan)) continue;
      if (!isSameDay(dates[i][0], date)) continue;
      lines.push('本日はCVの投稿予定日です ' + account.name + '_' + cvLabel(plan));
    }
  });
  return lines;
}

/** 「CVリール21」「CV21」などを「CV21」に揃える。番号がなければ企画名をそのまま使う */
function cvLabel(plan) {
  const m = plan.match(/CV\D*?(\d+)/i);
  return m ? 'CV' + m[1] : plan;
}

/**
 * セルの日付と対象日が同じ日か判定する。
 * 日付型のセルは年まで比較し、「03/16」のような文字列は月日だけで比較する。
 */
function isSameDay(cell, date) {
  const tz = Session.getScriptTimeZone();
  if (cell instanceof Date) {
    return Utilities.formatDate(cell, tz, 'yyyyMMdd') === Utilities.formatDate(date, tz, 'yyyyMMdd');
  }
  const m = String(cell).trim().match(/^(\d{1,2})\/(\d{1,2})/);
  if (!m) return false;
  const month = Number(Utilities.formatDate(date, tz, 'M'));
  const day = Number(Utilities.formatDate(date, tz, 'd'));
  return Number(m[1]) === month && Number(m[2]) === day;
}

/** LINE Messaging API のプッシュメッセージで送信する */
function pushLine(text) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  const to = props.getProperty('LINE_TO');
  if (!token || !to) throw new Error('スクリプトプロパティに LINE_CHANNEL_ACCESS_TOKEN と LINE_TO を設定してください');

  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ to: to, messages: [{ type: 'text', text: text }] }),
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('LINE送信に失敗しました (' + res.getResponseCode() + '): ' + res.getContentText());
  }
}

/** 初回に1度だけ実行：毎朝の通知トリガーを作成する（重複作成しない） */
function setupTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(function (t) { return t.getHandlerFunction() === 'notifyTodayCv'; })
    .forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('notifyTodayCv').timeBased().everyDays(1).atHour(NOTIFY_HOUR).create();
}

/** 動作確認：LINEにテスト送信する */
function testPush() {
  pushLine('テスト通知です（CV投稿日通知）');
}

/** 動作確認：指定日に通知される内容をログに出す（LINEには送らない） */
function testFind() {
  const target = new Date(2026, 7, 10); // 2026/8/10（ころさん CV3）
  Logger.log(findCvLines(target));
}
