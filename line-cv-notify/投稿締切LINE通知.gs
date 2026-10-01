// ============================================================
// 投稿管理スプシ 締切チェック → LINE通知
// 毎朝トリガーで checkDeadlines() を実行
// ============================================================

// このスクリプト導入日。これより前の締切は(未整理データが多いため)一切対象にしない
const CUTOFF_DATE = new Date(2026, 7, 10);

const CONFIG = {
  spreadsheets: [
    {
      id: '1CyZas3MhqxKDNdKnh8UHV4sU1XzIL1rlN3oRoEaslc0',
      sheetName: '腸活・投稿作成',
      titleCol: 2, // B: 企画題名
      tasks: [ // 「項目」列がある方(トーク系)
        { name: '原稿',      dateCol: 5,  statusCol: 6 },
        { name: '撮影',      dateCol: 8,  statusCol: 9 },
        { name: '編集',      dateCol: 11, statusCol: 12 },
        { name: 'サムネイル', dateCol: 14, statusCol: 15 },
      ],
    },
    {
      id: '1Jc9C0GUtwvafTgb9hMxiSzXb2HG47J1_IACMznmbyZY',
      sheetName: '腸活・投稿作成',
      titleCol: 2, // B: 企画題名
      tasks: [ // 「項目」列がない方(CV系)は1列ずつ前にずれる、サムネイル列なし
        { name: '原稿',      dateCol: 4,  statusCol: 5 },
        { name: '撮影',      dateCol: 7,  statusCol: 8 },
        { name: '編集',      dateCol: 10, statusCol: 11 },
      ],
    },
    {
      id: '1UuBSCoF6BbNL1TTa77Nj0_RgiynR5i2yxcYCGr_1e5A',
      sheetName: '投稿作成',
      titleCol: 2, // B: 企画題名
      tasks: [ // 「項目」列がない方、サムネイル列なし
        { name: '原稿',      dateCol: 4,  statusCol: 5 },
        { name: '撮影',      dateCol: 7,  statusCol: 8 },
        { name: '編集',      dateCol: 10, statusCol: 11 },
      ],
    },
  ],
};

// 「腸活ダイエット」シートで、企画(E列)に CV とある行の投稿日付(C列)が今日なら通知する
const CV_POST_CONFIG = {
  sheetName: '腸活ダイエット',
  dateCol: 3, // C: 投稿日付
  planCol: 5, // E: 企画
  accounts: [
    { name: 'かめさん', id: '1CyZas3MhqxKDNdKnh8UHV4sU1XzIL1rlN3oRoEaslc0' },
    { name: 'ころさん', id: '1Jc9C0GUtwvafTgb9hMxiSzXb2HG47J1_IACMznmbyZY' },
  ],
};

function checkDeadlines() {
  try {
    checkCvPostDays();
  } catch (e) {
    console.error(e); // CV通知が失敗しても締切チェックは続ける
  }

  const today = startOfToday_();
  const overdue = [];
  const dueToday = [];
  const errors = [];

  CONFIG.spreadsheets.forEach(conf => {
    let label = conf.id;
    try {
      const ss = SpreadsheetApp.openById(conf.id);
      label = ss.getName();
      const sheet = ss.getSheetByName(conf.sheetName);
      if (!sheet) throw new Error('シート「' + conf.sheetName + '」が見つかりません');

      const rows = sheet.getDataRange().getValues();
      rows.forEach(row => {
        const title = String(row[conf.titleCol - 1] || '').trim();
        if (!title) return; // 空行・月ラベル行はスキップ

        conf.tasks.forEach(task => {
          const deadline = parseDate_(row[task.dateCol - 1], today);
          if (!deadline) return; // 日付でないセル(見出し行など)はスキップ
          if (deadline < CUTOFF_DATE) return; // 導入日より前の締切は対象外(過去分の未整理データを無視)

          const status = String(row[task.statusCol - 1] || '').trim();
          const urlNote = String(row[task.dateCol + 1] || '').trim(); // URL欄に「不要」等の注記がある場合も完了扱い
          if (status.indexOf('完了') !== -1 || urlNote.indexOf('不要') !== -1) return;

          const diffDays = Math.round((today - deadline) / 86400000);
          const item = {
            label: label, title: title, task: task.name,
            deadline: deadline, diffDays: diffDays,
          };
          if (diffDays > 0) overdue.push(item);
          else if (diffDays === 0) dueToday.push(item);
        });
      });
    } catch (e) {
      errors.push('⚠️ 「' + label + '」を読み取れませんでした: ' + e.message);
    }
  });

  if (overdue.length === 0 && dueToday.length === 0 && errors.length === 0) return;

  sendLine_(buildMessage_(today, overdue, dueToday, errors));
}

// 例: 本日はCVの投稿予定日です かめさん_CV21
function checkCvPostDays() {
  const lines = findCvPostLines_(startOfToday_());
  if (lines.length > 0) sendLine_(lines.join('\n'));
}

function findCvPostLines_(today) {
  const lines = [];

  CV_POST_CONFIG.accounts.forEach(acc => {
    try {
      const sheet = SpreadsheetApp.openById(acc.id).getSheetByName(CV_POST_CONFIG.sheetName);
      if (!sheet) throw new Error('シート「' + CV_POST_CONFIG.sheetName + '」が見つかりません');

      const rows = sheet.getDataRange().getValues();
      rows.forEach(row => {
        const plan = String(row[CV_POST_CONFIG.planCol - 1] || '').trim();
        if (!/CV/i.test(plan)) return;
        const postDate = parseDate_(row[CV_POST_CONFIG.dateCol - 1], today);
        if (!postDate || postDate.getTime() !== today.getTime()) return;

        const m = plan.match(/CV\D*?(\d+)/i); // 「CVリール21」も「CV21」と表示
        lines.push('本日はCVの投稿予定日です ' + acc.name + '_' + (m ? 'CV' + m[1] : plan));
      });
    } catch (e) {
      lines.push('⚠️ ' + acc.name + 'の「' + CV_POST_CONFIG.sheetName + '」を読み取れませんでした: ' + e.message);
    }
  });

  return lines;
}

function buildMessage_(today, overdue, dueToday, errors) {
  const now = new Date();
  const hh = ('0' + now.getHours()).slice(-2);
  const mm = ('0' + now.getMinutes()).slice(-2);
  const lines = ['【投稿管理アラート ' + (today.getMonth() + 1) + '/' + today.getDate() + ' ' + hh + ':' + mm + '】'];

  if (overdue.length > 0) {
    overdue.sort((a, b) => b.diffDays - a.diffDays);
    lines.push('', '⚠️ 締切超過 (' + overdue.length + '件)');
    overdue.forEach(it => {
      lines.push('・' + it.label + ' ' + it.title + ' ' + it.task +
        '(期日' + fmtDate_(it.deadline) + '・' + it.diffDays + '日超過)');
    });
  }

  if (dueToday.length > 0) {
    lines.push('', '📅 本日締切 (' + dueToday.length + '件)');
    dueToday.forEach(it => {
      lines.push('・' + it.label + ' ' + it.title + ' ' + it.task);
    });
  }

  if (errors.length > 0) {
    lines.push('');
    errors.forEach(e => lines.push(e));
  }

  return lines.join('\n');
}

// "8/9（日）" のような文字列も、日付型セルも両方パースする
function parseDate_(cell, today) {
  if (cell instanceof Date && !isNaN(cell.getTime())) {
    return new Date(cell.getFullYear(), cell.getMonth(), cell.getDate());
  }
  const m = String(cell || '').match(/(\d{1,2})\/(\d{1,2})/);
  if (!m) return null;
  const d = new Date(today.getFullYear(), Number(m[1]) - 1, Number(m[2]));
  // 年またぎ補正: 今日から半年以上離れていたら前後の年とみなす
  const HALF_YEAR = 183 * 86400000;
  if (d - today > HALF_YEAR) d.setFullYear(d.getFullYear() - 1);
  else if (today - d > HALF_YEAR) d.setFullYear(d.getFullYear() + 1);
  return d;
}

function startOfToday_() {
  const now = new Date();
  return new Date(now.getFullYear(), now.getMonth(), now.getDate());
}

function fmtDate_(d) {
  return (d.getMonth() + 1) + '/' + d.getDate();
}

function sendLine_(text) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('LINE_TOKEN');
  if (!token) throw new Error('スクリプトプロパティ LINE_TOKEN を設定してください');

  // broadcast: このチャネルの友だち全員に送信(この用途専用ボットなので実質あなた宛て)
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/broadcast', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ messages: [{ type: 'text', text: text }] }),
  });
}

// ---- セットアップ用(1回だけ手動実行) ----

// LINE接続テスト
function testLine() {
  sendLine_('✅ テスト送信です。この通知が届けば設定OKです!');
}

// CV通知の確認(LINEには送らずログに出す)。日付はシートでCVが入っている日に変えて使う
function testCvPostDays() {
  console.log(findCvPostLines_(new Date(2026, 7, 10))); // 2026/8/10
}

// 毎朝9時のトリガーを設定
function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('checkDeadlines').timeBased().everyDays(1).atHour(9).create();
}

