/**
 * 不只讀書的讀書會｜療癒罪咎 —— 報名表後端
 * 貼到 Google 試算表的「擴充功能 → Apps Script」，部署成「網頁應用程式」。
 */

const CONFIG = {
  SHEET_NAME: '報名',              // 資料寫進這個工作表（沒有會自動建立）
  NOTIFY_EMAIL: '',                // 有人報名時通知你：填你的 email；留空＝不通知
  SEND_CONFIRMATION: true,         // 是否寄確認信給報名者
  COURSE_NAME: '不只讀書的讀書會｜療癒罪咎',
  COURSE_TIME: '每週五晚上 20:00–22:30，2026/11/27 起，共 16 週',
  COURSE_BREAKS: '停課：12/25、1/1，以及 2 月整月（3/5 恢復）',
};

// [欄位代號, 試算表欄名]：順序就是試算表的欄位順序
const FIELDS = [
  ['name',      '姓名'],
  ['nickname',  '課堂上希望被稱呼'],
  ['email',     'Email'],
  ['phone',     '手機'],
  ['line',      'LINE ID'],
  ['returning', '是否參加過讀書會'],
  ['source',    '從哪裡得知'],
  ['message',   '想說的話'],
  ['consent',   '錄影說明已了解'],
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');

    // 防機器人：隱藏欄位被填了就假裝成功、不寫入
    if (data.website) return json_({ ok: true });

    if (!data.name || !data.email || !data.phone) {
      return json_({ ok: false, error: '請填寫姓名、Email 與手機。' });
    }

    const sheet = getSheet_();
    const row = [new Date()].concat(FIELDS.map(f => clean_(data[f[0]])));
    sheet.appendRow(row);

    if (CONFIG.SEND_CONFIRMATION) sendConfirmation_(data);
    if (CONFIG.NOTIFY_EMAIL) {
      MailApp.sendEmail({
        to: CONFIG.NOTIFY_EMAIL,
        subject: '【新報名】' + CONFIG.COURSE_NAME + '：' + data.name,
        body: FIELDS.map(f => f[1] + '：' + (data[f[0]] || '')).join('\n'),
      });
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// 用瀏覽器打開網址時看到的訊息（用來測試部署是否成功）
function doGet() {
  return json_({ ok: true, service: CONFIG.COURSE_NAME + ' 報名表' });
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['報名時間'].concat(FIELDS.map(f => f[1])));
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, FIELDS.length + 1).setFontWeight('bold');
  }
  return sheet;
}

// 清理輸入：去頭尾空白、限長度；避免被當成公式；保留手機開頭的 0
function clean_(v) {
  if (v === undefined || v === null) return '';
  let s = String(v).trim().slice(0, 2000);
  if (/^[=+\-@]/.test(s) || /^0\d+$/.test(s)) s = "'" + s;
  return s;
}

function sendConfirmation_(data) {
  const who = data.nickname || data.name;
  const body = [
    who + '，你好：',
    '',
    '已收到你的報名，歡迎一起來讀《療癒罪咎》。',
    '',
    '時間：' + CONFIG.COURSE_TIME,
    CONFIG.COURSE_BREAKS,
    '',
    '上課連結與相關資訊，會在開課前另行通知。',
    '',
    '沒有誰是壞人，你也不是。',
    '',
    '不只讀書的讀書會',
  ].join('\n');
  MailApp.sendEmail({
    to: data.email,
    subject: '報名成功｜' + CONFIG.COURSE_NAME,
    body: body,
  });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
