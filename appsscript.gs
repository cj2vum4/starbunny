// =============================================
// 腸寶圖 POS - Google Apps Script
// =============================================

const LINE_TOKEN = 'xQjzGk3kDZn72uSiM8WRcPWXOalwkEbwrsIuwJXZixWvwy3rGhjqQrTxDeNT6A1UEzj+TPBmY3BKKJZUwvriQ3DQmKqvAAXFAx8oGFzLN91eIdq/N4PrD/ZiRRd9Uk2hwQe0g7pFjr/wla+gWHotEwdB04t89/1O/w1cDnyilFU=';
const SHEET_NAME = '訂單紀錄';

// ── GET：讀取今日訂單 ──────────────────────────
function doGet(e) {
  const sheet = getSheet();
  const today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd');
  const data = sheet.getDataRange().getValues();

  const orders = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    const timeStr = data[i][0] instanceof Date
      ? Utilities.formatDate(data[i][0], 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss')
      : data[i][0].toString();
    if (!timeStr.startsWith(today)) continue;
    orders.push({
      row:    i + 1,
      time:   timeStr,
      items:  data[i][1].toString(),
      total:  Number(data[i][2]),
      cash:   (data[i][3] !== '' && data[i][3] != null) ? Number(data[i][3]) : null,
      change: (data[i][4] !== '' && data[i][4] != null) ? Number(data[i][4]) : null,
    });
  }

  return ContentService
    .createTextOutput(JSON.stringify(orders))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── POST：新增 / 刪除 ─────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.action === 'delete') {
      const sheet = getSheet();
      sheet.deleteRow(Number(data.row));
      return ok();
    }

    // action === 'save'（預設）
    const sheet = getSheet();
    const now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
    const items = data.items.map(i => `${i.name}×${i.qty}($${i.subtotal})`).join('、');
    sheet.appendRow([now, items, data.total, data.cash ?? '', data.change ?? '']);

    return ok();
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── 每日報表（每天晚上 10 點觸發）─────────────
function sendDailyReport() {
  const sheet = getSheet();
  if (!sheet || sheet.getLastRow() <= 1) {
    broadcastLine('🌭 腸寶圖 每日報表\n今日尚無訂單紀錄。');
    return;
  }

  const today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd');
  const data = sheet.getDataRange().getValues();

  let total = 0, count = 0;
  const itemMap = {};

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    const t = data[i][0] instanceof Date ? Utilities.formatDate(data[i][0], 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss') : data[i][0].toString();
    if (!t.startsWith(today)) continue;
    total += Number(data[i][2]);
    count++;
    data[i][1].toString().split('、').forEach(p => {
      const m = p.match(/^(.+?)×(\d+)/);
      if (m) itemMap[m[1]] = (itemMap[m[1]] || 0) + parseInt(m[2]);
    });
  }

  const topItems = Object.entries(itemMap)
    .sort((a, b) => b[1] - a[1]).slice(0, 5)
    .map(([n, q]) => `  ${n}：${q} 份`).join('\n');

  broadcastLine([
    '🌭 腸寶圖 每日報表',
    `📅 ${today}`,
    `📦 訂單數：${count} 筆`,
    `💰 今日營業額：$${total}`,
    topItems ? `\n🏆 熱銷品項：\n${topItems}` : ''
  ].filter(Boolean).join('\n'));
}

// ── 月報表（每月 1 號晚上 10:10 觸發）──────────
function sendMonthlyReport() {
  const sheet = getSheet();
  const today = new Date();
  const lastMonth = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  const year = lastMonth.getFullYear();
  const month = String(lastMonth.getMonth() + 1).padStart(2, '0');
  const prefix = `${year}/${month}`;

  if (!sheet || sheet.getLastRow() <= 1) {
    broadcastLine(`🌭 腸寶圖 月報表\n📅 ${year}年${parseInt(month)}月\n尚無訂單紀錄。`);
    return;
  }

  const data = sheet.getDataRange().getValues();
  let total = 0, count = 0;
  const itemMap = {}, dailyMap = {};

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    const t = data[i][0] instanceof Date ? Utilities.formatDate(data[i][0], 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss') : data[i][0].toString();
    if (!t.startsWith(prefix)) continue;
    total += Number(data[i][2]);
    count++;
    const day = t.substring(0, 10);
    dailyMap[day] = (dailyMap[day] || 0) + Number(data[i][2]);
    data[i][1].toString().split('、').forEach(p => {
      const m = p.match(/^(.+?)×(\d+)/);
      if (m) itemMap[m[1]] = (itemMap[m[1]] || 0) + parseInt(m[2]);
    });
  }

  const topItems = Object.entries(itemMap)
    .sort((a, b) => b[1] - a[1]).slice(0, 5)
    .map(([n, q]) => `  ${n}：${q} 份`).join('\n');

  const avgDaily = Object.keys(dailyMap).length
    ? Math.round(total / Object.keys(dailyMap).length) : 0;

  broadcastLine([
    '🌭 腸寶圖 月報表',
    `📅 ${year}年${parseInt(month)}月`,
    `📦 總訂單數：${count} 筆`,
    `💰 月營業額：$${total}`,
    `📊 日均營業額：$${avgDaily}`,
    topItems ? `\n🏆 熱銷品項 TOP5：\n${topItems}` : ''
  ].filter(Boolean).join('\n'));
}

// ── 工具函式 ──────────────────────────────────
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['時間', '品項明細', '訂單合計', '收款金額', '找零']);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
  }
  return sheet;
}

function ok() {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function broadcastLine(message) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/broadcast', {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + LINE_TOKEN
    },
    payload: JSON.stringify({ messages: [{ type: 'text', text: message }] }),
    muteHttpExceptions: true
  });
}

function testDailyReport()   { sendDailyReport(); }
function testMonthlyReport() { sendMonthlyReport(); }

function testDoGet() {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    const raw = data[i][0];
    Logger.log('row=%s | type=%s | isDate=%s | toString=%s | formatted=%s',
      i + 1,
      typeof raw,
      raw instanceof Date,
      raw.toString(),
      raw instanceof Date ? Utilities.formatDate(raw, 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss') : raw
    );
  }
}
