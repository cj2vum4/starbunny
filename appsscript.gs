// =============================================
// 腸寶圖 POS - Google Apps Script
// =============================================

const LINE_TOKEN = 'xQjzGk3kDZn72uSiM8WRcPWXOalwkEbwrsIuwJXZixWvwy3rGhjqQrTxDeNT6A1UEzj+TPBmY3BKKJZUwvriQ3DQmKqvAAXFAx8oGFzLN91eIdq/N4PrD/ZiRRd9Uk2hwQe0g7pFjr/wla+gWHotEwdB04t89/1O/w1cDnyilFU=';
const SHEET_NAME = '訂單紀錄';

// ── 接收 POS 結帳資料 ──────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(['時間', '品項明細', '訂單合計']);
      sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    }

    const now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
    const items = data.items.map(i => `${i.name}×${i.qty}($${i.subtotal})`).join('、');
    sheet.appendRow([now, items, data.total]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── 每日報表（每天晚上 10 點觸發）─────────────
function sendDailyReport() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() <= 1) {
    broadcastLine('🌭 腸寶圖 每日報表\n今日尚無訂單紀錄。');
    return;
  }

  const today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd');
  const data = sheet.getDataRange().getValues();

  let total = 0;
  let count = 0;
  const itemMap = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const rowDate = row[0].toString().substring(0, 10);
    if (rowDate !== today) continue;

    total += Number(row[2]);
    count++;

    // 統計各品項銷售
    const parts = row[1].toString().split('、');
    parts.forEach(p => {
      const m = p.match(/^(.+?)×(\d+)/);
      if (m) {
        const name = m[1];
        const qty = parseInt(m[2]);
        itemMap[name] = (itemMap[name] || 0) + qty;
      }
    });
  }

  const topItems = Object.entries(itemMap)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5)
    .map(([name, qty]) => `  ${name}：${qty} 份`)
    .join('\n');

  const msg = [
    '🌭 腸寶圖 每日報表',
    `📅 ${today}`,
    `📦 訂單數：${count} 筆`,
    `💰 今日營業額：$${total}`,
    topItems ? `\n🏆 熱銷品項：\n${topItems}` : ''
  ].filter(Boolean).join('\n');

  broadcastLine(msg);
}

// ── 月報表（每月 1 號晚上 10:10 觸發）──────────
function sendMonthlyReport() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

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
  let total = 0;
  let count = 0;
  const itemMap = {};
  const dailyMap = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const rowStr = row[0].toString();
    if (!rowStr.startsWith(prefix)) continue;

    total += Number(row[2]);
    count++;

    const day = rowStr.substring(0, 10);
    dailyMap[day] = (dailyMap[day] || 0) + Number(row[2]);

    const parts = row[1].toString().split('、');
    parts.forEach(p => {
      const m = p.match(/^(.+?)×(\d+)/);
      if (m) {
        const name = m[1];
        const qty = parseInt(m[2]);
        itemMap[name] = (itemMap[name] || 0) + qty;
      }
    });
  }

  const topItems = Object.entries(itemMap)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5)
    .map(([name, qty]) => `  ${name}：${qty} 份`)
    .join('\n');

  const avgDaily = count > 0
    ? Math.round(total / Object.keys(dailyMap).length)
    : 0;

  const msg = [
    '🌭 腸寶圖 月報表',
    `📅 ${year}年${parseInt(month)}月`,
    `📦 總訂單數：${count} 筆`,
    `💰 月營業額：$${total}`,
    `📊 日均營業額：$${avgDaily}`,
    topItems ? `\n🏆 熱銷品項 TOP5：\n${topItems}` : ''
  ].filter(Boolean).join('\n');

  broadcastLine(msg);
}

// ── Line Broadcast ─────────────────────────────
function broadcastLine(message) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/broadcast', {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + LINE_TOKEN
    },
    payload: JSON.stringify({
      messages: [{ type: 'text', text: message }]
    }),
    muteHttpExceptions: true
  });
}

// ── 手動測試用（執行後確認 Line 有收到訊息）─────
function testDailyReport()   { sendDailyReport(); }
function testMonthlyReport() { sendMonthlyReport(); }
