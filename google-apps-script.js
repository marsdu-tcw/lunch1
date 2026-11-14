/**
 * 2026 聖誕季｜小組聖誕餐會報名系統 - Google 試算表後端腳本 (Google Apps Script)
 *
 * 【使用步驟】
 * 1. 新增一個空白的 Google 試算表 (Google Sheets)
 * 2. 點選上方選單「擴充功能」 -> 「Apps Script」
 * 3. 將本檔案所有程式碼貼上（覆蓋原本的 function myFunction()）
 * 4. 點選上方「部署」 -> 「新增部署作業」
 * 5. 左側齒輪選擇「網頁應用程式 (Web App)」：
 *    - 說明：聖誕餐會報名 API
 *    - 執行身分：「我 (Me)」
 *    - 誰可以存取：「所有人 (Anyone)」
 * 6. 點選「部署」，授權後複製產生的「網頁應用程式網址 (https://script.google.com/macros/s/.../exec)」
 * 7. 貼入網頁右上角「⚙️ 設定」或 index.html 的 GOOGLE_SCRIPT_URL 即可即時連動！
 */

const SHEET_NAME = '報名名單';

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const headers = [
      '報名時間',
      '姓名',
      '聯絡電話',
      '所屬小組',
      '大人人數',
      '小孩人數',
      '總人數',
      '是否攜帶新朋友',
      '新朋友人數',
      '飲食需求',
      '過敏或飲食備註'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#7c1d1d')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function doGet(e) {
  try {
    const sheet = getOrCreateSheet();
    const data = sheet.getDataRange().getValues();
    const rows = data.slice(1).map((row, idx) => ({
      id: 'row-' + (idx + 2),
      createdAt: row[0] ? Utilities.formatDate(new Date(row[0]), 'Asia/Taipei', 'yyyy/MM/dd HH:mm') : '',
      name: String(row[1] || ''),
      phone: String(row[2] || ''),
      group: String(row[3] || ''),
      adults: Number(row[4] || 0),
      children: Number(row[5] || 0),
      total: Number(row[6] || 0),
      bringGuest: String(row[7] || '否'),
      guestCount: Number(row[8] || 0),
      diet: String(row[9] || '葷食'),
      dietNote: String(row[10] || '')
    }));

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, registrations: rows.reverse() }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const sheet = getOrCreateSheet();
    const payload = JSON.parse(e.postData.contents);
    const adults = Number(payload.adults || 0);
    const children = Number(payload.children || 0);
    const guestCount = Number(payload.guestCount || 0);
    const total = adults + children;

    sheet.appendRow([
      new Date(),
      payload.name || '',
      payload.phone || '',
      payload.group || '',
      adults,
      children,
      total,
      payload.bringGuest || '否',
      guestCount,
      payload.diet || '葷食',
      payload.dietNote || ''
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
