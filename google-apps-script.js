/**
 * 2026 聖誕季｜小組聖誕餐會報名系統 - Google 試算表後端腳本 (Google Apps Script)
 *
 * 【欄位結構】
 * 報名時間 | 出席的弟兄或姊妹姓名 | 聯絡電話 | 所屬小組 | 總出席人數 | 是否攜帶新朋友 | 新朋友人數 | 新朋友姓名 | 備註（由哪位弟兄或姊妹帶來） | 飲食需求 | 過敏或飲食備註
 */

const SHEET_NAME = '報名名單';

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const headers = [
      '報名時間',
      '出席的弟兄或姊妹姓名',
      '聯絡電話',
      '所屬小組',
      '總出席人數',
      '是否攜帶新朋友',
      '新朋友人數',
      '新朋友姓名',
      '備註（由哪位弟兄或姊妹帶來）',
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
      memberName: String(row[1] || ''),
      phone: String(row[2] || ''),
      group: String(row[3] || ''),
      total: Number(row[4] || 1),
      bringGuest: String(row[5] || '否'),
      guestCount: Number(row[6] || 0),
      guestNames: String(row[7] || ''),
      broughtByNote: String(row[8] || ''),
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
    const memberName = payload.memberName || '';
    const guestCount = payload.bringGuest === '是' ? Number(payload.guestCount || 0) : 0;
    const guestNames = payload.bringGuest === '是' ? (payload.guestNames || '') : '';
    const total = 1 + guestCount;
    const broughtByNote = guestCount > 0
      ? `由 ${memberName} 帶來${guestNames ? '（新朋友：' + guestNames + '）' : '（共 ' + guestCount + ' 位新朋友）'}`
      : `${memberName} 本人出席（無攜伴）`;

    sheet.appendRow([
      new Date(),
      memberName,
      payload.phone || '',
      payload.group || '',
      total,
      payload.bringGuest || '否',
      guestCount,
      guestNames,
      broughtByNote,
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
