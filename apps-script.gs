/**
 * نظام حصر ومتابعة غياب الطالبات — المتوسطة التاسعة والستون
 * هذا الكود يُلصق كاملاً داخل محرر Apps Script المرتبط بجدول بيانات Google Sheets.
 * راجع ملف README.md لخطوات النشر كاملة.
 */

// غيّري هذا الرمز إلى رمز سري من اختيارك — هو نفسه "رمز الدخول" الذي تدخلينه في لوحة المتابعة
const ACCESS_KEY = 'hayat69school';

const SHEET_NAME = 'Responses';
const DRIVE_FOLDER_NAME = 'مرفقات غياب الطالبات';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.action === 'reset') {
      if (data.key !== ACCESS_KEY) return jsonResponse({ status: 'unauthorized' });
      const sheet = getSheet();
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
      return jsonResponse({ status: 'ok' });
    }

    const sheet = getSheet();
    let fileUrl = '';
    if (data.fileBase64 && data.fileName) {
      fileUrl = saveFile(data.fileBase64, data.fileName, data.fileMime);
    }
    sheet.appendRow([
      new Date(),
      data.date || '',
      data.grade || '',
      data.section || '',
      data.studentName || '',
      data.type || '',
      data.reason || '',
      fileUrl
    ]);
    return jsonResponse({ status: 'ok' });
  } catch (err) {
    return jsonResponse({ status: 'error', message: String(err) });
  }
}

function doGet(e) {
  const key = e.parameter.key;
  if (key !== ACCESS_KEY) {
    return jsonResponse({ status: 'unauthorized' });
  }
  const sheet = getSheet();
  const values = sheet.getDataRange().getValues();
  const rows = values.slice(1).filter(r => r[4]).map(r => ({
    timestamp: r[0],
    date: formatDate(r[1]),
    grade: r[2],
    section: r[3],
    studentName: r[4],
    type: r[5],
    reason: r[6],
    fileUrl: r[7]
  }));
  return jsonResponse({ status: 'ok', records: rows });
}

function formatDate(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return val;
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['الوقت', 'التاريخ', 'الصف', 'الفصل', 'اسم الطالبة', 'النوع', 'السبب', 'رابط المرفق']);
  }
  return sheet;
}

function saveFile(base64, name, mime) {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(DRIVE_FOLDER_NAME);
  const bytes = Utilities.base64Decode(base64);
  const blob = Utilities.newBlob(bytes, mime, name);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
