/**
 * 前日TEL不都合案件 - GAS Web App
 * Server-side code
 */

var SHEET_NAME = 'data';
var HEADERS = ['登録日時', '担当', '理由', '対応', '対応者', '備考', '更新日時', '更新者'];

/**
 * Web App entry point - serves HTML
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('前日TEL不都合案件')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Get or create the data sheet with headers
 */
function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  return sheet;
}

/**
 * Get the current user's email
 */
function getCurrentUser() {
  try {
    return Session.getActiveUser().getEmail();
  } catch (e) {
    return '';
  }
}

/**
 * Register a new record
 * @param {Object} formData - { 担当, 理由, 対応, 対応者, 備考 }
 * @return {Object} result
 */
function registerRecord(formData) {
  try {
    var sheet = getSheet_();
    var now = new Date();
    var user = getCurrentUser();

    var row = [
      now,                    // 登録日時
      formData['担当'],       // 担当
      formData['理由'],       // 理由
      formData['対応'],       // 対応
      formData['対応者'],     // 対応者
      formData['備考'] || '', // 備考
      now,                    // 更新日時
      user                    // 更新者
    ];

    sheet.appendRow(row);

    // Format datetime columns
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 1).setNumberFormat('yyyy/MM/dd HH:mm:ss');
    sheet.getRange(lastRow, 7).setNumberFormat('yyyy/MM/dd HH:mm:ss');

    return { success: true, message: '登録が完了しました。' };
  } catch (e) {
    return { success: false, message: '登録に失敗しました: ' + e.message };
  }
}

/**
 * Get all records from the sheet
 * @return {Object} { headers: [], data: [] }
 */
function getAllRecords() {
  try {
    var sheet = getSheet_();
    var lastRow = sheet.getLastRow();

    if (lastRow <= 1) {
      return { success: true, headers: HEADERS, data: [] };
    }

    var dataRange = sheet.getRange(2, 1, lastRow - 1, HEADERS.length);
    var values = dataRange.getValues();

    // Convert dates to formatted strings and reverse for newest first
    var data = values.map(function(row, index) {
      return {
        rowIndex: index + 2, // actual row number in sheet
        登録日時: row[0] instanceof Date ? Utilities.formatDate(row[0], Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss') : String(row[0]),
        担当: String(row[1]),
        理由: String(row[2]),
        対応: String(row[3]),
        対応者: String(row[4]),
        備考: String(row[5]),
        更新日時: row[6] instanceof Date ? Utilities.formatDate(row[6], Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss') : String(row[6]),
        更新者: String(row[7])
      };
    });

    // Sort by registration datetime descending (newest first)
    data.sort(function(a, b) {
      return b.登録日時.localeCompare(a.登録日時);
    });

    return { success: true, headers: HEADERS, data: data };
  } catch (e) {
    return { success: false, message: 'データの取得に失敗しました: ' + e.message, headers: HEADERS, data: [] };
  }
}

/**
 * Update an existing record
 * @param {number} rowIndex - the row number in the sheet
 * @param {Object} formData - { 担当, 理由, 対応, 対応者, 備考 }
 * @return {Object} result
 */
function updateRecord(rowIndex, formData) {
  try {
    var sheet = getSheet_();
    var now = new Date();
    var user = getCurrentUser();

    // Update editable columns (columns 2-6: 担当, 理由, 対応, 対応者, 備考)
    sheet.getRange(rowIndex, 2).setValue(formData['担当']);
    sheet.getRange(rowIndex, 3).setValue(formData['理由']);
    sheet.getRange(rowIndex, 4).setValue(formData['対応']);
    sheet.getRange(rowIndex, 5).setValue(formData['対応者']);
    sheet.getRange(rowIndex, 6).setValue(formData['備考'] || '');

    // Update timestamp and user
    sheet.getRange(rowIndex, 7).setValue(now);
    sheet.getRange(rowIndex, 7).setNumberFormat('yyyy/MM/dd HH:mm:ss');
    sheet.getRange(rowIndex, 8).setValue(user);

    return { success: true, message: '更新が完了しました。' };
  } catch (e) {
    return { success: false, message: '更新に失敗しました: ' + e.message };
  }
}

/**
 * Get unique values for filter dropdowns
 * @return {Object} { 担当: [], 理由: [], 対応: [], 対応者: [] }
 */
function getFilterOptions() {
  try {
    var sheet = getSheet_();
    var lastRow = sheet.getLastRow();

    if (lastRow <= 1) {
      return { success: true, 担当: [], 理由: [], 対応: [], 対応者: [] };
    }

    var data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();

    var tantouSet = {};
    var riyuuSet = {};
    var taiouSet = {};
    var taioushaSet = {};

    data.forEach(function(row) {
      if (row[1]) tantouSet[String(row[1])] = true;
      if (row[2]) riyuuSet[String(row[2])] = true;
      if (row[3]) taiouSet[String(row[3])] = true;
      if (row[4]) taioushaSet[String(row[4])] = true;
    });

    return {
      success: true,
      担当: Object.keys(tantouSet).sort(),
      理由: Object.keys(riyuuSet).sort(),
      対応: Object.keys(taiouSet).sort(),
      対応者: Object.keys(taioushaSet).sort()
    };
  } catch (e) {
    return { success: false, 担当: [], 理由: [], 対応: [], 対応者: [] };
  }
}
