// File: 02_Utils.js
/**
 * 共通ユーティリティ関数
 * 
 * シート操作、データ変換、ID生成などの共通処理を提供します。
 */

/**
 * シートを取得（存在しない場合は作成）
 * @param {string} sheetName - シート名
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} シートオブジェクト
 */
function getSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    Logger.log(`シート "${sheetName}" を作成しました`);
  }
  
  return sheet;
}

/**
 * シートを取得し、ヘッダーが存在しない場合は設定
 * @param {string} sheetName - シート名
 * @param {Array<string>} headers - ヘッダー配列
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} シートオブジェクト
 */
function getOrCreateSheet(sheetName, headers) {
  const sheet = getSheet(sheetName);
  
  // ヘッダー行が存在しない、または空の場合は設定
  const headerRow = sheet.getRange(1, 1, 1, headers.length);
  const existingHeaders = headerRow.getValues()[0];
  
  if (existingHeaders.every(h => !h || h === '')) {
    headerRow.setValues([headers]);
    // ヘッダー行を固定
    sheet.setFrozenRows(1);
    Logger.log(`シート "${sheetName}" のヘッダーを設定しました`);
  }
  
  return sheet;
}

/**
 * 電話番号を正規化（ハイフン・スペース除去）
 * @param {string} phone - 電話番号
 * @returns {string} 正規化された電話番号
 */
function normalizePhoneNumber(phone) {
  if (!phone || typeof phone !== 'string') {
    return '';
  }
  // ハイフン、スペース、括弧を除去
  return phone.replace(/[-\s()]/g, '');
}

/**
 * IDを生成（日時ベース）
 * @param {string} prefix - IDプレフィックス（例: 'CUST', 'CALL'）
 * @returns {string} 生成されたID（例: 'CUST_20240101_001'）
 */
function generateId(prefix) {
  const now = new Date();
  const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd');
  const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HHmmss');
  const random = Math.floor(Math.random() * 1000).toString().padStart(3, '0');
  return `${prefix}_${dateStr}_${timeStr}_${random}`;
}

/**
 * 日時をフォーマット
 * @param {Date} date - 日時オブジェクト
 * @param {string} format - フォーマット（'date' または 'datetime'）
 * @returns {string} フォーマットされた日時文字列
 */
function formatDateTime(date, format = 'datetime') {
  if (!date) {
    return '';
  }
  
  const timeZone = Session.getScriptTimeZone();
  const formatStr = format === 'date' 
    ? Config.DATE_FORMAT 
    : Config.DATETIME_FORMAT;
  
  return Utilities.formatDate(date, timeZone, formatStr);
}

/**
 * 配列（行データ）をオブジェクトに変換
 * @param {Array} row - 行データ配列
 * @param {Array<string>} headers - ヘッダー配列
 * @returns {Object} オブジェクト（{列名: 値}）
 */
function arrayToObject(row, headers) {
  const obj = {};
  headers.forEach((header, index) => {
    obj[header] = row[index] || '';
  });
  return obj;
}

/**
 * オブジェクトを配列（行データ）に変換
 * @param {Object} obj - オブジェクト
 * @param {Array<string>} headers - ヘッダー配列
 * @returns {Array} 行データ配列
 */
function objectToArray(obj, headers) {
  return headers.map(header => obj[header] || '');
}

/**
 * シートから全データをバッチ取得
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - シートオブジェクト
 * @param {number} startRow - 開始行（1始まり、ヘッダー行を除く）
 * @param {number} numRows - 取得行数（-1の場合は全行）
 * @returns {Array<Array>} データ配列
 */
function batchGetValues(sheet, startRow = 2, numRows = -1) {
  if (!sheet) {
    return [];
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) {
    return [];
  }
  
  const numCols = sheet.getLastColumn();
  if (numCols === 0) {
    return [];
  }
  
  const actualNumRows = numRows === -1 
    ? lastRow - startRow + 1 
    : Math.min(numRows, lastRow - startRow + 1);
  
  if (actualNumRows <= 0) {
    return [];
  }
  
  return sheet.getRange(startRow, 1, actualNumRows, numCols).getValues();
}

/**
 * シートにデータをバッチ書き込み
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - シートオブジェクト
 * @param {number} startRow - 開始行（1始まり）
 * @param {Array<Array>} values - 書き込むデータ配列
 */
function batchSetValues(sheet, startRow, values) {
  if (!sheet || !values || values.length === 0) {
    return;
  }
  
  const numRows = values.length;
  const numCols = values[0].length;
  
  if (numRows === 0 || numCols === 0) {
    return;
  }
  
  const range = sheet.getRange(startRow, 1, numRows, numCols);
  range.setValues(values);
}

/**
 * シートに行を追加（バッチ処理）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - シートオブジェクト
 * @param {Array<Array>} rows - 追加する行データ配列
 */
function batchAppendRows(sheet, rows) {
  if (!sheet || !rows || rows.length === 0) {
    return;
  }
  
  const numRows = rows.length;
  const numCols = rows[0].length;
  
  if (numRows === 0 || numCols === 0) {
    return;
  }
  
  const lastRow = sheet.getLastRow();
  const startRow = lastRow + 1;
  
  batchSetValues(sheet, startRow, rows);
}

/**
 * 日付文字列をDateオブジェクトに変換
 * @param {string|Date} dateValue - 日付文字列またはDateオブジェクト
 * @returns {Date|null} Dateオブジェクト（変換できない場合はnull）
 */
function parseDate(dateValue) {
  if (!dateValue) {
    return null;
  }
  
  if (dateValue instanceof Date) {
    return dateValue;
  }
  
  if (typeof dateValue === 'string') {
    // スプレッドシートの日付形式（yyyy-MM-dd や yyyy/MM/dd）をパース
    const date = new Date(dateValue);
    if (!isNaN(date.getTime())) {
      return date;
    }
  }
  
  return null;
}

/**
 * 今日の日付を取得（時刻は00:00:00）
 * @returns {Date} 今日の日付
 */
function getToday() {
  const now = new Date();
  now.setHours(0, 0, 0, 0);
  return now;
}

/**
 * 日付を比較（時刻を無視）
 * @param {Date} date1 - 日付1
 * @param {Date} date2 - 日付2
 * @returns {number} date1 < date2 なら -1, date1 > date2 なら 1, 等しいなら 0
 */
function compareDates(date1, date2) {
  const d1 = new Date(date1);
  d1.setHours(0, 0, 0, 0);
  const d2 = new Date(date2);
  d2.setHours(0, 0, 0, 0);
  
  if (d1 < d2) return -1;
  if (d1 > d2) return 1;
  return 0;
}

/**
 * 指定した列の値で行を検索
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - シートオブジェクト
 * @param {number} searchCol - 検索列（1始まり）
 * @param {*} searchValue - 検索値
 * @param {number} startRow - 検索開始行（1始まり、デフォルト: 2）
 * @returns {number} 見つかった行番号（見つからない場合は-1）
 */
function findRowByValue(sheet, searchCol, searchValue, startRow = 2) {
  if (!sheet) {
    return -1;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) {
    return -1;
  }
  
  const range = sheet.getRange(startRow, searchCol, lastRow - startRow + 1, 1);
  const values = range.getValues();
  
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === searchValue) {
      return startRow + i;
    }
  }
  
  return -1;
}

/**
 * 指定した列の値で複数行を検索
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - シートオブジェクト
 * @param {number} searchCol - 検索列（1始まり）
 * @param {*} searchValue - 検索値
 * @param {number} startRow - 検索開始行（1始まり、デフォルト: 2）
 * @returns {Array<number>} 見つかった行番号の配列
 */
function findAllRowsByValue(sheet, searchCol, searchValue, startRow = 2) {
  const rows = [];
  if (!sheet) {
    return rows;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) {
    return rows;
  }
  
  const range = sheet.getRange(startRow, searchCol, lastRow - startRow + 1, 1);
  const values = range.getValues();
  
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === searchValue) {
      rows.push(startRow + i);
    }
  }
  
  return rows;
}

/**
 * Utilsオブジェクト
 * すべてのユーティリティ関数をメソッドとして提供
 */
const Utils = {
  getSheet: getSheet,
  getOrCreateSheet: getOrCreateSheet,
  normalizePhoneNumber: normalizePhoneNumber,
  generateId: generateId,
  formatDateTime: formatDateTime,
  arrayToObject: arrayToObject,
  objectToArray: objectToArray,
  batchGetValues: batchGetValues,
  batchSetValues: batchSetValues,
  batchAppendRows: batchAppendRows,
  parseDate: parseDate,
  getToday: getToday,
  compareDates: compareDates,
  findRowByValue: findRowByValue,
  findAllRowsByValue: findAllRowsByValue
};
