// File: 07_CallManagementService.js
/**
 * ============================================================================
 * 【実行タイミング】手動実行（カスタムメニュー「架電管理シート設定」「キャンセルリスト同期」）
 * 【役割】インサイドセールス用架電管理シートの作成・管理
 * ============================================================================
 * 
 * キャンセルリスト（別スプレッドシート）からデータを取得し、
 * メイン管理ブック内の「IS_架電管理」シートに同期します。
 * 
 * 【実行方法】
 * - カスタムメニュー: 「IS管理システム」→「架電管理シート設定」
 * - カスタムメニュー: 「IS管理システム」→「キャンセルリスト同期」
 * 
 * 【依存関係】
 * - Config.gs に依存（シート名、列定義、キャンセルリスト設定など）
 * - Logger.gs に依存（ログ出力）
 * - Utils.gs に依存（シート操作、データ変換など）
 * 
 * 【処理フロー】
 * 1. 架電管理シートの作成・初期化（setupCallManagementSheet）
 * 2. キャンセルリストからデータを取得（getCancelListData_）
 * 3. 架電管理シートに同期（syncCancelListToCallSheet）
 * 
 * 【提供関数】
 * - setupCallManagementSheet(): 架電管理シートの作成・初期化
 * - syncCancelListToCallSheet(): キャンセルリストを架電管理シートに同期
 * - getCancelListData_(): キャンセルリストからデータを取得（内部関数）
 */

/**
 * 架電管理シートの作成・初期化
 * カスタムメニューから実行可能
 */
function setupCallManagementSheet() {
  const functionName = 'setupCallManagementSheet';
  logInfo('架電管理シートの設定を開始しました', functionName);
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = Config.SHEET_NAMES.CALL_MANAGEMENT;
    
    // シートの存在確認
    let sheet = ss.getSheetByName(sheetName);
    const isNewSheet = !sheet;
    
    if (isNewSheet) {
      // 新規作成
      sheet = ss.insertSheet(sheetName);
      logInfo(`シート "${sheetName}" を新規作成しました`, functionName);
    } else {
      logInfo(`シート "${sheetName}" は既に存在します`, functionName);
    }
    
    // ヘッダー行を設定
    const headers = Config.CALL_MANAGEMENT_HEADERS;
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    const existingHeaders = headerRange.getValues()[0];
    
    // 既存のヘッダーと比較して、不足している列があれば追加
    let needsUpdate = false;
    if (existingHeaders.length < headers.length) {
      needsUpdate = true;
    } else {
      for (let i = 0; i < headers.length; i++) {
        if (existingHeaders[i] !== headers[i]) {
          needsUpdate = true;
          break;
        }
      }
    }
    
    if (needsUpdate || isNewSheet) {
      headerRange.setValues([headers]);
      logInfo('ヘッダー行を設定しました', functionName);
    }
    
    // ヘッダー行を固定
    sheet.setFrozenRows(1);
    
    // データ検証（プルダウン）を設定
    setupDataValidation(sheet);
    
    // 列幅の調整（オプション）
    adjustColumnWidths(sheet);
    
    logInfo('架電管理シートの設定が完了しました', functionName);
    
    SpreadsheetApp.getUi().alert(
      '設定完了',
      `架電管理シート "${sheetName}" の設定が完了しました。`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (e) {
    logError('架電管理シートの設定でエラーが発生しました', functionName, e);
    SpreadsheetApp.getUi().alert(
      'エラー',
      `設定中にエラーが発生しました。\nログシート（LOGS）を確認してください。\n\nエラー: ${e.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    throw e;
  }
}

/**
 * データ検証（プルダウン）を設定
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - シートオブジェクト
 */
function setupDataValidation(sheet) {
  const functionName = 'setupDataValidation';
  
  try {
    // ステータス列（M列）にプルダウンを設定
    const statusCol = Config.CALL_MANAGEMENT_COLUMNS.STATUS + 1;
    const statusValues = Object.values(Config.CALL_MANAGEMENT_STATUS);
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(statusValues, true)
      .setAllowInvalid(false)
      .build();
    
    // 既存データがある場合は2行目以降に適用
    const lastRow = sheet.getLastRow();
    const dataStartRow = lastRow > 1 ? 2 : 2;
    const dataEndRow = Math.max(lastRow, 1000); // 最大1000行まで
    
    if (dataEndRow >= dataStartRow) {
      sheet.getRange(dataStartRow, statusCol, dataEndRow - dataStartRow + 1, 1)
        .setDataValidation(statusRule);
    }
    
    // ネタランク列（N列）にプルダウンを設定
    const noteRankCol = Config.CALL_MANAGEMENT_COLUMNS.NOTE_RANK + 1;
    const noteRankRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(Config.CALL_MANAGEMENT_NOTE_RANKS, true)
      .setAllowInvalid(false)
      .build();
    
    if (dataEndRow >= dataStartRow) {
      sheet.getRange(dataStartRow, noteRankCol, dataEndRow - dataStartRow + 1, 1)
        .setDataValidation(noteRankRule);
    }
    
    logInfo('データ検証（プルダウン）を設定しました', functionName);
    
  } catch (e) {
    logWarn('データ検証の設定でエラーが発生しました（処理は続行します）', functionName);
    logError('データ検証エラー', functionName, e);
  }
}

/**
 * 列幅を調整
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - シートオブジェクト
 */
function adjustColumnWidths(sheet) {
  try {
    // 各列の推奨幅を設定（オプション）
    const columnWidths = {
      A: 120,  // customer_id
      B: 150,  // customer_name
      C: 200,  // email
      D: 120,  // phone_number
      E: 100,  // list_added_date
      F: 100,  // source_type
      G: 150,  // lead_source_id
      H: 150,  // system_note
      I: 100,  // assigned_person
      J: 100,  // first_call_date
      K: 100,  // last_call_date
      L: 80,   // call_count
      M: 100,  // status
      N: 80,   // note_rank
      O: 100,  // next_call_date
      P: 300   // detail
    };
    
    for (const col in columnWidths) {
      const colIndex = col.charCodeAt(0) - 64; // A=1, B=2, ...
      sheet.setColumnWidth(colIndex, columnWidths[col]);
    }
    
  } catch (e) {
    // 列幅調整は失敗しても問題ない
    logWarn('列幅の調整でエラーが発生しました（処理は続行します）', 'adjustColumnWidths');
  }
}

/**
 * キャンセルリストからデータを取得（内部関数）
 * @returns {Array<Array>} キャンセルリストのデータ（二次元配列）
 */
function getCancelListData_() {
  const functionName = 'getCancelListData_';
  
  // 定数の確認
  if (Config.CANCEL_LIST_SPREADSHEET_ID === 'TODO: キャンセルリストのスプレッドシートID' ||
      Config.CANCEL_LIST_SHEET_NAME === 'TODO: キャンセルリストのシート名') {
    throw new Error(
      'キャンセルリストの設定が完了していません。\n' +
      'Config.jsのCANCEL_LIST_SPREADSHEET_IDとCANCEL_LIST_SHEET_NAMEを設定してください。'
    );
  }
  
  try {
    // 別スプレッドシートを開く
    const cancelListSs = SpreadsheetApp.openById(Config.CANCEL_LIST_SPREADSHEET_ID);
    const cancelListSheet = cancelListSs.getSheetByName(Config.CANCEL_LIST_SHEET_NAME);
    
    if (!cancelListSheet) {
      throw new Error(
        `キャンセルリストシート "${Config.CANCEL_LIST_SHEET_NAME}" が見つかりません。`
      );
    }
    
    // データを取得
    const lastRow = cancelListSheet.getLastRow();
    if (lastRow < 2) {
      logWarn('キャンセルリストにデータがありません', functionName);
      return [];
    }
    
    // ヘッダー行をスキップしてデータを取得
    // 想定される列構成（A列から）:
    // A列: 顧客ID(メール)
    // B列: 氏名
    // C列: 電話番号
    // D列: 商談予定日
    // E列: ステータス
    // F列: リスト追加日
    const dataRange = cancelListSheet.getRange(2, 1, lastRow - 1, 6);
    const rawData = dataRange.getValues();
    
    logInfo(`キャンセルリストから ${rawData.length} 件のデータを取得しました`, functionName);
    
    return rawData;
    
  } catch (e) {
    logError('キャンセルリストの取得でエラーが発生しました', functionName, e);
    throw e;
  }
}

/**
 * キャンセルリストを架電管理シートに同期
 * カスタムメニューから実行可能
 */
function syncCancelListToCallSheet() {
  const functionName = 'syncCancelListToCallSheet';
  logInfo('キャンセルリストの同期を開始しました', functionName);
  
  try {
    // 架電管理シートを取得（存在しない場合は作成）
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = Config.SHEET_NAMES.CALL_MANAGEMENT;
    let sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      logInfo('架電管理シートが存在しないため、作成します', functionName);
      setupCallManagementSheet();
      sheet = ss.getSheetByName(sheetName);
    }
    
    // キャンセルリストからデータを取得
    const cancelListData = getCancelListData_();
    
    if (cancelListData.length === 0) {
      logWarn('キャンセルリストにデータがありません', functionName);
      SpreadsheetApp.getUi().alert(
        '情報',
        'キャンセルリストにデータがありません。',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // 既存の架電管理シートのデータを取得（重複チェック用）
    const existingData = getExistingCallManagementData_(sheet);
    
    // 顧客マスタを取得（顧客IDの生成・検索用）
    const customerSheet = Utils.getSheet(Config.SHEET_NAMES.CUSTOMER);
    const allCustomers = customerSheet && customerSheet.getLastRow() >= 2
      ? batchGetValues(customerSheet, 2)
      : [];
    
    // 顧客マスタからメールアドレス/電話番号で顧客IDを検索するマップを作成
    const customerIdMap = createCustomerIdMap_(allCustomers);
    
    // 同期用のデータを準備
    const newRows = [];
    const updatedRows = [];
    let skippedCount = 0;
    
    for (const row of cancelListData) {
      try {
        // キャンセルリストの列マッピングに基づいてデータを取得
        const mapping = Config.CANCEL_LIST_COLUMN_MAPPING;
        const email = row[mapping.CUSTOMER_ID_EMAIL] || '';
        const fullName = row[mapping.FULL_NAME] || '';
        const phoneNumber = row[mapping.PHONE_NUMBER] || '';
        const appointmentDate = row[mapping.APPOINTMENT_DATE] || '';
        const status = row[mapping.STATUS] || '';
        const listAddedDate = row[mapping.LIST_ADDED_DATE] || '';
        
        // 必須項目チェック（メールアドレスまたは電話番号が必要）
        if (!email && !phoneNumber) {
          skippedCount++;
          logWarn(
            `メールアドレスと電話番号が両方空の行をスキップしました: ${fullName}`,
            functionName
          );
          continue;
        }
        
        // 顧客IDを取得または生成
        let customerId = findCustomerId_(email, phoneNumber, customerIdMap);
        if (!customerId) {
          // 顧客IDが見つからない場合は、メールアドレスまたは電話番号から生成
          customerId = generateCustomerId_(email, phoneNumber);
        }
        
        // 既存データをチェック
        const existingRowIndex = findExistingRow_(existingData, customerId, email, phoneNumber);
        
        // リスト追加日をパース
        const listAddedDateObj = listAddedDate ? parseDate(listAddedDate) : new Date();
        const listAddedDateStr = listAddedDateObj
          ? formatDateTime(listAddedDateObj, 'date')
          : formatDateTime(new Date(), 'date');
        
        // システム側でセットする列のデータ
        const systemData = [
          customerId,                    // A列: customer_id
          fullName,                       // B列: customer_name
          email,                          // C列: email
          phoneNumber,                    // D列: phone_number
          listAddedDateStr,               // E列: list_added_date
          'キャンセル',                    // F列: source_type
          '',                             // G列: lead_source_id（後で設定可能）
          `商談予定日: ${appointmentDate || '未設定'}, ステータス: ${status || '未設定'}`  // H列: system_note
        ];
        
        if (existingRowIndex >= 0) {
          // 既存行を更新（システム側の列のみ更新、担当者入力列は保持）
          const existingRow = existingData[existingRowIndex];
          const existingRowNumber = existingRowIndex + 2; // ヘッダー行を考慮
          
          // 既存の担当者入力列のデータを取得
          const userInputData = [
            existingRow[Config.CALL_MANAGEMENT_COLUMNS.ASSIGNED_PERSON] || '',  // I列: assigned_person
            existingRow[Config.CALL_MANAGEMENT_COLUMNS.FIRST_CALL_DATE] || '', // J列: first_call_date
            existingRow[Config.CALL_MANAGEMENT_COLUMNS.LAST_CALL_DATE] || '',   // K列: last_call_date
            existingRow[Config.CALL_MANAGEMENT_COLUMNS.CALL_COUNT] || '',       // L列: call_count
            existingRow[Config.CALL_MANAGEMENT_COLUMNS.STATUS] || '',          // M列: status
            existingRow[Config.CALL_MANAGEMENT_COLUMNS.NOTE_RANK] || '',        // N列: note_rank
            existingRow[Config.CALL_MANAGEMENT_COLUMNS.NEXT_CALL_DATE] || '',   // O列: next_call_date
            existingRow[Config.CALL_MANAGEMENT_COLUMNS.DETAIL] || ''            // P列: detail
          ];
          
          const updatedRow = [...systemData, ...userInputData];
          updatedRows.push({
            rowNumber: existingRowNumber,
            data: updatedRow
          });
          
        } else {
          // 新規行を追加
          const newRow = [
            ...systemData,
            '',  // I列: assigned_person（担当者が入力）
            '',  // J列: first_call_date（担当者が入力）
            '',  // K列: last_call_date（担当者が入力）
            '',  // L列: call_count（担当者が入力）
            Config.CALL_MANAGEMENT_STATUS.NOT_STARTED,  // M列: status（デフォルト: 未着手）
            '',  // N列: note_rank（担当者が入力）
            '',  // O列: next_call_date（担当者が入力）
            ''   // P列: detail（担当者が入力）
          ];
          newRows.push(newRow);
        }
        
      } catch (e) {
        logError(
          `キャンセルリストの行処理でエラーが発生しました: ${e.message}`,
          functionName,
          e
        );
        skippedCount++;
      }
    }
    
    // データを書き込み
    if (newRows.length > 0) {
      const lastRow = sheet.getLastRow();
      const startRow = lastRow + 1;
      batchAppendRows(sheet, newRows);
      logInfo(`${newRows.length} 件の新規レコードを追加しました`, functionName);
    }
    
    if (updatedRows.length > 0) {
      for (const item of updatedRows) {
        sheet.getRange(item.rowNumber, 1, 1, item.data.length).setValues([item.data]);
      }
      logInfo(`${updatedRows.length} 件の既存レコードを更新しました`, functionName);
    }
    
    // データ検証を再設定（新規行に適用）
    if (newRows.length > 0) {
      setupDataValidation(sheet);
    }
    
    logInfo(
      `キャンセルリストの同期が完了しました: 新規=${newRows.length}件, 更新=${updatedRows.length}件, スキップ=${skippedCount}件`,
      functionName
    );
    
    SpreadsheetApp.getUi().alert(
      '同期完了',
      `キャンセルリストの同期が完了しました。\n` +
      `新規: ${newRows.length}件\n` +
      `更新: ${updatedRows.length}件\n` +
      `スキップ: ${skippedCount}件`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (e) {
    logError('キャンセルリストの同期でエラーが発生しました', functionName, e);
    SpreadsheetApp.getUi().alert(
      'エラー',
      `同期中にエラーが発生しました。\nログシート（LOGS）を確認してください。\n\nエラー: ${e.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    throw e;
  }
}

/**
 * 既存の架電管理シートのデータを取得（重複チェック用）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - シートオブジェクト
 * @returns {Array<Array>} 既存データの配列
 */
function getExistingCallManagementData_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  
  return batchGetValues(sheet, 2);
}

/**
 * 顧客IDマップを作成（メールアドレス/電話番号 → 顧客ID）
 * @param {Array<Array>} allCustomers - 顧客マスタの全データ
 * @returns {Object} 顧客IDマップ {email: customerId, phone: customerId}
 */
function createCustomerIdMap_(allCustomers) {
  const map = {};
  
  for (const row of allCustomers) {
    const customerId = row[Config.CUSTOMER_COLUMNS.CUSTOMER_ID];
    const email = row[Config.CUSTOMER_COLUMNS.EMAIL];
    const phoneNumber = row[Config.CUSTOMER_COLUMNS.PHONE_NUMBER];
    
    if (email) {
      const normalizedEmail = email.toLowerCase().trim();
      map[`email:${normalizedEmail}`] = customerId;
    }
    
    if (phoneNumber) {
      const normalizedPhone = normalizePhoneNumber(phoneNumber);
      if (normalizedPhone) {
        map[`phone:${normalizedPhone}`] = customerId;
      }
    }
  }
  
  return map;
}

/**
 * 顧客IDを検索
 * @param {string} email - メールアドレス
 * @param {string} phoneNumber - 電話番号
 * @param {Object} customerIdMap - 顧客IDマップ
 * @returns {string|null} 顧客ID（見つからない場合はnull）
 */
function findCustomerId_(email, phoneNumber, customerIdMap) {
  // メールアドレスで検索
  if (email) {
    const normalizedEmail = email.toLowerCase().trim();
    const key = `email:${normalizedEmail}`;
    if (customerIdMap[key]) {
      return customerIdMap[key];
    }
  }
  
  // 電話番号で検索
  if (phoneNumber) {
    const normalizedPhone = normalizePhoneNumber(phoneNumber);
    if (normalizedPhone) {
      const key = `phone:${normalizedPhone}`;
      if (customerIdMap[key]) {
        return customerIdMap[key];
      }
    }
  }
  
  return null;
}

/**
 * 顧客IDを生成（メールアドレスまたは電話番号から）
 * @param {string} email - メールアドレス
 * @param {string} phoneNumber - 電話番号
 * @returns {string} 生成された顧客ID
 */
function generateCustomerId_(email, phoneNumber) {
  // メールアドレスがある場合はそれを使用
  if (email && email.trim()) {
    return `CUST_EMAIL_${email.toLowerCase().trim().replace(/[^a-z0-9]/g, '_')}`;
  }
  
  // 電話番号がある場合はそれを使用
  if (phoneNumber) {
    const normalizedPhone = normalizePhoneNumber(phoneNumber);
    if (normalizedPhone) {
      return `CUST_PHONE_${normalizedPhone}`;
    }
  }
  
  // どちらもない場合はタイムスタンプベース
  return generateId(Config.ID_PREFIXES.CUSTOMER);
}

/**
 * 既存行を検索
 * @param {Array<Array>} existingData - 既存データ
 * @param {string} customerId - 顧客ID
 * @param {string} email - メールアドレス
 * @param {string} phoneNumber - 電話番号
 * @returns {number} 既存行のインデックス（見つからない場合は-1）
 */
function findExistingRow_(existingData, customerId, email, phoneNumber) {
  const normalizedEmail = email ? email.toLowerCase().trim() : '';
  const normalizedPhone = phoneNumber ? normalizePhoneNumber(phoneNumber) : '';
  
  for (let i = 0; i < existingData.length; i++) {
    const row = existingData[i];
    const existingCustomerId = row[Config.CALL_MANAGEMENT_COLUMNS.CUSTOMER_ID];
    const existingEmail = row[Config.CALL_MANAGEMENT_COLUMNS.EMAIL] || '';
    const existingPhone = row[Config.CALL_MANAGEMENT_COLUMNS.PHONE_NUMBER] || '';
    
    // 顧客IDで一致
    if (existingCustomerId === customerId) {
      return i;
    }
    
    // メールアドレスで一致
    if (normalizedEmail && existingEmail) {
      if (existingEmail.toLowerCase().trim() === normalizedEmail) {
        return i;
      }
    }
    
    // 電話番号で一致
    if (normalizedPhone && existingPhone) {
      const normalizedExistingPhone = normalizePhoneNumber(existingPhone);
      if (normalizedExistingPhone === normalizedPhone) {
        return i;
      }
    }
  }
  
  return -1;
}

