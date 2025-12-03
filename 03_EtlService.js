// File: 03_EtlService.js
/**
 * ETL処理（データ取込・統合）
 * 
 * 複数の元リストから顧客データを取り込み、
 * M_CUSTOMERとM_LEAD_SOURCEに統合します。
 */

/**
 * デバッグ用：ソース設定と元データシートの存在確認
 * カスタムメニューから実行可能（開発用）
 */
function testSourceConfigs() {
  const functionName = 'testSourceConfigs';
  logInfo('ソース設定のテストを開始しました', functionName);
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceConfigs = Config.SOURCE_CONFIGS;
    
    let message = '=== ソース設定テスト結果 ===\n\n';
    let hasError = false;
    
    for (const sourceConfig of sourceConfigs) {
      message += `【${sourceConfig.name}】\n`;
      
      // シートの存在確認
      let sourceSheet;
      if (sourceConfig.spreadsheetId) {
        try {
          const sourceSs = SpreadsheetApp.openById(sourceConfig.spreadsheetId);
          sourceSheet = sourceSs.getSheetByName(sourceConfig.sheetName);
        } catch (e) {
          message += `  ❌ エラー: 別スプレッドシートにアクセスできません (ID: ${sourceConfig.spreadsheetId})\n`;
          hasError = true;
          continue;
        }
      } else {
        sourceSheet = ss.getSheetByName(sourceConfig.sheetName);
      }
      
      if (!sourceSheet) {
        message += `  ❌ シートが見つかりません: "${sourceConfig.sheetName}"\n`;
        hasError = true;
        continue;
      }
      
      message += `  ✅ シート存在: "${sourceConfig.sheetName}"\n`;
      
      // データ行数の確認
      const lastRow = sourceSheet.getLastRow();
      const dataRowCount = Math.max(0, lastRow - sourceConfig.dataStartRow + 1);
      message += `  📊 データ行数: ${dataRowCount}行\n`;
      
      if (dataRowCount === 0) {
        message += `  ⚠️  警告: データが0件です\n`;
        hasError = true;
      }
      
      // ヘッダー行の確認
      if (lastRow >= sourceConfig.headerRow) {
        const headerRow = sourceSheet.getRange(
          sourceConfig.headerRow,
          1,
          1,
          sourceSheet.getLastColumn()
        ).getValues()[0];
        
        message += `  📋 ヘッダー行の列数: ${headerRow.length}列\n`;
        
        // マッピング列の存在確認
        const missingColumns = [];
        Object.keys(sourceConfig.mapping).forEach(key => {
          const columnName = sourceConfig.mapping[key];
          if (columnName && !headerRow.includes(columnName)) {
            missingColumns.push(`${key} → "${columnName}"`);
          }
        });
        
        if (missingColumns.length > 0) {
          message += `  ❌ 見つからない列: ${missingColumns.join(', ')}\n`;
          hasError = true;
        } else {
          message += `  ✅ すべての列が見つかりました\n`;
        }
      } else {
        message += `  ❌ ヘッダー行が存在しません\n`;
        hasError = true;
      }
      
      message += '\n';
    }
    
    if (hasError) {
      message += '⚠️  エラーまたは警告があります。上記を確認してください。\n';
    } else {
      message += '✅ すべての設定が正常です。\n';
    }
    
    logInfo(message, functionName);
    
    SpreadsheetApp.getUi().alert(
      'ソース設定テスト',
      message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (e) {
    logError('ソース設定テストでエラーが発生しました', functionName, e);
    SpreadsheetApp.getUi().alert(
      'エラー',
      `テスト実行中にエラーが発生しました。\nログシート（LOGS）を確認してください。\n\nエラー: ${e.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * ETL処理のメイン関数（カスタムメニューから実行）
 */
function executeEtl() {
  const functionName = 'executeEtl';
  logInfo('ETL処理を開始しました', functionName);
  
  try {
    // 各ソースからデータを取込
    const sourceConfigs = Config.SOURCE_CONFIGS;
    let totalProcessed = 0;
    let totalNewCustomers = 0;
    let totalUpdatedCustomers = 0;
    let totalNewLeadSources = 0;
    
    for (const sourceConfig of sourceConfigs) {
      logInfo(`ソース "${sourceConfig.name}" の処理を開始`, functionName);
      
      try {
        const result = importFromSource(sourceConfig);
        totalProcessed += result.processed;
        totalNewCustomers += result.newCustomers;
        totalUpdatedCustomers += result.updatedCustomers;
        totalNewLeadSources += result.newLeadSources;
        
        logInfo(
          `ソース "${sourceConfig.name}" の処理完了: ` +
          `処理件数=${result.processed}, ` +
          `新規顧客=${result.newCustomers}, ` +
          `更新顧客=${result.updatedCustomers}, ` +
          `新規リードソース=${result.newLeadSources}`,
          functionName
        );
        
        // API制限を避けるため、少し待機
        Utilities.sleep(500);
        
      } catch (e) {
        logError(
          `ソース "${sourceConfig.name}" の処理でエラーが発生しました`,
          functionName,
          e
        );
        // エラーが発生しても次のソースの処理を続行
      }
    }
    
    logInfo(
      `ETL処理が完了しました: ` +
      `総処理件数=${totalProcessed}, ` +
      `新規顧客=${totalNewCustomers}, ` +
      `更新顧客=${totalUpdatedCustomers}, ` +
      `新規リードソース=${totalNewLeadSources}`,
      functionName
    );
    
    // 処理結果をスプレッドシートに表示（オプション）
    SpreadsheetApp.getUi().alert(
      'ETL処理が完了しました',
      `処理件数: ${totalProcessed}\n` +
      `新規顧客: ${totalNewCustomers}\n` +
      `更新顧客: ${totalUpdatedCustomers}\n` +
      `新規リードソース: ${totalNewLeadSources}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (e) {
    logError('ETL処理で致命的なエラーが発生しました', functionName, e);
    throw e;
  }
}

/**
 * 個別ソースからのデータ取込
 * @param {Object} sourceConfig - ソース設定（Config.SOURCE_CONFIGSの要素）
 * @returns {Object} 処理結果 {processed, newCustomers, updatedCustomers, newLeadSources}
 */
function importFromSource(sourceConfig) {
  const functionName = 'importFromSource';
  
  // 元データのスプレッドシートとシートを取得
  let sourceSheet;
  if (sourceConfig.spreadsheetId) {
    // 別スプレッドシートの場合
    const sourceSs = SpreadsheetApp.openById(sourceConfig.spreadsheetId);
    sourceSheet = sourceSs.getSheetByName(sourceConfig.sheetName);
  } else {
    // 同じスプレッドシート内の場合
    sourceSheet = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName(sourceConfig.sheetName);
  }
  
  if (!sourceSheet) {
    throw new Error(`ソースシート "${sourceConfig.sheetName}" が見つかりません`);
  }
  
  // 元データを取得
  const lastRow = sourceSheet.getLastRow();
  if (lastRow < sourceConfig.dataStartRow) {
    logWarn(`ソース "${sourceConfig.name}" にデータがありません`, functionName);
    return {
      processed: 0,
      newCustomers: 0,
      updatedCustomers: 0,
      newLeadSources: 0
    };
  }
  
  // ヘッダー行を取得して列インデックスをマッピング
  const headerRow = sourceSheet.getRange(
    sourceConfig.headerRow,
    1,
    1,
    sourceSheet.getLastColumn()
  ).getValues()[0];
  
  const columnMap = {};
  Object.keys(sourceConfig.mapping).forEach(key => {
    const columnName = sourceConfig.mapping[key];
    if (columnName) {
      const colIndex = headerRow.indexOf(columnName);
      if (colIndex >= 0) {
        columnMap[key] = colIndex;
      } else {
        logWarn(
          `列 "${columnName}" がソース "${sourceConfig.name}" に見つかりません`,
          functionName
        );
      }
    }
  });
  
  // データ行を取得
  const dataRows = batchGetValues(
    sourceSheet,
    sourceConfig.dataStartRow,
    lastRow - sourceConfig.dataStartRow + 1
  );
  
  let processed = 0;
  let newCustomers = 0;
  let updatedCustomers = 0;
  let newLeadSources = 0;
  
  // 各データ行を処理
  for (const row of dataRows) {
    try {
      // データをマッピング
      const customerData = {
        lineName: columnMap.lineName !== undefined ? row[columnMap.lineName] : '',
        fullName: columnMap.fullName !== undefined ? row[columnMap.fullName] : '',
        phoneNumber: columnMap.phoneNumber !== undefined ? row[columnMap.phoneNumber] : '',
        email: columnMap.email !== undefined ? row[columnMap.email] : '',
        sourceType: sourceConfig.sourceType,
        sourceDetail: columnMap.sourceDetail !== undefined ? row[columnMap.sourceDetail] : sourceConfig.name,
        listAddedDate: new Date(),  // デフォルトは今日
        eventDate: columnMap.eventDate !== undefined && row[columnMap.eventDate]
          ? parseDate(row[columnMap.eventDate])
          : null
      };
      
      // 必須項目のチェック
      if (!customerData.lineName && !customerData.phoneNumber) {
        logWarn(
          `LINE名と電話番号が両方空の行をスキップしました（行: ${processed + sourceConfig.dataStartRow}）`,
          functionName
        );
        continue;
      }
      
      // 顧客マスタへの統合
      const mergeResult = mergeCustomer(customerData);
      if (mergeResult.isNew) {
        newCustomers++;
      } else {
        updatedCustomers++;
      }
      
      // リードソースの追加
      const leadSourceResult = addLeadSource(mergeResult.customerId, {
        sourceType: customerData.sourceType,
        sourceDetail: customerData.sourceDetail,
        listAddedDate: customerData.listAddedDate,
        eventDate: customerData.eventDate
      });
      
      if (leadSourceResult.isNew) {
        newLeadSources++;
      }
      
      processed++;
      
      // 大量データ処理時のAPI制限対策
      if (processed % 100 === 0) {
        Utilities.sleep(200);
      }
      
    } catch (e) {
      logError(
        `データ行の処理でエラーが発生しました（行: ${processed + sourceConfig.dataStartRow}）`,
        functionName,
        e
      );
      // エラーが発生しても次の行の処理を続行
    }
  }
  
  return {
    processed,
    newCustomers,
    updatedCustomers,
    newLeadSources
  };
}

/**
 * 顧客マスタへの統合（重複判定・更新）
 * @param {Object} customerData - 顧客データ
 * @returns {Object} {customerId, isNew} - 顧客IDと新規フラグ
 */
function mergeCustomer(customerData) {
  const functionName = 'mergeCustomer';
  
  const customerSheet = Utils.getOrCreateSheet(
    Config.SHEET_NAMES.CUSTOMER,
    Config.CUSTOMER_HEADERS
  );
  
  // 重複判定: 電話番号（正規化後）またはLINE名で検索
  const normalizedPhone = normalizePhoneNumber(customerData.phoneNumber);
  let existingRow = -1;
  let existingCustomerId = null;
  
  if (normalizedPhone) {
    // 電話番号で検索
    const phoneCol = Config.CUSTOMER_COLUMNS.PHONE_NUMBER + 1;  // 1始まりに変換
    const allPhones = batchGetValues(customerSheet, 2);
    for (let i = 0; i < allPhones.length; i++) {
      const existingPhone = normalizePhoneNumber(allPhones[i][Config.CUSTOMER_COLUMNS.PHONE_NUMBER]);
      if (existingPhone && existingPhone === normalizedPhone) {
        existingRow = i + 2;  // 行番号（1始まり、ヘッダー行を考慮）
        existingCustomerId = allPhones[i][Config.CUSTOMER_COLUMNS.CUSTOMER_ID];
        break;
      }
    }
  }
  
  if (existingRow === -1 && customerData.lineName) {
    // LINE名で検索
    const lineNameCol = Config.CUSTOMER_COLUMNS.LINE_NAME + 1;
    const allLineNames = batchGetValues(customerSheet, 2);
    for (let i = 0; i < allLineNames.length; i++) {
      if (allLineNames[i][Config.CUSTOMER_COLUMNS.LINE_NAME] === customerData.lineName) {
        existingRow = i + 2;
        existingCustomerId = allLineNames[i][Config.CUSTOMER_COLUMNS.CUSTOMER_ID];
        break;
      }
    }
  }
  
  const now = new Date();
  const nowStr = formatDateTime(now);
  
  if (existingRow > 0) {
    // 既存レコードを更新
    const existingData = customerSheet.getRange(
      existingRow,
      1,
      1,
      Config.CUSTOMER_HEADERS.length
    ).getValues()[0];
    
    // 既存データとマージ（空欄の場合は既存値を保持）
    const updatedData = [
      existingData[Config.CUSTOMER_COLUMNS.CUSTOMER_ID],  // customer_id（変更なし）
      customerData.lineName || existingData[Config.CUSTOMER_COLUMNS.LINE_NAME],
      customerData.fullName || existingData[Config.CUSTOMER_COLUMNS.FULL_NAME],
      customerData.phoneNumber || existingData[Config.CUSTOMER_COLUMNS.PHONE_NUMBER],
      customerData.email || existingData[Config.CUSTOMER_COLUMNS.EMAIL],
      existingData[Config.CUSTOMER_COLUMNS.STATUS_OVERALL] || Config.STATUS_OVERALL.UNCONTACTED,  // status_overall（変更なし、空の場合は未接触）
      existingData[Config.CUSTOMER_COLUMNS.CREATED_AT],  // created_at（変更なし）
      nowStr  // updated_at
    ];
    
    customerSheet.getRange(
      existingRow,
      1,
      1,
      Config.CUSTOMER_HEADERS.length
    ).setValues([updatedData]);
    
    return {
      customerId: existingCustomerId,
      isNew: false
    };
    
  } else {
    // 新規レコードを追加
    const customerId = generateId(Config.ID_PREFIXES.CUSTOMER);
    const newData = [
      customerId,
      customerData.lineName || '',
      customerData.fullName || '',
      customerData.phoneNumber || '',
      customerData.email || '',
      Config.STATUS_OVERALL.UNCONTACTED,  // デフォルトは未接触
      nowStr,  // created_at
      nowStr   // updated_at
    ];
    
    customerSheet.appendRow(newData);
    
    return {
      customerId: customerId,
      isNew: true
    };
  }
}

/**
 * リードソースの追加
 * @param {string} customerId - 顧客ID
 * @param {Object} sourceData - リードソースデータ
 * @returns {Object} {leadSourceId, isNew} - リードソースIDと新規フラグ
 */
function addLeadSource(customerId, sourceData) {
  const functionName = 'addLeadSource';
  
  const leadSourceSheet = Utils.getOrCreateSheet(
    Config.SHEET_NAMES.LEAD_SOURCE,
    Config.LEAD_SOURCE_HEADERS
  );
  
  // 重複チェック: 同じ顧客ID + 同じソース種別 + 同じソース詳細の組み合わせが既に存在するか
  const allLeadSources = batchGetValues(leadSourceSheet, 2);
  let existingRow = -1;
  let existingLeadSourceId = null;
  
  for (let i = 0; i < allLeadSources.length; i++) {
    const row = allLeadSources[i];
    if (
      row[Config.LEAD_SOURCE_COLUMNS.CUSTOMER_ID] === customerId &&
      row[Config.LEAD_SOURCE_COLUMNS.SOURCE_TYPE] === sourceData.sourceType &&
      row[Config.LEAD_SOURCE_COLUMNS.SOURCE_DETAIL] === sourceData.sourceDetail
    ) {
      existingRow = i + 2;
      existingLeadSourceId = row[Config.LEAD_SOURCE_COLUMNS.LEAD_SOURCE_ID];
      break;
    }
  }
  
  const now = new Date();
  const nowStr = formatDateTime(now);
  const listAddedDateStr = sourceData.listAddedDate
    ? formatDateTime(sourceData.listAddedDate, 'date')
    : formatDateTime(now, 'date');
  const eventDateStr = sourceData.eventDate
    ? formatDateTime(sourceData.eventDate, 'date')
    : '';
  
  if (existingRow > 0) {
    // 既存レコードを更新（list_added_dateやevent_dateが更新される可能性がある）
    const updatedData = [
      existingLeadSourceId,  // lead_source_id（変更なし）
      customerId,
      sourceData.sourceType,
      sourceData.sourceDetail,
      listAddedDateStr,
      eventDateStr,
      allLeadSources[existingRow - 2][Config.LEAD_SOURCE_COLUMNS.CREATED_AT],  // created_at（変更なし）
      nowStr  // updated_at
    ];
    
    leadSourceSheet.getRange(
      existingRow,
      1,
      1,
      Config.LEAD_SOURCE_HEADERS.length
    ).setValues([updatedData]);
    
    return {
      leadSourceId: existingLeadSourceId,
      isNew: false
    };
    
  } else {
    // 新規レコードを追加
    const leadSourceId = generateId(Config.ID_PREFIXES.LEAD_SOURCE);
    const newData = [
      leadSourceId,
      customerId,
      sourceData.sourceType,
      sourceData.sourceDetail,
      listAddedDateStr,
      eventDateStr,
      nowStr,  // created_at
      nowStr   // updated_at
    ];
    
    leadSourceSheet.appendRow(newData);
    
    return {
      leadSourceId: leadSourceId,
      isNew: true
    };
  }
}

/**
 * キーで顧客を検索（重複判定用）
 * @param {string} phoneNumber - 電話番号
 * @param {string} lineName - LINE名
 * @returns {Object|null} 顧客データオブジェクト（見つからない場合はnull）
 */
function findCustomerByKey(phoneNumber, lineName) {
  const customerSheet = Utils.getSheet(Config.SHEET_NAMES.CUSTOMER);
  if (!customerSheet) {
    return null;
  }
  
  const normalizedPhone = normalizePhoneNumber(phoneNumber);
  const allCustomers = batchGetValues(customerSheet, 2);
  
  for (const row of allCustomers) {
    const existingPhone = normalizePhoneNumber(row[Config.CUSTOMER_COLUMNS.PHONE_NUMBER]);
    const existingLineName = row[Config.CUSTOMER_COLUMNS.LINE_NAME];
    
    if (
      (normalizedPhone && existingPhone === normalizedPhone) ||
      (lineName && existingLineName === lineName)
    ) {
      return {
        customerId: row[Config.CUSTOMER_COLUMNS.CUSTOMER_ID],
        lineName: existingLineName,
        fullName: row[Config.CUSTOMER_COLUMNS.FULL_NAME],
        phoneNumber: row[Config.CUSTOMER_COLUMNS.PHONE_NUMBER],
        email: row[Config.CUSTOMER_COLUMNS.EMAIL],
        statusOverall: row[Config.CUSTOMER_COLUMNS.STATUS_OVERALL]
      };
    }
  }
  
  return null;
}

