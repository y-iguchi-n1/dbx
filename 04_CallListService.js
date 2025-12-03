// File: 04_CallListService.js
/**
 * 今日の架電リスト生成処理
 * 
 * ISごとに「今日架電すべきリスト」を生成します。
 */

/**
 * デバッグ用：顧客マスタと架電対象の確認
 * カスタムメニューから実行可能（開発用）
 */
function debugCallTargets() {
  const functionName = 'debugCallTargets';
  logInfo('架電対象のデバッグを開始しました', functionName);
  
  try {
    const customerSheet = Utils.getSheet(Config.SHEET_NAMES.CUSTOMER);
    const callLogSheet = Utils.getSheet(Config.SHEET_NAMES.CALL_LOG);
    
    if (!customerSheet || customerSheet.getLastRow() < 2) {
      logWarn('M_CUSTOMERにデータがありません', functionName);
      SpreadsheetApp.getUi().alert(
        'デバッグ結果',
        'M_CUSTOMERにデータがありません。\nETL処理を実行してください。',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    const allCustomers = batchGetValues(customerSheet, 2);
    const allCallLogs = callLogSheet ? batchGetValues(callLogSheet, 2) : [];
    
    const today = getToday();
    const todayStr = formatDateTime(today, 'date');
    
    // ステータス別の集計
    const statusCounts = {};
    const targetStatuses = Config.TODAY_CALL_CONDITIONS.targetStatuses;
    
    for (const row of allCustomers) {
      const status = row[Config.CUSTOMER_COLUMNS.STATUS_OVERALL] || '(空欄)';
      statusCounts[status] = (statusCounts[status] || 0) + 1;
    }
    
    // 架電対象の候補を確認
    let candidateCount = 0;
    let excludedCount = 0;
    
    for (const customerRow of allCustomers) {
      const customerId = customerRow[Config.CUSTOMER_COLUMNS.CUSTOMER_ID];
      const statusOverall = customerRow[Config.CUSTOMER_COLUMNS.STATUS_OVERALL] || '';
      const effectiveStatus = statusOverall || Config.STATUS_OVERALL.UNCONTACTED;
      
      if (targetStatuses.includes(effectiveStatus)) {
        candidateCount++;
      } else {
        excludedCount++;
      }
    }
    
    let message = '=== 架電対象デバッグ結果 ===\n\n';
    message += `総顧客数: ${allCustomers.length}件\n`;
    message += `架電ログ数: ${allCallLogs.length}件\n\n`;
    message += `【ステータス別集計】\n`;
    for (const status in statusCounts) {
      const count = statusCounts[status];
      const isTarget = targetStatuses.includes(status) || (status === '(空欄)' && targetStatuses.includes(Config.STATUS_OVERALL.UNCONTACTED));
      message += `  ${status}: ${count}件${isTarget ? ' ✅' : ' ❌'}\n`;
    }
    message += `\n【架電対象候補】\n`;
    message += `  対象ステータス: ${candidateCount}件\n`;
    message += `  除外: ${excludedCount}件\n`;
    message += `\n【対象ステータス定義】\n`;
    message += `  ${targetStatuses.join(', ')}\n`;
    
    logInfo(message, functionName);
    
    SpreadsheetApp.getUi().alert(
      'デバッグ結果',
      message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (e) {
    logError('デバッグ処理でエラーが発生しました', functionName, e);
    SpreadsheetApp.getUi().alert(
      'エラー',
      `デバッグ処理中にエラーが発生しました。\nログシート（LOGS）を確認してください。\n\nエラー: ${e.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * 今日の架電リスト生成のメイン関数（カスタムメニューから実行）
 */
function generateTodayCallLists() {
  const functionName = 'generateTodayCallLists';
  logInfo('今日の架電リスト生成を開始しました', functionName);
  
  try {
    // 担当ISのリストを取得（T_CALL_LOGから抽出、または設定から取得）
    const assignedIsList = getAssignedIsList();
    
    if (assignedIsList.length === 0) {
      logWarn('担当ISが見つかりません', functionName);
      SpreadsheetApp.getUi().alert(
        'エラー',
        '担当ISが見つかりません。\nT_CALL_LOGに架電記録があるISを自動検出します。',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    let totalGenerated = 0;
    
    // 各ISごとにリストを生成
    for (const assignedIs of assignedIsList) {
      try {
        const targets = getTodayCallTargets(assignedIs);
        const sheetName = `TODAY_CALL_${assignedIs}`;
        
        createOrUpdateTodayCallSheet(assignedIs, targets);
        
        totalGenerated += targets.length;
        
        logInfo(
          `IS "${assignedIs}" のリスト生成完了: ${targets.length}件`,
          functionName
        );
        
        Utilities.sleep(200);
        
      } catch (e) {
        logError(
          `IS "${assignedIs}" のリスト生成でエラーが発生しました`,
          functionName,
          e
        );
      }
    }
    
    logInfo(
      `今日の架電リスト生成が完了しました: 総件数=${totalGenerated}`,
      functionName
    );
    
    SpreadsheetApp.getUi().alert(
      'リスト生成完了',
      `生成件数: ${totalGenerated}件\n担当IS数: ${assignedIsList.length}人`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (e) {
    logError('今日の架電リスト生成で致命的なエラーが発生しました', functionName, e);
    throw e;
  }
}

/**
 * 担当ISのリストを取得
 * @returns {Array<string>} 担当IS名の配列
 */
function getAssignedIsList() {
  const isSet = new Set();
  
  // 1. Config.gsで設定された担当ISリストを取得
  if (Config.ASSIGNED_IS_LIST && Config.ASSIGNED_IS_LIST.length > 0) {
    Config.ASSIGNED_IS_LIST.forEach(is => {
      if (is && is.trim() !== '') {
        isSet.add(is.trim());
      }
    });
  }
  
  // 2. T_CALL_LOGから担当ISを抽出（重複除去）
  const callLogSheet = Utils.getSheet(Config.SHEET_NAMES.CALL_LOG);
  if (callLogSheet && callLogSheet.getLastRow() >= 2) {
    const allCallLogs = batchGetValues(callLogSheet, 2);
    for (const row of allCallLogs) {
      const assignedIs = row[Config.CALL_LOG_COLUMNS.ASSIGNED_IS];
      if (assignedIs && assignedIs.trim() !== '') {
        isSet.add(assignedIs.trim());
      }
    }
  }
  
  const isList = Array.from(isSet);
  
  // 3. どちらからも取得できない場合は警告
  if (isList.length === 0) {
    logWarn(
      '担当ISが見つかりません。Config.gsのASSIGNED_IS_LISTに担当IS名を設定するか、T_CALL_LOGに架電記録を追加してください。',
      'getAssignedIsList'
    );
  }
  
  return isList;
}

/**
 * 担当者ごとの架電対象を取得
 * @param {string} assignedIs - 担当IS名
 * @returns {Array<Object>} 架電対象の配列
 */
function getTodayCallTargets(assignedIs) {
  const functionName = 'getTodayCallTargets';
  
  const customerSheet = Utils.getSheet(Config.SHEET_NAMES.CUSTOMER);
  const callLogSheet = Utils.getSheet(Config.SHEET_NAMES.CALL_LOG);
  const leadSourceSheet = Utils.getSheet(Config.SHEET_NAMES.LEAD_SOURCE);
  
  if (!customerSheet || customerSheet.getLastRow() < 2) {
    return [];
  }
  
  // 顧客マスタから全顧客を取得
  const allCustomers = batchGetValues(customerSheet, 2);
  
  // 架電ログから今日の架電記録を取得
  const today = getToday();
  const todayStr = formatDateTime(today, 'date');
  const allCallLogs = callLogSheet ? batchGetValues(callLogSheet, 2) : [];
  const todayCallLogs = allCallLogs.filter(row => {
    const callDate = parseDate(row[Config.CALL_LOG_COLUMNS.CALL_DATETIME]);
    if (!callDate) return false;
    const callDateStr = formatDateTime(callDate, 'date');
    return callDateStr === todayStr && row[Config.CALL_LOG_COLUMNS.ASSIGNED_IS] === assignedIs;
  });
  const todayCalledCustomerIds = new Set(todayCallLogs.map(row => row[Config.CALL_LOG_COLUMNS.CUSTOMER_ID]));
  
  // リードソースを取得（顧客IDでグループ化）
  const allLeadSources = leadSourceSheet ? batchGetValues(leadSourceSheet, 2) : [];
  const leadSourcesByCustomer = {};
  for (const row of allLeadSources) {
    const customerId = row[Config.LEAD_SOURCE_COLUMNS.CUSTOMER_ID];
    if (!leadSourcesByCustomer[customerId]) {
      leadSourcesByCustomer[customerId] = [];
    }
    leadSourcesByCustomer[customerId].push({
      sourceType: row[Config.LEAD_SOURCE_COLUMNS.SOURCE_TYPE],
      sourceDetail: row[Config.LEAD_SOURCE_COLUMNS.SOURCE_DETAIL]
    });
  }
  
  // 各顧客の最新架電ログを取得（顧客IDでグループ化）
  const latestCallLogsByCustomer = {};
  for (const row of allCallLogs) {
    const customerId = row[Config.CALL_LOG_COLUMNS.CUSTOMER_ID];
    if (!latestCallLogsByCustomer[customerId]) {
      latestCallLogsByCustomer[customerId] = row;
    } else {
      // より新しい架電ログを保持
      const existingDate = parseDate(latestCallLogsByCustomer[customerId][Config.CALL_LOG_COLUMNS.CALL_DATETIME]);
      const currentDate = parseDate(row[Config.CALL_LOG_COLUMNS.CALL_DATETIME]);
      if (currentDate && (!existingDate || currentDate > existingDate)) {
        latestCallLogsByCustomer[customerId] = row;
      }
    }
  }
  
  // 架電対象をフィルタリング
  const targets = [];
  let excludedCount = 0;
  let excludedReasons = {
    alreadyCalledToday: 0,
    wrongStatus: 0,
    futureNextAction: 0
  };
  
  for (const customerRow of allCustomers) {
    const customerId = customerRow[Config.CUSTOMER_COLUMNS.CUSTOMER_ID];
    const statusOverall = customerRow[Config.CUSTOMER_COLUMNS.STATUS_OVERALL] || '';
    
    // デバッグ: 除外理由を記録
    if (todayCalledCustomerIds.has(customerId)) {
      excludedReasons.alreadyCalledToday++;
      excludedCount++;
      continue;
    }
    
    const targetStatuses = Config.TODAY_CALL_CONDITIONS.targetStatuses;
    const effectiveStatus = statusOverall || Config.STATUS_OVERALL.UNCONTACTED;
    if (!targetStatuses.includes(effectiveStatus)) {
      excludedReasons.wrongStatus++;
      excludedCount++;
      continue;
    }
    
    const latestCallLog = latestCallLogsByCustomer[customerId];
    if (latestCallLog) {
      const nextActionDate = parseDate(latestCallLog[Config.CALL_LOG_COLUMNS.NEXT_ACTION_DATE]);
      if (nextActionDate) {
        const today = getToday();
        if (compareDates(nextActionDate, today) > 0) {
          excludedReasons.futureNextAction++;
          excludedCount++;
          continue;
        }
      }
    }
    
    // リストアップ条件をチェック
    if (shouldIncludeInTodayList(customerRow, latestCallLogsByCustomer[customerId], todayCalledCustomerIds)) {
      // 最新の架電ログを取得
      const latestCallLog = latestCallLogsByCustomer[customerId];
      
      // リードソースを取得（最初の1つを使用）
      const leadSources = leadSourcesByCustomer[customerId] || [];
      const primarySource = leadSources.length > 0 ? leadSources[0] : { sourceType: '', sourceDetail: '' };
      
      // 架電回数を計算
      const callCount = calculateCallCountForCustomer(customerId, allCallLogs);
      
      // 最終架電日を取得
      let lastCallDate = '';
      if (latestCallLog) {
        const callDate = parseDate(latestCallLog[Config.CALL_LOG_COLUMNS.CALL_DATETIME]);
        if (callDate) {
          lastCallDate = formatDateTime(callDate, 'date');
        }
      }
      
      targets.push({
        customerId: customerId,
        lineName: customerRow[Config.CUSTOMER_COLUMNS.LINE_NAME],
        fullName: customerRow[Config.CUSTOMER_COLUMNS.FULL_NAME],
        phoneNumber: customerRow[Config.CUSTOMER_COLUMNS.PHONE_NUMBER],
        sourceType: primarySource.sourceType,
        lastCallDate: lastCallDate,
        callCount: callCount,
        status: latestCallLog ? latestCallLog[Config.CALL_LOG_COLUMNS.STATUS] : '',
        noteRank: latestCallLog ? latestCallLog[Config.CALL_LOG_COLUMNS.NOTE_RANK] : '',
        nextActionDate: latestCallLog ? latestCallLog[Config.CALL_LOG_COLUMNS.NEXT_ACTION_DATE] : ''
      });
    }
  }
  
  // デバッグログ
  if (excludedCount > 0 || targets.length === 0) {
    logInfo(
      `架電対象フィルタリング結果: 対象=${targets.length}件, 除外=${excludedCount}件 ` +
      `(今日架電済み=${excludedReasons.alreadyCalledToday}, ` +
      `ステータス不一致=${excludedReasons.wrongStatus}, ` +
      `次回アクション日が未来=${excludedReasons.futureNextAction})`,
      functionName
    );
  }
  
  return targets;
}

/**
 * リストアップ条件を判定
 * @param {Array} customerRow - 顧客データ（行配列）
 * @param {Array|null} latestCallLog - 最新の架電ログ（行配列、存在しない場合はnull）
 * @param {Set<string>} todayCalledCustomerIds - 今日既に架電した顧客IDのセット
 * @returns {boolean} リストアップ対象の場合true
 */
function shouldIncludeInTodayList(customerRow, latestCallLog, todayCalledCustomerIds) {
  const customerId = customerRow[Config.CUSTOMER_COLUMNS.CUSTOMER_ID];
  const statusOverall = customerRow[Config.CUSTOMER_COLUMNS.STATUS_OVERALL] || '';
  
  // 今日既に架電している場合は除外
  if (todayCalledCustomerIds.has(customerId)) {
    return false;
  }
  
  // ステータスが対象外の場合は除外
  const targetStatuses = Config.TODAY_CALL_CONDITIONS.targetStatuses;
  // ステータスが空欄の場合は「未接触」として扱う
  const effectiveStatus = statusOverall || Config.STATUS_OVERALL.UNCONTACTED;
  if (!targetStatuses.includes(effectiveStatus)) {
    return false;
  }
  
  // next_action_dateをチェック
  if (latestCallLog) {
    const nextActionDate = parseDate(latestCallLog[Config.CALL_LOG_COLUMNS.NEXT_ACTION_DATE]);
    if (nextActionDate) {
      const today = getToday();
      // next_action_dateが未来の場合は除外
      if (compareDates(nextActionDate, today) > 0) {
        return false;
      }
    }
  }
  
  return true;
}

/**
 * 顧客の架電回数を計算
 * @param {string} customerId - 顧客ID
 * @param {Array<Array>} allCallLogs - 全架電ログ
 * @returns {number} 架電回数
 */
function calculateCallCountForCustomer(customerId, allCallLogs) {
  return allCallLogs.filter(row => row[Config.CALL_LOG_COLUMNS.CUSTOMER_ID] === customerId).length;
}

/**
 * 今日の架電リストシートを生成・更新
 * @param {string} assignedIs - 担当IS名
 * @param {Array<Object>} targets - 架電対象の配列
 */
function createOrUpdateTodayCallSheet(assignedIs, targets) {
  const functionName = 'createOrUpdateTodayCallSheet';
  const sheetName = `TODAY_CALL_${assignedIs}`;
  
  const sheet = Utils.getOrCreateSheet(sheetName, Config.TODAY_CALL_HEADERS);
  
  // 既存データをクリア（ヘッダー行を除く）
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }
  
  // 新しいデータを書き込み
  if (targets.length === 0) {
    logInfo(`IS "${assignedIs}" の架電対象が0件です`, functionName);
    return;
  }
  
  const rows = targets.map(target => [
    target.customerId,           // customer_id
    target.lineName,              // line_name
    target.fullName,              // full_name
    target.phoneNumber,           // phone_number
    target.sourceType,            // source_type
    target.lastCallDate,          // last_call_date
    target.callCount,             // call_count
    target.status || '',          // status（ISが入力）
    target.noteRank || '',        // note_rank（ISが入力）
    target.nextActionDate || '',  // next_action_date（ISが入力）
    '',                           // memo（ISが入力）
    '',                           // appointment_datetime（ISが入力）
    ''                            // registered（自動設定）
  ]);
  
  batchAppendRows(sheet, rows);
  
  // 参照用列（customer_id ～ call_count）を保護（編集不可にする場合はコメントアウトを解除）
  // const protection = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).protect();
  // protection.setDescription('参照用列は編集不可');
  
  logInfo(`IS "${assignedIs}" のリストシートを更新しました: ${targets.length}件`, functionName);
}

