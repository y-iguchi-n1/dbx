// File: 05_CallLogService.js
/**
 * 架電結果の登録処理
 * 
 * TODAY_CALL_*シートの編集を検知して、
 * T_CALL_LOGとT_APPOINTMENTに自動登録します。
 */

/**
 * onEditトリガー用エントリーポイント
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e - 編集イベント
 */
function onEdit(e) {
  const functionName = 'onEdit';
  
  try {
    const sheet = e.source.getActiveSheet();
    const sheetName = sheet.getName();
    
    // TODAY_CALL_*シートの編集のみ処理
    if (!sheetName.startsWith('TODAY_CALL_')) {
      return;
    }
    
    const row = e.range.getRow();
    const col = e.range.getColumn();
    
    // ヘッダー行は無視
    if (row === 1) {
      return;
    }
    
    // 登録済みフラグをチェック
    const registeredCol = Config.TODAY_CALL_COLUMNS.REGISTERED + 1;  // 1始まりに変換
    const registeredValue = sheet.getRange(row, registeredCol).getValue();
    if (registeredValue === '✓') {
      // 既に登録済みの場合はスキップ（再登録を防ぐ）
      return;
    }
    
    // 入力用列（status, note_rank, next_action_date, memo, appointment_datetime）の編集のみ処理
    const inputCols = [
      Config.TODAY_CALL_COLUMNS.STATUS + 1,
      Config.TODAY_CALL_COLUMNS.NOTE_RANK + 1,
      Config.TODAY_CALL_COLUMNS.NEXT_ACTION_DATE + 1,
      Config.TODAY_CALL_COLUMNS.MEMO + 1,
      Config.TODAY_CALL_COLUMNS.APPOINTMENT_DATETIME + 1
    ];
    
    if (!inputCols.includes(col)) {
      return;
    }
    
    // 担当IS名を取得（シート名から抽出）
    const assignedIs = sheetName.replace('TODAY_CALL_', '');
    
    // 行データを取得
    const rowData = sheet.getRange(row, 1, 1, Config.TODAY_CALL_HEADERS.length).getValues()[0];
    const customerId = rowData[Config.TODAY_CALL_COLUMNS.CUSTOMER_ID];
    
    if (!customerId) {
      logWarn(`顧客IDが空の行をスキップしました（行: ${row}）`, functionName);
      return;
    }
    
    // 架電結果を登録
    const inputData = {
      assignedIs: assignedIs,
      status: rowData[Config.TODAY_CALL_COLUMNS.STATUS] || '',
      noteRank: rowData[Config.TODAY_CALL_COLUMNS.NOTE_RANK] || '',
      nextActionDate: rowData[Config.TODAY_CALL_COLUMNS.NEXT_ACTION_DATE] || '',
      memo: rowData[Config.TODAY_CALL_COLUMNS.MEMO] || '',
      appointmentDatetime: rowData[Config.TODAY_CALL_COLUMNS.APPOINTMENT_DATETIME] || ''
    };
    
    // statusが入力されている場合のみ登録
    if (!inputData.status) {
      return;
    }
    
    registerCallResult(sheet, row, customerId, inputData);
    
  } catch (e) {
    logError('onEdit処理でエラーが発生しました', functionName, e);
    // エラーが発生してもスプレッドシートの操作を妨げないようにする
  }
}

/**
 * 架電結果を登録
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - シートオブジェクト
 * @param {number} row - 行番号
 * @param {string} customerId - 顧客ID
 * @param {Object} inputData - 入力データ
 */
function registerCallResult(sheet, row, customerId, inputData) {
  const functionName = 'registerCallResult';
  
  try {
    logInfo(`架電結果を登録します: customerId=${customerId}`, functionName);
    
    // 架電ログを作成
    const callId = createCallLog(customerId, inputData);
    
    // アポが発生した場合はアポイント情報を作成
    if (inputData.appointmentDatetime) {
      const appointmentDatetime = parseDate(inputData.appointmentDatetime);
      if (appointmentDatetime) {
        createAppointment(customerId, callId, appointmentDatetime);
      }
    }
    
    // 顧客のステータスを更新
    updateCustomerStatus(customerId, inputData.status);
    
    // 登録済みフラグを設定
    const registeredCol = Config.TODAY_CALL_COLUMNS.REGISTERED + 1;
    sheet.getRange(row, registeredCol).setValue('✓');
    
    logInfo(`架電結果の登録が完了しました: customerId=${customerId}`, functionName);
    
  } catch (e) {
    logError(`架電結果の登録でエラーが発生しました: customerId=${customerId}`, functionName, e);
    throw e;
  }
}

/**
 * 架電ログを作成
 * @param {string} customerId - 顧客ID
 * @param {Object} inputData - 入力データ
 * @returns {string} 生成されたcall_id
 */
function createCallLog(customerId, inputData) {
  const functionName = 'createCallLog';
  
  const callLogSheet = Utils.getOrCreateSheet(
    Config.SHEET_NAMES.CALL_LOG,
    Config.CALL_LOG_HEADERS
  );
  
  // 架電回数を計算
  const callCount = calculateCallCount(customerId) + 1;
  
  // リードソースIDを取得（最初の1つを使用）
  const leadSourceId = getPrimaryLeadSourceId(customerId);
  
  const now = new Date();
  const nowStr = formatDateTime(now);
  const nextActionDateStr = inputData.nextActionDate
    ? formatDateTime(parseDate(inputData.nextActionDate), 'date')
    : '';
  
  const callId = generateId(Config.ID_PREFIXES.CALL);
  const newData = [
    callId,                                    // call_id
    customerId,                                // customer_id
    leadSourceId || '',                        // lead_source_id
    inputData.assignedIs,                      // assigned_is
    nowStr,                                    // call_datetime
    callCount,                                 // call_count
    inputData.status,                          // status
    inputData.noteRank || '',                  // note_rank
    nextActionDateStr,                         // next_action_date
    inputData.memo || '',                      // memo
    nowStr,                                    // created_at
    nowStr                                     // updated_at
  ];
  
  callLogSheet.appendRow(newData);
  
  logInfo(`架電ログを作成しました: callId=${callId}, customerId=${customerId}`, functionName);
  
  return callId;
}

/**
 * アポイント情報を作成
 * @param {string} customerId - 顧客ID
 * @param {string} callId - 元架電ID
 * @param {Date} appointmentDatetime - アポイント日時
 * @returns {string} 生成されたappointment_id
 */
function createAppointment(customerId, callId, appointmentDatetime) {
  const functionName = 'createAppointment';
  
  const appointmentSheet = Utils.getOrCreateSheet(
    Config.SHEET_NAMES.APPOINTMENT,
    Config.APPOINTMENT_HEADERS
  );
  
  const now = new Date();
  const nowStr = formatDateTime(now);
  const appointmentDatetimeStr = formatDateTime(appointmentDatetime);
  
  const appointmentId = generateId(Config.ID_PREFIXES.APPOINTMENT);
  const newData = [
    appointmentId,                             // appointment_id
    customerId,                                // customer_id
    callId,                                    // from_call_id
    nowStr,                                    // appointment_created_datetime
    appointmentDatetimeStr,                    // meeting_datetime
    '',                                        // attendance_status（後で更新）
    '',                                        // deal_status（後で更新）
    '',                                        // deal_amount（後で更新）
    nowStr,                                    // created_at
    nowStr                                     // updated_at
  ];
  
  appointmentSheet.appendRow(newData);
  
  logInfo(
    `アポイント情報を作成しました: appointmentId=${appointmentId}, customerId=${customerId}`,
    functionName
  );
  
  return appointmentId;
}

/**
 * 顧客のステータスを更新
 * @param {string} customerId - 顧客ID
 * @param {string} callStatus - 架電ステータス（T_CALL_LOG.status）
 */
function updateCustomerStatus(customerId, callStatus) {
  const functionName = 'updateCustomerStatus';
  
  const customerSheet = Utils.getSheet(Config.SHEET_NAMES.CUSTOMER);
  if (!customerSheet) {
    logWarn(`顧客シートが見つかりません: customerId=${customerId}`, functionName);
    return;
  }
  
  // 顧客を検索
  const customerIdCol = Config.CUSTOMER_COLUMNS.CUSTOMER_ID + 1;
  const row = findRowByValue(customerSheet, customerIdCol, customerId, 2);
  
  if (row === -1) {
    logWarn(`顧客が見つかりません: customerId=${customerId}`, functionName);
    return;
  }
  
  // ステータス定義を参照して、status_overallを決定
  let newStatusOverall = null;
  
  // ステータス定義に基づいて判定
  if (callStatus === 'アポ調整' || callStatus === '通電') {
    // アポ調整や通電の場合は、アポイント情報をチェック
    const hasAppointment = hasActiveAppointment(customerId);
    if (hasAppointment) {
      newStatusOverall = Config.STATUS_OVERALL.APPOINTMENT;
    } else {
      newStatusOverall = Config.STATUS_OVERALL.CALLING;
    }
  } else if (callStatus === 'NG') {
    // NGの場合はクローズ
    newStatusOverall = Config.STATUS_OVERALL.CLOSED;
  } else if (callStatus === '最架電' || callStatus === '留守電' || callStatus === '不在' || callStatus === '話中' || callStatus === '不通') {
    // 再架電が必要な場合は架電中
    newStatusOverall = Config.STATUS_OVERALL.CALLING;
  }
  
  // ステータスを更新
  if (newStatusOverall) {
    const statusCol = Config.CUSTOMER_COLUMNS.STATUS_OVERALL + 1;
    const updatedAtCol = Config.CUSTOMER_COLUMNS.UPDATED_AT + 1;
    const nowStr = formatDateTime(new Date());
    
    customerSheet.getRange(row, statusCol).setValue(newStatusOverall);
    customerSheet.getRange(row, updatedAtCol).setValue(nowStr);
    
    logInfo(
      `顧客ステータスを更新しました: customerId=${customerId}, status=${newStatusOverall}`,
      functionName
    );
  }
}

/**
 * 顧客にアクティブなアポイントがあるかチェック
 * @param {string} customerId - 顧客ID
 * @returns {boolean} アクティブなアポイントがある場合true
 */
function hasActiveAppointment(customerId) {
  const appointmentSheet = Utils.getSheet(Config.SHEET_NAMES.APPOINTMENT);
  if (!appointmentSheet || appointmentSheet.getLastRow() < 2) {
    return false;
  }
  
  const allAppointments = batchGetValues(appointmentSheet, 2);
  const now = new Date();
  
  for (const row of allAppointments) {
    if (row[Config.APPOINTMENT_COLUMNS.CUSTOMER_ID] === customerId) {
      const meetingDate = parseDate(row[Config.APPOINTMENT_COLUMNS.MEETING_DATETIME]);
      if (meetingDate && meetingDate >= now) {
        // 未来のアポイントがある
        const attendanceStatus = row[Config.APPOINTMENT_COLUMNS.ATTENDANCE_STATUS];
        const dealStatus = row[Config.APPOINTMENT_COLUMNS.DEAL_STATUS];
        // 着席ステータスが未設定、または成約ステータスが未設定の場合はアクティブ
        if (!attendanceStatus || !dealStatus) {
          return true;
        }
      }
    }
  }
  
  return false;
}

/**
 * 顧客の架電回数を計算
 * @param {string} customerId - 顧客ID
 * @returns {number} 架電回数
 */
function calculateCallCount(customerId) {
  const callLogSheet = Utils.getSheet(Config.SHEET_NAMES.CALL_LOG);
  if (!callLogSheet || callLogSheet.getLastRow() < 2) {
    return 0;
  }
  
  const allCallLogs = batchGetValues(callLogSheet, 2);
  return allCallLogs.filter(row => row[Config.CALL_LOG_COLUMNS.CUSTOMER_ID] === customerId).length;
}

/**
 * 顧客の主要リードソースIDを取得
 * @param {string} customerId - 顧客ID
 * @returns {string|null} リードソースID（見つからない場合はnull）
 */
function getPrimaryLeadSourceId(customerId) {
  const leadSourceSheet = Utils.getSheet(Config.SHEET_NAMES.LEAD_SOURCE);
  if (!leadSourceSheet || leadSourceSheet.getLastRow() < 2) {
    return null;
  }
  
  const allLeadSources = batchGetValues(leadSourceSheet, 2);
  for (const row of allLeadSources) {
    if (row[Config.LEAD_SOURCE_COLUMNS.CUSTOMER_ID] === customerId) {
      return row[Config.LEAD_SOURCE_COLUMNS.LEAD_SOURCE_ID];
    }
  }
  
  return null;
}

