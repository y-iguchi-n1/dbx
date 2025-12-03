// File: 06_KpiService.js
/**
 * KPI集計処理
 * 
 * T_CALL_LOGとT_APPOINTMENTからKPIを計算し、
 * V_KPI_DAILYとV_KPI_BY_LISTに結果を書き込みます。
 */

/**
 * 日次KPI集計のメイン関数（カスタムメニューから実行）
 */
function calculateKpiDaily() {
  const functionName = 'calculateKpiDaily';
  logInfo('日次KPI集計を開始しました', functionName);
  
  try {
    const kpiSheet = Utils.getOrCreateSheet(
      Config.SHEET_NAMES.KPI_DAILY,
      Config.KPI_DAILY_HEADERS
    );
    
    // 集計対象の日付範囲を取得（直近30日）
    const today = getToday();
    const startDate = new Date(today);
    startDate.setDate(startDate.getDate() - 30);
    
    // 担当ISのリストを取得
    const assignedIsList = getAssignedIsList();
    if (assignedIsList.length === 0) {
      logWarn('担当ISが見つかりません', functionName);
      return;
    }
    
    // ソース種別のリストを取得
    const sourceTypes = Object.values(Config.SOURCE_TYPES);
    sourceTypes.push('ALL');  // 全ソース合計
    
    const results = [];
    
    // 日別 × 担当者別 × ソース別で集計
    for (let d = 0; d <= 30; d++) {
      const targetDate = new Date(startDate);
      targetDate.setDate(targetDate.getDate() + d);
      const dateStr = formatDateTime(targetDate, 'date');
      
      for (const assignedIs of assignedIsList) {
        for (const sourceType of sourceTypes) {
          const kpi = aggregateDailyKpi(targetDate, assignedIs, sourceType);
          if (kpi.callCount > 0 || kpi.appointmentCount > 0) {
            // データがある場合のみ追加
            results.push({
              date: dateStr,
              assignedIs: assignedIs,
              leadSourceType: sourceType,
              ...kpi
            });
          }
        }
      }
    }
    
    // 既存データをクリア（ヘッダー行を除く）
    const lastRow = kpiSheet.getLastRow();
    if (lastRow > 1) {
      kpiSheet.deleteRows(2, lastRow - 1);
    }
    
    // 新しいデータを書き込み
    if (results.length > 0) {
      const rows = results.map(result => [
        result.date,
        result.assignedIs,
        result.leadSourceType,
        result.callCount,
        result.connectedCount,
        result.connectionRate,
        result.appointmentCount,
        result.appointmentRate,
        result.attendanceCount,
        result.attendanceRate,
        result.dealCount,
        result.dealRate,
        formatDateTime(new Date())  // updated_at
      ]);
      
      batchAppendRows(kpiSheet, rows);
    }
    
    logInfo(`日次KPI集計が完了しました: ${results.length}件`, functionName);
    
    SpreadsheetApp.getUi().alert(
      'KPI集計完了',
      `日次KPI集計が完了しました\n集計件数: ${results.length}件`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (e) {
    logError('日次KPI集計で致命的なエラーが発生しました', functionName, e);
    throw e;
  }
}

/**
 * リスト別KPI集計のメイン関数（カスタムメニューから実行）
 */
function calculateKpiByList() {
  const functionName = 'calculateKpiByList';
  logInfo('リスト別KPI集計を開始しました', functionName);
  
  try {
    const kpiSheet = Utils.getOrCreateSheet(
      Config.SHEET_NAMES.KPI_BY_LIST,
      Config.KPI_BY_LIST_HEADERS
    );
    
    // リードソースを取得
    const leadSourceSheet = Utils.getSheet(Config.SHEET_NAMES.LEAD_SOURCE);
    if (!leadSourceSheet || leadSourceSheet.getLastRow() < 2) {
      logWarn('リードソースデータがありません', functionName);
      return;
    }
    
    const allLeadSources = batchGetValues(leadSourceSheet, 2);
    
    // ソース種別とソース詳細の組み合わせを取得
    const sourceGroups = {};
    for (const row of allLeadSources) {
      const sourceType = row[Config.LEAD_SOURCE_COLUMNS.SOURCE_TYPE];
      const sourceDetail = row[Config.LEAD_SOURCE_COLUMNS.SOURCE_DETAIL];
      const key = `${sourceType}::${sourceDetail}`;
      
      if (!sourceGroups[key]) {
        sourceGroups[key] = {
          sourceType: sourceType,
          sourceDetail: sourceDetail,
          customerIds: new Set()
        };
      }
      sourceGroups[key].customerIds.add(row[Config.LEAD_SOURCE_COLUMNS.CUSTOMER_ID]);
    }
    
    const results = [];
    const today = getToday();
    const periodStart = new Date(today);
    periodStart.setMonth(periodStart.getMonth() - 1);  // 直近1ヶ月
    const periodEnd = today;
    
    // 各ソースグループで集計
    for (const key in sourceGroups) {
      const group = sourceGroups[key];
      const customerIds = Array.from(group.customerIds);
      
      const kpi = aggregateListKpi(
        group.sourceType,
        group.sourceDetail,
        periodStart,
        periodEnd,
        customerIds
      );
      
      results.push({
        sourceType: group.sourceType,
        sourceDetail: group.sourceDetail,
        periodStart: formatDateTime(periodStart, 'date'),
        periodEnd: formatDateTime(periodEnd, 'date'),
        ...kpi
      });
    }
    
    // 既存データをクリア（ヘッダー行を除く）
    const lastRow = kpiSheet.getLastRow();
    if (lastRow > 1) {
      kpiSheet.deleteRows(2, lastRow - 1);
    }
    
    // 新しいデータを書き込み
    if (results.length > 0) {
      const rows = results.map(result => [
        result.sourceType,
        result.sourceDetail,
        result.periodStart,
        result.periodEnd,
        result.totalCustomers,
        result.callCount,
        result.connectedCount,
        result.connectionRate,
        result.appointmentCount,
        result.appointmentRate,
        result.attendanceCount,
        result.attendanceRate,
        result.dealCount,
        result.dealRate,
        formatDateTime(new Date())  // updated_at
      ]);
      
      batchAppendRows(kpiSheet, rows);
    }
    
    logInfo(`リスト別KPI集計が完了しました: ${results.length}件`, functionName);
    
    SpreadsheetApp.getUi().alert(
      'KPI集計完了',
      `リスト別KPI集計が完了しました\n集計件数: ${results.length}件`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (e) {
    logError('リスト別KPI集計で致命的なエラーが発生しました', functionName, e);
    throw e;
  }
}

/**
 * 日別×担当者別KPI集計
 * @param {Date} date - 集計日
 * @param {string} assignedIs - 担当IS名
 * @param {string} sourceType - ソース種別（'ALL'の場合は全ソース）
 * @returns {Object} KPIオブジェクト
 */
function aggregateDailyKpi(date, assignedIs, sourceType) {
  const callLogSheet = Utils.getSheet(Config.SHEET_NAMES.CALL_LOG);
  const appointmentSheet = Utils.getSheet(Config.SHEET_NAMES.APPOINTMENT);
  
  if (!callLogSheet || callLogSheet.getLastRow() < 2) {
    return createEmptyKpi();
  }
  
  const dateStr = formatDateTime(date, 'date');
  const allCallLogs = batchGetValues(callLogSheet, 2);
  
  // 対象日の架電ログをフィルタ
  let targetCallLogs = allCallLogs.filter(row => {
    const callDate = parseDate(row[Config.CALL_LOG_COLUMNS.CALL_DATETIME]);
    if (!callDate) return false;
    const callDateStr = formatDateTime(callDate, 'date');
    if (callDateStr !== dateStr) return false;
    if (row[Config.CALL_LOG_COLUMNS.ASSIGNED_IS] !== assignedIs) return false;
    return true;
  });
  
  // ソース種別でフィルタ（'ALL'の場合はフィルタしない）
  if (sourceType !== 'ALL') {
    const leadSourceSheet = Utils.getSheet(Config.SHEET_NAMES.LEAD_SOURCE);
    if (leadSourceSheet && leadSourceSheet.getLastRow() >= 2) {
      const allLeadSources = batchGetValues(leadSourceSheet, 2);
      const leadSourceMap = {};
      for (const row of allLeadSources) {
        if (row[Config.LEAD_SOURCE_COLUMNS.SOURCE_TYPE] === sourceType) {
          leadSourceMap[row[Config.LEAD_SOURCE_COLUMNS.LEAD_SOURCE_ID]] = true;
        }
      }
      
      targetCallLogs = targetCallLogs.filter(row => {
        const leadSourceId = row[Config.CALL_LOG_COLUMNS.LEAD_SOURCE_ID];
        return leadSourceId && leadSourceMap[leadSourceId];
      });
    }
  }
  
  // KPIを計算
  const callCount = targetCallLogs.length;
  const connectedCount = getConnectionCount(targetCallLogs);
  const connectionRate = callCount > 0 ? connectedCount / callCount : 0;
  
  // アポ件数を取得
  const customerIds = targetCallLogs.map(row => row[Config.CALL_LOG_COLUMNS.CUSTOMER_ID]);
  const appointmentCount = getAppointmentCount(customerIds, date, date);
  const appointmentRate = callCount > 0 ? appointmentCount / callCount : 0;
  
  // 着席件数と成約件数を取得
  const appointments = getAppointmentsByCustomerIds(customerIds, date, date);
  const attendanceCount = getAttendanceCount(appointments);
  const attendanceRate = appointmentCount > 0 ? attendanceCount / appointmentCount : 0;
  const dealCount = getDealCount(appointments);
  const dealRate = callCount > 0 ? dealCount / callCount : 0;
  
  return {
    callCount: callCount,
    connectedCount: connectedCount,
    connectionRate: connectionRate,
    appointmentCount: appointmentCount,
    appointmentRate: appointmentRate,
    attendanceCount: attendanceCount,
    attendanceRate: attendanceRate,
    dealCount: dealCount,
    dealRate: dealRate
  };
}

/**
 * リスト別KPI集計
 * @param {string} sourceType - ソース種別
 * @param {string} sourceDetail - ソース詳細
 * @param {Date} periodStart - 集計期間開始
 * @param {Date} periodEnd - 集計期間終了
 * @param {Array<string>} customerIds - 対象顧客IDの配列
 * @returns {Object} KPIオブジェクト
 */
function aggregateListKpi(sourceType, sourceDetail, periodStart, periodEnd, customerIds) {
  const callLogSheet = Utils.getSheet(Config.SHEET_NAMES.CALL_LOG);
  const appointmentSheet = Utils.getSheet(Config.SHEET_NAMES.APPOINTMENT);
  
  const totalCustomers = customerIds.length;
  
  if (!callLogSheet || callLogSheet.getLastRow() < 2) {
    return {
      totalCustomers: totalCustomers,
      ...createEmptyKpi()
    };
  }
  
  const allCallLogs = batchGetValues(callLogSheet, 2);
  
  // 対象期間・対象顧客の架電ログをフィルタ
  const targetCallLogs = allCallLogs.filter(row => {
    const customerId = row[Config.CALL_LOG_COLUMNS.CUSTOMER_ID];
    if (!customerIds.includes(customerId)) return false;
    
    const callDate = parseDate(row[Config.CALL_LOG_COLUMNS.CALL_DATETIME]);
    if (!callDate) return false;
    
    return callDate >= periodStart && callDate <= periodEnd;
  });
  
  // KPIを計算
  const callCount = targetCallLogs.length;
  const connectedCount = getConnectionCount(targetCallLogs);
  const connectionRate = callCount > 0 ? connectedCount / callCount : 0;
  
  // アポ件数を取得
  const targetCustomerIds = Array.from(new Set(targetCallLogs.map(row => row[Config.CALL_LOG_COLUMNS.CUSTOMER_ID])));
  const appointmentCount = getAppointmentCount(targetCustomerIds, periodStart, periodEnd);
  const appointmentRate = callCount > 0 ? appointmentCount / callCount : 0;
  
  // 着席件数と成約件数を取得
  const appointments = getAppointmentsByCustomerIds(targetCustomerIds, periodStart, periodEnd);
  const attendanceCount = getAttendanceCount(appointments);
  const attendanceRate = appointmentCount > 0 ? attendanceCount / appointmentCount : 0;
  const dealCount = getDealCount(appointments);
  const dealRate = callCount > 0 ? dealCount / callCount : 0;
  
  return {
    totalCustomers: totalCustomers,
    callCount: callCount,
    connectedCount: connectedCount,
    connectionRate: connectionRate,
    appointmentCount: appointmentCount,
    appointmentRate: appointmentRate,
    attendanceCount: attendanceCount,
    attendanceRate: attendanceRate,
    dealCount: dealCount,
    dealRate: dealRate
  };
}

/**
 * 通電件数を計算
 * @param {Array<Array>} callLogs - 架電ログの配列
 * @returns {number} 通電件数
 */
function getConnectionCount(callLogs) {
  let count = 0;
  for (const row of callLogs) {
    const status = row[Config.CALL_LOG_COLUMNS.STATUS];
    const statusDef = Config.STATUS_DEFINITIONS[status];
    if (statusDef && statusDef.connectedCount) {
      count++;
    }
  }
  return count;
}

/**
 * アポ件数を計算
 * @param {Array<string>} customerIds - 顧客IDの配列
 * @param {Date} periodStart - 期間開始
 * @param {Date} periodEnd - 期間終了
 * @returns {number} アポ件数
 */
function getAppointmentCount(customerIds, periodStart, periodEnd) {
  const appointmentSheet = Utils.getSheet(Config.SHEET_NAMES.APPOINTMENT);
  if (!appointmentSheet || appointmentSheet.getLastRow() < 2) {
    return 0;
  }
  
  const allAppointments = batchGetValues(appointmentSheet, 2);
  const customerIdSet = new Set(customerIds);
  
  let count = 0;
  for (const row of allAppointments) {
    const customerId = row[Config.APPOINTMENT_COLUMNS.CUSTOMER_ID];
    if (!customerIdSet.has(customerId)) continue;
    
    const appointmentDate = parseDate(row[Config.APPOINTMENT_COLUMNS.APPOINTMENT_CREATED_DATETIME]);
    if (!appointmentDate) continue;
    
    if (appointmentDate >= periodStart && appointmentDate <= periodEnd) {
      count++;
    }
  }
  
  return count;
}

/**
 * アポイント情報を取得
 * @param {Array<string>} customerIds - 顧客IDの配列
 * @param {Date} periodStart - 期間開始
 * @param {Date} periodEnd - 期間終了
 * @returns {Array<Array>} アポイント情報の配列
 */
function getAppointmentsByCustomerIds(customerIds, periodStart, periodEnd) {
  const appointmentSheet = Utils.getSheet(Config.SHEET_NAMES.APPOINTMENT);
  if (!appointmentSheet || appointmentSheet.getLastRow() < 2) {
    return [];
  }
  
  const allAppointments = batchGetValues(appointmentSheet, 2);
  const customerIdSet = new Set(customerIds);
  
  return allAppointments.filter(row => {
    const customerId = row[Config.APPOINTMENT_COLUMNS.CUSTOMER_ID];
    if (!customerIdSet.has(customerId)) return false;
    
    const appointmentDate = parseDate(row[Config.APPOINTMENT_COLUMNS.APPOINTMENT_CREATED_DATETIME]);
    if (!appointmentDate) return false;
    
    return appointmentDate >= periodStart && appointmentDate <= periodEnd;
  });
}

/**
 * 着席件数を計算
 * @param {Array<Array>} appointments - アポイント情報の配列
 * @returns {number} 着席件数
 */
function getAttendanceCount(appointments) {
  let count = 0;
  for (const row of appointments) {
    const attendanceStatus = row[Config.APPOINTMENT_COLUMNS.ATTENDANCE_STATUS];
    if (attendanceStatus === Config.ATTENDANCE_STATUS.ATTENDED) {
      count++;
    }
  }
  return count;
}

/**
 * 成約件数を計算
 * @param {Array<Array>} appointments - アポイント情報の配列
 * @returns {number} 成約件数
 */
function getDealCount(appointments) {
  let count = 0;
  for (const row of appointments) {
    const dealStatus = row[Config.APPOINTMENT_COLUMNS.DEAL_STATUS];
    if (dealStatus === Config.DEAL_STATUS.DEAL) {
      count++;
    }
  }
  return count;
}

/**
 * 空のKPIオブジェクトを作成
 * @returns {Object} 空のKPIオブジェクト
 */
function createEmptyKpi() {
  return {
    callCount: 0,
    connectedCount: 0,
    connectionRate: 0,
    appointmentCount: 0,
    appointmentRate: 0,
    attendanceCount: 0,
    attendanceRate: 0,
    dealCount: 0,
    dealRate: 0
  };
}

