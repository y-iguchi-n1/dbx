// File: 99_Main.js
/**
 * カスタムメニュー定義、エントリーポイント
 */

/**
 * スプレッドシートを開いたときにカスタムメニューを追加
 * @param {GoogleAppsScript.Events.SheetsOnOpen} e - オープンイベント
 */
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('IS管理システム')
    .addItem('データ取込・統合（ETL）実行', 'executeEtl')
    .addItem('【デバッグ】ソース設定テスト', 'testSourceConfigs')
    .addItem('【デバッグ】架電対象確認', 'debugCallTargets')
    .addSeparator()
    .addItem('今日の架電リスト生成', 'generateTodayCallLists')
    .addSeparator()
    .addItem('【架電管理】シート設定', 'setupCallManagementSheet')
    .addItem('【架電管理】キャンセルリスト同期', 'syncCancelListToCallSheet')
    .addSeparator()
    .addItem('KPI集計（日次）', 'calculateKpiDaily')
    .addItem('KPI集計（リスト別）', 'calculateKpiByList')
    .addSeparator()
    .addItem('全処理実行（ETL → リスト生成 → KPI集計）', 'executeAllProcesses')
    .addToUi();
}

/**
 * 全処理を順次実行
 */
function executeAllProcesses() {
  const functionName = 'executeAllProcesses';
  logInfo('全処理の実行を開始しました', functionName);
  
  try {
    // 1. ETL処理
    logInfo('ステップ1: ETL処理を実行します', functionName);
    executeEtl();
    Utilities.sleep(1000);
    
    // 2. 架電リスト生成
    logInfo('ステップ2: 架電リスト生成を実行します', functionName);
    generateTodayCallLists();
    Utilities.sleep(1000);
    
    // 3. KPI集計
    logInfo('ステップ3: KPI集計を実行します', functionName);
    calculateKpiDaily();
    Utilities.sleep(1000);
    calculateKpiByList();
    
    logInfo('全処理が完了しました', functionName);
    
    SpreadsheetApp.getUi().alert(
      '全処理完了',
      'すべての処理が正常に完了しました。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (e) {
    logError('全処理の実行でエラーが発生しました', functionName, e);
    SpreadsheetApp.getUi().alert(
      'エラー',
      '処理中にエラーが発生しました。\nログシート（LOGS）を確認してください。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    throw e;
  }
}

