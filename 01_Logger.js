// File: 01_Logger.js
/**
 * ログ出力ユーティリティ
 * 
 * すべてのメイン処理でこのLoggerを使用することで、
 * 統一されたログフォーマットでLOGSシートに記録されます。
 */

/**
 * 情報ログを出力
 * @param {string} message - ログメッセージ
 * @param {string} functionName - 関数名（オプション）
 */
function logInfo(message, functionName = '') {
  _writeLog(Config.LOG_LEVELS.INFO, message, functionName, null);
}

/**
 * 警告ログを出力
 * @param {string} message - ログメッセージ
 * @param {string} functionName - 関数名（オプション）
 */
function logWarn(message, functionName = '') {
  _writeLog(Config.LOG_LEVELS.WARN, message, functionName, null);
}

/**
 * エラーログを出力
 * @param {string} message - ログメッセージ
 * @param {string} functionName - 関数名（オプション）
 * @param {Error} error - エラーオブジェクト（オプション）
 */
function logError(message, functionName = '', error = null) {
  const stacktrace = error ? error.stack || error.toString() : '';
  _writeLog(Config.LOG_LEVELS.ERROR, message, functionName, stacktrace);
}

/**
 * ログをLOGSシートに書き込む（内部関数）
 * @param {string} level - ログレベル
 * @param {string} message - ログメッセージ
 * @param {string} functionName - 関数名
 * @param {string} stacktrace - スタックトレース（エラー時のみ）
 * @private
 */
function _writeLog(level, message, functionName, stacktrace) {
  try {
    const sheet = Utils.getOrCreateSheet(
      Config.SHEET_NAMES.LOGS,
      Config.LOGS_HEADERS
    );
    
    const now = new Date();
    const timestamp = Utilities.formatDate(
      now,
      Session.getScriptTimeZone(),
      Config.DATETIME_FORMAT
    );
    
    const row = [
      timestamp,
      functionName || '',
      level,
      message || '',
      stacktrace || ''
    ];
    
    // バッチ処理で追記（パフォーマンス最適化）
    sheet.appendRow(row);
    
    // コンソールにも出力（デバッグ用）
    console.log(`[${level}] ${functionName}: ${message}`);
    if (stacktrace) {
      console.error(stacktrace);
    }
    
  } catch (e) {
    // ログ出力自体が失敗した場合はコンソールに出力
    console.error('Logger failed:', e);
    console.error(`[${level}] ${functionName}: ${message}`);
  }
}

