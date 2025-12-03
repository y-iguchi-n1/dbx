// File: 00_Config.js
/**
 * システム全体の設定を一元管理するオブジェクト
 * 
 * このファイルを修正することで、ソース追加・ステータス追加・列追加に対応できます。
 */

const Config = {
  
  // ========================================
  // シート名定義
  // ========================================
  SHEET_NAMES: {
    CUSTOMER: 'M_CUSTOMER',
    LEAD_SOURCE: 'M_LEAD_SOURCE',
    CALL_LOG: 'T_CALL_LOG',
    APPOINTMENT: 'T_APPOINTMENT',
    KPI_DAILY: 'V_KPI_DAILY',
    KPI_BY_LIST: 'V_KPI_BY_LIST',
    CALL_MANAGEMENT: 'IS_架電管理',  // インサイドセールス用架電管理シート
    LOGS: 'LOGS'
  },

  // ========================================
  // M_CUSTOMER（顧客マスタ）の列定義
  // ========================================
  CUSTOMER_COLUMNS: {
    CUSTOMER_ID: 0,        // A列
    LINE_NAME: 1,          // B列
    FULL_NAME: 2,          // C列
    PHONE_NUMBER: 3,       // D列
    EMAIL: 4,              // E列
    STATUS_OVERALL: 5,     // F列
    CREATED_AT: 6,         // G列
    UPDATED_AT: 7          // H列
  },
  
  CUSTOMER_HEADERS: [
    'customer_id',
    'line_name',
    'full_name',
    'phone_number',
    'email',
    'status_overall',
    'created_at',
    'updated_at'
  ],

  // ========================================
  // M_LEAD_SOURCE（リードソースマスタ）の列定義
  // ========================================
  LEAD_SOURCE_COLUMNS: {
    LEAD_SOURCE_ID: 0,     // A列
    CUSTOMER_ID: 1,        // B列
    SOURCE_TYPE: 2,        // C列
    SOURCE_DETAIL: 3,      // D列
    LIST_ADDED_DATE: 4,    // E列
    EVENT_DATE: 5,         // F列
    CREATED_AT: 6,         // G列
    UPDATED_AT: 7          // H列
  },
  
  LEAD_SOURCE_HEADERS: [
    'lead_source_id',
    'customer_id',
    'source_type',
    'source_detail',
    'list_added_date',
    'event_date',
    'created_at',
    'updated_at'
  ],

  // ========================================
  // T_CALL_LOG（架電活動ログ）の列定義
  // ========================================
  CALL_LOG_COLUMNS: {
    CALL_ID: 0,            // A列
    CUSTOMER_ID: 1,        // B列
    LEAD_SOURCE_ID: 2,     // C列
    ASSIGNED_IS: 3,        // D列
    CALL_DATETIME: 4,      // E列
    CALL_COUNT: 5,         // F列
    STATUS: 6,             // G列
    NOTE_RANK: 7,          // H列
    NEXT_ACTION_DATE: 8,   // I列
    MEMO: 9,               // J列
    CREATED_AT: 10,        // K列
    UPDATED_AT: 11         // L列
  },
  
  CALL_LOG_HEADERS: [
    'call_id',
    'customer_id',
    'lead_source_id',
    'assigned_is',
    'call_datetime',
    'call_count',
    'status',
    'note_rank',
    'next_action_date',
    'memo',
    'created_at',
    'updated_at'
  ],

  // ========================================
  // T_APPOINTMENT（アポイント情報）の列定義
  // ========================================
  APPOINTMENT_COLUMNS: {
    APPOINTMENT_ID: 0,     // A列
    CUSTOMER_ID: 1,        // B列
    FROM_CALL_ID: 2,       // C列
    APPOINTMENT_CREATED_DATETIME: 3,  // D列
    MEETING_DATETIME: 4,   // E列
    ATTENDANCE_STATUS: 5,  // F列
    DEAL_STATUS: 6,       // G列
    DEAL_AMOUNT: 7,       // H列
    CREATED_AT: 8,         // I列
    UPDATED_AT: 9          // J列
  },
  
  APPOINTMENT_HEADERS: [
    'appointment_id',
    'customer_id',
    'from_call_id',
    'appointment_created_datetime',
    'meeting_datetime',
    'attendance_status',
    'deal_status',
    'deal_amount',
    'created_at',
    'updated_at'
  ],

  // ========================================
  // TODAY_CALL_*（今日の架電リスト）の列定義
  // ========================================
  TODAY_CALL_COLUMNS: {
    CUSTOMER_ID: 0,        // A列（参照用）
    LINE_NAME: 1,          // B列（参照用）
    FULL_NAME: 2,          // C列（参照用）
    PHONE_NUMBER: 3,       // D列（参照用）
    SOURCE_TYPE: 4,        // E列（参照用）
    LAST_CALL_DATE: 5,     // F列（参照用）
    CALL_COUNT: 6,         // G列（参照用）
    STATUS: 7,             // H列（ISが入力）
    NOTE_RANK: 8,          // I列（ISが入力）
    NEXT_ACTION_DATE: 9,   // J列（ISが入力）
    MEMO: 10,              // K列（ISが入力）
    APPOINTMENT_DATETIME: 11,  // L列（ISが入力、アポ時のみ）
    REGISTERED: 12         // M列（自動設定）
  },
  
  TODAY_CALL_HEADERS: [
    'customer_id',
    'line_name',
    'full_name',
    'phone_number',
    'source_type',
    'last_call_date',
    'call_count',
    'status',
    'note_rank',
    'next_action_date',
    'memo',
    'appointment_datetime',
    'registered'
  ],

  // ========================================
  // V_KPI_DAILY（KPI日次集計）の列定義
  // ========================================
  KPI_DAILY_HEADERS: [
    'date',
    'assigned_is',
    'lead_source_type',
    'call_count',
    'connected_count',
    'connection_rate',
    'appointment_count',
    'appointment_rate',
    'attendance_count',
    'attendance_rate',
    'deal_count',
    'deal_rate',
    'updated_at'
  ],

  // ========================================
  // V_KPI_BY_LIST（KPIリスト別集計）の列定義
  // ========================================
  KPI_BY_LIST_HEADERS: [
    'source_type',
    'source_detail',
    'period_start',
    'period_end',
    'total_customers',
    'call_count',
    'connected_count',
    'connection_rate',
    'appointment_count',
    'appointment_rate',
    'attendance_count',
    'attendance_rate',
    'deal_count',
    'deal_rate',
    'updated_at'
  ],

  // ========================================
  // LOGS（ログ）の列定義
  // ========================================
  LOGS_HEADERS: [
    'timestamp',
    'function_name',
    'level',
    'message',
    'stacktrace'
  ],

  // ========================================
  // ステータス定義
  // ========================================
  
  /**
   * status_overall（M_CUSTOMER）の定義
   */
  STATUS_OVERALL: {
    UNCONTACTED: '未接触',
    CALLING: '架電中',
    APPOINTMENT: 'アポ中',
    CLOSED: 'クローズ'
  },

  /**
   * status（T_CALL_LOG）の定義
   * 通電カウント対象: connectedCount = true のステータス
   */
  STATUS_DEFINITIONS: {
    '最架電': {
      connectedCount: true,
      description: '再度架電が必要'
    },
    '通電': {
      connectedCount: true,
      description: '電話がつながった'
    },
    'アポ調整': {
      connectedCount: true,
      description: 'アポイント調整中'
    },
    '留守電': {
      connectedCount: true,
      description: '留守電にメッセージを残した'
    },
    'NG': {
      connectedCount: true,
      description: '興味なし・断られた'
    },
    '不在': {
      connectedCount: false,
      description: '不在'
    },
    '話中': {
      connectedCount: false,
      description: '話中'
    },
    '不通': {
      connectedCount: false,
      description: '電話がつながらない'
    }
  },

  // ========================================
  // ソース種別定義
  // ========================================
  SOURCE_TYPES: {
    PRESENT: 'プレゼント受け取り',
    SEMINAR: 'セミナーアンケート',
    CANCEL: 'キャンセルリスト',
    OTHER: 'その他キャンペーン'
  },

  // ========================================
  // ソース設定（元データの場所とマッピングルール）
  // ========================================
  /**
   * 新しいソースを追加する場合は、この配列に設定を追加してください。
   * 
   * 設定項目:
   * - name: ソース名（識別用）
   * - sourceType: ソース種別（SOURCE_TYPESのいずれか）
   * - spreadsheetId: 元データのスプレッドシートID（同じスプレッドシートの場合は null）
   * - sheetName: 元データのシート名
   * - mapping: 列マッピング（元データの列名 → 内部カラム名）
   *   - lineName: LINE名の列名
   *   - fullName: 本名の列名（オプション）
   *   - phoneNumber: 電話番号の列名
   *   - email: メールアドレスの列名（オプション）
   *   - sourceDetail: ソース詳細（キャンペーン名など）の列名（オプション）
   *   - eventDate: イベント日の列名（オプション）
   * - headerRow: ヘッダー行の行番号（1始まり、デフォルト: 1）
   * - dataStartRow: データ開始行の行番号（1始まり、デフォルト: 2）
   */
  SOURCE_CONFIGS: [
    {
      name: 'プレゼント受け取りリスト',
      sourceType: 'プレゼント受け取り',  // SOURCE_TYPES.PRESENT の値
      spreadsheetId: null,  // 同じスプレッドシート内の場合は null
      sheetName: 'プレゼント受け取りリスト',  // 元データのシート名
      mapping: {
        lineName: 'LINE名',  // 元データの列名
        fullName: '本名',
        phoneNumber: '電話番号',
        email: 'メールアドレス',
        sourceDetail: 'キャンペーン名',
        eventDate: null  // このソースにはイベント日がない
      },
      headerRow: 1,
      dataStartRow: 2
    },
    {
      name: 'セミナーアンケート',
      sourceType: 'セミナーアンケート',  // SOURCE_TYPES.SEMINAR の値
      spreadsheetId: null,
      sheetName: 'セミナーアンケート',
      mapping: {
        lineName: 'LINE名',
        fullName: '氏名',
        phoneNumber: '電話番号',
        email: 'メール',
        sourceDetail: 'セミナー名',
        eventDate: 'セミナー実施日'
      },
      headerRow: 1,
      dataStartRow: 2
    },
    {
      name: 'キャンセルリスト',
      sourceType: 'キャンセルリスト',  // SOURCE_TYPES.CANCEL の値
      spreadsheetId: null,
      sheetName: 'キャンセルリスト',
      mapping: {
        lineName: 'LINE名',
        fullName: '名前',
        phoneNumber: '電話',
        email: null,
        sourceDetail: 'キャンセル理由',
        eventDate: null
      },
      headerRow: 1,
      dataStartRow: 2
    }
    // 新しいソースを追加する場合は、ここに設定を追加してください
    // sourceType には SOURCE_TYPES の値（文字列）を直接指定してください
  ],

  // ========================================
  // その他の定数
  // ========================================
  
  /**
   * ID生成のプレフィックス
   */
  ID_PREFIXES: {
    CUSTOMER: 'CUST',
    LEAD_SOURCE: 'LEAD',
    CALL: 'CALL',
    APPOINTMENT: 'APPT'
  },

  /**
   * 日時フォーマット
   */
  DATE_FORMAT: 'yyyy-MM-dd',
  DATETIME_FORMAT: 'yyyy-MM-dd HH:mm:ss',

  /**
   * ログレベル
   */
  LOG_LEVELS: {
    INFO: 'INFO',
    WARN: 'WARN',
    ERROR: 'ERROR'
  },

  /**
   * ネタランクの定義
   */
  NOTE_RANKS: ['A', 'B', 'C', 'D'],

  /**
   * アポイントメントの着席ステータス
   */
  ATTENDANCE_STATUS: {
    ATTENDED: '着席',
    NO_SHOW: 'no-show',
    CANCELLED: 'キャンセル'
  },

  /**
   * 成約ステータス
   */
  DEAL_STATUS: {
    DEAL: '成約',
    PASS: '見送り',
    CONSIDERING: '検討中'
  },

  /**
   * 今日の架電リスト生成条件
   */
  TODAY_CALL_CONDITIONS: {
    // ステータスが以下のいずれかの場合にリストアップ
    targetStatuses: [
      '未接触',  // STATUS_OVERALL.UNCONTACTED の値
      '架電中'   // STATUS_OVERALL.CALLING の値
    ],
    // next_action_date が今日以前、または空欄の場合にリストアップ
    includeNoNextAction: true
  },

  /**
   * 担当ISのリスト（手動設定）
   * 
   * T_CALL_LOGにデータがない場合でも、このリストから担当ISを取得します。
   * 実際の担当IS名をここに追加してください。
   * 
   * 例: ['田中花子', '佐藤太郎', '鈴木一郎']
   */
  ASSIGNED_IS_LIST: [
    // ここに担当IS名を追加してください
    // 例: '田中花子',
    // 例: '佐藤太郎',
    '岩谷優太'
  ],

  // ========================================
  // 架電管理シート用の設定
  // ========================================
  
  /**
   * キャンセルリスト用スプレッドシートID（別ブック）
   * TODO: 実際のスプレッドシートIDに置き換えてください
   */
  CANCEL_LIST_SPREADSHEET_ID: 'TODO: キャンセルリストのスプレッドシートID',
  
  /**
   * キャンセルリスト用シート名
   * TODO: 実際のシート名に置き換えてください
   */
  CANCEL_LIST_SHEET_NAME: 'TODO: キャンセルリストのシート名',

  /**
   * IS_架電管理シートの列定義
   * 【システム側でセットする列】（A〜H列）
   * 【担当者が入力する列】（I〜P列）
   */
  CALL_MANAGEMENT_COLUMNS: {
    // システム側でセットする列
    CUSTOMER_ID: 0,           // A列: 顧客ID
    CUSTOMER_NAME: 1,         // B列: 顧客名
    EMAIL: 2,                 // C列: メールアドレス
    PHONE_NUMBER: 3,          // D列: 電話番号
    LIST_ADDED_DATE: 4,      // E列: リスト追加日
    SOURCE_TYPE: 5,           // F列: 元リスト種別
    LEAD_SOURCE_ID: 6,       // G列: リードソースID / 名称
    SYSTEM_NOTE: 7,          // H列: システム備考
    
    // 担当者が入力する列
    ASSIGNED_PERSON: 8,       // I列: 担当者
    FIRST_CALL_DATE: 9,      // J列: 初回架電日
    LAST_CALL_DATE: 10,       // K列: 最終架電日
    CALL_COUNT: 11,           // L列: 架電回数
    STATUS: 12,               // M列: ステータス
    NOTE_RANK: 13,            // N列: ネタランク
    NEXT_CALL_DATE: 14,       // O列: 再架電予定日
    DETAIL: 15                // P列: 詳細（自由記載欄）
  },
  
  CALL_MANAGEMENT_HEADERS: [
    // システム側でセットする列（A〜H列）
    'customer_id',           // A列: 顧客ID
    'customer_name',         // B列: 顧客名
    'email',                 // C列: メールアドレス
    'phone_number',          // D列: 電話番号
    'list_added_date',       // E列: リスト追加日
    'source_type',           // F列: 元リスト種別
    'lead_source_id',        // G列: リードソースID / 名称
    'system_note',           // H列: システム備考
    
    // 担当者が入力する列（I〜P列）
    'assigned_person',       // I列: 担当者
    'first_call_date',       // J列: 初回架電日
    'last_call_date',        // K列: 最終架電日
    'call_count',            // L列: 架電回数
    'status',                // M列: ステータス
    'note_rank',             // N列: ネタランク
    'next_call_date',        // O列: 再架電予定日
    'detail'                 // P列: 詳細（自由記載欄）
  ],

  /**
   * 架電管理シートのステータス定義
   */
  CALL_MANAGEMENT_STATUS: {
    NOT_STARTED: '未着手',
    CALLING: '架電中',
    APPOINTMENT: 'アポ取得',
    LOST: '失注',
    HOLD: '保留'
  },

  /**
   * ネタランクの定義
   */
  CALL_MANAGEMENT_NOTE_RANKS: ['A', 'B', 'C', 'D', 'E'],

  /**
   * キャンセルリストの列マッピング
   * 実際のキャンセルリストシートの列構成に合わせて調整してください
   * 
   * 想定される列構成（A列から）:
   * - 顧客ID(メール)
   * - 氏名
   * - 電話番号
   * - 商談予定日
   * - ステータス
   * - リスト追加日
   */
  CANCEL_LIST_COLUMN_MAPPING: {
    CUSTOMER_ID_EMAIL: 0,    // A列: 顧客ID(メール)
    FULL_NAME: 1,            // B列: 氏名
    PHONE_NUMBER: 2,         // C列: 電話番号
    APPOINTMENT_DATE: 3,     // D列: 商談予定日
    STATUS: 4,               // E列: ステータス
    LIST_ADDED_DATE: 5       // F列: リスト追加日
  }
};

