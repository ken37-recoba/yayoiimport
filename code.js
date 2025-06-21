// =================================================================================
// ファイル名: コード.gs
// 役割: システム全体の制御と、主要なイベントハンドラを管理します。
// =================================================================================

/**************************************************************************************************
 * * 領収書・通帳OCRシステム (v6.1 Passbook Master Support)
 * * 概要:
 * - 複数通帳に対応するため、「通帳マスター」シートを導入。
 **************************************************************************************************/
/**************************************************************************************************
 * 1. グローバル設定 (Global Settings)
 **************************************************************************************************/
let CONFIG;
const STATUS = {
  PENDING: '未処理',
  PROCESSING: '処理中',
  PROCESSED: '処理済み',
  ERROR: 'エラー',
};

// ハイライト用の背景色
const DUPLICATE_HIGHLIGHT_COLOR = '#fff799'; // 明るい黄色
const CRITICAL_ERROR_HIGHLIGHT_COLOR = '#ffcccc'; // 明るい赤色

/**
 * スクリプト実行時に最初に呼び出され、設定を読み込む
 */
function loadConfig_() {
  if (CONFIG) return;
  
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('設定');
    if (!sheet) throw new Error('「設定」シートが見つかりません。初回設定が完了していない可能性があります。');
    
    const data = sheet.getRange('A2:B6').getValues();
    const settings = data.reduce((obj, row) => {
      if (row[0]) obj[row[0]] = row[1];
      return obj;
    }, {});

    CONFIG = {
      SPREADSHEET_ID: SpreadsheetApp.getActiveSpreadsheet().getId(),
      // --- 領収書関連 ---
      SOURCE_FOLDER_ID: settings['領収書データ化フォルダID'],
      EXPORT_FOLDER_ID: settings['クライアントフォルダID'],
      ARCHIVE_FOLDER_ID: settings['アーカイブ済みフォルダID'],
      // --- 通帳関連 ---
      PASSBOOK_SOURCE_FOLDER_ID: settings['通帳データ化フォルダID'],
      PASSBOOK_ARCHIVE_FOLDER_ID: settings['通帳アーカイブ済みフォルダID'],
      
      // --- 共通設定 ---
      EXECUTION_TIME_LIMIT_SECONDS: 300,
      MASTER_SHEET: '勘定科目マスター',
      LEARNING_SHEET: '学習データ',
      CONFIG_SHEET: '設定',
      ERROR_LOG_SHEET: 'エラーログ',
      GEMINI_MODEL: 'gemini-1.5-flash-preview-0520',
      THINKING_BUDGET: 10000,
      
      // --- シート名 ---
      FILE_LIST_SHEET: 'ファイルリスト',
      OCR_RESULT_SHEET: 'OCR結果',
      EXPORTED_SHEET: '出力済み',
      TOKEN_LOG_SHEET: 'トークンログ',
      PASSBOOK_FILE_LIST_SHEET: '通帳ファイルリスト',
      PASSBOOK_RESULT_SHEET: '通帳OCR結果',
      PASSBOOK_EXPORTED_SHEET: '通帳出力済み',
      PASSBOOK_MASTER_SHEET: '通帳マスター',
      
      // --- 弥生会計用設定 ---
      YAYOI: {
        SHIKIBETSU_FLAG: '2000',
        KASHIKATA_KAMOKU: '役員借入金',
        KASHIKATA_ZEIKUBUN: '対象外',
        KASHIKATA_ZEIGAKU: '0',
        TORIHIKI_TYPE: '0',
        CHOUSEI: 'no',
        CSV_COLUMNS: [
          '識別フラグ', '伝票NO', '決算整理仕訳', '取引日付', '借方勘定科目', '借方補助科目', '借方部門',
          '借方税区分', '借方金額', '借方税金額', '貸方勘定科目', '貸方補助科目', '貸方部門',
          '貸方税区分', '貸方金額', '貸方税金額', '摘要', '手形番号', '手形期日', '取引タイプ',
          '生成元', '仕訳メモ', '付箋1', '付箋2', '調整'
        ],
      },
      // --- ヘッダー定義 ---
      HEADERS: {
        FILE_LIST: ['ファイルID', 'ファイル名', 'ステータス', 'エラー詳細', '登録日時'],
        OCR_RESULT: [
          '取引ID', '処理日時', '取引日', '店名', '摘要', '勘定科目', '補助科目',
          '税率(%)', '金額(税込)', 'うち消費税', '登録番号',
          '消費税課税区分コード', 'ファイルへのリンク', '備考', '学習チェック'
        ],
        TOKEN_LOG: ['日時', 'ファイル名', '入力トークン', '思考トークン', '出力トークン', '合計トークン'],
        LEARNING: ['店名', '摘要（キーワード）', '金額条件', '金額', '勘定科目', '補助科目', '摘要のテンプレート', '学習登録日時', '取引ID'],
        ERROR_LOG: ['日時', '関数名', 'エラーメッセージ', '関連情報', 'スタックトレース'],
        PASSBOOK_FILE_LIST: ['ファイルID', 'ファイル名', '銀行タイプ', 'ステータス', 'エラー詳細', '登録日時'],
        PASSBOOK_RESULT: [
            '取引ID', '処理日時', '取引日', '摘要', '入金額', '出金額', '残高',
            '通帳勘定科目', '相手方勘定科目', '相手方補助科目',
            '借方税区分', '貸方税区分', 'ファイルへのリンク', '備考'
        ],
        PASSBOOK_MASTER: ['ファイル名キーワード', '弥生会計で使う勘定科目名']
      },
    };
    
    CONFIG.HEADERS.EXPORTED = [...CONFIG.HEADERS.OCR_RESULT, '出力日'];
    CONFIG.HEADERS.PASSBOOK_EXPORTED = [...CONFIG.HEADERS.PASSBOOK_RESULT, '出力日'];

    const requiredIds = ['SOURCE_FOLDER_ID', 'ARCHIVE_FOLDER_ID', 'PASSBOOK_SOURCE_FOLDER_ID', 'PASSBOOK_ARCHIVE_FOLDER_ID'];
    for (const idKey of requiredIds) {
        if (!CONFIG[idKey] || CONFIG[idKey].startsWith('【')) {
            throw new Error(`設定シートの「${idKey}」が正しく設定されていません。`);
        }
    }

  } catch (e) {
    SpreadsheetApp.getUi().alert('設定の読み込みエラー', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    throw e;
  }
}

/**************************************************************************************************
 * 2. セットアップ & メインプロセス (Setup & Main Process)
 **************************************************************************************************/
function onOpen() {
  try {
    loadConfig_();
    const menu = SpreadsheetApp.getUi().createMenu('自動データ化');
    
    const receiptMenu = SpreadsheetApp.getUi().createMenu('領収書処理');
    receiptMenu.addItem('手動で新規ファイルを処理', 'mainProcess');
    receiptMenu.addItem('弥生会計形式でエクスポート', 'exportForYayoi');
    receiptMenu.addItem('選択した取引をOCR結果に戻す', 'moveTransactionsBackToOcr');
    menu.addSubMenu(receiptMenu);

    const passbookMenu = SpreadsheetApp.getUi().createMenu('通帳処理');
    passbookMenu.addItem('手動で新規ファイルを処理', 'mainProcessPassbooks');
    passbookMenu.addItem('弥生会計形式でエクスポート', 'exportPassbookForYayoi');
    passbookMenu.addItem('選択した取引を「通帳OCR結果」に戻す', 'movePassbookTransactionsBackToOcr');
    menu.addSubMenu(passbookMenu);

    menu.addSeparator();

    const checkMenu = SpreadsheetApp.getUi().createMenu('チェックと修正');
    checkMenu.addItem('重複の可能性をチェック (領収書)', 'highlightDuplicates_');
    checkMenu.addItem('重大なエラーをチェック (領収書)', 'highlightCriticalErrors_');
    checkMenu.addItem('選択行のハイライトを解除 (領収書)', 'removeHighlight_');
    checkMenu.addSeparator();
    checkMenu.addItem('選択した取引を削除', 'deleteSelectedTransactions');
    menu.addSubMenu(checkMenu);

    menu.addSeparator();
    
    const settingsMenu = SpreadsheetApp.getUi().createMenu('その他・設定');
    settingsMenu.addItem('【初回/変更時】定期実行をセットアップ', 'createTimeBasedTrigger_');
    settingsMenu.addItem('フィルタをオンにする (現在のシート)', 'activateFilter');
    settingsMenu.addSeparator();
    settingsMenu.addItem('選択行の領収書をプレビュー', 'showReceiptPreview');
    settingsMenu.addItem('選択行の通帳をプレビュー', 'showPassbookPreview');
    menu.addSubMenu(settingsMenu);
    
    menu.addToUi();
  } catch (e) {
    logError_('onOpen', e);
    showError('スクリプトの起動中にエラーが発生しました。\n\n「設定」シートが正しく構成されているか確認してください。\n\n詳細: ' + e.message, '起動エラー');
  }
}

function onEdit(e) {
  try {
    loadConfig_();
    const range = e.range;
    const sheet = range.getSheet();
    const sheetName = sheet.getName();
    const row = range.getRow();
    const col = range.getColumn();
    
    if (sheetName === CONFIG.OCR_RESULT_SHEET && row > 1) {
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const learnCheckColIndex = headers.indexOf('学習チェック') + 1;
      const taxCodeColIndex = headers.indexOf('登録番号') + 1;
      if (col === learnCheckColIndex) {
        handleLearningCheck(sheet, row, col, headers);
      }
      if (col === taxCodeColIndex && range.isBlank()) {
        handleTaxCodeRemoval(sheet, row, headers);
      }
    }
  } catch (err) {
    logError_('onEdit', err);
    console.error("onEdit Error: " + err.toString());
  }
}

function mainProcess() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    console.log('領収書処理：別のプロセスが実行中のためスキップ。');
    return;
  }
  try {
    loadConfig_();
    const startTime = new Date();
    console.log('領収書処理のメインプロセスを開始します。');
    initializeEnvironment();
    processNewFiles();
    performOcrOnPendingFiles(startTime);
    highlightCriticalErrors_();
    highlightDuplicates_();
    console.log('領収書処理のメインプロセスが完了しました。');
  } catch (e) {
    logError_('mainProcess', e);
    showErrorSafe_('領収書処理中にエラーが発生しました。', e);
  } finally {
    lock.releaseLock();
    console.log('領収書処理：ロックを解放しました。');
  }
}

function mainProcessPassbooks() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(11000)) {
    console.log('通帳処理：別のプロセスが実行中のためスキップ。');
    return;
  }
  try {
    loadConfig_();
    const startTime = new Date();
    console.log('通帳処理のメインプロセスを開始します。');
    initializeEnvironment();
    processNewPassbookFiles();
    performOcrOnPassbookFiles(startTime);
    console.log('通帳処理のメインプロセスが完了しました。');
  } catch (e) {
    logError_('mainProcessPassbooks', e);
    showErrorSafe_('通帳処理中にエラーが発生しました。', e);
  } finally {
    lock.releaseLock();
    console.log('通帳処理：ロックを解放しました。');
  }
}

function showErrorSafe_(message, error) {
    console.error(`${message}: ${error.toString()}`);
    try {
        SpreadsheetApp.getActiveSpreadsheet().toast(`${message} 詳細はログを確認してください。`, 'エラー', 10);
    } catch (uiError) {
        console.error("UIの表示にも失敗しました。トリガー実行中の可能性があります。");
    }
}
