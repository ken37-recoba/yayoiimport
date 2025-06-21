// =================================================================================
// ファイル名: コード.gs
// 役割: システム全体の制御と、主要なイベントハンドラを管理します。
// =================================================================================

/**************************************************************************************************
 * * 領収書OCRシステム (v5.0 Refactored)
 * * 概要:
 * 機能ごとにファイルを分割し、メンテナンス性と可読性を向上。
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
    
    const data = sheet.getRange('A2:B10').getValues();
    const settings = data.reduce((obj, row) => {
      if (row[0]) obj[row[0]] = row[1];
      return obj;
    }, {});

    CONFIG = {
      SPREADSHEET_ID: SpreadsheetApp.getActiveSpreadsheet().getId(),
      SOURCE_FOLDER_ID: settings['領収書データ化フォルダID'],
      EXPORT_FOLDER_ID: settings['クライアントフォルダID'],
      ARCHIVE_FOLDER_ID: settings['アーカイブ済みフォルダID'],
      EXECUTION_TIME_LIMIT_SECONDS: 300,
      FILE_LIST_SHEET: 'ファイルリスト',
      OCR_RESULT_SHEET: 'OCR結果',
      EXPORTED_SHEET: '出力済み',
      TOKEN_LOG_SHEET: 'トークンログ',
      MASTER_SHEET: '勘定科目マスター',
      LEARNING_SHEET: '学習データ',
      CONFIG_SHEET: '設定',
      ERROR_LOG_SHEET: 'エラーログ',
      GEMINI_MODEL: 'gemini-2.5-flash-preview-05-20',
      THINKING_BUDGET: 10000,
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
      HEADERS: {
        FILE_LIST: ['ファイルID', 'ファイル名', 'ステータス', 'エラー詳細', '登録日時'],
        OCR_RESULT: [
          '取引ID', '処理日時', '取引日', '店名', '摘要', '勘定科目', '補助科目',
          '税率(%)', '金額(税込)', 'うち消費税', '登録番号',
          '消費税課税区分コード', 'ファイルへのリンク', '備考', '学習チェック'
        ],
        EXPORTED: [
          '取引ID', '処理日時', '取引日', '店名', '摘要', '勘定科目', '補助科目',
          '税率(%)', '金額(税込)', 'うち消費税', '登録番号',
          '消費税課税区分コード', 'ファイルへのリンク', '備考', '学習チェック',
          '出力日'
        ],
        TOKEN_LOG: ['日時', 'ファイル名', '入力トークン', '思考トークン', '出力トークン', '合計トークン'],
        LEARNING: ['店名', '摘要（キーワード）', '金額条件', '金額', '勘定科目', '補助科目', '摘要のテンプレート', '学習登録日時', '取引ID'],
        ERROR_LOG: ['日時', '関数名', 'エラーメッセージ', '関連情報', 'スタックトレース'],
      },
    };

    if (!CONFIG.SOURCE_FOLDER_ID || !CONFIG.EXPORT_FOLDER_ID || !CONFIG.ARCHIVE_FOLDER_ID) {
      throw new Error('設定シートに必要なフォルダIDが設定されていません。');
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
    const menu = SpreadsheetApp.getUi().createMenu('領収書OCR');
    
    menu.addItem('手動で新規ファイルを処理', 'mainProcess');
    menu.addSeparator();
    menu.addItem('【初回/変更時】定期実行をセットアップ', 'createTimeBasedTrigger_');
    menu.addSeparator();
    menu.addItem('重複の可能性をチェック', 'highlightDuplicates_');
    menu.addItem('重大なエラーをチェック', 'highlightCriticalErrors_');
    menu.addItem('選択行のハイライトを解除', 'removeHighlight_');
    menu.addSeparator();
    menu.addItem('選択行の領収書をプレビュー', 'showReceiptPreview');
    menu.addSeparator();
    menu.addItem('弥生会計形式でエクスポート', 'exportForYayoi');
    menu.addItem('選択した取引をOCR結果に戻す', 'moveTransactionsBackToOcr');
    menu.addSeparator();
    menu.addItem('フィルタをオンにする', 'activateFilter');
    menu.addItem('選択した行にダミー番号を挿入', 'insertDummyInvoiceNumber');
    menu.addItem('選択した取引を削除', 'deleteSelectedTransactions');
    
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
    const row = range.getRow();
    const col = range.getColumn();

    if (sheet.getName() !== CONFIG.OCR_RESULT_SHEET || row <= 1) {
      return;
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const learnCheckColIndex = headers.indexOf('学習チェック') + 1;
    const taxCodeColIndex = headers.indexOf('登録番号') + 1;

    // 学習チェックボックスの操作
    if (col === learnCheckColIndex) {
      handleLearningCheck(sheet, row, col, headers);
    }
    // 登録番号の削除
    if (col === taxCodeColIndex && range.isBlank()) {
      handleTaxCodeRemoval(sheet, row, headers);
    }
  } catch (err) {
    logError_('onEdit', err);
    console.error("onEdit Error: " + err.toString());
  }
}

function mainProcess() {
  const lock = LockService.getScriptLock();
  const gotLock = lock.tryLock(10000);

  if (gotLock) {
    try {
      loadConfig_();
      const startTime = new Date();
      console.log('メインプロセスを開始します。');
      initializeEnvironment();
      processNewFiles();
      performOcrOnPendingFiles(startTime);
      highlightCriticalErrors_();
      highlightDuplicates_();
      console.log('メインプロセスが完了しました。');
    } catch (e) {
      logError_('mainProcess', e);
      console.error("メインプロセスの実行中にエラー: " + e.toString());
      try {
          showError('処理中にエラーが発生しました。\n\n詳細: ' + e.message);
      } catch (uiError) {
          console.error("UIの表示にも失敗しました。トリガー実行中の可能性があります。");
      }
    } finally {
      lock.releaseLock();
      console.log('ロックを解放しました。');
    }
  } else {
    console.log('別のプロセスが実行中のため、今回の実行はスキップされました。');
  }
}
