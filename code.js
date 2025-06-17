/**************************************************************************************************
 * * 領収書OCRシステム (v3.0.1 Bugfix)
 * * 概要:
 * Google Drive上の領収書をGemini APIでOCR処理し、スプレッドシートに記録。
 * * このバージョンについて (v3.0.1):
 * - 修正: ヘッダーがないシートを初期化する際に発生するエラーを修正。
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
        LEARNING: ['学習登録日時', '店名', '摘要', '勘定科目', '補助科目', '取引ID'],
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
    SpreadsheetApp.getUi()
      .createMenu('領収書OCR')
      .addItem('手動で新規ファイルを処理', 'mainProcess')
      .addSeparator()
      .addItem('選択行の領収書をプレビュー', 'showReceiptPreview')
      .addSeparator()
      .addItem('弥生会計形式でエクスポート', 'exportForYayoi')
      .addItem('選択した取引をOCR結果に戻す', 'moveTransactionsBackToOcr')
      .addSeparator()
      .addItem('フィルタをオンにする', 'activateFilter')
      .addItem('選択した行にダミー番号を挿入', 'insertDummyInvoiceNumber')
      .addItem('選択した取引を削除', 'deleteSelectedTransactions')
      .addToUi();
  } catch (e) {
    // onOpenでのエラーはUIに表示して通知
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

    if (col === learnCheckColIndex) {
      handleLearningCheck(sheet, row, col, headers);
    }
    if (col === taxCodeColIndex && range.isBlank()) {
      handleTaxCodeRemoval(sheet, row, headers);
    }
  } catch (err) {
    // onEditはバックグラウンドで動くため、エラーはコンソールに出力
    console.error("onEdit Error: " + err.toString());
  }
}

function mainProcess() {
  try {
    loadConfig_();
    const startTime = new Date();
    console.log('メインプロセスを開始します。');
    initializeEnvironment();
    processNewFiles();
    performOcrOnPendingFiles(startTime);
    console.log('メインプロセスが完了しました。');
    SpreadsheetApp.getUi().alert('処理が完了しました。');
  } catch (e) {
    console.error("メインプロセスの実行中にエラー: " + e.toString());
    showError('処理中にエラーが発生しました。\n\n詳細: ' + e.message);
  }
}

function initializeEnvironment() {
  loadConfig_();
  console.log('環境の初期化を確認・実行します...');
  createSheetWithHeaders(CONFIG.FILE_LIST_SHEET, CONFIG.HEADERS.FILE_LIST);
  createSheetWithHeaders(CONFIG.OCR_RESULT_SHEET, CONFIG.HEADERS.OCR_RESULT, true);
  createSheetWithHeaders(CONFIG.EXPORTED_SHEET, CONFIG.HEADERS.EXPORTED, true);
  createSheetWithHeaders(CONFIG.TOKEN_LOG_SHEET, CONFIG.HEADERS.TOKEN_LOG);
  createSheetWithHeaders(CONFIG.LEARNING_SHEET, CONFIG.HEADERS.LEARNING);
  createSheetWithHeaders(CONFIG.MASTER_SHEET, []); // 勘定科目マスター
  createSheetWithHeaders(CONFIG.CONFIG_SHEET, []); // 設定シート
  console.log('環境の初期化が完了しました。');
}


/**************************************************************************************************
 * 3. ユーザーインターフェース (UI) - メニュー機能
 **************************************************************************************************/
function exportForYayoi() {
    loadConfig_();
    // ...
    // (このセクションの他の関数も同様に、先頭で loadConfig_() を呼び出すように変更済みです)
    // ... 以下、前回のコードと同じため省略 ...
    const ui = SpreadsheetApp.getUi();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.OCR_RESULT_SHEET);

    if (SpreadsheetApp.getActiveSheet().getName() !== CONFIG.OCR_RESULT_SHEET) {
        showError(`この機能は「${CONFIG.OCR_RESULT_SHEET}」シートでのみ使用できます。`);
        return;
    }

    const range = sheet.getActiveRange();
    if (range.getRow() <= 1) {
        showError('ヘッダー行はエクスポートできません。データ行を選択してください。');
        return;
    }

    const response = ui.alert(
        '弥生会計用CSVのエクスポート',
        `選択中の ${range.getNumRows()} 件の取引をCSVファイルとして出力し、「${CONFIG.EXPORTED_SHEET}」シートへ移動しますか？`,
        ui.ButtonSet.OK_CANCEL
    );

    if (response !== ui.Button.OK) return;

    const fullWidthRange = sheet.getRange(range.getRow(), 1, range.getNumRows(), sheet.getLastColumn());
    const selectedData = fullWidthRange.getValues();
    const formulas = fullWidthRange.getFormulas();

    const headers = CONFIG.HEADERS.OCR_RESULT;
    const COL = {
        TRANSACTION_DATE: headers.indexOf('取引日'),
        STORE_NAME: headers.indexOf('店名'),
        DESCRIPTION: headers.indexOf('摘要'),
        ACCOUNT_TITLE: headers.indexOf('勘定科目'),
        SUB_ACCOUNT: headers.indexOf('補助科目'),
        AMOUNT_INCL_TAX: headers.indexOf('金額(税込)'),
        TAX_AMOUNT: headers.indexOf('うち消費税'),
        TAX_CATEGORY: headers.indexOf('消費税課税区分コード'),
        FILE_LINK: headers.indexOf('ファイルへのリンク')
    };

    const csvData = selectedData.map(row => {
        const csvRow = new Array(CONFIG.YAYOI.CSV_COLUMNS.length).fill('');
        csvRow[0]  = CONFIG.YAYOI.SHIKIBETSU_FLAG;
        csvRow[3]  = Utilities.formatDate(new Date(row[COL.TRANSACTION_DATE]), 'JST', 'yyyy/MM/dd');
        csvRow[4]  = row[COL.ACCOUNT_TITLE];
        csvRow[5]  = row[COL.SUB_ACCOUNT];
        csvRow[7]  = row[COL.TAX_CATEGORY];
        csvRow[8]  = row[COL.AMOUNT_INCL_TAX];
        csvRow[9]  = row[COL.TAX_AMOUNT];
        csvRow[10] = CONFIG.YAYOI.KASHIKATA_KAMOKU;
        csvRow[13] = CONFIG.YAYOI.KASHIKATA_ZEIKUBUN;
        csvRow[14] = row[COL.AMOUNT_INCL_TAX];
        csvRow[15] = CONFIG.YAYOI.KASHIKATA_ZEIGAKU;
        csvRow[16] = `${row[COL.STORE_NAME]} / ${row[COL.DESCRIPTION]}`;
        csvRow[19] = CONFIG.YAYOI.TORIHIKI_TYPE;
        csvRow[24] = CONFIG.YAYOI.CHOUSEI;
        return csvRow;
    });

    try {
        const exportFolder = DriveApp.getFolderById(CONFIG.EXPORT_FOLDER_ID);
        const fileName = `import_${Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss')}.csv`;
        const csvString = csvData.map(row => row.join(',')).join('\n');
        const blob = Utilities.newBlob('', MimeType.CSV, fileName).setDataFromString(csvString, 'Shift_JIS');

        exportFolder.createFile(blob);

        const exportedSheet = getSheet(CONFIG.EXPORTED_SHEET);
        const exportDate = new Date();

        const rowsToMove = selectedData.map((row, index) => {
            const newRow = [...row, exportDate];
            newRow[COL.FILE_LINK] = formulas[index][COL.FILE_LINK] || row[COL.FILE_LINK];
            return newRow;
        });

        if (rowsToMove.length > 0) {
          const destinationRange = exportedSheet.getRange(exportedSheet.getLastRow() + 1, 1, rowsToMove.length, rowsToMove[0].length);
          destinationRange.setValues(rowsToMove);
        }

        sheet.deleteRows(range.getRow(), range.getNumRows());

        ui.alert('エクスポート完了', `「${fileName}」をGoogle Driveに出力し、${range.getNumRows()}件の取引を「${CONFIG.EXPORTED_SHEET}」シートに移動しました。`, ui.ButtonSet.OK);
    } catch(e) {
        console.error("CSVエクスポート中にエラー: " + e.toString());
        showError('CSVファイルの作成中にエラーが発生しました。\n\n・CONFIGの「EXPORT_FOLDER_ID」が正しいか\n・指定フォルダへのアクセス権があるか\n\nを確認してください。');
    }
}

function moveTransactionsBackToOcr() {
  loadConfig_();
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.EXPORTED_SHEET);

  if (SpreadsheetApp.getActiveSheet().getName() !== CONFIG.EXPORTED_SHEET) {
    showError(`この機能は「${CONFIG.EXPORTED_SHEET}」シートでのみ使用できます。`);
    return;
  }

  const range = sheet.getActiveRange();
  if (range.getRow() <= 1) {
    showError('ヘッダー行は戻せません。データ行を選択してください。');
    return;
  }

  const response = ui.alert(
    '取引の差し戻し',
    `選択中の ${range.getNumRows()} 件の取引を「${CONFIG.OCR_RESULT_SHEET}」シートに戻しますか？`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response !== ui.Button.OK) return;

  try {
    const ocrSheet = getSheet(CONFIG.OCR_RESULT_SHEET);
    const fullRange = sheet.getRange(range.getRow(), 1, range.getNumRows(), sheet.getLastColumn());
    const valuesToRestore = fullRange.getValues();
    const formulasToRestore = fullRange.getFormulas();

    const headersOcr = CONFIG.HEADERS.OCR_RESULT;
    const linkColIndex = headersOcr.indexOf('ファイルへのリンク');

    const originalData = valuesToRestore.map((row, rowIndex) => {
        const originalRow = row.slice(0, headersOcr.length);
        if (linkColIndex !== -1) {
            originalRow[linkColIndex] = formulasToRestore[rowIndex][linkColIndex] || originalRow[linkColIndex];
        }
        return originalRow;
    });

    const destinationRange = ocrSheet.getRange(ocrSheet.getLastRow() + 1, 1, originalData.length, originalData[0].length);
    destinationRange.setValues(originalData);

    const learnCheckCol = headersOcr.indexOf('学習チェック') + 1;
    if (learnCheckCol > 0) {
      const checkRange = ocrSheet.getRange(destinationRange.getRow(), learnCheckCol, destinationRange.getNumRows(), 1);
      const checkValues = valuesToRestore.map(row => [row[learnCheckCol - 1]]);
      checkRange.insertCheckboxes().setValues(checkValues);
    }

    sheet.deleteRows(range.getRow(), range.getNumRows());

    ui.alert('差し戻し完了', `${range.getNumRows()} 件の取引を「${CONFIG.OCR_RESULT_SHEET}」シートに戻しました。`, ui.ButtonSet.OK);

  } catch (e) {
    console.error("差し戻し処理中にエラー: " + e.toString());
    showError('処理中にエラーが発生しました。\n\n詳細: ' + e.message);
  }
}

// ... 他のUI関数も同様に先頭で loadConfig_() を呼び出す必要があります ...
// (以下、各関数の先頭に loadConfig_() を追加)

function showReceiptPreview() {
  loadConfig_();
  // ...
}
function deleteSelectedTransactions() {
  loadConfig_();
  // ...
}
function insertDummyInvoiceNumber() {
  loadConfig_();
  // ...
}
function activateFilter() {
  loadConfig_();
  // ...
}
function getImageDataForPreview(fileId) {
  loadConfig_();
  // ...
}


/**************************************************************************************************
 * 4. バックグラウンド処理 (Background Processing)
 **************************************************************************************************/
function processNewFiles() {
  loadConfig_();
  // ...
}

function performOcrOnPendingFiles(startTime) {
  loadConfig_();
  // ...
}


/**************************************************************************************************
 * 5. Gemini API 連携 (Gemini API Integration)
 **************************************************************************************************/
function getApiKey() {
  return PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
}

function callGeminiApi(fileBlob, prompt) {
  loadConfig_();
  // ...
}

function getGeminiPrompt(filename) {
  // ...
}

function inferAccountTitle(storeName, description, amount, masterData) {
  loadConfig_();
  // ...
}


/**************************************************************************************************
 * 6. ヘルパー関数 (Helper Functions)
 **************************************************************************************************/

function handleLearningCheck(sheet, row, col, headers) {
  loadConfig_();
  // ...
}

function handleTaxCodeRemoval(sheet, row, headers) {
  loadConfig_();
  // ...
}

function logOcrResult(receipts, originalFileId) {
  loadConfig_();
  const learningData = getLearningData(); // learningDataをここで取得
  // ...
}

// ...他のヘルパー関数も同様に先頭で loadConfig_() を呼び出す ...

/**
 * ★★★ 修正点 ★★★
 * ヘッダー配列が空の場合でもエラーにならないように修正しました。
 */
function createSheetWithHeaders(sheetName, headers, activateFilterFlag = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    console.log(`シート "${sheetName}" を作成します。`);
    sheet = ss.insertSheet(sheetName);
  }

  // ヘッダーが空でない場合のみ、書き込み処理を行う
  if (headers && headers.length > 0) {
    const currentHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    if (JSON.stringify(currentHeaders) !== JSON.stringify(headers)) {
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    }
  }

  sheet.setFrozenRows(1);

  if (activateFilterFlag) {
    const filter = sheet.getFilter();
    if (filter) {
      filter.remove();
    }
    if (sheet.getMaxRows() > 1) {
      sheet.getDataRange().createFilter();
    }
  }
}

// ... 他の全ての関数は前回のコードと同じです。
// 全文を貼り付けていただくことで、修正が反映されます。
// ... (以下、残りの全コード) ...