/**************************************************************************************************
 * * 領収書OCRシステム (v4.5 Model Pinning Corrected)
 * * 概要:
 * Google Drive上の領収書をGemini APIでOCR処理し、スプレッドシートに記録。
 * * このバージョンについて (v4.5):
 * - 修正: ユーザー環境で動作が確認された正しいモデル名 `gemini-2.5-flash-preview-05-20` に修正。
 * - 新機能: エラーログ機能を実装。
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
      ERROR_LOG_SHEET: 'エラーログ',
      GEMINI_MODEL: 'gemini-2.5-flash-preview-05-20', // ★★★ 修正: 正しいモデルバージョンに修正
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
        ERROR_LOG: ['日時', '関数名', 'エラーメッセージ', '関連情報', 'スタックトレース'],
      },
    };

    if (!CONFIG.SOURCE_FOLDER_ID || !CONFIG.EXPORT_FOLDER_ID || !CONFIG.ARCHIVE_FOLDER_ID) {
      throw new Error('設定シートに必要なフォルダIDが設定されていません。');
    }

  } catch (e) {
    // この段階のエラーはログシートに書き込めない可能性が高いため、UIにのみ表示
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

    if (col === learnCheckColIndex) {
      handleLearningCheck(sheet, row, col, headers);
    }
    if (col === taxCodeColIndex && range.isBlank()) {
      handleTaxCodeRemoval(sheet, row, headers);
    }
  } catch (err) {
    logError_('onEdit', err);
    console.error("onEdit Error: " + err.toString()); // ユーザー操作を妨げないようUIエラーは表示しない
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
    logError_('mainProcess', e);
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
  createSheetWithHeaders(CONFIG.MASTER_SHEET, []); 
  createSheetWithHeaders(CONFIG.CONFIG_SHEET, []); 
  createSheetWithHeaders(CONFIG.ERROR_LOG_SHEET, CONFIG.HEADERS.ERROR_LOG);
  console.log('環境の初期化が完了しました。');
}
/**************************************************************************************************
 * 3. ユーザーインターフェース (UI) - メニュー機能
 **************************************************************************************************/
function exportForYayoi() {
    loadConfig_();
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

    try {
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
        logError_('exportForYayoi', e);
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
    logError_('moveTransactionsBackToOcr', e);
    console.error("差し戻し処理中にエラー: " + e.toString());
    showError('処理中にエラーが発生しました。\n\n詳細: ' + e.message);
  }
}

function showReceiptPreview() {
  loadConfig_();
  const ui = SpreadsheetApp.getUi();
  let fileInfo = 'N/A';
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    if (sheet.getName() !== CONFIG.OCR_RESULT_SHEET && sheet.getName() !== CONFIG.EXPORTED_SHEET) {
      showError(`この機能は「${CONFIG.OCR_RESULT_SHEET}」または「${CONFIG.EXPORTED_SHEET}」シートで実行してください。`);
      return;
    }

    const range = sheet.getActiveRange();
    const startRow = range.getRow();
    if (startRow <= 1) {
      showError('データ行を選択してください。');
      return;
    }

    const fileId = getFileIdFromCell(sheet, startRow);
    if (!fileId) return;
    fileInfo = `File ID: ${fileId}`;

    const htmlTemplate = HtmlService.createTemplateFromFile('Preview');
    htmlTemplate.fileId = fileId;

    const htmlOutput = htmlTemplate.evaluate().setWidth(700).setHeight(800);
    ui.showModalDialog(htmlOutput, `領収書プレビュー`);

  } catch (e) {
    logError_('showReceiptPreview', e, fileInfo);
    console.error('プレビュー表示中にエラーが発生しました: ' + e.toString());
    showError('プレビューの表示中にエラーが発生しました。\n\n詳細: ' + e.message);
  }
}

function deleteSelectedTransactions() {
  loadConfig_();
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const sheetName = sheet.getName();
  let contextInfo = `Sheet: ${sheetName}`;

  try {
    if (sheetName !== CONFIG.OCR_RESULT_SHEET && sheetName !== CONFIG.EXPORTED_SHEET) {
      showError(`この機能は「${CONFIG.OCR_RESULT_SHEET}」または「${CONFIG.EXPORTED_SHEET}」シートでのみ使用できます。`);
      return;
    }

    const range = sheet.getActiveRange();
    const startRow = range.getRow();
    contextInfo += `, Range: ${range.getA1Notation()}`;

    if (startRow <= 1) {
      showError('ヘッダー行は削除できません。データ行を選択してください。');
      return;
    }

    const response = ui.alert(
      '選択した取引の完全削除',
      `選択中の ${range.getNumRows()} 件の取引を完全に削除しますか？\n\n学習済みの取引が含まれている場合、関連する学習データも完全に削除されます。この操作は元に戻せません。`,
      ui.ButtonSet.OK_CANCEL
    );

    if (response !== ui.Button.OK) return;

    const headers = (sheetName === CONFIG.OCR_RESULT_SHEET) ? CONFIG.HEADERS.OCR_RESULT : CONFIG.HEADERS.EXPORTED;
    const transactionIdColIndex = headers.indexOf('取引ID');

    const fullRange = sheet.getRange(startRow, 1, range.getNumRows(), sheet.getLastColumn());
    const selectedRows = fullRange.getValues();

    const transactionIdsToDelete = selectedRows.map(row => row[transactionIdColIndex]).filter(id => id);
    contextInfo += `, Transaction IDs: ${transactionIdsToDelete.join(', ')}`;

    let learnedDeletedCount = 0;
    if (transactionIdsToDelete.length > 0) {
      learnedDeletedCount = deleteLearningDataByIds(transactionIdsToDelete);
    }

    sheet.deleteRows(startRow, range.getNumRows());

    ui.alert('処理完了', `${range.getNumRows()}件の取引を完全に削除しました。\n(うち、${learnedDeletedCount}件の学習データも関連して削除されました。)`, ui.ButtonSet.OK);
  
  } catch(e) {
    logError_('deleteSelectedTransactions', e, contextInfo);
    showError('削除処理中に予期せぬエラーが発生しました。\n\n詳細: ' + e.message);
  }
}

function insertDummyInvoiceNumber() {
    loadConfig_();
    try {
        const sheet = SpreadsheetApp.getActiveSheet();
        const sheetName = sheet.getName();

        if (sheetName !== CONFIG.OCR_RESULT_SHEET) {
            showError(`この機能は「${CONFIG.OCR_RESULT_SHEET}」シートでのみ使用できます。`);
            return;
        }

        const range = sheet.getActiveRange();
        if (range.getRow() <= 1) {
            showError('ヘッダー行には適用できません。データ行を選択してください。');
            return;
        }

        const headers = CONFIG.HEADERS.OCR_RESULT;
        const taxCodeCol = headers.indexOf('登録番号') + 1;
        const taxRateCol = headers.indexOf('税率(%)') + 1;
        const taxCategoryCol = headers.indexOf('消費税課税区分コード') + 1;

        if ([taxCodeCol, taxRateCol, taxCategoryCol].includes(0)) {
            showError('必要な列（登録番号、税率(%)、消費税課税区分コード）が見つかりません。');
            return;
        }

        const dataRange = sheet.getRange(range.getRow(), 1, range.getNumRows(), sheet.getLastColumn());
        const values = dataRange.getValues();
        let updatedCount = 0;

        values.forEach((row, i) => {
            if (!row[taxCodeCol - 1]) {
                const dummyNumber = 'T' + Math.random().toString().slice(2, 15);
                row[taxCodeCol - 1] = dummyNumber;
                row[taxCategoryCol - 1] = getTaxCategoryCode(row[taxRateCol - 1], dummyNumber);

                const taxCodeCell = sheet.getRange(range.getRow() + i, taxCodeCol);
                taxCodeCell.setValue(dummyNumber).setFontColor("#0000FF");
                sheet.getRange(range.getRow() + i, taxCategoryCol).setValue(row[taxCategoryCol - 1]);
                updatedCount++;
            }
        });

        if (updatedCount > 0) {
            SpreadsheetApp.getUi().alert('処理完了', `${updatedCount}件の取引にダミーの登録番号を挿入し、税区分を更新しました。`, SpreadsheetApp.getUi().ButtonSet.OK);
        } else {
            showError('処理対象なし', '選択された行に、登録番号が空欄の取引はありませんでした。');
        }
    } catch (e) {
        logError_('insertDummyInvoiceNumber', e, `Sheet: ${SpreadsheetApp.getActiveSheet().getName()}`);
        showError('ダミー番号の挿入中にエラーが発生しました。\n\n詳細: ' + e.message);
    }
}

function activateFilter() {
  loadConfig_();
  const sheet = SpreadsheetApp.getActiveSheet();
  if (sheet) {
    if (sheet.getFilter()) {
      sheet.getFilter().remove();
    }
    sheet.getDataRange().createFilter();
    SpreadsheetApp.getUi().alert(`シート「${sheet.getName()}」にフィルタをオンにしました。`);
  }
}

function getImageDataForPreview(fileId) {
  loadConfig_();
  try {
    const file = DriveApp.getFileById(fileId);
    const originalBlob = file.getBlob();
    const imageBlob = (originalBlob.getContentType() === MimeType.PDF)
      ? originalBlob.getAs('image/png')
      : originalBlob;

    const dataUrl = `data:${imageBlob.getContentType()};base64,${Utilities.base64Encode(imageBlob.getBytes())}`;

    return {
      success: true,
      fileName: file.getName(),
      dataUrl: dataUrl
    };
  } catch (e) {
    logError_('getImageDataForPreview', e, `File ID: ${fileId}`);
    console.error('画像データの取得中にエラー: ' + e.toString());
    return { success: false, error: e.message };
  }
}
/**************************************************************************************************
 * 4. バックグラウンド処理 (Background Processing)
 **************************************************************************************************/
function processNewFiles() {
  loadConfig_();
  console.log('ステップ1: 新規ファイルの処理を開始...');
  try {
    const sourceFolder = DriveApp.getFolderById(CONFIG.SOURCE_FOLDER_ID);
    const fileListSheet = getSheet(CONFIG.FILE_LIST_SHEET);
    
    const lastRow = fileListSheet.getLastRow();
    let existingFileIds = [];
    if (lastRow >= 2) {
        existingFileIds = fileListSheet.getRange(2, 1, lastRow - 1, 1).getValues()
            .flat().filter(id => id);
    }

    const files = sourceFolder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      const fileId = file.getId();

      if (existingFileIds.includes(fileId)) continue;

      const mimeType = file.getMimeType();
      if (mimeType === MimeType.PDF || mimeType.startsWith('image/')) {
        fileListSheet.appendRow([fileId, file.getName(), STATUS.PENDING, '', new Date()]);
        existingFileIds.push(fileId);
      }
    }
    console.log('ステップ1: 新規ファイルの処理が完了しました。');
  } catch(e) {
    logError_('processNewFiles', e);
    throw e; // メインプロセスにエラーを伝播させる
  }
}

function performOcrOnPendingFiles(startTime) {
  loadConfig_();
  console.log('ステップ2: OCR処理を開始...');
  const fileListSheet = getSheet(CONFIG.FILE_LIST_SHEET);
  const archiveFolder = DriveApp.getFolderById(CONFIG.ARCHIVE_FOLDER_ID);
  const data = fileListSheet.getDataRange().getValues();
  const learningData = getLearningData();

  for (let i = 1; i < data.length; i++) {
    const elapsedTime = (new Date().getTime() - startTime.getTime()) / 1000;
    if (elapsedTime > CONFIG.EXECUTION_TIME_LIMIT_SECONDS) {
      console.log(`実行時間が上限(${CONFIG.EXECUTION_TIME_LIMIT_SECONDS}秒)に近づいたため、処理を中断します。`);
      break;
    }

    const rowData = data[i];
    if (rowData[2] === STATUS.PENDING) {
      const fileId = rowData[0];
      const fileName = rowData[1];
      const rowNum = i + 1;
      const contextInfo = `File: ${fileName} (ID: ${fileId})`;

      try {
        fileListSheet.getRange(rowNum, 3).setValue(STATUS.PROCESSING);
        SpreadsheetApp.flush();
        console.log(`OCR処理を開始: ${fileName}`);

        const file = DriveApp.getFileById(fileId);
        const result = callGeminiApi(file.getBlob(), getGeminiPrompt(fileName));

        Utilities.sleep(1500);

        if (result.success) {
          const ocrData = JSON.parse(result.data);
          if (ocrData && ocrData.length > 0) {
            logOcrResult(ocrData, file.getId());
            logTokenUsage(fileName, result.usage);
            fileListSheet.getRange(rowNum, 3, 1, 2).setValues([[STATUS.PROCESSED, '']]);
            file.moveTo(archiveFolder);
            console.log(`OCR処理成功: ${fileName}`);
          } else {
            console.log(`ファイル ${fileName} から領収書は検出されませんでした。処理待ちに戻します。`);
            fileListSheet.getRange(rowNum, 3).setValue(STATUS.PENDING);
          }
        } else {
          throw new Error(result.error);
        }
      } catch (e) {
        const errorMessage = e.message || e.toString();
        logError_('performOcrOnPendingFiles', e, contextInfo);
        console.error(`OCR処理中にエラー: ${fileName}, Error: ${errorMessage}`);
        fileListSheet.getRange(rowNum, 3, 1, 2).setValues([[STATUS.ERROR, errorMessage]]);
      } finally {
        fileListSheet.getRange(rowNum, 5).setValue(new Date());
      }
    }
  }
  console.log('ステップ2: OCR処理が完了しました。');
}
/**************************************************************************************************
 * 5. Gemini API 連携 (Gemini API Integration)
 **************************************************************************************************/
function getApiKey() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    const error = new Error('APIキーがスクリプトプロパティに設定されていません。プロジェクトの設定を確認してください。');
    logError_('getApiKey', error);
    throw error;
  }
  return apiKey;
}

function callGeminiApi(fileBlob, prompt) {
  loadConfig_();
  try {
    const apiKey = getApiKey();
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${CONFIG.GEMINI_MODEL}:generateContent?key=${apiKey}`;

    const payload = {
      "contents": [{
        "parts": [
          { "text": prompt },
          { "inline_data": { "mime_type": fileBlob.getContentType(), "data": Utilities.base64Encode(fileBlob.getBytes()) }}
        ]
      }],
      "generationConfig": {
        "responseMimeType": "application/json",
        "temperature": 0.1,
        "thinkingConfig": { "thinkingBudget": CONFIG.THINKING_BUDGET }
      }
    };

    const options = {
      'method': 'post',
      'contentType': 'application/json',
      'payload': JSON.stringify(payload),
      'muteHttpExceptions': true
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      const jsonResponse = JSON.parse(responseBody);
      if (jsonResponse.candidates && jsonResponse.candidates[0].content && jsonResponse.candidates[0].content.parts) {
        return { success: true, data: jsonResponse.candidates[0].content.parts[0].text, usage: jsonResponse.usageMetadata };
      } else {
        const errorMsg = "APIからのレスポンスが予期した形式ではありません。";
        logError_('callGeminiApi', new Error(errorMsg), `Response: ${responseBody}`);
        return { success: false, error: errorMsg, usage: jsonResponse.usageMetadata || null };
      }
    } else {
      const errorMsg = `API Error ${responseCode}: ${responseBody}`;
      logError_('callGeminiApi', new Error(errorMsg), `File Type: ${fileBlob.getContentType()}`);
      return { success: false, error: errorMsg };
    }
  } catch(e) {
    logError_('callGeminiApi', e, `File Type: ${fileBlob.getContentType()}`);
    return { success: false, error: e.message };
  }
}

function getGeminiPrompt(filename) {
  return `
processing_context:
  processing_date: "${new Date()}"
role_and_responsibility:
  role: プロの経理
  task: 会計ソフトへデータを入力するために領収書の情報をルールに従い間違いなく整理する必要がある
input_characteristics:
  file_types:
    - 画像 (PNG, JPEG等)
    - PDF
  quality: 解像度や鮮明度は様々
  layout: 領収書のレイアウトやデザインは多岐にわたる
output_specifications:
  format: JSON
  instruction: |-
    JSON形式のみを出力し、他の文言を入れないようにしてください。
    【最重要ルール】もし1枚の領収書に複数の消費税率（例: 10%と8%）が混在している場合は、必ず税率ごとに別のReceiptオブジェクトとして分割して生成してください。
    例えば、10%と8%の品目が含まれる領収書は、2つのReceiptオブジェクト（1つは10%用、もう1つは8%用）に分けてください。
  type_name: Receipts
  type_definition: |
    interface Receipt {
      date: string; // 取引日を「yyyy/mm/dd」形式で出力。和暦は西暦に変換。どうしても日付が読み取れない場合のみnullとする。
      storeName: string; // 領収書の発行者（取引相手）
      description: string; // その税率に該当する取引の明細内容。複数ある場合は " | " で区切る。
      tax_rate: number; // 消費税率（例: 10, 8, 0）。免税・非課税は0とする。
      amount: number; // その税率に対応する「税込」の合計金額。
      tax_amount: number; // amountのうち、消費税額に該当する金額。
      tax_code: string; // 登録番号（Tから始まる13桁の番号）。同じ領収書なら同じ番号を記載。
      filename: string; // 入力されたデータのファイル名。同じ領収書なら同じファイル名を記載。
      note: string; // 特記事項（読取り、消費税、登録番号の不備について記載する）
    }
    type Receipts = Receipt[];
  filename_rule:
    placeholder: "${filename}"
    instruction: filenameフィールドに必ず ${filename} の値を設定する
  duplication_rule:
    description: "1枚の領収書から複数の税率の行を生成する場合、date, storeName, tax_code, filenameはすべての行で同じ内容を記載してください。"
data_rules:
  date:
    formats: ['YYYY/MM/DD', 'YYYY-MM-DD', '和暦']
    year_interpretation:
      condition: 年度表記が不明瞭な場合
      rule: processing_context.processing_date を基準に最も近い過去の日付として年を補完する。
  amount_rules:
    - 金額はカンマや通貨記号なしの数字のみとする
  character_rules:
    - 環境依存文字は使用しない
    - 内容にカンマ、句読点は含めない
    - 複数の内容を列挙する場合の区切り文字は " | " とする
  missing_data_rule:
    description: 記載すべき内容が存在しない場合、数値は0、文字列はnullとする。
special_note_examples:
  - condition: 手書き文字等で文字認識に懸念がある場合
    note_content: "印字不鮮明"
  - condition: 金額の合計や日付形式に矛盾がある場合
    note_content: "合計金額の不一致"
`;
}

function inferAccountTitle(storeName, description, amount, masterData) {
  loadConfig_();
  const contextInfo = `Store: ${storeName}, Desc: ${description}, Amount: ${amount}`;
  try {
    const apiKey = getApiKey();
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${CONFIG.GEMINI_MODEL}:generateContent?key=${apiKey}`;

    const masterListWithKeywords = masterData.map(row => ({ title: row[0], keywords: row[1] || "特になし" }));
    const masterTitleList = masterData.map(row => row[0]);

    const prompt = `あなたは、日本の会計基準に精通したベテランの経理専門家です。あなたの任務は、与えられた領収書の情報と、社内ルールを含む勘定科目マスターを基に、最も可能性の高い勘定科目を特定することです。# 指示\n1. 以下の「領収書情報」と「勘定科目マスター」を注意深く分析してください。\n2. 特に「勘定科目マスター」の**キーワード/ルール**は重要です。例えば、「2万円未満の飲食代は会議費」といった金額に基づくルールが含まれている場合があります。\n3. すべての情報を総合的に判断し、「勘定科目マスター」のリストの中から最も適切だと考えられる勘定科目を**1つだけ**選択してください。\n4. あなたの回答は、必ず指定されたJSON形式に従ってください。# 領収書情報\n- 店名: ${storeName}\n- 摘要: ${description}\n- 金額(税込): ${amount}円\n# 勘定科目マスター（キーワード/ルールを含む）\n${JSON.stringify(masterListWithKeywords)}`;

    const payload = {
      "contents": [{"parts": [{ "text": prompt }]}],
      "generationConfig": {
        "responseMimeType": "application/json",
        "temperature": 0,
        "responseSchema": {
          "type": "OBJECT",
          "properties": { "accountTitle": { "type": "STRING", "enum": masterTitleList }},
          "required": ["accountTitle"]
        }
      }
    };

    const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true };
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      const inferredText = JSON.parse(responseBody).candidates?.[0]?.content?.parts?.[0]?.text;
      if (inferredText) {
        const finalAnswer = JSON.parse(inferredText);
        if (finalAnswer.accountTitle && masterTitleList.includes(finalAnswer.accountTitle)) {
           return finalAnswer.accountTitle;
        }
      }
      const errorMsg = "AIからのJSONレスポンスの形式が不正です。";
      logError_('inferAccountTitle', new Error(errorMsg), `${contextInfo}, Response: ${responseBody}`);
      return "【形式エラー】";
    } else {
      const errorMsg = `勘定科目推測APIエラー [${responseCode}]: ${responseBody}`;
      logError_('inferAccountTitle', new Error(errorMsg), contextInfo);
      return `【APIエラー ${responseCode}】`;
    }
  } catch (e) {
    logError_('inferAccountTitle', e, contextInfo);
    return "【推測エラー】";
  }
}
/**************************************************************************************************
 * 6. ヘルパー関数 (Helper Functions)
 **************************************************************************************************/
/**
 * エラー情報をスプレッドシートに記録する
 * @param {string} functionName - エラーが発生した関数名
 * @param {Error} error - 発生したErrorオブジェクト
 * @param {string} [contextInfo=''] - エラーの文脈情報（例: ファイルID、シート名など）
 */
function logError_(functionName, error, contextInfo = '') {
    try {
        if (!CONFIG) { // CONFIGが未ロードの場合はログをスキップ
            console.error(`[${functionName}] CONFIG not loaded. Error: ${error.stack || error.message}`);
            return;
        }
        const sheet = getSheet(CONFIG.ERROR_LOG_SHEET);
        if (!sheet) {
            console.error(`[${functionName}] Error log sheet not found. Error: ${error.stack || error.message}`);
            return;
        }
        sheet.appendRow([
            new Date(),
            functionName,
            error.message,
            contextInfo,
            error.stack || 'N/A'
        ]);
    } catch (logErr) {
        console.error(`Failed to write to error log. Original error in ${functionName}: ${error.stack || error.message}. Logging error: ${logErr.message}`);
    }
}

function handleLearningCheck(sheet, row, col, headers) {
  loadConfig_();
  const transactionId = sheet.getRange(row, headers.indexOf('取引ID') + 1).getValue();
  let contextInfo = `Transaction ID: ${transactionId}, Cell: ${sheet.getRange(row, col).getA1Notation()}`;
  try {
    const range = sheet.getRange(row, col);
    if (!transactionId) return;

    if (range.isChecked()) {
      if (range.getNote().includes('学習済み')) return;

      const dataRow = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
      const storeName = dataRow[headers.indexOf('店名')];
      const description = dataRow[headers.indexOf('摘要')];
      const kanjo = dataRow[headers.indexOf('勘定科目')];
      const hojo = dataRow[headers.indexOf('補助科目')];

      getSheet(CONFIG.LEARNING_SHEET).appendRow([new Date(), storeName, description, kanjo, hojo, transactionId]);

      range.setNote(`学習済み (ID: ${transactionId})`);
      sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground('#e6f4ea');
      SpreadsheetApp.getActiveSpreadsheet().toast(`「${storeName}」の勘定科目を学習しました。`);
    } else {
      const deletedCount = deleteLearningDataByIds([transactionId]);
      if (deletedCount > 0) {
        range.clearNote();
        sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground(null);
        SpreadsheetApp.getActiveSpreadsheet().toast(`取引ID ${transactionId} の学習データを取り消しました。`);
      }
    }
  } catch (e) {
    logError_('handleLearningCheck', e, contextInfo);
  }
}

function handleTaxCodeRemoval(sheet, row, headers) {
  loadConfig_();
  let contextInfo = `Sheet: ${sheet.getName()}, Row: ${row}`;
  try {
    const taxRateCol = headers.indexOf('税率(%)') + 1;
    const taxCategoryCol = headers.indexOf('消費税課税区分コード') + 1;

    if (taxRateCol > 0 && taxCategoryCol > 0) {
      const taxRate = sheet.getRange(row, taxRateCol).getValue();
      const newTaxCategory = getTaxCategoryCode(taxRate, "");
      sheet.getRange(row, taxCategoryCol).setValue(newTaxCategory);
      SpreadsheetApp.getActiveSpreadsheet().toast(`行 ${row} の登録番号が削除されたため、税区分を更新しました。`);
    }
  } catch(e) {
    logError_('handleTaxCodeRemoval', e, contextInfo);
  }
}

function logOcrResult(receipts, originalFileId) {
  loadConfig_();
  let contextInfo = `File ID: ${originalFileId}`;
  try {
    const sheet = getSheet(CONFIG.OCR_RESULT_SHEET);
    const originalFile = DriveApp.getFileById(originalFileId);
    const masterData = getMasterData();
    const learningData = getLearningData();
    
    const learnedKeys = Object.keys(learningData).sort((a, b) => b.length - a.length);

    const newRows = receipts.map(r => {
      let kanjo = null, hojo = null;
      let isLearned = false;

      const normalizedOcrName = normalizeStoreName(r.storeName);
      for (const learnedKey of learnedKeys) {
        if (normalizedOcrName.includes(learnedKey)) {
          const learned = learningData[learnedKey];
          kanjo = learned.kanjo;
          hojo = learned.hojo;
          isLearned = true;
          console.log(`学習データを適用: OCR店名「${r.storeName}」が学習済み店名「${learned.raw}」(正規化: ${learnedKey})に一致しました。`);
          break;
        }
      }

      if (!isLearned) {
        kanjo = inferAccountTitle(r.storeName, r.description, r.amount, masterData);
        hojo = "";
      }

      return [
        Utilities.getUuid(), new Date(), r.date, r.storeName, r.description,
        kanjo, hojo, r.tax_rate, r.amount, r.tax_amount, r.tax_code,
        getTaxCategoryCode(r.tax_rate, r.tax_code),
        `=HYPERLINK("${originalFile.getUrl()}","${r.filename || originalFile.getName()}")`,
        r.note
      ];
    });

    if (newRows.length > 0) {
      const startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, newRows.length, newRows[0].length).setValues(newRows);

      const headers = CONFIG.HEADERS.OCR_RESULT;
      const learnCheckCol = headers.indexOf('学習チェック') + 1;
      if (learnCheckCol > 0) {
        sheet.getRange(startRow, learnCheckCol, newRows.length).insertCheckboxes();
      }
    }
  } catch (e) {
    logError_('logOcrResult', e, contextInfo);
    throw e; // エラーを呼び出し元に伝播
  }
}

function logTokenUsage(fileName, usage) {
  loadConfig_();
  try {
    const sheet = getSheet(CONFIG.TOKEN_LOG_SHEET);
    sheet.appendRow([
      new Date(), fileName,
      usage.promptTokenCount || 0, usage.thoughtsTokenCount || 0,
      usage.candidatesTokenCount || 0, usage.totalTokenCount || 0
    ]);
  } catch (e) {
    logError_('logTokenUsage', e, `File: ${fileName}`);
  }
}

function deleteLearningDataByIds(transactionIds) {
  loadConfig_();
  let contextInfo = `Transaction IDs: ${transactionIds.join(', ')}`;
  try {
    const learningSheet = getSheet(CONFIG.LEARNING_SHEET);
    if (!learningSheet || learningSheet.getLastRow() < 2) return 0;

    const data = learningSheet.getRange(2, 1, learningSheet.getLastRow() - 1, CONFIG.HEADERS.LEARNING.length).getValues();
    const idCol = CONFIG.HEADERS.LEARNING.indexOf('取引ID');
    let deletedCount = 0;

    for (let i = data.length - 1; i >= 0; i--) {
      if (transactionIds.includes(data[i][idCol])) {
        learningSheet.deleteRow(i + 2);
        deletedCount++;
      }
    }
    return deletedCount;
  } catch (e) {
    logError_('deleteLearningDataByIds', e, contextInfo);
    return 0; // エラー時は0を返す
  }
}

function getTaxCategoryCode(taxRate, taxCode) {
  loadConfig_();
  const hasInvoiceNumber = taxCode && taxCode.match(/^T\d{13}$/);
  if (taxRate === 10) return hasInvoiceNumber ? '共対仕入内10%適格' : '共対仕入内10%区分80%';
  if (taxRate === 8) return hasInvoiceNumber ? '共対仕入内軽減8%適格' : '共対仕入内軽減8%区分80%';
  return '対象外';
}

function getLearningData() {
  loadConfig_();
  const learningData = {};
  try {
    const sheet = getSheet(CONFIG.LEARNING_SHEET);
    if (sheet && sheet.getLastRow() > 1) {
      const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, CONFIG.HEADERS.LEARNING.length).getValues();
      for (let i = data.length - 1; i >= 0; i--) {
        const row = data[i];
        const rawStoreName = row[1];
        const normalizedStoreName = normalizeStoreName(rawStoreName);
        
        if (normalizedStoreName && !learningData[normalizedStoreName]) {
          learningData[normalizedStoreName] = { 
            kanjo: row[3], 
            hojo: row[4],
            raw: rawStoreName
          };
        }
      }
    }
  } catch(e) {
    logError_("getLearningData", e);
    console.error("学習データの取得に失敗: " + e.toString());
  }
  return learningData;
}

function getMasterData() {
  loadConfig_();
  try {
    const sheet = getSheet(CONFIG.MASTER_SHEET);
    if (!sheet || sheet.getLastRow() < 2) return [];
    return sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().filter(row => row[0]);
  } catch (e) {
    logError_("getMasterData", e);
    console.error(e);
    showError(`シート「${CONFIG.MASTER_SHEET}」からデータを取得できませんでした。`);
    return [];
  }
}

function getFileIdFromCell(sheet, row) {
  loadConfig_();
  let contextInfo = `Sheet: ${sheet.getName()}, Row: ${row}`;
  try {
    const headers = (sheet.getName() === CONFIG.OCR_RESULT_SHEET) ? CONFIG.HEADERS.OCR_RESULT : CONFIG.HEADERS.EXPORTED;
    const linkCol = headers.indexOf('ファイルへのリンク') + 1;
    if (linkCol === 0) {
      showError('「ファイルへのリンク」列が見つかりません。');
      return null;
    }

    const cellFormula = sheet.getRange(row, linkCol).getFormula();
    if (!cellFormula) {
      showError('選択した行にファイルへのリンクがありません。');
      return null;
    }
    contextInfo += `, Formula: ${cellFormula}`;

    const urlMatch = cellFormula.match(/HYPERLINK\("([^"]+)"/);
    if (!urlMatch || !urlMatch[1]) {
      showError('リンクの形式が正しくありません。');
      return null;
    }

    const fileUrl = urlMatch[1];
    const idMatch = fileUrl.match(/d\/([a-zA-Z0-9_-]{28,})/) || fileUrl.match(/id=([a-zA-Z0-9_-]{28,})/);

    if (!idMatch || !idMatch[1]) {
      showError('ファイルURLからIDを抽出できませんでした。URL: ' + fileUrl);
      return null;
    }
    return idMatch[1];
  } catch(e) {
    logError_('getFileIdFromCell', e, contextInfo);
    showError('リンクの解析中に予期せぬエラーが発生しました。');
    return null;
  }
}

function normalizeStoreName(name) {
  if (!name || typeof name !== 'string') {
    return '';
  }
  return name
    .toLowerCase()
    .replace(/[Ａ-Ｚａ-ｚ０-９！＂＃＄％＆＇（）＊＋，－．／：；＜＝＞？＠［＼］＾＿｀｛｜｝～]/g, s =>
      String.fromCharCode(s.charCodeAt(0) - 0xFEE0)
    )
    .replace(/\s|　/g, '')
    .replace(/-|－|—|ｰ/g, '')
    .replace(/株式会社|有限会社|\(株\)|\（株\)|\(有\)|\（有\）/g, '');
}

function getSheet(name) {
  loadConfig_();
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function createSheetWithHeaders(sheetName, headers, activateFilterFlag = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    console.log(`シート "${sheetName}" を作成します。`);
    sheet = ss.insertSheet(sheetName);
  }

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

function showError(message, title = 'エラー') {
  SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
}
