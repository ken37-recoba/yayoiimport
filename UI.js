// =================================================================================
// ファイル名: UI.gs
// 役割: メニュー操作など、ユーザーインターフェースに関連する関数を管理します。
// =================================================================================

function exportForYayoi() {
    loadConfig_();
    const ui = SpreadsheetApp.getUi();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.OCR_RESULT_SHEET);
    if (!sheet) { showError(`シート「${CONFIG.OCR_RESULT_SHEET}」が見つかりません。`); return; }

    const range = sheet.getActiveRange();
    if (range.getRow() <= 1) {
        showError('ヘッダー行はエクスポートできません。データ行を選択してください。');
        return;
    }

    const response = ui.alert('弥生会計用CSVのエクスポート', `選択中の ${range.getNumRows()} 件の取引をCSVファイルとして出力し、「${CONFIG.EXPORTED_SHEET}」シートへ移動しますか？`, ui.ButtonSet.OK_CANCEL);
    if (response !== ui.Button.OK) return;

    try {
        const fullWidthRange = sheet.getRange(range.getRow(), 1, range.getNumRows(), sheet.getLastColumn());
        const selectedData = fullWidthRange.getValues();
        const formulas = fullWidthRange.getFormulas();

        const headers = CONFIG.HEADERS.OCR_RESULT;
        const COL = headers.reduce((acc, header, i) => ({...acc, [header]: i}), {});

        const csvData = selectedData.map(row => {
            const csvRow = new Array(CONFIG.YAYOI.CSV_COLUMNS.length).fill('');
            csvRow[0]  = CONFIG.YAYOI.SHIKIBETSU_FLAG;
            csvRow[3]  = Utilities.formatDate(new Date(row[COL['取引日']]), 'JST', 'yyyy/MM/dd');
            csvRow[4]  = row[COL['勘定科目']];
            csvRow[5]  = row[COL['補助科目']];
            csvRow[7]  = row[COL['消費税課税区分コード']];
            csvRow[8]  = row[COL['金額(税込)']];
            csvRow[9]  = row[COL['うち消費税']];
            csvRow[10] = CONFIG.YAYOI.KASHIKATA_KAMOKU;
            csvRow[13] = CONFIG.YAYOI.KASHIKATA_ZEIKUBUN;
            csvRow[14] = row[COL['金額(税込)']];
            csvRow[15] = CONFIG.YAYOI.KASHIKATA_ZEIGAKU;
            csvRow[16] = `${row[COL['店名']]} / ${row[COL['摘要']]}`;
            csvRow[19] = CONFIG.YAYOI.TORIHIKI_TYPE;
            csvRow[24] = CONFIG.YAYOI.CHOUSEI;
            return csvRow;
        });

        const exportFolder = DriveApp.getFolderById(CONFIG.YAYOI_EXPORT_FOLDER_ID);
        const fileName = `receipt_import_${Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss')}.csv`;
        const csvString = csvData.map(row => row.join(',')).join('\n');
        const blob = Utilities.newBlob('', MimeType.CSV, fileName).setDataFromString(csvString, 'Shift_JIS');
        exportFolder.createFile(blob);

        const exportedSheet = getSheet(CONFIG.EXPORTED_SHEET);
        const exportDate = new Date();
        const rowsToMove = selectedData.map((row, index) => {
            const newRow = [...row, exportDate];
            newRow[COL['ファイルへのリンク']] = formulas[index][COL['ファイルへのリンク']] || row[COL['ファイルへのリンク']];
            return newRow;
        });

        if (rowsToMove.length > 0) {
          exportedSheet.getRange(exportedSheet.getLastRow() + 1, 1, rowsToMove.length, rowsToMove[0].length).setValues(rowsToMove);
        }
        sheet.deleteRows(range.getRow(), range.getNumRows());
        ui.alert('エクスポート完了', `「${fileName}」をGoogle Driveに出力し、${range.getNumRows()}件の取引を「${CONFIG.EXPORTED_SHEET}」シートに移動しました。`, ui.ButtonSet.OK);
    } catch(e) {
        logError_('exportForYayoi', e);
        showError('CSVファイルの作成中にエラーが発生しました。\n\n詳細: ' + e.message);
    }
}

function exportPassbookForYayoi() {
    loadConfig_();
    const ui = SpreadsheetApp.getUi();
    const sheet = SpreadsheetApp.getActiveSheet();
    if (sheet.getName() !== CONFIG.PASSBOOK_RESULT_SHEET) {
        showError(`この機能は「${CONFIG.PASSBOOK_RESULT_SHEET}」シートでのみ使用できます。`);
        return;
    }
    
    const range = sheet.getActiveRange();
    if (range.getRow() <= 1) {
        showError('ヘッダー行はエクスポートできません。データ行を選択してください。');
        return;
    }

    const response = ui.alert('弥生会計用CSVのエクスポート (通帳)', `選択中の ${range.getNumRows()} 件の取引をCSVファイルとして出力し、「${CONFIG.PASSBOOK_EXPORTED_SHEET}」シートへ移動しますか？`, ui.ButtonSet.OK_CANCEL);
    if (response !== ui.Button.OK) return;

    try {
        const fullWidthRange = sheet.getRange(range.getRow(), 1, range.getNumRows(), sheet.getLastColumn());
        const selectedData = fullWidthRange.getValues();
        const formulas = fullWidthRange.getFormulas();

        const headers = CONFIG.HEADERS.PASSBOOK_RESULT;
        const COL = headers.reduce((acc, header, i) => ({...acc, [header]: i}), {});

        const csvData = selectedData.map((row, i) => {
            const csvRow = new Array(CONFIG.YAYOI.CSV_COLUMNS.length).fill('');
            const isDeposit = Number(row[COL['入金額']]) > 0;
            const passbookAccountName = row[COL['通帳勘定科目']];
            if (!passbookAccountName || passbookAccountName === '（未設定）') {
                throw new Error(`行 ${range.getRow() + i} の「通帳勘定科目」が不明です。「通帳マスター」の設定を確認し、ファイル名にキーワードが含まれているか確認してください。`);
            }

            csvRow[0]  = CONFIG.YAYOI.SHIKIBETSU_FLAG;
            csvRow[3]  = Utilities.formatDate(new Date(row[COL['取引日']]), 'JST', 'yyyy/MM/dd');
            
            if (isDeposit) {
                const taxAmount = calculateTaxAmount_(row[COL['入金額']], row[COL['貸方税区分']]);
                csvRow[4]  = passbookAccountName;
                csvRow[7]  = '対象外';
                csvRow[8]  = row[COL['入金額']];
                csvRow[9]  = 0;
                csvRow[10] = row[COL['相手方勘定科目']];
                csvRow[11] = row[COL['相手方補助科目']];
                csvRow[13] = row[COL['貸方税区分']];
                csvRow[14] = row[COL['入金額']];
                csvRow[15] = taxAmount;
            } else {
                const taxAmount = calculateTaxAmount_(row[COL['出金額']], row[COL['借方税区分']]);
                csvRow[4]  = row[COL['相手方勘定科目']];
                csvRow[5]  = row[COL['相手方補助科目']];
                csvRow[7]  = row[COL['借方税区分']];
                csvRow[8]  = row[COL['出金額']];
                csvRow[9]  = taxAmount;
                csvRow[10] = passbookAccountName;
                csvRow[13] = '対象外';
                csvRow[14] = row[COL['出金額']];
                csvRow[15] = 0;
            }
            
            csvRow[16] = row[COL['摘要']];
            csvRow[19] = CONFIG.YAYOI.TORIHIKI_TYPE;
            csvRow[24] = CONFIG.YAYOI.CHOUSEI;
            return csvRow;
        });

        const exportFolder = DriveApp.getFolderById(CONFIG.YAYOI_EXPORT_FOLDER_ID);
        const fileName = `passbook_import_${Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss')}.csv`;
        const csvString = csvData.map(row => row.join(',')).join('\n');
        const blob = Utilities.newBlob('', MimeType.CSV, fileName).setDataFromString(csvString, 'Shift_JIS');
        exportFolder.createFile(blob);

        const exportedSheet = getSheet(CONFIG.PASSBOOK_EXPORTED_SHEET);
        const exportDate = new Date();
        const rowsToMove = selectedData.map((row, index) => {
            const newRow = [...row, exportDate];
            newRow[COL['ファイルへのリンク']] = formulas[index][COL['ファイルへのリンク']] || row[COL['ファイルへのリンク']];
            return newRow;
        });
        if (rowsToMove.length > 0) {
          exportedSheet.getRange(exportedSheet.getLastRow() + 1, 1, rowsToMove.length, rowsToMove[0].length).setValues(rowsToMove);
        }

        sheet.deleteRows(range.getRow(), range.getNumRows());
        ui.alert('エクスポート完了 (通帳)', `「${fileName}」を出力し、${range.getNumRows()}件の取引を「${CONFIG.PASSBOOK_EXPORTED_SHEET}」シートに移動しました。`, ui.ButtonSet.OK);
    } catch(e) {
        logError_('exportPassbookForYayoi', e);
        showError('CSVファイルの作成中にエラーが発生しました。\n\n詳細: ' + e.message);
    }
}

function moveTransactionsBackToOcr() {
  // 実装は後続のタスク
}

function movePassbookTransactionsBackToOcr() {
  // 実装は後続のタスク
}

function showReceiptPreview() {
  // 実装は後続のタスク
}

function showPassbookPreview() {
  // 実装は後続のタスク
}

function deleteSelectedTransactions() {
  // 実装は後続のタスク
}

function insertDummyInvoiceNumber() {
  // 実装は後続のタスク
}

function activateFilter() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (sheet) {
    if (sheet.getFilter()) {
      sheet.getFilter().remove();
    }
    sheet.getDataRange().createFilter();
    SpreadsheetApp.getUi().alert(`シート「${sheet.getName()}」にフィルタをオンにしました。`);
  }
}

function highlightDuplicates_() {
  loadConfig_();
  const sheet = getSheet(CONFIG.OCR_RESULT_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return;
  
  const range = sheet.getDataRange();
  const values = range.getValues();
  const headers = values[0];
  const data = values.slice(1);
  const dateCol = headers.indexOf('取引日');
  const amountCol = headers.indexOf('金額(税込)');
  if (dateCol === -1 || amountCol === -1) return;

  const backgroundColors = range.getBackgrounds();
  for (let i = 1; i < backgroundColors.length; i++) {
    for (let j = 0; j < backgroundColors[i].length; j++) {
      if (backgroundColors[i][j] === DUPLICATE_HIGHLIGHT_COLOR) {
        backgroundColors[i][j] = null;
      }
    }
  }

  const counts = {};
  const transactionMap = {};

  data.forEach((row, index) => {
    if (row.every(cell => cell === '')) return;
    const date = new Date(row[dateCol]).toLocaleDateString();
    const amount = row[amountCol];
    const key = `${date}_${amount}`;
    counts[key] = (counts[key] || 0) + 1;
    if (!transactionMap[key]) transactionMap[key] = [];
    transactionMap[key].push(index + 2);
  });

  let highlightedCount = 0;
  for (const key in counts) {
    if (counts[key] > 1) {
      const rowsToHighlight = transactionMap[key];
      rowsToHighlight.forEach(rowNum => {
        if (backgroundColors[rowNum - 1][0] !== CRITICAL_ERROR_HIGHLIGHT_COLOR) {
            for (let j = 0; j < backgroundColors[rowNum - 1].length; j++) {
                backgroundColors[rowNum - 1][j] = DUPLICATE_HIGHLIGHT_COLOR;
            }
            highlightedCount++;
        }
      });
    }
  }
  range.setBackgrounds(backgroundColors);
}

function removeHighlight_() {
  loadConfig_();
  const sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== CONFIG.OCR_RESULT_SHEET && sheet.getName() !== CONFIG.PASSBOOK_RESULT_SHEET) {
      showError('この機能は「OCR結果」または「通帳OCR結果」シートで実行してください。');
      return;
  }
  const activeRange = sheet.getActiveRange();
  if (activeRange.getRow() <= 1) { showError('データ行を選択してください。'); return; }
  
  sheet.getRange(activeRange.getRow(), 1, activeRange.getNumRows(), sheet.getLastColumn()).setBackground(null);
  SpreadsheetApp.getActiveSpreadsheet().toast(`${activeRange.getNumRows()}行のハイライトを解除しました。`);
}

function highlightCriticalErrors_() {
  loadConfig_();
  const sheet = getSheet(CONFIG.OCR_RESULT_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return;

  const range = sheet.getDataRange();
  const values = range.getValues();
  const noteCol = values[0].indexOf('備考');
  if (noteCol === -1) return;
  
  const backgroundColors = range.getBackgrounds();
  values.slice(1).forEach((row, index) => {
    if (row[noteCol] && row[noteCol].includes("【要確認")) {
      for (let j = 0; j < backgroundColors[index + 1].length; j++) {
        backgroundColors[index + 1][j] = CRITICAL_ERROR_HIGHLIGHT_COLOR;
      }
    }
  });
  range.setBackgrounds(backgroundColors);
}

function highlightPassbookCriticalErrors_() {
  loadConfig_();
  console.log('通帳の重大なエラーのチェックを開始します...');
  const sheet = getSheet(CONFIG.PASSBOOK_RESULT_SHEET);
  if (!sheet || sheet.getLastRow() < 2) {
    console.log('チェック対象のデータがありません。');
    return;
  }

  const range = sheet.getDataRange();
  const values = range.getValues();
  const headers = values[0];
  const data = values.slice(1);

  const noteCol = headers.indexOf('備考');
  if (noteCol === -1) {
    console.error('「備考」列が見つかりません。');
    return;
  }
  
  const backgroundColors = range.getBackgrounds();
  
  data.forEach((row, index) => {
    const note = row[noteCol];
    if (note && note.includes("【要確認")) {
      for (let j = 0; j < backgroundColors[index + 1].length; j++) {
        backgroundColors[index + 1][j] = CRITICAL_ERROR_HIGHLIGHT_COLOR;
      }
    }
  });

  range.setBackgrounds(backgroundColors);
  console.log('通帳の重大なエラーのチェックが完了しました。');
}
