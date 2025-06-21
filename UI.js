// =================================================================================
// ファイル名: UI.gs
// 役割: メニュー操作など、ユーザーインターフェースに関連する関数を管理します。
// =================================================================================

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

function highlightDuplicates_() {
  loadConfig_();
  console.log('重複チェックを開始します...');
  const sheet = getSheet(CONFIG.OCR_RESULT_SHEET);
  if (!sheet || sheet.getLastRow() < 2) {
    console.log('チェック対象のデータがありません。');
    return;
  }
  
  const range = sheet.getDataRange();
  const values = range.getValues();
  const headers = values[0];
  const data = values.slice(1);

  const dateCol = headers.indexOf('取引日');
  const amountCol = headers.indexOf('金額(税込)');

  if (dateCol === -1 || amountCol === -1) {
    console.error('「取引日」または「金額(税込)」列が見つかりません。');
    return;
  }

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
  console.log(`重複チェック完了。${highlightedCount}件の重複の可能性がある取引をハイライトしました。`);
}

function removeHighlight_() {
  loadConfig_();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.OCR_RESULT_SHEET);
  if (!sheet) {
    showError(`シート「${CONFIG.OCR_RESULT_SHEET}」が見つかりません。`);
    return;
  }

  const activeRange = sheet.getActiveRange();
  if (activeRange.getRow() <= 1) {
    showError('データ行を選択してください。');
    return;
  }
  
  const startRow = activeRange.getRow();
  const numRows = activeRange.getNumRows();
  const lastColumn = sheet.getLastColumn();
  const fullRowRange = sheet.getRange(startRow, 1, numRows, lastColumn);

  fullRowRange.setBackground(null);
  SpreadsheetApp.getActiveSpreadsheet().toast(`${numRows}行のハイライトを解除しました。`);
}

function highlightCriticalErrors_() {
  loadConfig_();
  console.log('重大なエラーのチェックを開始します...');
  const sheet = getSheet(CONFIG.OCR_RESULT_SHEET);
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
  console.log('重大なエラーのチェックが完了しました。');
}
