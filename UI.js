// =================================================================================
// ファイル名: UI.gs
// 役割: メニュー操作など、ユーザーインターフェースに関連する関数を管理します。
// =================================================================================

/**
 * 領収書データを弥生会計形式でエクスポートする
 */
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

    const response = ui.alert('弥生会計用CSVのエクスポート', `選択中の ${range.getNumRows()} 件の取引をCSVファイルとして出力し、「${CONFIG.EXPORTED_SHEET}」シートへ移動しますか？`, ui.ButtonSet.OK_CANCEL);
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
          exportedSheet.getRange(exportedSheet.getLastRow() + 1, 1, rowsToMove.length, rowsToMove[0].length).setValues(rowsToMove);
        }
        sheet.deleteRows(range.getRow(), range.getNumRows());
        ui.alert('エクスポート完了', `「${fileName}」をGoogle Driveに出力し、${range.getNumRows()}件の取引を「${CONFIG.EXPORTED_SHEET}」シートに移動しました。`, ui.ButtonSet.OK);
    } catch(e) {
        logError_('exportForYayoi', e);
        showError('CSVファイルの作成中にエラーが発生しました。\n\n詳細: ' + e.message);
    }
}

/**
 * 通帳データを弥生会計形式でエクスポートする
 */
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
                csvRow[4]  = passbookAccountName;
                csvRow[7]  = row[COL['借方税区分']];
                csvRow[8]  = row[COL['入金額']];
                csvRow[10] = row[COL['相手方勘定科目']];
                csvRow[11] = row[COL['相手方補助科目']];
                csvRow[13] = row[COL['貸方税区分']];
                csvRow[14] = row[COL['入金額']];
            } else {
                csvRow[4]  = row[COL['相手方勘定科目']];
                csvRow[5]  = row[COL['相手方補助科目']];
                csvRow[7]  = row[COL['借方税区分']];
                csvRow[8]  = row[COL['出金額']];
                csvRow[10] = passbookAccountName;
                csvRow[13] = row[COL['貸方税区分']];
                csvRow[14] = row[COL['出金額']];
            }
            
            csvRow[16] = row[COL['摘要']];
            csvRow[9] = 0; csvRow[15] = 0; csvRow[19] = CONFIG.YAYOI.TORIHIKI_TYPE; csvRow[24] = CONFIG.YAYOI.CHOUSEI;
            return csvRow;
        });

        const exportFolder = DriveApp.getFolderById(CONFIG.EXPORT_FOLDER_ID);
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


function moveTransactionsBackToOcr() { /* ... 実装は後続のタスク ... */ }
function movePassbookTransactionsBackToOcr() { /* ... 実装は後続のタスク ... */ }
function showReceiptPreview() { /* ... 実装は後続のタスク ... */ }
function showPassbookPreview() { /* ... 実装は後続のタスク ... */ }
function deleteSelectedTransactions() { /* ... 実装は後続のタスク ... */ }
function insertDummyInvoiceNumber() { /* ... 実装は後続のタスク ... */ }
function activateFilter() { /* ... 実装は後続のタスク ... */ }
function highlightDuplicates_() { /* ... 実装は後続のタスク ... */ }
function removeHighlight_() { /* ... 実装は後続のタスク ... */ }
function highlightCriticalErrors_() { /* ... 実装は後続のタスク ... */ }
