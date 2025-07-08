// =================================================================================
// ファイル名: Processing.gs
// 役割: ファイルの検出やOCR処理など、中核となるバックグラウンド処理を管理します。
// =================================================================================

function processNewFiles() {
  loadConfig_();
  console.log('ステップ1: 新規の領収書ファイルを処理...');
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
    console.log('ステップ1: 新規の領収書ファイルの処理が完了。');
  } catch(e) {
    logError_('processNewFiles', e);
    throw e; 
  }
}

function performOcrOnPendingFiles(startTime) {
  loadConfig_();
  console.log('ステップ2: 領収書のOCR処理を開始...');
  const fileListSheet = getSheet(CONFIG.FILE_LIST_SHEET);
  const archiveFolder = DriveApp.getFolderById(CONFIG.ARCHIVE_FOLDER_ID);
  const data = fileListSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const elapsedTime = (new Date().getTime() - startTime.getTime()) / 1000;
    if (elapsedTime > CONFIG.EXECUTION_TIME_LIMIT_SECONDS) {
      console.log(`実行時間が上限(${CONFIG.EXECUTION_TIME_LIMIT_SECONDS}秒)に近づいたため、領収書処理を中断します。`);
      break;
    }

    const rowData = data[i];
    if (rowData[2] === STATUS.PENDING || rowData[2] === STATUS.PROCESSING) {
      const fileId = rowData[0];
      const fileName = rowData[1];
      const rowNum = i + 1;
      const contextInfo = `Receipt File: ${fileName} (ID: ${fileId})`;

      try {
        fileListSheet.getRange(rowNum, 3).setValue(STATUS.PROCESSING);
        SpreadsheetApp.flush();
        console.log(`領収書OCR処理を開始: ${fileName}`);

        const file = DriveApp.getFileById(fileId);
        const result = callGeminiApi(file.getBlob(), getGeminiPrompt(fileName));
        Utilities.sleep(1500);

        if (result.success) {
          const ocrData = JSON.parse(result.data);
          if (ocrData && ocrData.length > 0) {
            logOcrResult(ocrData, fileId);
            logTokenUsage(fileName, result.usage);
            
            const firstTransaction = ocrData[0];
            const newFileName = generateNewFileName_(firstTransaction, fileName);
            file.setName(newFileName);
            console.log(`ファイル名を変更しました: ${newFileName}`);

            fileListSheet.getRange(rowNum, 3, 1, 2).setValues([[STATUS.PROCESSED, '']]);
            file.moveTo(archiveFolder);
            console.log(`領収書OCR処理成功: ${fileName}`);
          } else {
            const msg = `ファイル ${fileName} から領収書データは検出されませんでした。`;
            console.log(msg);
            fileListSheet.getRange(rowNum, 3, 1, 2).setValues([[STATUS.ERROR, msg]]);
          }
        } else {
          throw new Error(result.error);
        }
      } catch (e) {
        const errorMessage = e.message || e.toString();
        logError_('performOcrOnPendingFiles', e, contextInfo);
        console.error(`領収書OCR処理中にエラー: ${fileName}, Error: ${errorMessage}`);
        fileListSheet.getRange(rowNum, 3, 1, 2).setValues([[STATUS.ERROR, errorMessage]]);
      } finally {
        fileListSheet.getRange(rowNum, 5).setValue(new Date());
      }
    }
  }
  console.log('ステップ2: 領収書のOCR処理が完了。');
}

function processNewPassbookFiles() {
  loadConfig_();
  console.log('ステップ1: 新規の通帳ファイルを処理...');
  try {
    const sourceFolder = DriveApp.getFolderById(CONFIG.PASSBOOK_SOURCE_FOLDER_ID);
    const fileListSheet = getSheet(CONFIG.PASSBOOK_FILE_LIST_SHEET);
    
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
      if (mimeType.startsWith('image/')) {
        const fileName = file.getName().toUpperCase();
        let bankType = 'STANDARD';
        if (fileName.includes('UFJ')) bankType = 'MUFG';
        if (fileName.includes('OSAKA_SHINKIN')) bankType = 'OSAKA_SHINKIN';
        
        fileListSheet.appendRow([fileId, file.getName(), bankType, STATUS.PENDING, '', '', new Date()]);
        existingFileIds.push(fileId);
      }
    }
    console.log('ステップ1: 新規の通帳ファイルの処理が完了。');
  } catch(e) {
    logError_('processNewPassbookFiles', e);
    throw e; 
  }
}

function performOcrOnPassbookFiles(startTime) {
  loadConfig_();
  console.log('ステップ2: 通帳のOCR処理を開始...');
  const fileListSheet = getSheet(CONFIG.PASSBOOK_FILE_LIST_SHEET);
  const archiveFolder = DriveApp.getFolderById(CONFIG.PASSBOOK_ARCHIVE_FOLDER_ID);
  const data = fileListSheet.getDataRange().getValues();
  const headers = data[0];
  const COL = headers.reduce((acc, h, i) => ({...acc, [h]: i}), {});
  
  for (let i = 1; i < data.length; i++) {
    const elapsedTime = (new Date().getTime() - startTime.getTime()) / 1000;
    if (elapsedTime > CONFIG.EXECUTION_TIME_LIMIT_SECONDS) {
      console.log(`実行時間が上限(${CONFIG.EXECUTION_TIME_LIMIT_SECONDS}秒)に近づいたため、通帳処理を中断します。`);
      break;
    }

    const rowData = data[i];
    if (rowData[COL['ステータス']] === STATUS.PENDING || rowData[COL['ステータス']] === STATUS.PROCESSING) {
      const fileId = rowData[COL['ファイルID']];
      const fileName = rowData[COL['ファイル名']];
      const bankType = rowData[COL['銀行タイプ']];
      const rowNum = i + 1;
      const contextInfo = `Passbook File: ${fileName} (ID: ${fileId})`;

      try {
        fileListSheet.getRange(rowNum, COL['ステータス'] + 1).setValue(STATUS.PROCESSING);
        SpreadsheetApp.flush();
        console.log(`通帳OCR処理を開始: ${fileName}`);

        const file = DriveApp.getFileById(fileId);
        const result = callPassbookGeminiApi(file.getBlob(), bankType);
        Utilities.sleep(1500);

        if (result.success) {
          const transactions = JSON.parse(result.data);
          
          if (transactions && Array.isArray(transactions) && transactions.length > 0) {
            const rowCount = transactions.length;
            const passbookAccountName = logPassbookResult(transactions, fileId, fileName);
            logTokenUsage(fileName, result.usage);
            
            const newFileName = generateNewPassbookFileName_(passbookAccountName, fileName);
            file.setName(newFileName);
            console.log(`ファイル名を変更しました: ${newFileName}`);

            fileListSheet.getRange(rowNum, COL['ステータス'] + 1, 1, 2).setValues([[STATUS.PROCESSED, rowCount]]);
            file.moveTo(archiveFolder);
            console.log(`通帳OCR処理成功: ${fileName}`);
          } else {
            const msg = `ファイル ${fileName} から取引データは検出されませんでした。`;
            console.log(msg);
            fileListSheet.getRange(rowNum, COL['ステータス'] + 1, 1, 2).setValues([[STATUS.ERROR, 0]]);
          }
        } else {
          throw new Error(result.error);
        }
      } catch (e) {
        const errorMessage = e.message || e.toString();
        logError_('performOcrOnPassbookFiles', e, contextInfo);
        console.error(`通帳OCR処理中にエラー: ${fileName}, Error: ${errorMessage}`);
        fileListSheet.getRange(rowNum, COL['ステータス'] + 1, 1, 2).setValues([[STATUS.ERROR, 0]]);
      } finally {
        fileListSheet.getRange(rowNum, COL['登録日時'] + 1).setValue(new Date());
      }
    }
  }
  console.log('ステップ2: 通帳のOCR処理が完了。');
}
