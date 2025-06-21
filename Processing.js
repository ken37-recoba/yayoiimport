// =================================================================================
// ファイル名: Processing.gs
// 役割: ファイルの検出やOCR処理など、中核となるバックグラウンド処理を管理します。
// =================================================================================

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
    throw e; 
  }
}

function performOcrOnPendingFiles(startTime) {
  loadConfig_();
  console.log('ステップ2: OCR処理を開始...');
  const fileListSheet = getSheet(CONFIG.FILE_LIST_SHEET);
  const archiveFolder = DriveApp.getFolderById(CONFIG.ARCHIVE_FOLDER_ID);
  const data = fileListSheet.getDataRange().getValues();
  
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
            
            const firstTransaction = ocrData[0];
            const newFileName = generateNewFileName_(firstTransaction, fileName);
            file.setName(newFileName);
            console.log(`ファイル名を変更しました: ${newFileName}`);

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
