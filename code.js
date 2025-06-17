/****************************************************************
 * 領収書OCRシステム (弥生会計連携/全機能搭載)
 * * 【必須設定】
 * 1. 以下のCONFIGオブジェクトにIDとAPIキー、CSV出力先フォルダIDを設定してください。
 * 2. このバージョンでは、拡張サービスを有効にする必要は一切ありません。
 * 3. 「勘定科目マスター」という名前のシートを作成し、A列に勘定科目、B列に摘要キーワードを記載してください。
 ****************************************************************/
const CONFIG = {
  //【要設定】スプレッドシートのID (URLから取得)
  SPREADSHEET_ID: '1BpqUIgIV-PkeimeJa05x4yFJqKe4UpHhS9-5cubZ1cw', // ご自身のIDに書き換えてください

  //【要設定】最初に領収書をアップロードするフォルダのID
  SOURCE_FOLDER_ID: '1x6k_iC7ws8YyMW31DgQKtObWbDddKair', // ご自身のIDに書き換えてください
  
  // ★追加：【要設定】弥生会計用CSVを出力するフォルダのID
  EXPORT_FOLDER_ID: '1gPUmeOungbwWPB4KPsQCxKSK-3xgKnI8',

  //【要設定】Gemini APIキー
  GEMINI_API_KEY: 'AIzaSyA52gv5dZCx06uvVbLyXlJwaA8WWQodzMM', // ご自身のAPIキーに書き換えてください

  // 実行時間対策：スクリプトの実行を安全に停止するまでの秒数 (5分 = 300秒)
  EXECUTION_TIME_LIMIT_SECONDS: 300,

  // 以下は自動生成されるフォルダ・シートの名前（変更可）
  ARCHIVE_FOLDER_NAME: '[OCR] アーカイブ済み',
  FILE_LIST_SHEET: 'ファイルリスト',
  OCR_RESULT_SHEET: 'OCR結果',
  TOKEN_LOG_SHEET: 'トークンログ',
  MASTER_SHEET: '勘定科目マスター',
  LEARNING_SHEET: '学習データ',

  // Gemini API設定 (読取り精度に関わるため変更しないでください)
  GEMINI_MODEL: 'gemini-2.5-flash-preview-05-20',
  THINKING_BUDGET: 10000,
};

// 状態管理用の定数
const STATUS = {
  PENDING: '未処理',
  PROCESSING: '処理中',
  PROCESSED: '処理済み',
  ERROR: 'エラー',
};


/****************************************************************
 * メイン処理 & トリガー用関数
 ****************************************************************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('領収書OCR')
    .addItem('手動で新規ファイルを処理', 'mainProcess')
    .addSeparator()
    .addItem('選択行の領収書をプレビュー', 'showReceiptPreview')
    .addSeparator()
    .addItem('弥生会計形式でエクスポート', 'exportForYayoi')
    .addSeparator() 
    .addItem('フィルタをオンにする', 'activateFilter')
    .addItem('選択した行にダミー番号を挿入', 'insertDummyInvoiceNumber')
    .addItem('選択した取引を削除', 'deleteSelectedTransactions')
    .addToUi();
}

/**
 * チェックボックスのON/OFFや登録番号の削除をトリガーに関数を実行する
 */
function onEdit(e) {
  try {
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
    const transactionIdColIndex = headers.indexOf('取引ID') + 1;

    // --- Case 1: 「学習チェック」列が編集された場合 ---
    if (col === learnCheckColIndex) {
      const transactionId = sheet.getRange(row, transactionIdColIndex).getValue();
      if (!transactionId) return; 
      
      const isChecked = range.isChecked();
      
      if (isChecked) { // 学習登録
        if (range.getNote().includes('学習済み')) return;
        
        const learningSheet = getSheet(CONFIG.LEARNING_SHEET);
        const dataRow = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
        
        const storeName = dataRow[headers.indexOf('店名')];
        const description = dataRow[headers.indexOf('摘要')];
        const kanjo = dataRow[headers.indexOf('勘定科目')];
        const hojo = dataRow[headers.indexOf('補助科目')];

        learningSheet.appendRow([new Date(), storeName, description, kanjo, hojo, transactionId]);
        
        range.setNote(`学習済み (ID: ${transactionId})`);
        sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground('#e6f4ea');
        SpreadsheetApp.getActiveSpreadsheet().toast(`「${storeName}」の勘定科目を学習しました。`);
      
      } else { // 学習取消
        const learningSheet = getSheet(CONFIG.LEARNING_SHEET);
        if (!learningSheet || learningSheet.getLastRow() < 2) return;

        const learningData = learningSheet.getRange(2, 1, learningSheet.getLastRow() - 1, 6).getValues();
        const learningIdCol = 5; 
        
        for (let i = learningData.length - 1; i >= 0; i--) {
          if (learningData[i][learningIdCol] === transactionId) {
            learningSheet.deleteRow(i + 2);
            range.clearNote();
            sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground(null);
            SpreadsheetApp.getActiveSpreadsheet().toast(`取引ID ${transactionId} の学習データを取り消しました。`);
            return;
          }
        }
      }
    }

    // --- Case 2: 「登録番号」列が削除（空に）された場合 ---
    if (col === taxCodeColIndex && range.isBlank()) {
      range.setFontColor(null); 
      
      const taxRateColIndex = headers.indexOf('税率(%)') + 1;
      const taxCategoryColIndex = headers.indexOf('消費税課税区分コード') + 1;

      if (taxRateColIndex > 0 && taxCategoryColIndex > 0) {
        const taxRate = sheet.getRange(row, taxRateColIndex).getValue();
        const newTaxCategory = getTaxCategoryCode(taxRate, ""); 
        sheet.getRange(row, taxCategoryColIndex).setValue(newTaxCategory);
        SpreadsheetApp.getActiveSpreadsheet().toast(`行 ${row} の登録番号が削除されたため、税区分を更新しました。`);
      }
    }
  } catch (err) {
    console.error("onEdit Error: " + err.toString());
    SpreadsheetApp.getActiveSpreadsheet().toast("エラーが発生しました: " + err.message);
  }
}

function mainProcess() {
  const startTime = new Date();
  console.log('メインプロセスを開始します。');
  initializeEnvironment();
  processNewFiles();
  performOcrOnPendingFiles(startTime);
  console.log('メインプロセスが完了しました。');
}


/**
 * 選択された行をOCR結果シートと学習データシートから安全に削除する関数
 */
function deleteSelectedTransactions() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.OCR_RESULT_SHEET);
  
  if (SpreadsheetApp.getActiveSheet().getName() !== CONFIG.OCR_RESULT_SHEET) {
    ui.alert('この機能は「OCR結果」シートでのみ使用できます。');
    return;
  }
  
  const range = sheet.getActiveRange();
  const startRow = range.getRow();
  
  if (startRow <= 1) {
      ui.alert('ヘッダー行は削除できません。データ行を選択してください。');
      return;
  }
  
  const response = ui.alert(
    '選択した取引の削除',
    `選択中の ${range.getNumRows()} 件の取引を削除しますか？\n\n学習済みの取引が含まれている場合、関連する学習データも完全に削除されます。この操作は元に戻せません。`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response !== ui.Button.OK) {
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const transactionIdColIndex = headers.indexOf('取引ID') + 1;

  const transactionIdsToDelete = [];
  const fullRange = sheet.getRange(startRow, 1, range.getNumRows(), sheet.getLastColumn());
  const selectedRows = fullRange.getValues();

  for(let i = 0; i < selectedRows.length; i++){
      const id = selectedRows[i][transactionIdColIndex -1];
      if (id) {
          transactionIdsToDelete.push(id);
      }
  }
  
  const learningSheet = getSheet(CONFIG.LEARNING_SHEET);
  let learnedDeletedCount = 0;
  if (learningSheet && learningSheet.getLastRow() > 1) {
    const learningData = learningSheet.getRange(2, 1, learningSheet.getLastRow() - 1, 6).getValues();
    const learningIdCol = 5;

    for (let i = learningData.length - 1; i >= 0; i--) {
        if(transactionIdsToDelete.includes(learningData[i][learningIdCol])){
            learningSheet.deleteRow(i + 2);
            learnedDeletedCount++;
        }
    }
  }

  sheet.deleteRows(startRow, range.getNumRows());
  
  ui.alert('処理完了', `${range.getNumRows()}件の取引を削除しました。\n(うち、${learnedDeletedCount}件の学習データも関連して削除されました。)`, ui.ButtonSet.OK);
}

/**
 * 選択された行にダミーのインボイス番号を挿入し、税区分コードを更新する
 */
function insertDummyInvoiceNumber() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.OCR_RESULT_SHEET);

  if (SpreadsheetApp.getActiveSheet().getName() !== CONFIG.OCR_RESULT_SHEET) {
    ui.alert('この機能は「OCR結果」シートでのみ使用できます。');
    return;
  }
  
  const range = sheet.getActiveRange();
  if (range.getRow() <= 1) {
    ui.alert('ヘッダー行には適用できません。データ行を選択してください。');
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const taxCodeColIndex = headers.indexOf('登録番号') + 1;
  const taxRateColIndex = headers.indexOf('税率(%)') + 1;
  const taxCategoryColIndex = headers.indexOf('消費税課税区分コード') + 1;

  if (taxCodeColIndex === 0 || taxRateColIndex === 0 || taxCategoryColIndex === 0) {
    ui.alert('必要な列（登録番号、税率(%)、消費税課税区分コード）が見つかりません。');
    return;
  }

  let updatedCount = 0;
  const startRow = range.getRow();

  for (let i = 0; i < range.getNumRows(); i++) {
    const currentRow = startRow + i;
    const taxCodeCell = sheet.getRange(currentRow, taxCodeColIndex);
    
    if (taxCodeCell.isBlank()) {
      const dummyNumber = 'T' + Math.random().toString().slice(2,15);
      taxCodeCell.setValue(dummyNumber)
                 .setFontColor("#0000FF"); 

      const taxRate = sheet.getRange(currentRow, taxRateColIndex).getValue();
      const newTaxCategory = getTaxCategoryCode(taxRate, dummyNumber);
      
      sheet.getRange(currentRow, taxCategoryColIndex).setValue(newTaxCategory);
      updatedCount++;
    }
  }

  if (updatedCount > 0) {
    ui.alert('処理完了', `${updatedCount}件の取引にダミーの登録番号を挿入し、税区分を更新しました。`, ui.ButtonSet.OK);
  } else {
    ui.alert('処理対象なし', '選択された行に、登録番号が空欄の取引はありませんでした。', ui.ButtonSet.OK);
  }
}

/**
 * 選択された行を弥生会計インポート用のCSVとしてエクスポートする
 */
function exportForYayoi() {
    const ui = SpreadsheetApp.getUi();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.OCR_RESULT_SHEET);

    if (SpreadsheetApp.getActiveSheet().getName() !== CONFIG.OCR_RESULT_SHEET) {
        ui.alert('この機能は「OCR結果」シートでのみ使用できます。');
        return;
    }

    const range = sheet.getActiveRange();
    if (range.getRow() <= 1) {
        ui.alert('ヘッダー行はエクスポートできません。データ行を選択してください。');
        return;
    }

    const response = ui.alert(
        '弥生会計用CSVのエクスポート',
        `選択中の ${range.getNumRows()} 件の取引をCSVファイルとして出力しますか？\n\n出力された行は「エクスポート済み」としてマーキングされます。`,
        ui.ButtonSet.OK_CANCEL
    );

    if (response !== ui.Button.OK) {
        return;
    }
    
    const startRow = range.getRow();
    const numRows = range.getNumRows();
    const lastCol = sheet.getLastColumn();
    const fullWidthRange = sheet.getRange(startRow, 1, numRows, lastCol);
    const selectedData = fullWidthRange.getDisplayValues(); 

    const COL = {
        TRANSACTION_DATE: 2,
        STORE_NAME: 3,
        DESCRIPTION: 4,
        ACCOUNT_TITLE: 5,
        SUB_ACCOUNT: 6,
        AMOUNT_INCL_TAX: 8,
        TAX_AMOUNT: 9,
        TAX_CATEGORY: 11,
    };

    const csvData = [];
    
    selectedData.forEach(row => {
        const csvRow = new Array(25).fill(''); 

        csvRow[0]  = '2000';
        csvRow[3]  = row[COL.TRANSACTION_DATE];
        csvRow[4]  = row[COL.ACCOUNT_TITLE];
        csvRow[5]  = row[COL.SUB_ACCOUNT];
        csvRow[7]  = row[COL.TAX_CATEGORY];
        csvRow[8]  = row[COL.AMOUNT_INCL_TAX];
        csvRow[9]  = row[COL.TAX_AMOUNT];
        csvRow[10] = '役員借入金';
        csvRow[13] = '対象外';
        csvRow[14] = row[COL.AMOUNT_INCL_TAX];
        csvRow[15] = '0';
        csvRow[16] = `${row[COL.STORE_NAME]} / ${row[COL.DESCRIPTION]}`;
        csvRow[19] = '0';
        csvRow[24] = 'no';

        csvData.push(csvRow);
    });

    try {
        const exportFolder = DriveApp.getFolderById(CONFIG.EXPORT_FOLDER_ID);
        const fileName = `import_${Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss')}.csv`;
        const csvString = csvData.map(row => row.join(',')).join('\n');
        
        const blob = Utilities.newBlob('', MimeType.CSV, fileName).setDataFromString(csvString, 'Shift_JIS');
        
        exportFolder.createFile(blob);

        fullWidthRange.setBackground('#f3f3f3'); 
        
        ui.alert('エクスポート完了', `「${fileName}」をGoogle Driveの指定フォルダに出力しました。`, ui.ButtonSet.OK);

    } catch(e) {
        console.error("CSVエクスポート中にエラー: " + e.toString());
        ui.alert('エクスポート失敗', 'CSVファイルの作成中にエラーが発生しました。\n\n・CONFIGの「EXPORT_FOLDER_ID」が正しいか\n・指定フォルダへのアクセス権があるか\n\nを確認してください。', ui.ButtonSet.OK);
    }
}


/****************************************************************
 * ステップ1: 新規ファイルの取り込みとリスト化
 ****************************************************************/
function processNewFiles() {
  console.log('ステップ1: 新規ファイルの処理を開始...');
  const sourceFolder = DriveApp.getFolderById(CONFIG.SOURCE_FOLDER_ID);
  const fileListSheet = getSheet(CONFIG.FILE_LIST_SHEET);
  const existingFileIds = fileListSheet.getRange(2, 1, fileListSheet.getLastRow(), 1).getValues()
    .flat().filter(id => id);

  const files = sourceFolder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    const fileId = file.getId();
    const mimeType = file.getMimeType();

    if (existingFileIds.includes(fileId)) {
      continue;
    }

    if (mimeType === MimeType.PDF || mimeType.startsWith('image/')) {
      try {
        console.log(`新規処理対象ファイルを発見: ${file.getName()}`);
        fileListSheet.appendRow([fileId, file.getName(), STATUS.PENDING, '', new Date()]);
        existingFileIds.push(fileId);
      } catch (e) {
        console.error(`ファイルリストへの追加中にエラー: ${file.getName()}, Error: ${e.toString()}`);
      }
    } else {
      console.log(`サポート外のファイル形式のためスキップ: ${file.getName()} (${mimeType})`);
    }
  }
  console.log('ステップ1: 新規ファイルの処理が完了しました。');
}


/****************************************************************
 * ステップ2: OCR処理の実行
 ****************************************************************/
function performOcrOnPendingFiles(startTime) {
  console.log('ステップ2: OCR処理を開始...');
  const fileListSheet = getSheet(CONFIG.FILE_LIST_SHEET);
  const archiveFolder = getFolderByName(CONFIG.ARCHIVE_FOLDER_NAME);
  const data = fileListSheet.getDataRange().getValues();

  const learningData = getLearningData();

  for (let i = 1; i < data.length; i++) {
    const currentTime = new Date();
    const elapsedTime = (currentTime.getTime() - startTime.getTime()) / 1000;

    if (elapsedTime > CONFIG.EXECUTION_TIME_LIMIT_SECONDS) {
      console.log(`実行時間が上限(${CONFIG.EXECUTION_TIME_LIMIT_SECONDS}秒)に近づいたため、処理を安全に中断します。`);
      break;
    }

    const row = data[i];
    if (row[2] === STATUS.PENDING) {
      const fileId = row[0];
      const fileName = row[1];
      const rowNum = i + 1;

      try {
        fileListSheet.getRange(rowNum, 3).setValue(STATUS.PROCESSING);
        SpreadsheetApp.flush();
        console.log(`OCR処理を開始: ${fileName} (経過時間: ${Math.round(elapsedTime)}秒)`);

        const file = DriveApp.getFileById(fileId);
        const fileBlob = file.getBlob();
        
        const prompt = getGeminiPrompt(fileName);
        const result = callGeminiApi(fileBlob, prompt);
        
        Utilities.sleep(1500);

        if (result.success) {
          const ocrData = JSON.parse(result.data);
          Logger.log(ocrData);
          if (ocrData && ocrData.length > 0) {
            
            logOcrResult(ocrData, fileId, learningData);

            logTokenUsage(fileName, result.usage);
            fileListSheet.getRange(rowNum, 3).setValue(STATUS.PROCESSED);
            fileListSheet.getRange(rowNum, 5).setValue(new Date());
            console.log(`OCR処理成功: ${fileName}`);
            
            file.moveTo(archiveFolder);
            console.log(`ファイル ${fileName} をアーカイブしました。`);

          } else {
            console.log(`ファイル ${fileName} から領収書は検出されませんでした。`);
            fileListSheet.getRange(rowNum, 3).setValue(STATUS.PENDING);
            console.log(`ファイル ${fileName} はアーカイブされませんでした。処理待ちに戻ります。`);
          }

        } else {
          throw new Error(result.error);
        }

      } catch (e) {
        const errorMessage = e.message || e.toString();
        console.error(`OCR処理中にエラー: ${fileName}, Error: ${errorMessage}`);
        fileListSheet.getRange(rowNum, 3).setValue(STATUS.ERROR);
        fileListSheet.getRange(rowNum, 4).setValue(errorMessage);
        fileListSheet.getRange(rowNum, 5).setValue(new Date());
      }
    }
  }
  console.log('ステップ2: OCR処理が完了しました。');
}

/****************************************************************
 * Gemini API 関連
 ****************************************************************/
function callGeminiApi(fileBlob, prompt) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${CONFIG.GEMINI_MODEL}:generateContent?key=${CONFIG.GEMINI_API_KEY}`;
  
  const payload = {
    "contents": [{
      "parts": [
        { "text": prompt },
        {
          "inline_data": {
            "mime_type": fileBlob.getContentType(),
            "data": Utilities.base64Encode(fileBlob.getBytes())
          }
        }
      ]
    }],
    "generationConfig": {
      "responseMimeType": "application/json",
      "temperature": 0.1,
      "thinkingConfig": {
        "thinkingBudget": CONFIG.THINKING_BUDGET
      }
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
      return {
        success: true,
        data: jsonResponse.candidates[0].content.parts[0].text,
        usage: jsonResponse.usageMetadata,
        error: null
      };
    } else {
       return {
        success: false,
        data: null,
        usage: jsonResponse.usageMetadata || null,
        error: "APIからのレスポンスが予期した形式ではありません。"
      };
    }
  } else {
      console.error(`API Error Response [${responseCode}]: ${responseBody}`);
    return {
      success: false,
      data: null,
      usage: null,
      error: `API Error ${responseCode}: ${responseBody}`
    };
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


/****************************************************************
 * ログ記録 & シート操作
 ****************************************************************/
function logOcrResult(receipts, originalFileId, learningData) {
  const sheet = getSheet(CONFIG.OCR_RESULT_SHEET);
  const originalFile = DriveApp.getFileById(originalFileId);
  const originalFileName = originalFile.getName();
  const fileUrl = originalFile.getUrl();

  const masterData = getMasterData();

  receipts.forEach(r => {
    const taxCategoryCode = getTaxCategoryCode(r.tax_rate, r.tax_code);
    const transactionId = Utilities.getUuid();

    let kanjo = null;
    let hojo = null;

    const learned = learningData[r.storeName];
    if (learned) {
      kanjo = learned.kanjo;
      hojo = learned.hojo;
      console.log(`学習データを適用: ${r.storeName} -> 勘定科目:${kanjo}, 補助科目:${hojo}`);
    } else {
      console.log(`AIによる勘定科目推測を開始: ${r.storeName}`);
      try {
        kanjo = inferAccountTitle(r.storeName, r.description, r.amount, masterData);
        console.log(`AI推測結果: ${kanjo}`);
      } catch(e) {
        console.error(`勘定科目の推測中にエラー: ${e.toString()}`);
        kanjo = "【推測エラー】";
      }
      hojo = ""; 
    }

    const fileLink = `=HYPERLINK("${fileUrl}","${r.filename || originalFileName}")`;
    
    const newRow = [
      transactionId,
      new Date(),
      r.date,
      r.storeName,
      r.description,
      kanjo,
      hojo,
      r.tax_rate,
      r.amount,
      r.tax_amount,
      r.tax_code,
      taxCategoryCode,
      fileLink,
      r.note,
    ];
    sheet.appendRow(newRow);

    const lastRow = sheet.getLastRow();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const learnCheckColIndex = headers.indexOf('学習チェック') + 1;
    if(learnCheckColIndex > 0){
       sheet.getRange(lastRow, learnCheckColIndex).insertCheckboxes();
    }
  });
}

function logTokenUsage(fileName, usage) {
  const sheet = getSheet(CONFIG.TOKEN_LOG_SHEET);
  sheet.appendRow([
    new Date(),
    fileName,
    usage.promptTokenCount || 0,
    usage.thoughtsTokenCount || 0,
    usage.candidatesTokenCount || 0,
    usage.totalTokenCount || 0
  ]);
}

/****************************************************************
 * 勘定科目推測とAI学習のための関数群
 ****************************************************************/

function getMasterData() {
  try {
    const sheet = getSheet(CONFIG.MASTER_SHEET);
    if (!sheet) {
      SpreadsheetApp.getUi().alert(`エラー: シート「${CONFIG.MASTER_SHEET}」が見つかりません。シートを作成してください。`);
      throw new Error(`シート「${CONFIG.MASTER_SHEET}」が見つかりません。`);
    }
    if (sheet.getLastRow() < 2) return [];
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    return data.filter(row => row[0]);
  } catch (e) {
    console.error(e);
    return [];
  }
}

function getLearningData() {
  const learningData = {};
  try {
    const sheet = getSheet(CONFIG.LEARNING_SHEET);
    if (sheet && sheet.getLastRow() > 1) {
      const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
      for (let i = data.length - 1; i >= 0; i--) {
        const row = data[i];
        const storeName = row[1];
        if (storeName && !learningData[storeName]) {
          learningData[storeName] = {
            description: row[2],
            kanjo: row[3],
            hojo: row[4]
          };
        }
      }
    }
  } catch(e){
      console.error("学習データの取得に失敗: " + e.toString());
  }
  return learningData;
}


function inferAccountTitle(storeName, description, amount, masterData) {
  if (!masterData || masterData.length === 0) {
    return "【マスター未設定】";
  }

  const masterListWithKeywords = masterData.map(row => {
    return { title: row[0], keywords: row[1] || "特になし" };
  });
  const masterTitleList = masterData.map(row => row[0]);

  const prompt = `
あなたは、日本の会計基準に精通したベテランの経理専門家です。あなたの任務は、与えられた領収書の情報と、社内ルールを含む勘定科目マスターを基に、最も可能性の高い勘定科目を特定することです。

# 指示
1.  以下の「領収書情報」と「勘定科目マスター」を注意深く分析してください。
2.  特に「勘定科目マスター」の**キーワード/ルール**は重要です。例えば、「2万円未満の飲食代は会議費」といった金額に基づくルールが含まれている場合があります。
3.  すべての情報を総合的に判断し、「勘定科目マスター」のリストの中から最も適切だと考えられる勘定科目を**1つだけ**選択してください。
4.  あなたの回答は、必ず指定されたJSON形式に従ってください。

# 領収書情報
- 店名: ${storeName}
- 摘要: ${description}
- 金額(税込): ${amount}円

# 勘定科目マスター（キーワード/ルールを含む）
${JSON.stringify(masterListWithKeywords)}
`;

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${CONFIG.GEMINI_MODEL}:generateContent?key=${CONFIG.GEMINI_API_KEY}`;
  
  const payload = {
    "contents": [{"parts": [{ "text": prompt }]}],
    "generationConfig": {
      "responseMimeType": "application/json",
      "temperature": 0,
      "responseSchema": {
        "type": "OBJECT",
        "properties": {
          "accountTitle": {
            "type": "STRING",
            "enum": masterTitleList
          }
        },
        "required": ["accountTitle"]
      }
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
    try {
      const jsonResponse = JSON.parse(responseBody);
      const inferredText = jsonResponse.candidates?.[0]?.content?.parts?.[0]?.text;
      if (inferredText) {
        const finalAnswer = JSON.parse(inferredText);
        if (finalAnswer.accountTitle && masterTitleList.includes(finalAnswer.accountTitle)) {
           return finalAnswer.accountTitle;
        }
      }
      console.error("AIからのJSONレスポンスの形式が不正です。", responseBody);
      return "【形式エラー】";
    } catch (e) {
      console.error("AIからのJSONレスポンスの解析に失敗しました。", e.toString(), responseBody);
      return "【解析エラー】";
    }
  } else {
    console.error(`勘定科目推測APIエラー [${responseCode}]: ${responseBody}`);
    return `【APIエラー ${responseCode}】`;
  }
}

/****************************************************************
 * ユーティリティ & 初期化
 ****************************************************************/

/**
 * ★追加：選択された行の領収書画像をプレビュー表示する
 */
function showReceiptPreview() {
  const ui = SpreadsheetApp.getUi();
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    if (sheet.getName() !== CONFIG.OCR_RESULT_SHEET) {
      ui.alert('この機能は「OCR結果」シートで実行してください。');
      return;
    }

    const range = sheet.getActiveRange();
    const startRow = range.getRow();
    if (startRow <= 1) {
      ui.alert('データ行を選択してください。');
      return;
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const linkColIndex = headers.indexOf('ファイルへのリンク') + 1;
    if (linkColIndex === 0) {
      ui.alert('「ファイルへのリンク」列が見つかりません。');
      return;
    }

    const cellFormula = sheet.getRange(startRow, linkColIndex).getFormula();
    if (!cellFormula) {
      ui.alert('選択した行にファイルへのリンクがありません。');
      return;
    }
    
    const urlMatch = cellFormula.match(/HYPERLINK\("([^"]+)"/);
    if (!urlMatch || !urlMatch[1]) {
      ui.alert('リンクの形式が正しくありません。');
      return;
    }

    const fileUrl = urlMatch[1];
    let fileId = null;

    // Google Driveの標準的なURL形式からIDを抽出
    const idMatch1 = fileUrl.match(/d\/([a-zA-Z0-9_-]{28,})/);
    if (idMatch1 && idMatch1[1]) {
      fileId = idMatch1[1];
    } else {
      // 代替のURL形式 (id=...) からIDを抽出
      const idMatch2 = fileUrl.match(/id=([a-zA-Z0-9_-]{28,})/);
      if (idMatch2 && idMatch2[1]) {
        fileId = idMatch2[1];
      }
    }
    
    if (!fileId) {
      ui.alert('ファイルURLからIDを抽出できませんでした。URL: ' + fileUrl);
      return;
    }

    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    const dataUrl = `data:${blob.getContentType()};base64,${Utilities.base64Encode(blob.getBytes())}`;
    const fileName = file.getName();

    const htmlTemplate = HtmlService.createTemplateFromFile('Preview');
    htmlTemplate.fileName = fileName;
    htmlTemplate.dataUrl = dataUrl;

    const htmlOutput = htmlTemplate.evaluate()
        .setWidth(700)
        .setHeight(800);
    ui.showModalDialog(htmlOutput, `領収書プレビュー: ${fileName}`);

  } catch (e) {
    console.error('プレビュー表示中にエラーが発生しました: ' + e.toString());
    // ★★★【修正点】★★★
    // ui.alertの正しいシグネチャ `alert(title, prompt, buttons)` に合わせて修正
    ui.alert('エラー', 'プレビューの表示中にエラーが発生しました。\n\n詳細: ' + e.message, ui.ButtonSet.OK);
  }
}

function getTaxCategoryCode(taxRate, taxCode) {
  const hasInvoiceNumber = taxCode && taxCode.match(/^T\d{13}$/);

  if (taxRate === 10) {
    return hasInvoiceNumber ? '課対仕入内10%適格' : '課対仕入内10%区分80%';
  } else if (taxRate === 8) {
    return hasInvoiceNumber ? '課対仕入内軽減8%適格' : '課対仕入内軽減8%区分80%';
  } else {
    return '対象外';
  }
}

function initializeEnvironment() {
  console.log('環境の初期化を確認・実行します...');
  getFolderByName(CONFIG.ARCHIVE_FOLDER_NAME, true);

  const fileListHeaders = ['ファイルID', 'ファイル名', 'ステータス', 'エラー詳細', '登録日時'];
  
  const ocrResultHeaders = [
    '取引ID', '処理日時', '取引日', '店名', '摘要', '勘定科目', '補助科目', 
    '税率(%)', '金額(税込)', 'うち消費税', '登録番号', 
    '消費税課税区分コード', 'ファイルへのリンク', '備考', '学習チェック'
  ];
  const tokenLogHeaders = ['日時', 'ファイル名', '入力トークン', '思考トークン', '出力トークン', '合計トークン'];
  
  const learningHeaders = ['学習登録日時', '店名', '摘要', '勘定科目', '補助科目', '取引ID'];

  createSheetWithHeaders(CONFIG.FILE_LIST_SHEET, fileListHeaders);
  createSheetWithHeaders(CONFIG.OCR_RESULT_SHEET, ocrResultHeaders, true);
  createSheetWithHeaders(CONFIG.TOKEN_LOG_SHEET, tokenLogHeaders);
  createSheetWithHeaders(CONFIG.LEARNING_SHEET, learningHeaders);

  console.log('環境の初期化が完了しました。');
}

function activateFilter() {
  const sheet = getSheet(CONFIG.OCR_RESULT_SHEET);
  if (sheet) {
    if (sheet.getFilter()) {
      sheet.getFilter().remove();
    }
    sheet.getDataRange().createFilter();
    SpreadsheetApp.getUi().alert('フィルタをオンにしました。');
  }
}

function getSpreadsheet() {
  try {
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    SpreadsheetApp.getUi().alert("スクリプトエラー", "スプレッドシートにアクセスできません。");
    console.error("スプレッドシートにアクセスできませんでした: " + e.toString());
    throw e;
  }
}

function getSheet(name) {
  return getSpreadsheet().getSheetByName(name);
}

function getFolderByName(name, createIfNotExist = false) {
  const sourceFolder = DriveApp.getFolderById(CONFIG.SOURCE_FOLDER_ID);
  const parents = sourceFolder.getParents();

  if (parents.hasNext()) {
    const parentFolder = parents.next();
    const folders = parentFolder.getFoldersByName(name);
    if (folders.hasNext()) {
      return folders.next();
    }
    if (createIfNotExist) {
      console.log(`フォルダ「${name}」を「${parentFolder.getName()}」内に作成します。`);
      return parentFolder.createFolder(name);
    }
  } else {
    const rootFolders = DriveApp.getFoldersByName(name);
     if (rootFolders.hasNext()) {
       return rootFolders.next();
     }
     if (createIfNotExist) {
       console.log(`フォルダ「${name}」をマイドライブ直下に作成します。`);
       return DriveApp.createFolder(name);
     }
  }
  
  return null;
}

function createSheetWithHeaders(sheetName, headers, activateFilterFlag = false) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    console.log(`シート "${sheetName}" を作成します。`);
    const newSheet = ss.insertSheet(sheetName);
    newSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    newSheet.setFrozenRows(1);
    if(activateFilterFlag && newSheet.getLastRow() > 0) {
      newSheet.getDataRange().createFilter();
    }
  } else {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  }
}


/*
★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
【追加手順】: 以下のHTMLコードで「Preview.html」という名前の
HTMLファイルを新規作成してください。

1. Apps Scriptエディタの左側「ファイル」の横にある「+」をクリック
2. 「HTML」を選択
3. ファイル名を「Preview」と入力してEnter
4. 作成されたファイルの中身を、以下のコードで完全に上書きします。
★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★

<!DOCTYPE html>
<html>
  <head>
    <base target="_top">
    <style>
      body { 
        font-family: 'Helvetica Neue', Arial, sans-serif;
        margin: 0; 
        padding: 0;
        background-color: #f0f2f5;
        display: flex;
        justify-content: center;
        align-items: center;
        height: 100vh;
      }
      #container {
        max-width: 95%;
        max-height: 95vh;
        overflow: auto;
        background-color: white;
        border-radius: 8px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
        display: flex;
        flex-direction: column;
      }
      h3 {
        padding: 16px 24px;
        margin: 0;
        background-color: #ffffff;
        border-bottom: 1px solid #e0e0e0;
        color: #333;
        font-size: 16px;
        font-weight: 600;
        border-top-left-radius: 8px;
        border-top-right-radius: 8px;
        text-align: center;
      }
      #image-wrapper {
        padding: 24px;
        text-align: center;
        overflow: auto;
      }
      img {
        max-width: 100%;
        height: auto;
        border: 1px solid #ddd;
        border-radius: 4px;
      }
    </style>
  </head>
  <body>
    <div id="container">
      <h3><?= fileName ?></h3>
      <div id="image-wrapper">
        <img src="<?= dataUrl ?>" alt="領収書プレビュー">
      </div>
    </div>
  </body>
</html>

*/
