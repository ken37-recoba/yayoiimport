// =================================================================================
// ファイル名: Utilities.gs
// 役割: 様々な場所から呼び出される補助的な便利関数を管理します。
// =================================================================================

function logError_(functionName, error, contextInfo = '') {
    try {
        if (!CONFIG) {
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
  const range = sheet.getRange(row, col);
  const transactionId = sheet.getRange(row, headers.indexOf('取引ID') + 1).getValue();
  if (!transactionId) return;

  let contextInfo = `Transaction ID: ${transactionId}, Cell: ${range.getA1Notation()}`;
  try {
    if (range.isChecked()) {
      if (range.getNote().includes('学習済み')) return;

      const dataRow = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
      const storeName = dataRow[headers.indexOf('店名')];
      const description = dataRow[headers.indexOf('摘要')];
      const kanjo = dataRow[headers.indexOf('勘定科目')];
      const hojo = dataRow[headers.indexOf('補助科目')];

      getSheet(CONFIG.LEARNING_SHEET).appendRow([
        storeName,
        description, 
        '', 
        '', 
        kanjo,
        hojo,
        '', // 摘要テンプレート用の空欄
        new Date(),
        transactionId
      ]);

      range.setNote(`学習済み (ID: ${transactionId})`);
      sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground('#e6f4ea');
      SpreadsheetApp.getActiveSpreadsheet().toast(`「${storeName}」のルールを作成しました。「学習データ」シートで詳細を編集できます。`);
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
    const learningRules = getLearningData();
    
    const newRows = receipts.map(r => {
      let kanjo = null, hojo = null;
      let isLearned = false;
      let finalDescription = r.description || '';

      // --- 高度な学習ルールによる判定 ---
      for (const rule of learningRules) {
        const ocrData = {
          storeName: normalizeStoreName(r.storeName),
          description: r.description || '',
          amount: Number(r.amount) || 0
        };
        
        const storeMatch = !rule.storeName || ocrData.storeName.includes(rule.storeName) || rule.storeName.includes(ocrData.storeName);
        const descMatch = !rule.descriptionKeyword || ocrData.description.includes(rule.descriptionKeyword);
        
        let amountMatch = true;
        if (rule.amountCondition && rule.amountValue !== null) { // 金額が0の場合も考慮
            if (rule.amountCondition === '以上') {
                amountMatch = ocrData.amount >= rule.amountValue;
            } else if (rule.amountCondition === '未満') {
                amountMatch = ocrData.amount < rule.amountValue;
            }
        }
        
        if (storeMatch && descMatch && amountMatch) {
          kanjo = rule.kanjo;
          hojo = rule.hojo;
          // 摘要テンプレートの適用
          if (rule.descriptionTemplate) {
            finalDescription = rule.descriptionTemplate
              .replace(/【日付】/g, r.date || '')
              .replace(/【店名】/g, r.storeName || '')
              .replace(/【金額】/g, Math.trunc(r.amount || 0));
          }
          isLearned = true;
          console.log(`学習ルールを適用: OCRデータ(店名:${r.storeName}, 摘要:${ocrData.description}, 金額:${ocrData.amount}) がルール(店名:${rule.rawStoreName}, 摘要キーワード:${rule.descriptionKeyword}, 金額条件:${rule.amountCondition}${rule.amountValue})に一致しました。`);
          break;
        }
      }

      // --- AIによる推測 (学習ルールに一致しなかった場合) ---
      if (!isLearned) {
        console.log("学習ルールに一致しなかったため、AIによる推測を実行します。");
        kanjo = inferAccountTitle(r.storeName, r.description, r.amount, masterData);
        hojo = "";
      }

      const truncatedAmount = Math.trunc(r.amount || 0);
      const truncatedTaxAmount = Math.trunc(r.tax_amount || 0);

      return [
        Utilities.getUuid(), new Date(), r.date, r.storeName, finalDescription,
        kanjo, hojo, r.tax_rate, truncatedAmount, truncatedTaxAmount, r.tax_code,
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
    throw e;
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
    return 0;
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
  const learningRules = [];
  try {
    const sheet = getSheet(CONFIG.LEARNING_SHEET);
    if (!sheet || sheet.getLastRow() < 2) return [];

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, CONFIG.HEADERS.LEARNING.length).getValues();
    const headers = CONFIG.HEADERS.LEARNING;
    const COL = {
        STORE_NAME: headers.indexOf('店名'),
        DESC_KEYWORD: headers.indexOf('摘要（キーワード）'),
        AMOUNT_COND: headers.indexOf('金額条件'),
        AMOUNT_VAL: headers.indexOf('金額'),
        KANJO: headers.indexOf('勘定科目'),
        HOJO: headers.indexOf('補助科目'),
        DESC_TEMPLATE: headers.indexOf('摘要のテンプレート'),
    };

    for (const row of data) {
      if (!row[COL.KANJO]) continue;
      
      const amountValue = row[COL.AMOUNT_VAL];
      
      learningRules.push({
        rawStoreName: row[COL.STORE_NAME] || '',
        storeName: normalizeStoreName(row[COL.STORE_NAME] || ''),
        descriptionKeyword: row[COL.DESC_KEYWORD] || '',
        amountCondition: row[COL.AMOUNT_COND] || '',
        amountValue: (amountValue !== '' && !isNaN(amountValue)) ? Number(amountValue) : null,
        kanjo: row[COL.KANJO],
        hojo: row[COL.HOJO] || '',
        descriptionTemplate: row[COL.DESC_TEMPLATE] || '',
      });
    }
    console.log(`学習データを ${learningRules.length} 件読み込みました。`);

  } catch(e) {
    logError_("getLearningData", e);
    console.error("学習データの取得に失敗: " + e.toString());
  }
  return learningRules;
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

function generateNewFileName_(transaction, originalFileName) {
  try {
    const date = new Date(transaction.date);
    const formattedDate = Utilities.formatDate(date, "JST", "yyyyMMdd");
    
    // ファイル名に使えない文字を置換
    const safeStoreName = (transaction.storeName || '不明な店名').replace(/[\\/:*?"<>|]/g, '-');
    
    const amount = Math.trunc(transaction.amount || 0);
    
    const extensionMatch = originalFileName.match(/\.([^.]+)$/);
    const extension = extensionMatch ? extensionMatch[1] : 'jpg';

    return `${formattedDate}_${safeStoreName}_${amount}円.${extension}`;
  } catch (e) {
    console.error(`新しいファイル名の生成に失敗しました: ${e.toString()}`);
    // エラーが発生した場合は、元のファイル名をそのまま返す
    return originalFileName;
  }
}
