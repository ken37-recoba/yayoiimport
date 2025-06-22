// =================================================================================
// ファイル名: Utilities.gs
// 役割: 様々な場所から呼び出される補助的な便利関数を管理します。
// =================================================================================

function calculateTaxAmount_(amount, taxCategory) {
  if (!amount || !taxCategory) return 0;
  
  const taxStr = taxCategory.toString();
  if (taxStr.includes('対象外') || taxStr.includes('不課税') || taxStr.includes('非課税売上')) {
    return 0;
  }
  
  if (taxStr.includes('軽減8%')) {
    return Math.floor(amount * 8 / 108);
  }
  
  return Math.floor(amount * 10 / 110);
}

function logError_(functionName, error, contextInfo = '') {
    try {
        if (!CONFIG) return;
        const sheet = getSheet(CONFIG.ERROR_LOG_SHEET);
        if (!sheet) return;
        sheet.appendRow([ new Date(), functionName, error.message, contextInfo, error.stack || 'N/A' ]);
    } catch (logErr) {
        console.error(`Failed to write to error log. Original error in ${functionName}: ${error.stack || error.message}. Logging error: ${logErr.message}`);
    }
}

function handleLearningCheck(sheet, row, col, headers) {
  loadConfig_();
  const range = sheet.getRange(row, col);
  const transactionId = sheet.getRange(row, headers.indexOf('取引ID') + 1).getValue();
  if (!transactionId) return;

  try {
    if (range.isChecked()) {
      if (range.getNote().includes('学習済み')) return;
      const dataRow = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
      const COL = headers.reduce((acc, h, i) => ({...acc, [h]: i}), {});
      
      const taxCategory = dataRow[COL['消費税課税区分コード']];

      getSheet(CONFIG.LEARNING_SHEET).appendRow([
        dataRow[COL['店名']], dataRow[COL['摘要']], '', '', '', 
        dataRow[COL['勘定科目']], dataRow[COL['補助科目']], taxCategory,
        '', new Date(), transactionId
      ]);
      range.setNote(`学習済み (ID: ${transactionId})`);
      sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground('#e6f4ea');
    } else {
      const deletedCount = deleteLearningDataByIds([transactionId]);
      if (deletedCount > 0) {
        range.clearNote();
        sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground(null);
      }
    }
  } catch (e) {
    logError_('handleLearningCheck', e);
  }
}

function handlePassbookLearningCheck(sheet, row, col, headers) {
  loadConfig_();
  const range = sheet.getRange(row, col);
  const transactionId = sheet.getRange(row, headers.indexOf('取引ID') + 1).getValue();
  if (!transactionId) return;

  try {
    if (range.isChecked()) {
      if (range.getNote().includes('学習済み')) return;
      const dataRow = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
      const COL = headers.reduce((acc, h, i) => ({...acc, [h]: i}), {});

      const isDeposit = Number(dataRow[COL['入金額']]) > 0;
      const taxCategory = isDeposit ? dataRow[COL['貸方税区分']] : dataRow[COL['借方税区分']];
      const passbookAccountName = dataRow[COL['通帳勘定科目']];
      
      getSheet(CONFIG.LEARNING_SHEET).appendRow([
        '', dataRow[COL['摘要']], passbookAccountName, '', '',
        dataRow[COL['相手方勘定科目']], dataRow[COL['相手方補助科目']], taxCategory,
        '', new Date(), transactionId
      ]);
      range.setNote(`学習済み (ID: ${transactionId})`);
      sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground('#e6f4ea');
    } else {
      const deletedCount = deleteLearningDataByIds([transactionId]);
       if (deletedCount > 0) {
        range.clearNote();
        sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground(null);
      }
    }
  } catch (e) {
    logError_('handlePassbookLearningCheck', e);
  }
}

function handleTaxCodeRemoval(sheet, row, headers) {
  loadConfig_();
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
    logError_('handleTaxCodeRemoval', e, `Row: ${row}`);
  }
}

function logOcrResult(receipts, originalFileId) {
  loadConfig_();
  const contextInfo = `File ID: ${originalFileId}`;
  try {
    const sheet = getSheet(CONFIG.OCR_RESULT_SHEET);
    const originalFile = DriveApp.getFileById(originalFileId);
    const masterData = getMasterData();
    const learningRules = getLearningData();
    
    const newRows = receipts.map(r => {
      let kanjo = null, hojo = null;
      let isLearned = false;
      let finalDescription = r.description || '';

      for (const rule of learningRules) {
        if (!rule.storeName) continue;
        const ocrData = { storeName: normalizeStoreName(r.storeName), description: r.description || '', amount: Number(r.amount) || 0 };
        const storeMatch = ocrData.storeName.includes(rule.storeName) || rule.storeName.includes(ocrData.storeName);
        const descMatch = !rule.descriptionKeyword || ocrData.description.includes(rule.descriptionKeyword);
        let amountMatch = true;
        if (rule.amountCondition && rule.amountValue !== null) {
            amountMatch = rule.amountCondition === '以上' ? ocrData.amount >= rule.amountValue : ocrData.amount < rule.amountValue;
        }
        if (storeMatch && descMatch && amountMatch) {
          kanjo = rule.kanjo;
          hojo = rule.hojo;
          if (rule.descriptionTemplate) {
            finalDescription = rule.descriptionTemplate.replace(/【日付】/g, r.date || '').replace(/【店名】/g, r.storeName || '').replace(/【金額】/g, Math.trunc(r.amount || 0));
          }
          isLearned = true;
          break;
        }
      }

      if (!isLearned) {
        kanjo = inferAccountTitle(r.storeName, r.description, r.amount, masterData);
        hojo = "";
      }

      const truncatedAmount = Math.trunc(r.amount || 0);
      const truncatedTaxAmount = Math.trunc(r.tax_amount || 0);

      return [
        Utilities.getUuid(), new Date(), r.date, r.storeName, finalDescription,
        kanjo, hojo, r.tax_rate, truncatedAmount, truncatedTaxAmount, r.tax_code,
        getTaxCategoryCode(r.tax_rate, r.tax_code),
        `=HYPERLINK("${originalFile.getUrl()}","${r.filename || originalFile.getName()}")`, r.note
      ];
    });

    if (newRows.length > 0) {
      const startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, newRows.length, newRows[0].length).setValues(newRows);
      const learnCheckCol = CONFIG.HEADERS.OCR_RESULT.indexOf('学習チェック') + 1;
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
    if (usage) {
      sheet.appendRow([ new Date(), fileName, usage.promptTokenCount || 0, usage.thoughtsTokenCount || 0, usage.candidatesTokenCount || 0, usage.totalTokenCount || 0 ]);
    }
  } catch (e) {
    logError_('logTokenUsage', e, `File: ${fileName}`);
  }
}

function deleteLearningDataByIds(transactionIds) {
  loadConfig_();
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
    logError_('deleteLearningDataByIds', e, `IDs: ${transactionIds.join(', ')}`);
    return 0;
  }
}

function getTaxCategoryCode(taxRate, taxCode) {
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
    const COL = CONFIG.HEADERS.LEARNING.reduce((acc, h, i) => ({...acc, [h]: i}), {});
    for (const row of data) {
      if (!row[COL['勘定科目']]) continue;
      const amountValue = row[COL['金額']];
      learningRules.push({
        rawStoreName: row[COL['店名']] || '',
        storeName: normalizeStoreName(row[COL['店名']] || ''),
        descriptionKeyword: row[COL['摘要（キーワード）']] || '',
        passbookAccountName: row[COL['通帳勘定科目']] || '',
        amountCondition: row[COL['金額条件']] || '',
        amountValue: (amountValue !== '' && !isNaN(amountValue)) ? Number(amountValue) : null,
        kanjo: row[COL['勘定科目']],
        hojo: row[COL['補助科目']] || '',
        taxCategory: row[COL['税区分']] || '対象外',
        descriptionTemplate: row[COL['摘要のテンプレート']] || '',
      });
    }
  } catch(e) {
    logError_("getLearningData", e);
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
    showError(`シート「${CONFIG.MASTER_SHEET}」からデータを取得できませんでした。`);
    return [];
  }
}

function getFileIdFromCell(sheet, row) {
  loadConfig_();
  const sheetName = sheet.getName();
  let headers;
  if (sheetName === CONFIG.OCR_RESULT_SHEET || sheetName === CONFIG.EXPORTED_SHEET) {
    headers = CONFIG.HEADERS.OCR_RESULT;
  } else if (sheetName === CONFIG.PASSBOOK_RESULT_SHEET || sheetName === CONFIG.PASSBOOK_EXPORTED_SHEET) {
    headers = CONFIG.HEADERS.PASSBOOK_RESULT;
  } else {
    showError('このシートではプレビュー機能は利用できません。');
    return null;
  }
  const linkCol = headers.indexOf('ファイルへのリンク') + 1;
  if (linkCol === 0) return null;
  const cellFormula = sheet.getRange(row, linkCol).getFormula();
  if (!cellFormula) return null;
  const urlMatch = cellFormula.match(/HYPERLINK\("([^"]+)"/);
  if (!urlMatch || !urlMatch[1]) return null;
  const fileUrl = urlMatch[1];
  const idMatch = fileUrl.match(/d\/([a-zA-Z0-9_-]{28,})/);
  return idMatch ? idMatch[1] : null;
}

function normalizeStoreName(name) {
  if (!name || typeof name !== 'string') return '';
  return name.toLowerCase().replace(/[Ａ-Ｚａ-ｚ０-９！＂＃＄％＆＇（）＊＋，－．／：；＜＝＞？＠［＼］＾＿｀｛｜｝～]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0)).replace(/\s|　/g, '').replace(/-|－|—|ｰ/g, '').replace(/株式会社|有限会社|\(株\)|\（株\)|\(有\)|\（有\）/g, '');
}

function getSheet(name) {
  loadConfig_();
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function createSheetWithHeaders(sheetName, headers, activateFilterFlag = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
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
    if (sheet.getFilter()) sheet.getFilter().remove();
    if (sheet.getMaxRows() > 1) sheet.getDataRange().createFilter();
  }
}

function showError(message, title = 'エラー') {
  SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
}

function generateNewFileName_(transaction, originalFileName) {
  try {
    const date = new Date(transaction.date);
    const formattedDate = Utilities.formatDate(date, "JST", "yyyyMMdd");
    const safeStoreName = (transaction.storeName || '不明').replace(/[\\/:*?"<>|]/g, '-');
    const amount = Math.trunc(transaction.amount || 0);
    const extension = originalFileName.includes('.') ? originalFileName.split('.').pop() : 'jpg';
    return `${formattedDate}_${safeStoreName}_${amount}円.${extension}`;
  } catch (e) {
    return originalFileName;
  }
}

function getPassbookMasterData() {
  loadConfig_();
  const masterData = [];
  try {
    const sheet = getSheet(CONFIG.PASSBOOK_MASTER_SHEET);
    if (!sheet || sheet.getLastRow() < 2) return [];

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    for (const row of data) {
      if (row[0] && row[1]) {
        masterData.push({ keyword: row[0].toLowerCase(), accountName: row[1] });
      }
    }
  } catch(e) {
    logError_("getPassbookMasterData", e);
  }
  return masterData;
}

function logPassbookResult(transactions, originalFileId, originalFileName) {
  loadConfig_();
  const contextInfo = `Passbook File ID: ${originalFileId}`;
  try {
    const sheet = getSheet(CONFIG.PASSBOOK_RESULT_SHEET);
    const originalFile = DriveApp.getFileById(originalFileId);
    
    const passbookMaster = getPassbookMasterData();
    const learningRules = getLearningData();
    const fileNameLower = originalFileName.toLowerCase();
    let passbookAccountName = '（未設定）';
    for (const master of passbookMaster) {
      if (fileNameLower.includes(master.keyword)) {
        passbookAccountName = master.accountName;
        break;
      }
    }

    const filteredTransactions = transactions.filter(tx => !((tx.取引内容 || '').includes('繰越') && (Number(tx.入金額) || 0) === 0 && (Number(tx.出金額) || 0) === 0));
    let verifiedTransactions = verifyAndCorrectPassbookBalances(filteredTransactions);

    const newRows = verifiedTransactions.map(tx => {
      let isLearned = false;
      let inferred = {};

      for (const rule of learningRules) {
        if (rule.storeName !== '') continue;
        
        const keywordMatch = !rule.descriptionKeyword || (tx.取引内容 || '').includes(rule.descriptionKeyword);
        const passbookMatch = !rule.passbookAccountName || rule.passbookAccountName === passbookAccountName;

        if (keywordMatch && passbookMatch) {
          inferred = {
            accountTitle: rule.kanjo,
            subAccount: rule.hojo,
            taxCategory: rule.taxCategory
          };
          isLearned = true;
          break;
        }
      }

      if (!isLearned) {
        inferred = inferPassbookAccountTitle(tx.取引内容);
      }
      
      const isDeposit = Number(tx.入金額) > 0;
      let debitTaxCategory = '対象外', creditTaxCategory = '対象外';
      if (isDeposit) creditTaxCategory = inferred.taxCategory;
      else debitTaxCategory = inferred.taxCategory;

      return [
        Utilities.getUuid(), new Date(), tx.取引日, tx.取引内容,
        tx.入金額, tx.出金額, tx.残高,
        passbookAccountName, inferred.accountTitle, inferred.subAccount,
        debitTaxCategory, creditTaxCategory,
        `=HYPERLINK("${originalFile.getUrl()}","${originalFileName}")`, tx.備考 || ''
      ];
    });

    if (newRows.length > 0) {
      const startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, newRows.length, newRows[0].length).setValues(newRows);
      const learnCheckCol = CONFIG.HEADERS.PASSBOOK_RESULT.indexOf('学習チェック') + 1;
      if (learnCheckCol > 0) {
        sheet.getRange(startRow, learnCheckCol, newRows.length).insertCheckboxes();
      }
    }
  } catch (e) {
    logError_('logPassbookResult', e, contextInfo);
    throw e;
  }
}

function verifyAndCorrectPassbookBalances(transactions) {
  if (!transactions || transactions.length < 2) return transactions;
  
  for (let i = 1; i < transactions.length; i++) {
    const prev = transactions[i - 1];
    const curr = transactions[i];
    
    const prevBalance = Number(prev.残高) || 0;
    const deposit = Number(curr.入金額) || 0;
    const withdrawal = Number(curr.出金額) || 0;
    const currentBalance = Number(curr.残高) || 0;
    
    const expectedBalance = prevBalance - withdrawal + deposit;
    
    if (currentBalance !== expectedBalance) {
      const swappedBalance = prevBalance - deposit + withdrawal;
      if (currentBalance === swappedBalance && (deposit > 0 || withdrawal > 0)) {
        curr.入金額 = withdrawal;
        curr.出金額 = deposit;
        curr.備考 = (curr.備考 || '') + '[入出金自動補正]';
      }
    }
  }
  return transactions;
}
