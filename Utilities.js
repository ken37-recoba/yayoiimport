// =================================================================================
// ファイル名: Utilities.js (記号誤認識 対策版)
// 役割: 様々な場所から呼び出される補助的な便利関数を管理します。
// =================================================================================

function verifyInvoiceNumber_(rawInvoiceNumber) {
  // 国税庁API連携を無効化。常にOCRの結果を正とする。
  const formattedNumber = rawInvoiceNumber ? rawInvoiceNumber.trim().toUpperCase() : '';
  return { 
    isValid: false, 
    officialName: null, 
    formattedNumber: formattedNumber, 
    note: ''
  };
}

function correctDate_(ocrDateString) {
  if (!ocrDateString || typeof ocrDateString !== 'string') {
    return { correctedDate: ocrDateString, wasCorrected: false, note: '【要確認：日付不明】' };
  }

  try {
    const processingDate = new Date();
    let dateStrToParse = ocrDateString.replace(/\s/g, '');
    const warekiMatch = dateStrToParse.match(/^(令和|平成|昭和|大正|明治)?(\d+|元)[年](\d+)[月](\d+)[日]/);
    
    if (warekiMatch) {
      let year = parseInt(warekiMatch[2] === '元' ? 1 : warekiMatch[2], 10);
      const era = warekiMatch[1];
      if (era === '令和' || (!era && year < 10)) {
        year += 2018;
      } else if (era === '平成' || (!era && year > 10)) {
        year += 1988;
      }
      dateStrToParse = `${year}-${warekiMatch[3]}-${warekiMatch[4]}`;
    }

    let ocrDate = new Date(dateStrToParse);

    if (isNaN(ocrDate.getTime())) {
       if (new Date(ocrDateString).getTime() === 0) {
         ocrDate = new Date();
       } else {
         return { correctedDate: ocrDateString, wasCorrected: false, note: '【要確認：日付形式不正】' };
       }
    }

    const ocrYear = ocrDate.getFullYear();
    const processingYear = processingDate.getFullYear();
    let correctedDate = new Date(ocrDate);
    let wasCorrected = false;

    if (ocrYear < processingYear - 1 || ocrYear > processingYear) {
      correctedDate.setFullYear(processingYear);
      wasCorrected = true;
    }

    const bufferProcessingDate = new Date();
    bufferProcessingDate.setDate(bufferProcessingDate.getDate() + 1);

    if (correctedDate > bufferProcessingDate) {
      correctedDate.setFullYear(correctedDate.getFullYear() - 1);
      wasCorrected = true;
    }
    
    const finalDateString = Utilities.formatDate(correctedDate, 'JST', 'yyyy/MM/dd');
    const note = wasCorrected ? '[日付を自動補正]' : '';

    return { correctedDate: finalDateString, wasCorrected: wasCorrected, note: note };

  } catch (e) {
    return { correctedDate: ocrDateString, wasCorrected: false, note: '【要確認：日付処理エラー】' };
  }
}

function normalizeText_(text) {
  if (!text || typeof text !== 'string') return '';
  
  let result = text;

  result = result.replace(/[Ａ-Ｚａ-ｚ０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));

  const hankakuKatakana = {
    'ｶﾞ': 'ガ', 'ｷﾞ': 'ギ', 'ｸﾞ': 'グ', 'ｹﾞ': 'ゲ', 'ｺﾞ': 'ゴ', 'ｻﾞ': 'ザ', 'ｼﾞ': 'ジ', 'ｽﾞ': 'ズ', 'ｾﾞ': 'ゼ', 'ｿﾞ': 'ゾ',
    'ﾀﾞ': 'ダ', 'ﾁﾞ': 'ヂ', 'ﾂﾞ': 'ヅ', 'ﾃﾞ': 'デ', 'ﾄﾞ': 'ド', 'ﾊﾞ': 'バ', 'ﾋﾞ': 'ビ', 'ﾌﾞ': 'ブ', 'ﾍﾞ': 'ベ', 'ﾎﾞ': 'ボ',
    'ﾊﾟ': 'パ', 'ﾋﾟ': 'ピ', 'ﾌﾟ': 'プ', 'ﾍﾟ': 'ペ', 'ﾎﾟ': 'ポ', 'ｳﾞ': 'ヴ', 'ﾜﾞ': 'ヷ', 'ｦﾞ': 'ヺ', 'ｱ': 'ア', 'ｲ': 'イ',
    'ｳ': 'ウ', 'ｴ': 'エ', 'ｵ': 'オ', 'ｶ': 'カ', 'ｷ': 'キ', 'ｸ': 'ク', 'ｹ': 'ケ', 'ｺ': 'コ', 'ｻ': 'サ', 'ｼ': 'シ', 'ｽ': 'ス',
    'ｾ': 'セ', 'ｿ': 'ソ', 'ﾀ': 'タ', 'ﾁ': 'チ', 'ﾂ': 'ツ', 'ﾃ': 'テ', 'ﾄ': 'ト', 'ﾅ': 'ナ', 'ﾆ': 'ニ', 'ﾇ': 'ヌ', 'ﾈ': 'ネ',
    'ﾉ': 'ノ', 'ﾊ': 'ハ', 'ﾋ': 'ヒ', 'ﾌ': 'フ', 'ﾍ': 'ヘ', 'ﾎ': 'ホ', 'ﾏ': 'マ', 'ﾐ': 'ミ', 'ﾑ': 'ム', 'ﾒ': 'メ', 'ﾓ': 'モ',
    'ﾔ': 'ヤ', 'ﾕ': 'ユ', 'ﾖ': 'ヨ', 'ﾗ': 'ラ', 'ﾘ': 'リ', 'ﾙ': 'ル', 'ﾚ': 'レ', 'ﾛ': 'ロ', 'ﾜ': 'ワ', 'ｦ': 'ヲ', 'ﾝ': 'ン',
    'ｧ': 'ァ', 'ｨ': 'ィ', 'ｩ': 'ゥ', 'ｪ': 'ェ', 'ｫ': 'ォ', 'ｯ': 'ッ', 'ｬ': 'ャ', 'ｭ': 'ュ', 'ｮ': 'ョ', '｡': '。', '､': '、',
    'ｰ': 'ー', '｢': '「', '｣': '」', '･': '・'
  };

  const reg = new RegExp('(' + Object.keys(hankakuKatakana).join('|') + ')', 'g');
  result = result.replace(reg, s => hankakuKatakana[s]);

  const hankakuSymbols = {
    '!': '！', '"': '”', '#': '＃', '$': '＄', '%': '％', '&': '＆', "'": '’', '(': '（', ')': '）', '*': '＊', '+': '＋',
    ',': '、', '-': '－', '.': '．', '/': '／', ':': '：', ';': '；', '<': '＜', '=': '＝', '>': '＞', '?': '？', '@': '＠',
    '[': '［', '\\': '￥', ']': '］', '^': '＾', '_': '＿', '`': '‘', '{': '｛', '|': '｜', '}': '｝', '~': '～'
  };
  const symbolReg = new RegExp('(' + Object.keys(hankakuSymbols).map(k => k.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|') + ')', 'g');
  result = result.replace(symbolReg, s => hankakuSymbols[s]);
  
  result = result.replace(/(\u30ab|\u30ad|\u30af|\u30b1|\u30b3|\u30b5|\u30b7|\u30b9|\u30bb|\u30bd|\u30bf|\u30c1|\u30c4|\u30c6|\u30c8|\u30cf|\u30d2|\u30d5|\u30d8|\u30db|\u30a6)(\u3099)/g, s => String.fromCharCode(s.charCodeAt(0) + 1));
  result = result.replace(/(\u30cf|\u30d2|\u30d5|\u30d8|\u30db)(\u309a)/g, s => String.fromCharCode(s.charCodeAt(0) + 2));

  return result;
}


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
      let finalStoreName = r.storeName || '';
      let finalTaxCode = r.tax_code || '';

      let finalAmount = Math.trunc(r.amount || 0);
      let finalTaxAmount = Math.trunc(r.tax_amount || 0);
      let taxRate = r.tax_rate || 0;
      let finalNote = r.note || '';

      const dateCorrectionResult = correctDate_(r.date);
      const finalDate = dateCorrectionResult.correctedDate;
      if (dateCorrectionResult.note) {
        finalNote = `${finalNote} ${dateCorrectionResult.note}`.trim();
      }

      if (finalAmount > 0 && finalTaxAmount > 0 && (taxRate === 10 || taxRate === 8)) {
        const calculatedTaxFromInclusive = finalAmount * taxRate / (100 + taxRate);
        const calculatedTaxFromExclusive = finalAmount * taxRate / 100;
        const isAlreadyInclusive = Math.abs(calculatedTaxFromInclusive - finalTaxAmount) <= 1;
        const isExclusive = Math.abs(calculatedTaxFromExclusive - finalTaxAmount) <= 1;

        if (!isAlreadyInclusive && isExclusive) {
          const correctedAmount = finalAmount + finalTaxAmount;
          console.log(`金額を自動補正しました。ファイル: ${originalFile.getName()}, 旧金額: ${finalAmount}, 新金額: ${correctedAmount}`);
          finalAmount = correctedAmount;
          finalNote = `${finalNote} [金額を自動補正]`.trim();
        }
      }

      // ▼▼▼【改善箇所】高額取引の警告機能を追加 ▼▼▼
      const HIGH_AMOUNT_THRESHOLD = 50000; // 5万円をしきい値とする
      if (finalAmount > HIGH_AMOUNT_THRESHOLD) {
        finalNote = `${finalNote} [【要確認：高額取引】]`.trim();
      }
      // ▲▲▲ 改善箇所 ▲▲▲

      if (finalTaxCode) {
        const verificationResult = verifyInvoiceNumber_(finalTaxCode);
        finalTaxCode = verificationResult.formattedNumber;
        if (verificationResult.isValid) {
          finalStoreName = verificationResult.officialName;
        } else if (verificationResult.note) {
          finalNote = `${finalNote} ${verificationResult.note}`.trim();
        }
      }

      for (const rule of learningRules) {
        if (!rule.storeName) continue;
        const ocrData = { storeName: normalizeStoreName(finalStoreName), description: finalDescription, amount: finalAmount };
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
            finalDescription = rule.descriptionTemplate.replace(/【日付】/g, finalDate).replace(/【店名】/g, finalStoreName).replace(/【金額】/g, finalAmount);
          }
          isLearned = true;
          break;
        }
      }

      if (!isLearned) {
        kanjo = inferAccountTitle(finalStoreName, finalDescription, finalAmount, masterData);
        hojo = "";
      }

      let finalTaxCategory = getTaxCategoryCode(taxRate, finalTaxCode);

      if (taxRate === 0 && finalTaxAmount === 0 && finalAmount > 0) {
        const exemptTitles = ['租税公課', '諸会費', '保険料'];
        if (kanjo && !exemptTitles.includes(kanjo)) {
          taxRate = 10;
          finalTaxAmount = Math.floor(finalAmount * 10 / 110);
          finalTaxCategory = getTaxCategoryCode(taxRate, finalTaxCode);
          finalNote = `${finalNote} [消費税を自動計算]`.trim();
          console.log(`消費税を自動計算しました。ファイル: ${originalFile.getName()}, 勘定科目: ${kanjo}, 金額: ${finalAmount}, 計算された消費税: ${finalTaxAmount}`);
        }
      }

      const normalizedStoreName = normalizeText_(finalStoreName);
      const normalizedFinalDescription = normalizeText_(finalDescription);

      return [
        Utilities.getUuid(), new Date(), finalDate, normalizedStoreName, normalizedFinalDescription,
        kanjo, hojo, taxRate, finalAmount, finalTaxAmount, finalTaxCode,
        finalTaxCategory,
        `=HYPERLINK("${originalFile.getUrl()}","${r.filename || originalFile.getName()}")`, finalNote
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

function generateNewPassbookFileName_(passbookAccountName, originalFileName) {
  try {
    const safeAccountName = (passbookAccountName || '不明な通帳').replace(/[\\/:*?"<>|]/g, '_');
    const formattedDate = Utilities.formatDate(new Date(), "JST", "yyyyMMdd");
    const extension = originalFileName.includes('.') ? originalFileName.split('.').pop() : 'jpg';
    return `${safeAccountName}_${formattedDate}.${extension}`;
  } catch (e) {
    console.error(`新しい通帳ファイル名の生成に失敗: ${e.toString()}`);
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

    const filteredTransactions = transactions.filter(tx => {
      const isBroughtForward = (tx.取引内容 || '').includes('繰越');
      const isTransactionEmpty = (Number(tx.入金額) || 0) === 0 && (Number(tx.出金額) || 0) === 0;
      const isDescriptionEmpty = !(tx.取引内容 || '').trim();
      return !(isBroughtForward && isTransactionEmpty) && !(isDescriptionEmpty && isTransactionEmpty);
    });

    let verifiedTransactions = verifyAndCorrectPassbookBalances(filteredTransactions);
    verifiedTransactions = complementMufgBalance_(verifiedTransactions);

    const newRows = verifiedTransactions.map(tx => {
      let isLearned = false;
      let inferred = {};
      const normalizedDescription = normalizeText_(tx.取引内容 || '');

      for (const rule of learningRules) {
        if (rule.storeName !== '') continue;
        
        const keywordMatch = !rule.descriptionKeyword || normalizedDescription.includes(rule.descriptionKeyword);
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
        inferred = inferPassbookAccountTitle(normalizedDescription);
      }
      
      const isDeposit = Number(tx.入金額) > 0;
      let debitTaxCategory = '対象外', creditTaxCategory = '対象外';
      if (isDeposit) creditTaxCategory = inferred.taxCategory;
      else debitTaxCategory = inferred.taxCategory;

      return [
        Utilities.getUuid(), new Date(), tx.取引日, normalizedDescription,
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
    return passbookAccountName;
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
      } else {
        curr.備考 = (curr.備考 || '') + '[【要確認：残高不整合】]';
      }
    }
  }
  return transactions;
}

function complementMufgBalance_(transactions) {
  if (!transactions || transactions.length < 1) return transactions;

  return transactions.map((tx, i, arr) => {
    const isBalanceEmpty = !tx.残高 || Number(tx.残高) === 0;

    if (i > 0 && tx.取引日 === arr[i-1].取引日 && isBalanceEmpty) {
      const prevTx = arr[i-1];
      const prevBalance = Number(prevTx.残高) || 0;
      const deposit = Number(tx.入金額) || 0;
      const withdrawal = Number(tx.出金額) || 0;
      
      const newBalance = prevBalance - withdrawal + deposit;
      tx.残高 = newBalance;
      tx.備考 = (tx.備考 || '') + '[残高印字なし]';
    }
    return tx;
  });
}
