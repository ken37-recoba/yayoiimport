// =================================================================================
// ファイル名: Gemini.gs
// 役割: Gemini APIとの連携に特化した関数を管理します。
// =================================================================================

function getApiKey() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    const error = new Error('APIキーがスクリプトプロパティに設定されていません。');
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
      return { success: false, error: errorMsg, usage: null };
    }
  } catch(e) {
    logError_('callGeminiApi', e, `File Type: ${fileBlob.getContentType()}`);
    return { success: false, error: e.message, usage: null };
  }
}

function getGeminiPrompt(filename) {
  return `
# 指示
この画像から領収書情報を抽出し、指定されたJSON形式で出力してください。
- 1枚の画像に複数の税率が混在する場合、税率ごとに別のオブジェクトを生成してください。
- 日付は西暦 (yyyy/mm/dd) に変換してください。
- 金額は数値のみで出力してください。
- 読み取れない項目は null または 0 としてください。
- 特記事項があれば note に記載してください。（例: 【要確認：日付エラー】）
# JSON形式
\`\`\`json
[
  {
    "date": "2025/06/21",
    "storeName": "株式会社サンプル",
    "description": "品代として",
    "tax_rate": 10,
    "amount": 1100,
    "tax_amount": 100,
    "tax_code": "T1234567890123",
    "filename": "${filename}",
    "note": ""
  }
]
\`\`\``;
}

function inferAccountTitle(storeName, description, amount, masterData) {
  loadConfig_();
  const contextInfo = `Store: ${storeName}, Desc: ${description}, Amount: ${amount}`;
  try {
    const apiKey = getApiKey();
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${CONFIG.GEMINI_MODEL}:generateContent?key=${apiKey}`;

    const masterListWithKeywords = masterData.map(row => ({ title: row[0], keywords: row[1] || "特になし" }));
    const masterTitleList = masterData.map(row => row[0]);

    const prompt = `あなたは日本の経理専門家です。与えられた情報と勘定科目マスターを基に、最も可能性の高い勘定科目を1つだけJSON形式で返してください。\n# 領収書情報\n- 店名: ${storeName}\n- 摘要: ${description}\n- 金額(税込): ${amount}円\n# 勘定科目マスター\n${JSON.stringify(masterListWithKeywords)}`;

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
    }
    logError_('inferAccountTitle', new Error(`API Error ${responseCode}: ${responseBody}`), contextInfo);
    return "【推測エラー】";
  } catch (e) {
    logError_('inferAccountTitle', e, contextInfo);
    return "【推測エラー】";
  }
}

function callPassbookGeminiApi(fileBlob, bankType) {
    const prompt = getPassbookGeminiPrompt(bankType);
    return callGeminiApi(fileBlob, prompt);
}

function getPassbookGeminiPrompt(bankType) {
    const basePrompt = `# 指示\n提供された通帳の画像から取引履歴を正確に抽出し、以下のJSON形式の配列として結果を返してください。\n- 金額は必ず**数値(Number)型**で出力してください。\n- 日付は必ず**'yyyy-mm-dd'形式の西暦文字列**に統一してください。\n# JSON出力形式\n\`\`\`json\n[\n  {\n    "取引日": "yyyy-mm-dd",\n    "出金額": 0,\n    "入金額": 50000,\n    "残高": 1050000,\n    "取引内容": "振込 タナカ タロウ",\n    "備考": ""\n  }\n]\n\`\`\``;
    
    let bankSpecificInstructions = '';
    if (bankType === 'MUFG') {
        bankSpecificInstructions = `\n# 三菱UFJ銀行の特別ルール\n- 日付は \`年-月日\` の形式です。例: \`07-428\` は令和7年4月28日です。\n- 「お支払金額」列は必ず『出金額』、「お預り金額」列は必ず『入金額』としてください。`;
    } else if (bankType === 'OSAKA_SHINKIN') {
        bankSpecificInstructions = `\n# 大阪信用金庫の特別ルール\n- 「差引残高」がアスタリスク(***)の行は、その直前の行の「摘要」の続きです。その行の摘要を直前の行の取引内容に連結し、アスタリスクの行自体は出力しないでください。`;
    }

    return basePrompt + bankSpecificInstructions;
}

function inferPassbookAccountTitle(description) {
  loadConfig_();
  const contextInfo = `Passbook Desc: ${description}`;
  try {
    const apiKey = getApiKey();
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${CONFIG.GEMINI_MODEL}:generateContent?key=${apiKey}`;
    const masterData = getMasterData();
    const masterTitleList = masterData.map(row => row[0]);

    const prompt = `あなたは日本の経理専門家です。以下の「摘要」に最も適した「勘定科目」「補助科目」「標準税区分」をJSONで返してください。\n# 摘要\n${description}\n# 勘定科目マスター\n${JSON.stringify(masterData)}`;

    const payload = {
      "contents": [{"parts": [{ "text": prompt }]}],
      "generationConfig": {
        "responseMimeType": "application/json",
        "temperature": 0.1,
        "responseSchema": {
          "type": "OBJECT",
          "properties": {
            "accountTitle": { "type": "STRING", "enum": masterTitleList },
            "subAccount": { "type": "STRING" },
            "taxCategory": { "type": "STRING" }
          },
          "required": ["accountTitle", "subAccount", "taxCategory"]
        }
      }
    };

    const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true };
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      const inferredText = JSON.parse(responseBody).candidates?.[0]?.content?.parts?.[0]?.text;
      if (inferredText) return JSON.parse(inferredText);
    }
    
    logError_('inferPassbookAccountTitle', new Error(`API Error ${responseCode}: ${responseBody}`), contextInfo);
    return { accountTitle: '【推測エラー】', subAccount: '', taxCategory: '対象外' };
  } catch (e) {
    logError_('inferPassbookAccountTitle', e, contextInfo);
    return { accountTitle: '【推測エラー】', subAccount: '', taxCategory: '対象外' };
  }
}
