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
    const currentYear = new Date().getFullYear();
    const reiwaStartYear = 2019;
    const currentReiwaYear = currentYear - reiwaStartYear + 1;

    const basePrompt = `
# 役割
あなたは、日本の銀行通帳の読取りに特化した高精度OCRエンジンです。あなたの任務は、提供された通帳の画像から取引履歴を正確に抽出し、指定されたJSON形式の配列として結果を返すことです。

# 全体ルール
- 画像から読み取れるすべての取引行を、1行も漏らさず抽出してください。
- 金額は必ず**数値(Number)型**で出力してください。通貨記号(¥)やカンマ(,)は含めないでください。
- 日付は必ず**'yyyy-mm-dd'形式の西暦文字列**に統一してください。

# 手書き文字の扱い
- 通帳には手書きのメモや摘要が追記されている場合があります。印刷された文字だけでなく、これらの**手書き文字も読み取り対象とし、取引内容に含めてください。**

# 日付の解釈ルール (最重要)
- 通帳の年は「7」や「07」のように2桁で記載されている場合があります。これは和暦の「令和」を指します。
- **現在の年は西暦${currentYear}年（令和${currentReiwaYear}年）です。**
- したがって、通帳に「7」と記載があれば、それは「令和7年」を意味し、西暦では「${reiwaStartYear - 1 + 7}年」となります。必ずこのルールに従って西暦へ変換してください。
- **必ず、処理実行日に最も近い過去の日付になるように西暦を判断してください。** 例えば「平成7年(1995年)」や「昭和7年(1932年)」のように、古すぎる年として解釈しないでください。

# 除外ルール (重要)
- 「摘要」が「繰越」や「繰越残高」となっており、かつ「お支払金額」と「お預り金額」の両方が空欄または0の行は、実際の取引ではないため、JSONの出力に**含めないでください**。
- 同様に、「摘要」が完全に空欄で、かつ「お支払金額」と「お預り金額」の両方が0の行も、ページ先頭の繰越残高行とみなし、JSON出力に**含めないでください**。

# 特記事項ルール (重要)
- 日付や金額の読み取りが困難な場合や、残高の計算に矛盾がある場合は、その内容を「備考」欄に「【要確認：金額不整合】」のように具体的に記載してください。
- 何らかの理由で必須項目（取引日、金額など）が読み取れなかった場合も、「備考」欄に「【要確認：必須項目欠落】」と記録してください。

# JSON出力形式
\`\`\`json
[
  {
    "取引日": "yyyy-mm-dd",
    "出金額": 0,
    "入金額": 50000,
    "残高": 1050000,
    "取引内容": "振込 タナカ タロウ（手書きメモ）",
    "備考": ""
  }
]
\`\`\`
`;
    
    let bankSpecificInstructions = '';
    if (bankType === 'MUFG') {
        bankSpecificInstructions = `
# 三菱UFJ銀行の特別ルール (最優先事項)
- **入出金の厳格なルール:** 通帳の「お支払金額」列にある数値は**絶対に『出金額』**としてください。「お預り金額」列にある数値は**絶対に『入金額』**としてください。この列の位置に基づくルールは、他のどの指示よりも優先されます。
- **残高の扱い:** 同日内の複数取引において、2行目以降の「残高」が空欄の場合があります。その場合は**無理に数字を読み取らず、JSONの\`残高\`フィールドを\`0\`または\`null\`として出力**してください。これは意図した動作なので、この件について備考欄に【要確認】と記載する必要はありません。
- **日付形式:** 日付は \`年-月日\` の形式です。例: \`07-428\` は令和7年4月28日です。
`;
    } else if (bankType === 'OSAKA_SHINKIN') {
        bankSpecificInstructions = `\n# 大阪信用金庫の特別ルール\n- 「差引残高」がアスタリスク(***)のみで埋められている行は、その直前の行の「摘要」の続きです。その行の摘要を直前の行の取引内容に連結し、アスタリスクの行自体は出力しないでください。`;
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

    const prompt = `あなたは日本の経理専門家です。以下の「摘要」に最も適した「勘定科目」および「標準税区分」を推測してください。補助科目は推測せず、JSONにも含めないでください。\n\n# 摘要\n${description}\n\n# 勘定科目マスター\n${JSON.stringify(masterData)}`;

    const payload = {
      "contents": [{"parts": [{ "text": prompt }]}],
      "generationConfig": {
        "responseMimeType": "application/json",
        "temperature": 0.1,
        "responseSchema": {
          "type": "OBJECT",
          "properties": {
            "accountTitle": { "type": "STRING", "enum": masterTitleList },
            "taxCategory": { "type": "STRING" }
          },
          "required": ["accountTitle", "taxCategory"]
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
        const result = JSON.parse(inferredText);
        return { 
          accountTitle: result.accountTitle, 
          subAccount: '',
          taxCategory: result.taxCategory 
        };
      }
    }
    
    logError_('inferPassbookAccountTitle', new Error(`API Error ${responseCode}: ${responseBody}`), contextInfo);
    return { accountTitle: '【推測エラー】', subAccount: '', taxCategory: '対象外' };
  } catch (e) {
    logError_('inferPassbookAccountTitle', e, contextInfo);
    return { accountTitle: '【推測エラー】', subAccount: '', taxCategory: '対象外' };
  }
}
