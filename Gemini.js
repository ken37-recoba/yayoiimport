// =================================================================================
// ファイル名: Gemini.js (記号誤認識 対策版)
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
        "temperature": 0.1
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

// ▼▼▼【改善箇所】AIへの指示に通貨記号の無視と金額の妥当性チェックを追加 ▼▼▼
function getGeminiPrompt(filename) {
  const today = Utilities.formatDate(new Date(), "JST", "yyyy-MM-dd");
  return `
# 指示
この画像から領収書情報を抽出し、指定されたJSON形式で出力してください。
- **【金額の読取りルール】**
  - 金額の前に手書きまたは印字の「￥」や「¥」マークがある場合、**これらの記号は形が数字の「7」に似ていても、絶対に数字として解釈せず、完全に無視してください。**
  - 最終的な支払額は、「合計」「ご請求額」「お会計」「総額」といったキーワードの近くに記載されていることが多いです。これらのキーワードを探し、最も下に記載されている最大の金額を支払額としてください。
  - 「小計」や「税抜合計」と「合計」の両方がある場合、**必ず「合計」と書かれた隣の金額を amount に採用してください**。
  - 抽出した金額が、取引内容（例：食事代）に対して常識的な範囲か考慮してください。喫茶店の食事代で7万円を超えるなど、不自然に高額な場合は、記号の誤認識を疑って再確認してください。
- **【日付の読取りルール】**
  - 現在の日付は ${today} です。この情報を参考に、領収書の日付の年（西暦）を正しく判断してください。特に「令和7年」や「7年」のような和暦や年が省略されている場合は、現在の日付に最も近い過去の日付となるように西暦を決定してください。
- **【その他のルール】**
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
// ▲▲▲ 改善箇所 ▲▲▲

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
あなたは、日本の銀行通帳の読取りに特化した高精度OCRエンジンです。あなたの任務は、提供された通帳の画像から取引履歴を正確に抽出し、指定されたJSON形式の**配列**として結果を返すことです。

# 全体ルール
- 「お支払金額」列に記載された数値は**必ず『出金額』**としてください。
- 「お預り金額」列に記載された数値は**必ず『入金額』**としてください。
- 金額は必ず**数値(Number)型**で出力してください。通貨記号(¥)やカンマ(,)は含めないでください。
- 日付は必ず**'yyyy-mm-dd'形式の西暦文字列**に統一してください。

# 文字整形ルール
- **カタカナ:** 半角カタカナは全角カタカナに変換してください。ひらがなや漢字は変換しないでください。
- **英数字:** すべて半角で出力してください。
- **濁点・半濁点の結合:** 濁点(ﾞ)や半濁点(ﾟ)は、前の文字と結合し、必ず1文字として表現してください。

# 日付の解釈ルール
- 通帳の年は「7」や「07」のように2桁で記載されている場合があります。これは和暦の「令和」を指します。
- **現在の年は西暦${currentYear}年（令和${currentReiwaYear}年）です。**
- したがって、通帳に「7」と記載があれば、それは「令和7年」を意味し、西暦では「${reiwaStartYear - 1 + 7}年」となります。必ずこのルールに従って西暦へ変換してください。
- **必ず、処理実行日に最も近い過去の日付になるように西暦を判断してください。**

# 除外ルール
- 「摘要」が「繰越」や「繰越残高」となっており、かつ「お支払金額」と「お預り金額」の両方が0の行は、出力に**含めないでください**。
- 同様に、「摘要」が完全に空欄で、かつ「お支払金額」と「お預り金額」の両方が0の行も、出力に**含めないでください**。

# JSON出力形式
\`\`\`json
[
  {
    "取引日": "yyyy-mm-dd",
    "出金額": 0,
    "入金額": 50000,
    "残高": 1050000,
    "取引内容": "振込 タナカ タロウ",
    "備考": ""
  }
]
\`\`\`
`;
    
    let bankSpecificInstructions = '';
    if (bankType === 'MUFG') {
        bankSpecificInstructions = `
# 三菱UFJ銀行の特別ルール
- **【最重要】列の定義:** この通帳には「お支払金額」と「お預り金額」という2つの金額列があります。
  1. まず、画像の上部にある「お支払金額」というヘッダーを探してください。この列にある数値は**すべて『出金額』**です。
  2. 次に、「お預り金額」というヘッダーを探してください。この列にある数値は**すべて『入金額』**です。
- **【絶対厳守】:** このルールは絶対です。残高の増減など、他の情報から類推してこのルールを曲げてはいけません。たとえ残高計算が合わないように見えても、列の物理的な位置を最優先してください。物理的に「お支払金額」の列にある数字は、必ずJSONの \`出金額\` に入れてください。
- **残高の扱い:** 同日内の複数取引において、2行目以降の「残高」が空欄の場合があります。その場合は**無理に数字を読み取らず、JSONの\`残高\`フィールドを\`0\`または\`null\`として出力**してください。
- **日付形式:** 日付は \`年-月日\` の形式です。例: \`07-428\` は令和7年4月28日です。
`;
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
