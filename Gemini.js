// =================================================================================
// ファイル名: Gemini.gs
// 役割: Gemini APIとの連携に特化した関数を管理します。
// =================================================================================

function getApiKey() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    const error = new Error('APIキーがスクリプトプロパティに設定されていません。プロジェクトの設定を確認してください。');
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
      return { success: false, error: errorMsg };
    }
  } catch(e) {
    logError_('callGeminiApi', e, `File Type: ${fileBlob.getContentType()}`);
    return { success: false, error: e.message };
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
      note: string; // 【重要】特記事項。以下の special_note_rules に従い、問題点を具体的に記述する。
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
  - rule: 領収書が白紙、または文字が著しく不鮮明で内容が全く読み取れない場合。
    note_content: "【要確認：読み取り不可】"
  - rule: dateが未来の日付、または暦上存在しない日付（例: 2月31日）になっている場合。
    note_content: "【要確認：日付エラー】"
  - rule: amountやtax_amountがマイナス、または消費税額が合計金額を上回るなど、金額の計算に矛盾がある場合。
    note_content: "【要確認：金額不整合】"
  - rule: amountまたはdateが読み取れずnullになった場合。
    note_content: "【要確認：必須項目欠落】"
  - rule: tax_codeがTで始まるが13桁ではないなど、形式が明らかに不正な場合。
    note_content: "【要確認：登録番号形式エラー】"
  - rule: 上記以外で軽微な懸念がある場合（例: 一部の文字が不鮮明）。
    note_content: "印字不鮮明"
`;
}

function inferAccountTitle(storeName, description, amount, masterData) {
  loadConfig_();
  const contextInfo = `Store: ${storeName}, Desc: ${description}, Amount: ${amount}`;
  try {
    const apiKey = getApiKey();
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${CONFIG.GEMINI_MODEL}:generateContent?key=${apiKey}`;

    const masterListWithKeywords = masterData.map(row => ({ title: row[0], keywords: row[1] || "特になし" }));
    const masterTitleList = masterData.map(row => row[0]);

    const prompt = `あなたは、日本の会計基準に精通したベテランの経理専門家です。あなたの任務は、与えられた領収書の情報と、社内ルールを含む勘定科目マスターを基に、最も可能性の高い勘定科目を特定することです。# 指示\n1. 以下の「領収書情報」と「勘定科目マスター」を注意深く分析してください。\n2. 特に「勘定科目マスター」の**キーワード/ルール**は重要です。例えば、「2万円未満の飲食代は会議費」といった金額に基づくルールが含まれている場合があります。\n3. すべての情報を総合的に判断し、「勘定科目マスター」のリストの中から最も適切だと考えられる勘定科目を**1つだけ**選択してください。\n4. あなたの回答は、必ず指定されたJSON形式に従ってください。# 領収書情報\n- 店名: ${storeName}\n- 摘要: ${description}\n- 金額(税込): ${amount}円\n# 勘定科目マスター（キーワード/ルールを含む）\n${JSON.stringify(masterListWithKeywords)}`;

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
      const errorMsg = "AIからのJSONレスポンスの形式が不正です。";
      logError_('inferAccountTitle', new Error(errorMsg), `${contextInfo}, Response: ${responseBody}`);
      return "【形式エラー】";
    } else {
      const errorMsg = `勘定科目推測APIエラー [${responseCode}]: ${responseBody}`;
      logError_('inferAccountTitle', new Error(errorMsg), contextInfo);
      return `【APIエラー ${responseCode}】`;
    }
  } catch (e) {
    logError_('inferAccountTitle', e, contextInfo);
    return "【推測エラー】";
  }
}
