// =================================================================================
// ファイル名: Setup.gs
// 役割: 環境の初期化やトリガー設定など、システムのセットアップに関する関数を管理します。
// =================================================================================

function initializeEnvironment() {
  loadConfig_();
  console.log('環境の初期化を確認・実行します...');
  
  createSheetWithHeaders(CONFIG.FILE_LIST_SHEET, CONFIG.HEADERS.FILE_LIST);
  createSheetWithHeaders(CONFIG.OCR_RESULT_SHEET, CONFIG.HEADERS.OCR_RESULT, true);
  createSheetWithHeaders(CONFIG.EXPORTED_SHEET, CONFIG.HEADERS.EXPORTED, true);
  
  createSheetWithHeaders(CONFIG.PASSBOOK_FILE_LIST_SHEET, CONFIG.HEADERS.PASSBOOK_FILE_LIST);
  createSheetWithHeaders(CONFIG.PASSBOOK_RESULT_SHEET, CONFIG.HEADERS.PASSBOOK_RESULT, true);
  createSheetWithHeaders(CONFIG.PASSBOOK_EXPORTED_SHEET, CONFIG.HEADERS.PASSBOOK_EXPORTED, true);
  
  createSheetWithHeaders(CONFIG.PASSBOOK_MASTER_SHEET, CONFIG.HEADERS.PASSBOOK_MASTER);
  createSheetWithHeaders(CONFIG.LEARNING_SHEET, CONFIG.HEADERS.LEARNING);
  
  createSheetWithHeaders(CONFIG.TOKEN_LOG_SHEET, CONFIG.HEADERS.TOKEN_LOG);
  createSheetWithHeaders(CONFIG.MASTER_SHEET, ['勘定科目', 'キーワード/ルール']); 
  createSheetWithHeaders(CONFIG.CONFIG_SHEET, []); 
  createSheetWithHeaders(CONFIG.ERROR_LOG_SHEET, CONFIG.HEADERS.ERROR_LOG);
  console.log('環境の初期化が完了しました。');
}

function createTimeBasedTrigger_() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
    console.log(`既存のトリガーを ${triggers.length} 件削除しました。`);

    ScriptApp.newTrigger('mainProcess')
      .timeBased()
      .everyMinutes(15)
      .create();
    
    ScriptApp.newTrigger('mainProcessPassbooks')
      .timeBased()
      .everyMinutes(15)
      .create();
    
    ui.alert('設定完了', '「領収書処理」と「通帳処理」をそれぞれ15分ごとに自動実行する設定が完了しました。', ui.ButtonSet.OK);
  
  } catch (e) {
    logError_('createTimeBasedTrigger_', e);
    ui.alert('トリガーの作成に失敗しました。\n\nスクリプトの実行権限を許可する必要があるかもしれません。\n詳細: ' + e.message);
  }
}
