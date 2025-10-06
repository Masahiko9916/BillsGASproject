/**
 * Main.gs V4.1 - メインエントリーポイント
 * - 新しいメニュー体系と操作ロジックに全面刷新
 */

/**
 * GASメニュー作成（スプレッドシート起動時に自動実行）
 */
function onOpen() {
  const ss = SpreadsheetApp.getActive();
  if (ss.getId() !== CONFIG.regSpreadsheetId) return;
  
  const ui = SpreadsheetApp.getUi();

  // メイン操作メニュー
  ui.createMenu('【メイン操作】')
    .addItem('✅ カレンダーに登録する', 'menuRequestRegister')
    .addItem('🔄 最新の状態にする（顧客情報補完 & 同期）', 'menuRequestUpdate')
    .addItem('❌ 予定を取り消す', 'menuRequestCancel')
    .addSeparator()
    .addItem('⏸️ 保留にする', 'menuRequestHold')
    .addToUi();
    
  // 開発ツールメニュー
  ui.createMenu('【開発ツール】')
    .addItem('登録シートの初期セットアップ', 'setupRegSheet_')
    .addItem('スポット回収シートの初期セットアップ', 'setupSpotSheets_')
    .addSeparator()
    .addItem('⭐ カレンダー権限を一括付与', 'grantCalendarAccessToAdmin')
    .addItem('アクセス可能なカレンダーを確認', 'listAccessibleCalendars')
    .addSeparator()
    .addItem('⭐ 自動処理用トリガー作成（5分おき）', 'installScheduledTasksTrigger')
    .addItem('ポーリング用トリガー作成（5分おき）', 'installTimeTrigger')
    .addToUi();
}

// ========================================
// 新しいメニュー関数
// ========================================

/**
 * 選択行を「保留中」ステータスに変更します
 */
function menuRequestHold() {
  const sh = SpreadsheetApp.getActiveSheet();
  const row = sh.getActiveRange().getRow();
  if (row === 1) return SpreadsheetApp.getUi().alert('データ行を選択してください。');

  setCell_(sh, row, 'ステータス', STATUS.HOLD);
  setLastUpdated_(sh, row, 'sheet');
  SpreadsheetApp.getActive().toast('ステータスを「保留中」に変更しました。');
}

/**
 * 選択行のカレンダー登録を予約します（事前チェック付き）
 */
function menuRequestRegister() {
  const sh = SpreadsheetApp.getActiveSheet();
  const row = sh.getActiveRange().getRow();
  if (row === 1) return SpreadsheetApp.getUi().alert('データ行を選択してください。');

  try {
    // --- 事前チェック ---
    const d = getCell_(sh, row, '回収予定日');
    const s = getCell_(sh, row, '開始時間');
    const e = getCell_(sh, row, '終了時間');
    const assignee = nz_(getCell_(sh, row, '回収担当者'));
    const area = nz_(getCell_(sh, row, 'エリア'));

    if (!d || !s || !e || !assignee || !area) {
      throw new Error('「回収予定日」「開始時間」「終了時間」「回収担当者」「エリア」は必須です。');
    }
    // 定期回収シートの場合のみ休診日チェックを実行
    if (sh.getName() === CONFIG.regSheetName && !validateAvailability_(sh, row, d, s, e)) {
      throw new Error('回収不可能日（休診日など）に設定されているため、登録できません。');
    }
    const eventId = getCell_(sh, row, 'カレンダーイベントID');
    if (eventId) {
      throw new Error('既にカレンダーに登録済みです。内容を更新する場合は「最新の状態にする」を選択してください。');
    }
  } catch (err) {
    SpreadsheetApp.getUi().alert('登録エラー', err.message, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // ステータスを設定
  setCell_(sh, row, 'ステータス', STATUS.CALENDAR_REGISTER);
  setLastUpdated_(sh, row, 'sheet');
  SpreadsheetApp.getUi().alert('✅ 受付完了', 'カレンダーへの登録を予約しました。\n数分以内に自動で反映されます。', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * 顧客情報を補完し、カレンダーの更新を予約します
 */
function menuRequestUpdate() {
  const sh = SpreadsheetApp.getActiveSheet();
  const row = sh.getActiveRange().getRow();
  if (row === 1) return SpreadsheetApp.getUi().alert('データ行を選択してください。');

  let message = '';
  let customerInfoApplied = false;

  try {
    // 顧客情報・ディーラー情報の補完を実行
    const sheetName = sh.getName();
    if (sheetName === CONFIG.regSheetName || sheetName === SPOT_CONFIG.sheets.clinic) {
      customerInfoApplied = applyCustomerMasterToRegRow_(sh, row);
    } else if (sheetName === SPOT_CONFIG.sheets.dealer) {
      const customerApplied = applyCustomerMasterToRegRow_(sh, row);
      const dealerApplied = applyDealerMasterToRegRow_(sh, row);
      customerInfoApplied = customerApplied || dealerApplied;
    }
    
    if (customerInfoApplied) {
      message += '顧客・ディーラー情報を補完しました。\n';
    } else {
      message += '補完する顧客・ディーラー情報はありませんでした。\n';
    }

  } catch (e) {
    console.error('顧客情報補完エラー:', e);
    SpreadsheetApp.getUi().alert('顧客情報の補完中にエラーが発生しました: ' + e.message);
    return;
  }

  // カレンダーに登録済みの場合は、更新待ちステータスにする
  const eventId = getCell_(sh, row, 'カレンダーイベントID');
  if (eventId) {
    setCell_(sh, row, 'ステータス', STATUS.RESYNC_REGISTER);
    setLastUpdated_(sh, row, 'sheet');
    message += 'カレンダーの更新を予約しました。\n数分以内に自動で反映されます。';
  } else {
    message += 'カレンダーには未登録です。登録する場合は「カレンダーに登録する」を実行してください。';
  }
  
  SpreadsheetApp.getUi().alert('✅ 処理完了', message, SpreadsheetApp.getUi().ButtonSet.OK);
}


/**
 * 選択行のカレンダーからの削除を予約します
 */
function menuRequestCancel() {
  const sh = SpreadsheetApp.getActiveSheet();
  const row = sh.getActiveRange().getRow();
  if (row === 1) return SpreadsheetApp.getUi().alert('データ行を選択してください。');

  const eventId = getCell_(sh, row, 'カレンダーイベントID');
  if (!eventId) {
    // 既に削除済み、または未登録の場合
    const currentStatus = getCell_(sh, row, 'ステータス');
    if (currentStatus !== STATUS.CANCEL_COMPLETE) {
      setCell_(sh, row, 'ステータス', STATUS.CANCEL_COMPLETE);
      setLastUpdated_(sh, row, 'sheet');
    }
    SpreadsheetApp.getActive().toast('この予定はカレンダーに存在しないため、ステータスを「取消済み」にしました。');
    return;
  }

  setCell_(sh, row, 'ステータス', STATUS.CANCEL_REGISTER);
  setLastUpdated_(sh, row, 'sheet');
  SpreadsheetApp.getUi().alert('✅ 受付完了', 'カレンダーからの削除を予約しました。\n数分以内に自動で反映されます。', SpreadsheetApp.getUi().ButtonSet.OK);
}


// ========================================
// トリガー関数
// ========================================

/**
 * セル編集時トリガー（担当者変更のみを検知）
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;

    const sh = e.range.getSheet();
    const row = e.range.getRow();
    const sheetName = sh.getName();

    const validSheets = [CONFIG.regSheetName, SPOT_CONFIG.sheets.clinic, SPOT_CONFIG.sheets.dealer];
    if (row === 1 || validSheets.indexOf(sheetName) === -1) {
      return;
    }

    const editedCol = e.range.getColumn();
    const assigneeCol = colByHeader_(sh, '回収担当者');

    if (assigneeCol && editedCol === assigneeCol) {
      const eventId = getCell_(sh, row, 'カレンダーイベントID');
      if (eventId) {
        setCell_(sh, row, 'ステータス', STATUS.ASSIGNEE_CHANGE_REGISTER);
        setLastUpdated_(sh, row, 'sheet');
        SpreadsheetApp.getActive().toast('担当者を変更しました。カレンダーへは数分以内に自動で反映されます。', '情報', 10);
      }
    }
  } catch (err) { 
    console.error('onEdit error:', err); 
  }
}

/**
 * フォーム送信時トリガー（自動実行）
 */
function onFormSubmit(e) {
  try {
    let formId = '';
    if (e && e.source && typeof e.source.getId === 'function') {
      formId = e.source.getId();
    }
    
    let targetSheetName, formType;
    
    if (formId === SPOT_CONFIG.forms.clinic) {
      targetSheetName = SPOT_CONFIG.sheets.clinic;
      formType = 'spot_clinic';
    } else if (formId === SPOT_CONFIG.forms.dealer) {
      targetSheetName = SPOT_CONFIG.sheets.dealer;
      formType = 'spot_dealer';
    } else {
      targetSheetName = CONFIG.regSheetName;
      formType = 'regular';
    }
    
    processFormSubmission_(e, targetSheetName, formType);
    
  } catch (err) {
    postChat_('エラー', '受付処理でエラー: ' + err.message, 'regular');
    throw err;
  }
}

/**
 * フォーム送信処理の共通ロジック
 */
function processFormSubmission_(e, targetSheetName, formType) {
  let srcHeaders = [], srcValues = [];
  if (e && e.range && typeof e.range.getSheet === 'function') {
    const srcSh = e.range.getSheet();
    const lastCol = srcSh.getLastColumn();
    srcHeaders = srcSh.getRange(1, 1, 1, lastCol).getValues()[0];
    srcValues = srcSh.getRange(e.range.getRow(), 1, 1, lastCol).getValues()[0];
  } else if (e && e.namedValues) {
    srcHeaders = Object.keys(e.namedValues);
    srcValues = srcHeaders.map(h =>
      Array.isArray(e.namedValues[h]) ? e.namedValues[h].join(' ') : (e.namedValues[h] ?? '')
    );
  } else {
    throw new Error('onFormSubmit: イベントから回答が取得できません。');
  }
  
  const ssReg = SpreadsheetApp.openById(CONFIG.regSpreadsheetId);
  const shReg = ssReg.getSheetByName(targetSheetName);
  if (!shReg) {
    throw new Error('シートが見つかりません: ' + targetSheetName);
  }
  
  const dstHeaders = sheetHeaders_(shReg);
  let rowArr, isFirstUse;
  
  if (formType === 'spot_clinic') {
    rowArr = buildSpotClinicRowByMapping_(dstHeaders, srcHeaders, srcValues);
    isFirstUse = detectFirstUseFromSubmission_(srcHeaders, srcValues);
  } else if (formType === 'spot_dealer') {
    rowArr = buildSpotDealerRowByMapping_(dstHeaders, srcHeaders, srcValues);
    isFirstUse = detectDealerFirstUse_(srcHeaders, srcValues);
  } else {
    rowArr = buildRowByMapping_(dstHeaders, srcHeaders, srcValues);
    isFirstUse = detectFirstUseFromSubmission_(srcHeaders, srcValues);
  }
  
  shReg.appendRow(rowArr);
  const lastRow = shReg.getLastRow();
  
  // 電話番号の文字列固定化
  if (formType === 'spot_clinic') {
    const telIdx1 = dstHeaders.indexOf('医院電話番号');
    if (telIdx1 >= 0) {
      const tel1 = pickValueByAliases_('医院電話番号', srcHeaders, srcValues, SPOT_CLINIC_ALIASES, {});
      forceTelAsText_(shReg, lastRow, '医院電話番号', tel1);
    }
    const telIdx2 = dstHeaders.indexOf('代表電話番号');
    if (telIdx2 >= 0) {
      const tel2 = pickValueByAliases_('代表電話番号', srcHeaders, srcValues, SPOT_CLINIC_ALIASES, {});
      forceTelAsText_(shReg, lastRow, '代表電話番号', tel2);
    }
  } else if (formType === 'spot_dealer') {
    const telIdx1 = dstHeaders.indexOf('医院電話番号');
    if (telIdx1 >= 0) {
      const tel1 = pickValueByAliases_('医院電話番号', srcHeaders, srcValues, SPOT_DEALER_ALIASES, {});
      forceTelAsText_(shReg, lastRow, '医院電話番号', tel1);
    }
    const telIdx2 = dstHeaders.indexOf('ディーラー電話番号');
    if (telIdx2 >= 0) {
      const tel2 = pickValueByAliases_('ディーラー電話番号', srcHeaders, srcValues, SPOT_DEALER_ALIASES, {});
      forceTelAsText_(shReg, lastRow, 'ディーラー電話番号', tel2);
    }
  } else {
    const telIdx = dstHeaders.indexOf('電話番号');
    if (telIdx >= 0) {
      const mappedTel = pickOriginalByAliases_('電話番号', srcHeaders, srcValues);
      forceTelAsText_(shReg, lastRow, '電話番号', mappedTel);
    }
  }
  
  const newId = getCell_(shReg, lastRow, '受付ID');
  
  sortRegSheetIfNeeded_(shReg);
  
  const r = findRowByColumnValue_(shReg, '受付ID', newId) || shReg.getLastRow();
  
  // 顧客マスタ処理
  try {
    if (formType === 'spot_clinic') {
      handleCustomerMasterOnSubmit_(shReg, r, isFirstUse);
    } else if (formType === 'spot_dealer') {
      handleDealerMasterOnSubmit_(shReg, r, isFirstUse);
      handleCustomerMasterOnSubmit_(shReg, r, false);
    } else {
      handleCustomerMasterOnSubmit_(shReg, r, isFirstUse);
    }
  } catch (mfErr) {
    console.error('顧客マスタ処理エラー:', mfErr);
  }
  
  // 行URL付与
  const rowUrl = `${ssReg.getUrl()}#gid=${shReg.getSheetId()}&range=${r}:${r}`;
  setCell_(shReg, r, '行URL', rowUrl);
  
  // 受付通知
  let summary;
  if (formType === 'spot_clinic') {
    summary = buildSpotClinicReceiptSummary_(shReg, r);
    postChat_('受付', summary + '\n' + rowUrl, 'spot_clinic');
  } else if (formType === 'spot_dealer') {
    summary = buildSpotDealerReceiptSummary_(shReg, r);
    postChat_('受付', summary + '\n' + rowUrl, 'spot_dealer');
  } else {
    summary = buildReceiptSummary_(shReg, r);
    postChat_('受付', summary + '\n' + rowUrl, 'regular');
  }
}


// ========================================
// 開発ツールメニュー関数
// ========================================

function setupRegSheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.regSpreadsheetId);
  let sh = ss.getSheetByName(CONFIG.regSheetName);
  if (!sh) sh = ss.insertSheet(CONFIG.regSheetName);

  // ★★★ 修正箇所 ★★★
  const expected = MGMT_HEADERS.concat(FORM_HEADERS).concat(SYS_HEADERS_MAIN).concat(CALENDAR_HEADERS);
  const curHeaders = sh.getRange(1,1,1,Math.max(1, sh.getLastColumn())).getValues()[0]
    .map(v => String(v).trim()).filter(Boolean);

  if (curHeaders.length === 0) {
    sh.getRange(1,1,1,expected.length).setValues([expected]);
    sh.setFrozenRows(1);
  } else {
    const setCur = new Set(curHeaders);
    let lastCol = sh.getLastColumn();
    expected.forEach(h => {
      if (!setCur.has(h)) {
        lastCol += 1;
        sh.getRange(1, lastCol).setValue(h);
      }
    });
    if (sh.getFrozenRows() === 0) sh.setFrozenRows(1);
  }

  setupDataValidation_(sh);
  sh.setFrozenColumns(MGMT_HEADERS.length);
  SpreadsheetApp.getUi().alert('定期回収シートのセットアップが完了しました。');
}

function setupSpotSheets_() {
  const ss = SpreadsheetApp.openById(CONFIG.regSpreadsheetId);
  
  // スポット医院シート
  let clinicSh = ss.getSheetByName(SPOT_CONFIG.sheets.clinic);
  if (!clinicSh) clinicSh = ss.insertSheet(SPOT_CONFIG.sheets.clinic);
  
  // ★★★ 修正箇所 ★★★
  const clinicExpected = MGMT_HEADERS.concat(SPOT_CLINIC_HEADERS).concat(SYS_HEADERS_MAIN).concat(CALENDAR_HEADERS);
  const clinicCurHeaders = clinicSh.getRange(1,1,1,Math.max(1, clinicSh.getLastColumn())).getValues()[0]
    .map(v => String(v).trim()).filter(Boolean);
  
  if (clinicCurHeaders.length === 0) {
    clinicSh.getRange(1,1,1,clinicExpected.length).setValues([clinicExpected]);
    clinicSh.setFrozenRows(1);
  } else {
    const setCur = new Set(clinicCurHeaders);
    let lastCol = clinicSh.getLastColumn();
    clinicExpected.forEach(h => {
      if (!setCur.has(h)) {
        lastCol += 1;
        clinicSh.getRange(1, lastCol).setValue(h);
      }
    });
    if (clinicSh.getFrozenRows() === 0) clinicSh.setFrozenRows(1);
  }
  clinicSh.setFrozenColumns(MGMT_HEADERS.length);
  setupDataValidation_(clinicSh);
  
  // スポットディーラーシート
  let dealerSh = ss.getSheetByName(SPOT_CONFIG.sheets.dealer);
  if (!dealerSh) dealerSh = ss.insertSheet(SPOT_CONFIG.sheets.dealer);
  
  // ★★★ 修正箇所 ★★★
  const dealerExpected = MGMT_HEADERS.concat(SPOT_DEALER_HEADERS).concat(SYS_HEADERS_MAIN).concat(CALENDAR_HEADERS);
  const dealerCurHeaders = dealerSh.getRange(1,1,1,Math.max(1, dealerSh.getLastColumn())).getValues()[0]
    .map(v => String(v).trim()).filter(Boolean);
  
  if (dealerCurHeaders.length === 0) {
    dealerSh.getRange(1,1,1,dealerExpected.length).setValues([dealerExpected]);
    dealerSh.setFrozenRows(1);
  } else {
    const setCur = new Set(dealerCurHeaders);
    let lastCol = dealerSh.getLastColumn();
    dealerExpected.forEach(h => {
      if (!setCur.has(h)) {
        lastCol += 1;
        dealerSh.getRange(1, lastCol).setValue(h);
      }
    });
    if (dealerSh.getFrozenRows() === 0) dealerSh.setFrozenRows(1);
  }
  dealerSh.setFrozenColumns(MGMT_HEADERS.length);
  setupDataValidation_(dealerSh);
  
  SpreadsheetApp.getUi().alert('スポット回収シートのセットアップが完了しました。');
}

/**
 * データ検証ルールの設定
 */
function setupDataValidation_(sh) {
  const statusCol = colByHeader_(sh,'ステータス');
  if (statusCol) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(Object.values(STATUS), true)
      .build();
    sh.getRange(2,statusCol,Math.max(1, sh.getMaxRows()-1),1).setDataValidation(rule);
  }

  const areaCol = colByHeader_(sh,'エリア');
  if (areaCol) {
    const master = ensureAreaMasterSheet_();
    const dvRange = master.getRange(2,1,Math.max(1000, master.getMaxRows()-1),1);
    const rule = SpreadsheetApp.newDataValidation().requireValueInRange(dvRange, true).build();
    sh.getRange(2, areaCol, Math.max(1, sh.getMaxRows()-1), 1).setDataValidation(rule);
  }

  const contactCol = colByHeader_(sh,'受付連絡ステータス');
  if (contactCol) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['未連絡','連絡済み'], true).build();
    sh.getRange(2,contactCol,Math.max(1, sh.getMaxRows()-1),1).setDataValidation(rule);
    
    const lastRow = sh.getLastRow();
    if (lastRow >= 2) {
      const vals = sh.getRange(2, contactCol, lastRow-1, 1).getValues();
      let changed = false;
      for (let i=0;i<vals.length;i++){ 
        if (!vals[i][0]) { 
          vals[i][0] = '未連絡'; 
          changed = true; 
        } 
      }
      if (changed) sh.getRange(2, contactCol, lastRow-1, 1).setValues(vals);
    }
  }

  const assigneeCol = colByHeader_(sh, '回収担当者');
  if (assigneeCol) {
    try {
      const ash = getAssigneeSheet_();
      const dvRange = ash.getRange(2,1,Math.max(1, ash.getLastRow()-1),1);
      const rule = SpreadsheetApp.newDataValidation().requireValueInRange(dvRange, true).build();
      sh.getRange(2, assigneeCol, Math.max(1, sh.getMaxRows()-1), 1).setDataValidation(rule);
    } catch (e) {
      console.error('担当者マスタ検証ルール設定エラー:', e);
    }
  }
}

function installTimeTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'syncCalendarToSheet')
    .forEach(t => ScriptApp.deleteTrigger(t));
    
  ScriptApp.newTrigger('syncCalendarToSheet')
    .timeBased()
    .everyMinutes(5)
    .create();
  
  SpreadsheetApp.getUi().alert('5分おきのポーリング用トリガーを作成しました。');
}

/**
 * 自動処理用トリガーのインストール
 */
function installScheduledTasksTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'processScheduledTasks')
    .forEach(t => ScriptApp.deleteTrigger(t));
    
  ScriptApp.newTrigger('processScheduledTasks')
    .timeBased()
    .everyMinutes(5)
    .create();
  
  SpreadsheetApp.getUi().alert('⭐ 自動処理用トリガー（5分おき）を作成しました。');
}