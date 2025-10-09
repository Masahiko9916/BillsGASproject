/**
 * DealerMaster.gs V4 - ディーラーマスタ連携
 * 
 * 【Ver4の変更点】
 * - STATUS定数の使用
 * - エラーハンドリングの改善
 * 
 * 【このファイルの役割】
 * - ディーラーマスタとの連携処理（検索・追加・補完）
 * - 初回利用判定・ディーラー情報自動補完
 * - ディーラー電話番号をキーとしたデータ管理
 * 
 * 【依存関係】
 * - 依存元：Config.gs（DEALER_CONFIG）、SheetUtils.gs、StringUtils.gs
 * - 依存先：なし
 * - 使用者：Main.gs（フォーム処理・メニュー操作）
 */

/**
 * ディーラー初回利用判定
 * 【説明】フォーム回答から初回利用かどうかを判定
 * @param {Array} srcHeaders - フォームヘッダー配列
 * @param {Array} srcValues - フォーム値配列
 * @returns {boolean} 初回利用の場合true
 */
function detectDealerFirstUse_(srcHeaders, srcValues) {
  const targetNorm = normalizeHeaderKey_(SPOT_CONFIG.firstUseQuestionHeader);
  let idx = -1;
  
  for (let i=0; i<srcHeaders.length; i++){
    if (normalizeHeaderKey_(srcHeaders[i]) === targetNorm) { 
      idx = i; 
      break; 
    }
  }
  
  if (idx === -1) return false;
  
  const v = String(srcValues[idx] || '').trim();
  return v === SPOT_CONFIG.firstUseFirstLabel;
}

/**
 * フォーム送信時のディーラーマスタ処理
 * 【説明】初回ならディーラーマスタに追加、既存なら登録シートに情報補完
 * @param {GoogleAppsScript.Spreadsheet.Sheet} regSh - 登録シート
 * @param {number} row - 対象行番号
 * @param {boolean} isFirstUse - 初回利用フラグ
 */
function handleDealerMasterOnSubmit_(regSh, row, isFirstUse) {
  if (isFirstUse) { 
    try { 
      insertDealerRowFromReg_(regSh, row); 
    } catch (e) { 
      console.error('ディーラーマスタINSERT失敗:', e); 
    } 
  }
  
  try { 
    applyDealerMasterToRegRow_(regSh, row); 
  } catch (e) { 
    console.error('ディーラー補完失敗:', e); 
  }
}

/**
 * ディーラーシートの取得
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} ディーラーマスタシート
 */
function getDealerSheet_() {
  const ss = SpreadsheetApp.openById(DEALER_CONFIG.spreadsheetId);
  const sh = ss.getSheetByName(DEALER_CONFIG.sheetName);
  if (!sh) {
    throw new Error('ディーラーマスタのシートが見つかりません: ' + DEALER_CONFIG.sheetName);
  }
  return sh;
}

/**
 * 登録シートからディーラーマスタへ新規追加
 * 【説明】初回利用時にディーラーマスタに新しいディーラー情報を追加
 * @param {GoogleAppsScript.Spreadsheet.Sheet} regSh - 登録シート
 * @param {number} regRow - 登録シートの行番号
 */
function insertDealerRowFromReg_(regSh, regRow) {
  const telKey = normalizeTelForKey_(getCell_(regSh, regRow, 'ディーラー電話番号'));
  if (!telKey) return;
  
  const dealerSh = getDealerSheet_();
  const dealerHeaders = sheetHeaders_(dealerSh).map(h => String(h).trim());
  const lastCol = dealerSh.getLastColumn() || 1;
  const out = new Array(lastCol).fill('');
  
  // ヘッダーが存在する場合のみ値を設定
  function setIfExists(header, value) {
    const idx = dealerHeaders.indexOf(header);
    if (idx >= 0) {
      out[idx] = (header === 'ディーラー電話番号') ? normalizeTelForKey_(value) : value;
    }
  }
  
  // ディーラー電話番号（キー）
  setIfExists('ディーラー電話番号', telKey);
  
  // 補完対象フィールドをコピー
  DEALER_FILL_HEADERS.forEach(h => {
    setIfExists(h, getCell_(regSh, regRow, h));
  });
  
  // メールアドレス
  setIfExists('メールアドレス', getCell_(regSh, regRow, 'メールアドレス'));
  
  const before = dealerSh.getLastRow();
  dealerSh.appendRow(out);
  const writtenRow = before + 1;
  
  // ディーラー電話番号列を'付きテキストとして保存
  const telColIdx = dealerHeaders.indexOf('ディーラー電話番号');
  if (telColIdx >= 0) {
    dealerSh.getRange(writtenRow, telColIdx + 1).setValue("'" + telKey);
  }
}

/**
 * ディーラーインデックスの構築
 * 【説明】ディーラーマスタをディーラー電話番号キーで検索可能な形に構築
 * @returns {Object} ディーラーインデックス（sheet, headers, index）
 */
function buildDealerIndex_() {
  const sh = getDealerSheet_();
  const headers = sheetHeaders_(sh).map(h => String(h).trim());
  const keyIdx = headers.indexOf(DEALER_CONFIG.keyHeader);
  
  if (keyIdx < 0) {
    throw new Error(`ディーラーマスタに「${DEALER_CONFIG.keyHeader}」列がありません。`);
  }
  
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const index = new Map();
  
  if (lastRow >= 2) {
    const values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
    
    for (let i=0; i<values.length; i++){
      const rowIx = 2 + i;
      const keyRaw = values[i][keyIdx];
      const key = normalizeTelForKey_(keyRaw);
      
      if (!key) continue;
      
      const rec = {};
      DEALER_MASTER_HEADERS.forEach(h => {
        const colIdx = headers.indexOf(h);
        if (colIdx >= 0) rec[h] = values[i][colIdx];
      });
      
      index.set(key, { row: rowIx, record: rec });
    }
  }
  
  return { sheet: sh, headers, index };
}

/**
 * ディーラーマスタから登録シートへの情報補完
 * 【説明】ディーラー電話番号でディーラーマスタを検索し、空欄を補完
 * @param {GoogleAppsScript.Spreadsheet.Sheet} regSh - 登録シート
 * @param {number} regRow - 対象行番号
 * @param {Object} prebuiltIdx - 事前構築済みインデックス（省略可）
 */
function applyDealerMasterToRegRow_(regSh, regRow, prebuiltIdx) {
  const telKey = normalizeTelForKey_(getCell_(regSh, regRow, 'ディーラー電話番号'));
  if (!telKey) return;
  
  const idx = prebuiltIdx || buildDealerIndex_();
  const hit = idx.index.get(telKey);
  if (!hit) return;
  
  const rec = hit.record;
  const updates = {};

  // 補完対象フィールドの処理
  DEALER_FILL_HEADERS.forEach(h => {
    const cur = getCell_(regSh, regRow, h);
    if (!cur && rec.hasOwnProperty(h) && rec[h] !== '' && rec[h] != null) {
      if (h === 'ディーラー電話番号') {
        updates['ディーラー電話番号'] = formatTelForSheet_(rec[h]);
      } else {
        updates[h] = rec[h];
      }
    }
  });

  // メールアドレスの補完
  if (!getCell_(regSh, regRow, 'メールアドレス') && rec['メールアドレス']) {
    updates['メールアドレス'] = rec['メールアドレス'];
  }

  if (Object.keys(updates).length > 0) {
    setLastUpdated_(regSh, regRow, 'sheet', updates);
  }
}

/**
 * メニュー：選択行をディーラー情報で補完
 */
function menuFillSelectedFromDealer_() {
  const sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== SPOT_CONFIG.sheets.dealer) {
    return SpreadsheetApp.getUi().alert(`「${SPOT_CONFIG.sheets.dealer}」で実行してください。`);
  }
  
  const row = sh.getActiveRange().getRow();
  if (row === 1) {
    return SpreadsheetApp.getUi().alert('データ行を選択してください。');
  }
  
  try { 
    applyDealerMasterToRegRow_(sh, row); 
    SpreadsheetApp.getUi().alert('ディーラー情報で空欄を補完しました。'); 
  }
  catch (e) { 
    SpreadsheetApp.getUi().alert('補完に失敗: ' + e.message); 
  }
}

/**
 * メニュー：全行をディーラー情報で一括補完
 */
function menuFillAllFromDealer_() {
  const sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== SPOT_CONFIG.sheets.dealer) {
    return SpreadsheetApp.getUi().alert(`「${SPOT_CONFIG.sheets.dealer}」で実行してください。`);
  }
  
  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    return SpreadsheetApp.getUi().alert('補完対象がありません。');
  }
  
  try {
    const built = buildDealerIndex_();
    for (let r=2; r<=lastRow; r++) {
      applyDealerMasterToRegRow_(sh, r, built);
    }
    SpreadsheetApp.getUi().alert('全行の空欄をディーラー情報で補完しました。');
  } catch (e) { 
    SpreadsheetApp.getUi().alert('一括補完に失敗: ' + e.message); 
  }
}