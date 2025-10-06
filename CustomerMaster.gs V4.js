/**
 * CustomerMaster.gs V4 - 顧客マスタ連携
 * 
 * 【Ver4の変更点】
 * - 変更なし（整合性保持のためV4化）
 * 
 * 【このファイルの役割】
 * - 顧客マスタとの連携処理（検索・追加・補完）
 * - 初回利用判定・顧客情報自動補完
 * - 電話番号をキーとした顧客データ管理
 * 
 * 【依存関係】
 * - 依存元：Config.gs（顧客マスタ設定）、SheetUtils.gs、StringUtils.gs
 * - 依存先：なし
 * - 使用者：Main.gs（フォーム処理・メニュー操作）
 */

/**
 * 初回利用判定
 * 【説明】フォーム回答から初回利用かどうかを判定
 * @param {Array} srcHeaders - フォームヘッダー配列
 * @param {Array} srcValues - フォーム値配列
 * @returns {boolean} 初回利用の場合true
 */
function detectFirstUseFromSubmission_(srcHeaders, srcValues) {
  const targetNorm = normalizeHeaderKey_(CONFIG.firstUseQuestionHeader);
  let idx = -1;
  
  for (let i=0; i<srcHeaders.length; i++){
    if (normalizeHeaderKey_(srcHeaders[i]) === targetNorm) { 
      idx = i; 
      break; 
    }
  }
  
  if (idx === -1) return false;
  
  const v = String(srcValues[idx] || '').trim();
  return v === CONFIG.firstUseFirstLabel;
}

/**
 * フォーム送信時の顧客マスタ処理
 * 【説明】初回なら顧客マスタに追加、既存なら登録シートに情報補完
 * @param {GoogleAppsScript.Spreadsheet.Sheet} regSh - 登録シート
 * @param {number} row - 対象行番号
 * @param {boolean} isFirstUse - 初回利用フラグ
 */
function handleCustomerMasterOnSubmit_(regSh, row, isFirstUse) {
  if (isFirstUse) { 
    try { 
      insertCustomerRowFromReg_(regSh, row); 
    } catch (e) { 
      console.error('顧客マスタINSERT失敗:', e); 
    } 
  }
  
  try { 
    applyCustomerMasterToRegRow_(regSh, row); 
  } catch (e) { 
    console.error('顧客補完失敗:', e); 
  }
}

/**
 * 顧客シートの取得
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} 顧客マスタシート
 */
function getCustomerSheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.customerSpreadsheetId);
  const sh = ss.getSheetByName(CONFIG.customerSheetName);
  if (!sh) {
    throw new Error('顧客マスタのシートが見つかりません: ' + CONFIG.customerSheetName);
  }
  return sh;
}

/**
 * 登録シートから顧客マスタへ新規追加
 * 【説明】初回利用時に顧客マスタに新しい顧客情報を追加
 * @param {GoogleAppsScript.Spreadsheet.Sheet} regSh - 登録シート
 * @param {number} regRow - 登録シートの行番号
 */
function insertCustomerRowFromReg_(regSh, regRow) {
  // シート名に応じて電話番号列名を判定
  const sheetName = regSh.getName();
  let telHeaderName = '電話番号';
  
  if (sheetName === SPOT_CONFIG.sheets.clinic || sheetName === SPOT_CONFIG.sheets.dealer) {
    telHeaderName = '医院電話番号';
  }
  
  const telKey = normalizeTelForKey_(getCell_(regSh, regRow, telHeaderName));
  if (!telKey) return;
  
  const custSh = getCustomerSheet_();
  const custHeaders = sheetHeaders_(custSh).map(h => String(h).trim());
  const lastCol = custSh.getLastColumn() || 1;
  const out = new Array(lastCol).fill('');
  
  // ヘッダーが存在する場合のみ値を設定
  function setIfExists(header, value) {
    const idx = custHeaders.indexOf(header);
    if (idx >= 0) {
      out[idx] = (header === '電話番号') ? normalizeTelForKey_(value) : value;
    }
  }
  
  // 電話番号（キー）
  setIfExists('電話番号', telKey);
  
  // 補完対象フィールドをコピー
  CUSTOMER_FILL_HEADERS.forEach(h => {
    if (h === '電話番号') {
      // 定期回収の場合のみ電話番号をコピー
      if (sheetName === CONFIG.regSheetName) {
        setIfExists(h, getCell_(regSh, regRow, h));
      }
    } else {
      setIfExists(h, getCell_(regSh, regRow, h));
    }
  });
  
  // メールアドレス
  setIfExists('メールアドレス', getCell_(regSh, regRow, 'メールアドレス'));
  
  const before = custSh.getLastRow();
  custSh.appendRow(out);
  const writtenRow = before + 1;
  
  // 電話番号列を'付きテキストとして保存
  const telColIdx = custHeaders.indexOf('電話番号');
  if (telColIdx >= 0) {
    custSh.getRange(writtenRow, telColIdx + 1).setValue("'" + telKey);
  }
}

/**
 * エイリアス付きヘッダーインデックス検索
 * 【説明】メインヘッダーで見つからない場合、エイリアスでも検索
 * @param {Array} headers - ヘッダー配列
 * @param {string} main - メインヘッダー名
 * @param {Array} aliases - エイリアス配列
 * @returns {number} 見つかったインデックス（見つからない場合は-1）
 */
function headerIndexWithAliases_(headers, main, aliases) {
  let idx = headers.indexOf(main);
  if (idx >= 0) return idx;
  
  for (const a of (aliases || [])) {
    idx = headers.indexOf(a);
    if (idx >= 0) return idx;
  }
  return -1;
}

/**
 * 顧客インデックスの構築
 * 【説明】顧客マスタを電話番号キーで検索可能な形に構築
 * @returns {Object} 顧客インデックス（sheet, headers, index）
 */
function buildCustomerIndex_() {
  const sh = getCustomerSheet_();
  const headers = sheetHeaders_(sh).map(h => String(h).trim());
  const keyIdx = headers.indexOf(CUSTOMER_MASTER_KEY_HEADER);
  
  if (keyIdx < 0) {
    throw new Error(`顧客マスタに「${CUSTOMER_MASTER_KEY_HEADER}」列がありません。`);
  }
  
  // フィールドエイリアス定義
  const alias = {
    '休診日': ['回収が不可能な日（休診日など）を選択してください。 [休診日]'],
    '午前休診': ['回収が不可能な日（休診日など）を選択してください。 [午前休診]'],
    '午後休診': ['回収が不可能な日（休診日など）を選択してください。 [午後休診]']
  };
  
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
      CUSTOMER_MASTER_HEADERS.forEach(h => {
        const colIdx = (h in alias)
          ? headerIndexWithAliases_(headers, h, alias[h])
          : headers.indexOf(h);
        if (colIdx >= 0) rec[h] = values[i][colIdx];
      });
      
      index.set(key, { row: rowIx, record: rec });
    }
  }
  
  return { sheet: sh, headers, index };
}

/**
 * 顧客マスタから登録シートへの情報補完
 * 【説明】電話番号で顧客マスタを検索し、空欄を補完
 * - 定期回収：「電話番号」列をキーに検索
 * - スポット回収：「医院電話番号」列をキーに検索
 * @param {GoogleAppsScript.Spreadsheet.Sheet} regSh - 登録シート
 * @param {number} regRow - 対象行番号
 * @param {Object} prebuiltIdx - 事前構築済みインデックス（省略可）
 */
function applyCustomerMasterToRegRow_(regSh, regRow, prebuiltIdx) {
  // シート名に応じて電話番号列名を判定
  const sheetName = regSh.getName();
  let telHeaderName = '電話番号';
  
  if (sheetName === SPOT_CONFIG.sheets.clinic || sheetName === SPOT_CONFIG.sheets.dealer) {
    telHeaderName = '医院電話番号';
  }
  
  const telKey = normalizeTelForKey_(getCell_(regSh, regRow, telHeaderName));
  if (!telKey) return;
  
  const idx = prebuiltIdx || buildCustomerIndex_();
  const hit = idx.index.get(telKey);
  if (!hit) return;
  
  const rec = hit.record;
  let changed = false;
  
  // 補完対象フィールドの処理
  CUSTOMER_FILL_HEADERS.forEach(h => {
    const cur = getCell_(regSh, regRow, h);
    if (!cur && rec.hasOwnProperty(h) && rec[h] !== '' && rec[h] != null) {
      if (h === '電話番号') {
        // 定期回収シートの場合のみ電話番号を補完
        if (sheetName === CONFIG.regSheetName) {
          forceTelAsText_(regSh, regRow, '電話番号', rec[h]); 
        }
      } else {
        setCell_(regSh, regRow, h, rec[h]);
      }
      changed = true;
    }
  });
  
  // スポット回収シートの場合、医院電話番号を補完
  if ((sheetName === SPOT_CONFIG.sheets.clinic || sheetName === SPOT_CONFIG.sheets.dealer) && 
      !getCell_(regSh, regRow, '医院電話番号') && rec['電話番号']) {
    forceTelAsText_(regSh, regRow, '医院電話番号', rec['電話番号']);
    changed = true;
  }
  
  // メールアドレスの補完
  if (!getCell_(regSh, regRow, 'メールアドレス') && rec['メールアドレス']) {
    setCell_(regSh, regRow, 'メールアドレス', rec['メールアドレス']);
    changed = true;
  }
  
  if (changed) setLastUpdated_(regSh, regRow, 'sheet');
}

/**
 * メニュー：選択行を顧客情報で補完
 */
function menuFillSelectedFromCustomer_() {
  const sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== CONFIG.regSheetName) {
    return SpreadsheetApp.getUi().alert(`「${CONFIG.regSheetName}」で実行してください。`);
  }
  
  const row = sh.getActiveRange().getRow();
  if (row === 1) {
    return SpreadsheetApp.getUi().alert('データ行を選択してください。');
  }
  
  try { 
    applyCustomerMasterToRegRow_(sh, row); 
    SpreadsheetApp.getUi().alert('顧客情報で空欄を補完しました。'); 
  }
  catch (e) { 
    SpreadsheetApp.getUi().alert('補完に失敗: ' + e.message); 
  }
}

/**
 * メニュー：全行を顧客情報で一括補完
 */
function menuFillAllFromCustomer_() {
  const sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== CONFIG.regSheetName) {
    return SpreadsheetApp.getUi().alert(`「${CONFIG.regSheetName}」で実行してください。`);
  }
  
  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    return SpreadsheetApp.getUi().alert('補完対象がありません。');
  }
  
  try {
    const built = buildCustomerIndex_();
    for (let r=2; r<=lastRow; r++) {
      applyCustomerMasterToRegRow_(sh, r, built);
    }
    SpreadsheetApp.getUi().alert('全行の空欄を顧客情報で補完しました。');
  } catch (e) { 
    SpreadsheetApp.getUi().alert('一括補完に失敗: ' + e.message); 
  }
}