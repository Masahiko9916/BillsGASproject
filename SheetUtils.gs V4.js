/**
 * SheetUtils.gs V4 - スプレッドシート操作ユーティリティ
 * 
 * 【Ver4の変更点】
 * - STATUS定数を使用
 * - 整合性保持
 * 
 * 【このファイルの役割】
 * - スプレッドシートの基本操作（セル読み書き、ヘッダー操作等）
 * - 行検索、ソート、シート作成等の共通処理
 * - 他の全ファイルで使用される基盤機能
 * 
 * 【依存関係】
 * - 依存元：Config.gs（設定値を使用）
 * - 依存先：StringUtils.gs（文字列処理）
 * - 使用者：Main.gs、CustomerMaster.gs、RegularCollection.gs等
 */

/**
 * シートのヘッダー行を取得
 * 【説明】1行目の全列データを配列で返す
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh - 対象シート
 * @returns {Array} ヘッダー配列
 */
const HEADER_INFO_CACHE_ = {};
const ROW_VALUE_CACHE_ = {};

function sheetKey_(sh) {
  try {
    const parent = sh.getParent();
    const parentId = parent && typeof parent.getId === 'function' ? parent.getId() : '';
    return parentId + '::' + sh.getSheetId();
  } catch (e) {
    return 'unknown::' + sh.getSheetId();
  }
}

function getHeaderInfo_(sh) {
  const key = sheetKey_(sh);
  const lastCol = sh.getLastColumn();
  const cached = HEADER_INFO_CACHE_[key];
  if (cached && cached.lastCol === lastCol) {
    return cached;
  }

  if (lastCol === 0) {
    const emptyInfo = { headers: [], map: {}, lastCol: 0, signature: '' };
    HEADER_INFO_CACHE_[key] = emptyInfo;
    ROW_VALUE_CACHE_[key] = { signature: '', rows: {} };
    return emptyInfo;
  }

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const headerMap = {};
  headers.forEach((h, idx) => {
    if (h && headerMap[h] === undefined) {
      headerMap[h] = idx;
    }
  });
  const info = {
    headers,
    map: headerMap,
    lastCol,
    signature: headers.join('\u0000'),
  };
  HEADER_INFO_CACHE_[key] = info;

  const rowCache = ROW_VALUE_CACHE_[key];
  if (!rowCache || rowCache.signature !== info.signature) {
    ROW_VALUE_CACHE_[key] = { signature: info.signature, rows: {} };
  }

  return info;
}

function sheetHeaders_(sh) {
  return getHeaderInfo_(sh).headers;
}

/**
 * ヘッダー名から列インデックスを取得
 * 【説明】指定したヘッダー名が何列目にあるかを調べる（0始まり）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh - 対象シート
 * @param {string} headerName - 検索するヘッダー名
 * @returns {number} 列インデックス（見つからない場合は-1）
 */
function headerIndex_(sh, headerName) {
  const headerInfo = getHeaderInfo_(sh);
  const idx = headerInfo.map[headerName];
  return idx === undefined ? -1 : idx;
}

/**
 * ヘッダー名から列番号を取得
 * 【説明】指定したヘッダー名が何列目にあるかを調べる（1始まり）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh - 対象シート
 * @param {string} headerName - 検索するヘッダー名
 * @returns {number|null} 列番号（見つからない場合はnull）
 */
function colByHeader_(sh, headerName) {
  const idx = headerIndex_(sh, headerName);
  return idx >= 0 ? idx + 1 : null;
}

function getRowCache_(sh, headerInfo) {
  const key = sheetKey_(sh);
  let cache = ROW_VALUE_CACHE_[key];
  if (!cache || cache.signature !== headerInfo.signature) {
    cache = { signature: headerInfo.signature, rows: {} };
    ROW_VALUE_CACHE_[key] = cache;
  }
  return cache;
}

function getRowValuesFromCache_(sh, row, headerInfo) {
  if (row <= 0) return null;
  if (!headerInfo.lastCol) return null;
  const cache = getRowCache_(sh, headerInfo);
  let rowValues = cache.rows[row];
  if (!rowValues) {
    rowValues = sh.getRange(row, 1, 1, headerInfo.lastCol).getValues()[0];
    cache.rows[row] = rowValues;
  }
  return rowValues;
}

function getRowValuesByHeaders_(sh, row, headerNames) {
  const headerInfo = getHeaderInfo_(sh);
  if (!headerNames || headerNames.length === 0) return {};
  const rowValues = getRowValuesFromCache_(sh, row, headerInfo);
  if (!rowValues) return {};

  const result = {};
  headerNames.forEach(name => {
    const idx = headerInfo.map[name];
    if (idx !== undefined) {
      result[name] = rowValues[idx];
    }
  });
  return result;
}

function updateRowCacheValue_(sh, row, headerInfo, colIdx, value) {
  const cache = getRowCache_(sh, headerInfo);
  const rowValues = cache.rows[row];
  if (rowValues) {
    rowValues[colIdx] = value;
  }
}

/**
 * ヘッダー名でセル値を取得
 * 【説明】行番号とヘッダー名を指定してセルの値を取得
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh - 対象シート
 * @param {number} row - 行番号
 * @param {string} headerName - ヘッダー名
 * @returns {any} セルの値（列が見つからない場合は空文字）
 */
function getCell_(sh, row, headerName){
  const headerInfo = getHeaderInfo_(sh);
  const colIdx = headerInfo.map[headerName];
  if (colIdx === undefined) return '';
  const rowValues = getRowValuesFromCache_(sh, row, headerInfo);
  if (!rowValues) return '';
  const value = rowValues[colIdx];
  return value === undefined ? '' : value;
}

/**
 * ヘッダー名でセル値を設定
 * 【説明】行番号とヘッダー名を指定してセルに値を設定
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh - 対象シート
 * @param {number} row - 行番号
 * @param {string} headerName - ヘッダー名
 * @param {any} val - 設定する値
 */
function setCell_(sh, row, headerName, val){
  if (!headerName) return;
  const updates = {};
  updates[headerName] = val;
  setRowValues_(sh, row, updates);
}

function setRowValues_(sh, row, valueMap) {
  if (!valueMap) return;
  const headerInfo = getHeaderInfo_(sh);
  if (!headerInfo.lastCol) return;

  const targetHeaders = Object.keys(valueMap);
  if (targetHeaders.length === 0) return;

  const cache = getRowCache_(sh, headerInfo);
  let rowValues = getRowValuesFromCache_(sh, row, headerInfo);
  if (!rowValues) {
    // 既存値を一度だけ取得（閲覧権限があれば読み取りは可能）
    rowValues = sh.getRange(row, 1, 1, headerInfo.lastCol).getValues()[0];
    cache.rows[row] = rowValues;
  }

  const updates = [];
  targetHeaders.forEach(name => {
    const idx = headerInfo.map[name];
    if (idx !== undefined) {
      updates.push({ idx, value: valueMap[name] });
    }
  });

  if (updates.length === 0) return;

  updates.sort((a, b) => a.idx - b.idx);

  let cursor = 0;
  while (cursor < updates.length) {
    let end = cursor + 1;
    while (end < updates.length && updates[end].idx === updates[end - 1].idx + 1) {
      end++;
    }

    const segment = updates.slice(cursor, end);
    const startCol = segment[0].idx + 1;
    const values = segment.map(item => item.value);
    sh.getRange(row, startCol, 1, values.length).setValues([values]);

    if (rowValues) {
      segment.forEach((item, offset) => {
        rowValues[item.idx] = values[offset];
      });
    }

    cursor = end;
  }
}

/**
 * 最終更新情報を設定
 * 【説明】指定行の最終更新者・時刻・元を自動設定し、必要に応じて追加列もまとめて更新
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh - 対象シート
 * @param {number} row - 行番号
 * @param {string} source - 更新元（'sheet', 'calendar', 'form'等）
 * @param {Object=} extraValues - 併せて更新したい列のマップ（{ヘッダー名: 値}）
 */
function setLastUpdated_(sh, row, source, extraValues){
  const updates = Object.assign({}, extraValues || {});
  if (updates['最終更新者'] === undefined) {
    updates['最終更新者'] = getActiveEmail_();
  }
  if (updates['最終更新時刻'] === undefined) {
    updates['最終更新時刻'] = Utilities.formatDate(new Date(), CONFIG.tz, 'yyyy-MM-dd HH:mm:ss');
  }
  if (source && updates['最終更新元'] === undefined) {
    updates['最終更新元'] = source;
  }
  setRowValues_(sh, row, updates);
}

/**
 * 必要に応じて自動ソート実行
 * 【説明】CONFIG.sortAfterSubmitがtrueの場合、指定された順序でソート
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh - 対象シート
 */
function sortRegSheetIfNeeded_(sh) {
  if (!CONFIG.sortAfterSubmit) return;
  
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow <= 2) return; // データが1行以下なら何もしない
  
  const headers = sheetHeaders_(sh).map(h => String(h).trim());
  const specs = [];
  
  // CONFIG.sortSpecに基づいてソート仕様を構築
  (CONFIG.sortSpec || []).forEach(s => {
    const idx = headers.indexOf(s.header);
    if (idx >= 0) {
      specs.push({ column: idx + 1, ascending: !!s.ascending });
    }
  });
  
  if (specs.length === 0) return;
  
  // ヘッダー行を除いてソート
  sh.getRange(2, 1, lastRow - 1, lastCol).sort(specs);
}

/**
 * 列値で行を検索
 * 【説明】指定した列の値で行番号を検索
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh - 対象シート
 * @param {string} headerName - 検索対象のヘッダー名
 * @param {any} val - 検索値
 * @returns {number|null} 見つかった行番号（見つからない場合はnull）
 */
function findRowByColumnValue_(sh, headerName, val) {
  const col = colByHeader_(sh, headerName);
  if (!col) return null;
  
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;
  
  const vals = sh.getRange(2, col, lastRow - 1, 1).getValues();
  const target = String(val);
  
  for (let i = 0; i < vals.length; i++) { 
    if (String(vals[i][0]) === target) return 2 + i; 
  }
  return null;
}

/**
 * エリアマスタシートの確保
 * 【説明】「エリアマスタ」シートが存在しない場合は作成し、ヘッダーを設定
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} エリアマスタシート
 */
function ensureAreaMasterSheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.regSpreadsheetId);
  let sh = ss.getSheetByName(CONFIG.areaMasterSheetName);
  if (!sh) sh = ss.insertSheet(CONFIG.areaMasterSheetName);
  
  // A1セルにヘッダーを設定
  if (String(sh.getRange(1,1).getValue() || '').trim() !== 'エリア候補') {
    sh.getRange(1,1).setValue('エリア候補');
  }
  return sh;
}

/**
 * 電話番号をテキスト形式で強制設定
 * 【説明】電話番号の先頭ゼロ保持のため、'（アポストロフィ）付きで設定
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh - 対象シート
 * @param {number} row - 行番号
 * @param {string} headerName - ヘッダー名
 * @param {any} raw - 元の値
 */
function forceTelAsText_(sh, row, headerName, raw){
  const updates = {};
  updates[headerName] = formatTelForSheet_(raw);
  setRowValues_(sh, row, updates);
}

/**
 * ヘッダー名でセル値を設定（汎用版）
 * 【説明】任意のヘッダー配列を使ってセル値を設定
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh - 対象シート
 * @param {number} row - 行番号
 * @param {Array} headers - ヘッダー配列
 * @param {string} headerName - ヘッダー名
 * @param {any} value - 設定する値
 */
function setByHeader_(sh, row, headers, headerName, value) {
  const idx = headers.indexOf(headerName);
  if (idx >= 0) sh.getRange(row, idx+1).setValue(value);
}

/**
 * アクティブユーザーのメールアドレス取得
 * 【説明】現在のGASユーザーのメールアドレスを取得
 * @returns {string} メールアドレス
 */
function getActiveEmail_() {
  try { 
    return Session.getActiveUser().getEmail() || ''; 
  } catch (e) {
    try { 
      return Session.getEffectiveUser().getEmail() || ''; 
    } catch (e2){ 
      return ''; 
    }
  }
}

/**
 * 文字列の空チェック・正規化
 * 【説明】nullや空文字を空文字に統一
 * @param {any} v - チェック対象
 * @returns {string} 正規化された文字列
 */
function nz_(v) { 
  return (v === null || v === undefined) ? '' : String(v).trim(); 
}