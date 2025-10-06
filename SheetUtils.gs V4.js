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
function sheetHeaders_(sh) {
  const lastCol = sh.getLastColumn();
  if (lastCol === 0) return [];
  return sh.getRange(1,1,1,lastCol).getValues()[0];
}

/**
 * ヘッダー名から列インデックスを取得
 * 【説明】指定したヘッダー名が何列目にあるかを調べる（0始まり）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh - 対象シート
 * @param {string} headerName - 検索するヘッダー名
 * @returns {number} 列インデックス（見つからない場合は-1）
 */
function headerIndex_(sh, headerName) {
  const headers = sheetHeaders_(sh);
  for (let i=0; i<headers.length; i++){
    if (String(headers[i]).trim() === headerName) return i;
  }
  return -1;
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
  return idx >= 0 ? idx+1 : null;
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
  const col = colByHeader_(sh, headerName);
  if (!col) return '';
  return sh.getRange(row, col).getValue();
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
  const col = colByHeader_(sh, headerName);
  if (!col) return;
  sh.getRange(row, col).setValue(val);
}

/**
 * 最終更新情報を設定
 * 【説明】指定行の最終更新者・時刻・元を自動設定
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh - 対象シート
 * @param {number} row - 行番号
 * @param {string} source - 更新元（'sheet', 'calendar', 'form'等）
 */
function setLastUpdated_(sh, row, source){
  setCell_(sh, row, '最終更新者', getActiveEmail_());
  setCell_(sh, row, '最終更新時刻', Utilities.formatDate(new Date(), CONFIG.tz, 'yyyy-MM-dd HH:mm:ss'));
  if (source) setCell_(sh, row, '最終更新元', source);
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
  const col = colByHeader_(sh, headerName);
  if (!col) return;
  
  const s = formatTelForSheet_(raw);
  sh.getRange(row, col).setValue(s);
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