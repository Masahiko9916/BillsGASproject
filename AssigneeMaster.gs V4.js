/**
 * AssigneeMaster.gs V4 - 担当者マスタ連携
 * 
 * 【Ver4の変更点】
 * - エラー処理の強化
 * - キャッシュ機能の検討（将来的な実装）
 * 
 * 【このファイルの役割】
 * - 担当者マスタの検索・管理
 * - 担当者名⇔メールアドレスの相互変換
 * - カレンダーID（メール）から担当者名の逆引き
 * 
 * 【依存関係】
 * - 依存元：Config.gs、SheetUtils.gs
 * - 依存先：なし
 * - 使用者：RegularCollection.gs、Main.gs
 */

/**
 * 担当者マスタシートの取得
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} 担当者マスタシート
 */
function getAssigneeSheet_(){
  const ss = SpreadsheetApp.openById(CONFIG.regSpreadsheetId);
  const sh = ss.getSheetByName(CONFIG.assigneeSheetName);
  if (!sh) {
    throw new Error('担当者マスタのシートが見つかりません: ' + CONFIG.assigneeSheetName);
  }
  return sh;
}

/**
 * 担当者インデックスの構築（担当者名→メール）
 * @returns {Map} 担当者名をキーとするマップ
 */
function buildAssigneeIndex_(){
  const sh = getAssigneeSheet_();
  const headers = sheetHeaders_(sh).map(h => String(h).trim());
  const nameIdx = headers.indexOf(CONFIG.assigneeNameHeader);
  const mailIdx = headers.indexOf(CONFIG.assigneeEmailHeader);
  
  if (nameIdx < 0 || mailIdx < 0) {
    throw new Error('担当者マスタのヘッダーが不足しています（担当者名・メールアドレス）。');
  }
  
  const lastRow = sh.getLastRow();
  const map = new Map();
  
  if (lastRow >= 2){
    const vals = sh.getRange(2, 1, lastRow-1, Math.max(nameIdx, mailIdx)+1).getValues();
    for (let i=0; i<vals.length; i++){
      const name = String(vals[i][nameIdx]||'').trim();
      const mail = String(vals[i][mailIdx]||'').trim();
      if (name && mail) {
        // 通常の名前で登録
        map.set(name, mail);
        // スペースを除去した名前でも登録（検索精度向上）
        const nameNoSpace = name.replace(/\s+/g, '');
        if (nameNoSpace !== name) {
          map.set(nameNoSpace, mail);
        }
      }
    }
  }
  return map;
}

/**
 * 逆引きインデックスの構築（メール→担当者名）
 * @returns {Map} メールをキーとするマップ
 */
function buildAssigneeEmailToNameIndex_(){
  const sh = getAssigneeSheet_();
  const headers = sheetHeaders_(sh).map(h => String(h).trim());
  const nameIdx = headers.indexOf(CONFIG.assigneeNameHeader);
  const mailIdx = headers.indexOf(CONFIG.assigneeEmailHeader);
  
  if (nameIdx < 0 || mailIdx < 0) {
    throw new Error('担当者マスタのヘッダーが不足しています（担当者名・メールアドレス）。');
  }
  
  const lastRow = sh.getLastRow();
  const map = new Map();
  
  if (lastRow >= 2){
    const vals = sh.getRange(2, 1, lastRow-1, Math.max(nameIdx, mailIdx)+1).getValues();
    for (let i=0; i<vals.length; i++){
      const name = String(vals[i][nameIdx]||'').trim();
      const mail = String(vals[i][mailIdx]||'').trim();
      if (name && mail) map.set(mail, name);
    }
  }
  return map;
}

/**
 * 担当者名からメールアドレスを取得
 * @param {string} name - 担当者名
 * @param {Map} prebuilt - 事前構築済みインデックス（省略可）
 * @returns {string} メールアドレス（見つからない場合は空文字）
 */
function getAssigneeEmailByName_(name, prebuilt){
  const idx = prebuilt || buildAssigneeIndex_();
  const searchName = String(name||'').trim();
  
  // まず完全一致で検索
  let result = idx.get(searchName);
  if (result) return result;
  
  // スペースを除去して再検索
  const searchNameNoSpace = searchName.replace(/\s+/g, '');
  result = idx.get(searchNameNoSpace);
  if (result) return result;
  
  // 全角スペースを半角に変換して再検索
  const searchNameHalfSpace = searchName.replace(/　/g, ' ');
  result = idx.get(searchNameHalfSpace);
  if (result) return result;
  
  // 全てのスペースを除去して再検索
  const searchNameAllNoSpace = searchNameHalfSpace.replace(/\s+/g, '');
  result = idx.get(searchNameAllNoSpace);
  if (result) return result;
  
  return '';
}

/**
 * メールアドレスから担当者名を取得
 * @param {string} email - メールアドレス
 * @param {Map} prebuilt - 事前構築済みインデックス（省略可）
 * @returns {string} 担当者名（見つからない場合は空文字）
 */
function getAssigneeNameByEmail_(email, prebuilt){
  const idx = prebuilt || buildAssigneeEmailToNameIndex_();
  return idx.get(String(email||'').trim()) || '';
}

/**
 * 全担当者のメールアドレスリストを取得【V4追加】
 * @returns {Array<string>} メールアドレスの配列
 */
function getAllAssigneeEmails_() {
  try {
    const sh = getAssigneeSheet_();
    const headers = sheetHeaders_(sh).map(h => String(h).trim());
    const mailIdx = headers.indexOf(CONFIG.assigneeEmailHeader);
    
    if (mailIdx < 0) {
      throw new Error('担当者マスタにメールアドレス列がありません。');
    }
    
    const lastRow = sh.getLastRow();
    if (lastRow < 2) {
      return [];
    }
    
    const values = sh.getRange(2, mailIdx + 1, lastRow - 1, 1).getValues();
    const emails = [];
    
    values.forEach(row => {
      const email = String(row[0] || '').trim();
      if (email && email.includes('@')) {
        emails.push(email);
      }
    });
    
    return [...new Set(emails)]; // 重複を削除
    
  } catch (e) {
    console.error('担当者リスト取得エラー:', e);
    return [];
  }
}