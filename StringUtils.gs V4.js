/**
 * StringUtils.gs V4 - 文字列処理ユーティリティ
 * 
 * 【Ver4の変更点】
 * - 変更なし（整合性保持のためV4化）
 * 
 * 【このファイルの役割】
 * - 全角→半角変換、正規化、フォーマット等の文字列処理
 * - ヘッダー名の正規化・マッピング処理
 * - 電話番号、数値等の特殊フォーマット処理
 * 
 * 【依存関係】
 * - 依存元：Config.gs（DEST_ALIASES等の設定値を使用）
 * - 依存先：なし（基盤ユーティリティ）
 * - 使用者：Main.gs、CustomerMaster.gs、SheetUtils.gs等
 */

/**
 * 全角文字を半角に変換
 * 【説明】全角英数字・記号を半角に変換、全角スペースも半角に
 * @param {string} str - 変換対象文字列
 * @returns {string} 半角変換済み文字列
 */
function toHalfWidth_(str) {
  return String(str).replace(/[！-～]/g, ch => 
    String.fromCharCode(ch.charCodeAt(0) - 0xFEE0)
  ).replace(/　/g, ' ');
}

/**
 * 丸数字を通常数字に変換
 * 【説明】①②③等の丸数字を1,2,3等に変換
 * @param {string} str - 変換対象文字列
 * @returns {string} 変換済み文字列
 */
function replaceCircledNums_(str) {
  const map = {
    '①':'1','②':'2','③':'3','④':'4','⑤':'5',
    '⑥':'6','⑦':'7','⑧':'8','⑨':'9','⑩':'10'
  };
  return String(str).replace(/[①②③④⑤⑥⑦⑧⑨⑩]/g, m => map[m] || m);
}

/**
 * ヘッダー名の正規化
 * 【説明】フォーム質問文と内部ヘッダーの照合用に文字列を正規化
 * - 括弧内文字削除、記号削除、大文字小文字統一等
 * @param {string} s - 正規化対象文字列
 * @returns {string} 正規化済み文字列
 */
function normalizeHeaderKey_(s) {
  if (s == null) return '';
  
  let x = String(s).trim();
  
  // 括弧内文字を削除
  x = x.replace(/（.*?）/g, '').replace(/\(.*?\)/g, '');
  x = x.replace(/【.*?】/g, '').replace(/\[.*?\]/g, '');
  
  // 半角変換
  x = toHalfWidth_(x);
  
  // リットル表記の統一
  x = x.replace(/ℓ/gi, 'L').replace(/Ｌ/gi, 'L').replace(/ﾘｯﾄﾙ|リットル/gi, 'L');
  
  // 丸数字を通常数字に
  x = replaceCircledNums_(x);
  
  // 記号・空白を削除
  x = x.replace(/[ \u3000\t\r\n、。，,.:：;；\/／\-－—–—]/g, '');
  
  return x.toLowerCase();
}

/**
 * 電話番号のチャット表示用フォーマット
 * 【説明】Chat通知時の電話番号表示形式
 * @param {any} v - 元の電話番号値
 * @returns {string} フォーマット済み電話番号
 */
function formatTelForChat_(v) {
  if (v === null || v === undefined) return '';
  
  let s = String(v).trim();
  if (s === '') return '';
  
  // 先頭のアポストロフィを削除
  if (s.charAt(0) === "'") s = s.substring(1);
  
  s = toHalfWidth_(s);
  if (s === '0') return '';
  
  return s;
}

/**
 * 文字列を整数数量に変換
 * 【説明】フォーム入力値を数値に変換（エラー時は0を返す）
 * @param {any} v - 変換対象値
 * @returns {number} 変換された整数（エラー時は0）
 */
function toIntQty_(v) {
  if (v === null || v === undefined) return 0;
  
  let s = toHalfWidth_(String(v)).replace(/[^\d-]/g, '');
  if (s === '' || s === '-') return 0;
  
  const n = parseInt(s, 10);
  return isNaN(n) ? 0 : n;
}

/**
 * 電話番号の検索キー正規化
 * 【説明】顧客マスタ検索用に電話番号を正規化
 * @param {any} tel - 電話番号
 * @returns {string} 正規化済み電話番号（数字のみ）
 */
function normalizeTelForKey_(tel) {
  if (tel == null) return '';
  
  let s = toHalfWidth_(String(tel)).trim();
  
  // 先頭のアポストロフィを削除
  if (s.charAt(0) === "'") s = s.substring(1);
  
  // 数字以外を削除
  s = s.replace(/[^\d]/g, '');
  
  return s;
}

/**
 * フォームヘッダーのエイリアス検索
 * 【説明】DEST_ALIASESを使って元の値を取得
 * @param {string} headerName - 検索するヘッダー名
 * @param {Array} srcHeaders - フォームヘッダー配列
 * @param {Array} srcValues - フォーム値配列
 * @returns {any} 見つかった値（見つからない場合は空文字）
 */
function pickOriginalByAliases_(headerName, srcHeaders, srcValues){
  const candidates = [headerName].concat(DEST_ALIASES[headerName] || []);
  
  for (const cand of candidates) {
    let idx = srcHeaders.indexOf(cand);
    
    // 完全一致で見つからない場合は正規化して検索
    if (idx === -1) {
      const normCand = normalizeHeaderKey_(cand || '');
      idx = srcHeaders.findIndex(sh => normalizeHeaderKey_(sh).includes(normCand));
    }
    
    if (idx !== -1 && idx < srcValues.length) {
      return srcValues[idx];
    }
  }
  return '';
}

/**
 * フォーム回答→内部形式変換
 * 【説明】フォーム回答をスプレッドシートの行データに変換
 * @param {Array} dstHeaders - 出力先ヘッダー配列
 * @param {Array} srcHeaders - フォームヘッダー配列
 * @param {Array} srcValues - フォーム値配列
 * @returns {Array} 変換された行データ
 */
function buildRowByMapping_(dstHeaders, srcHeaders, srcValues) {
  const now = new Date();
  const op = getActiveEmail_();
  const uuid = Utilities.getUuid();
  
  // フォームヘッダーの正規化インデックス作成
  const srcNormToIndex = {};
  for (let i = 0; i < srcHeaders.length; i++) {
    const norm = normalizeHeaderKey_(srcHeaders[i]);
    if (!(norm in srcNormToIndex)) srcNormToIndex[norm] = i;
  }
  
  // 管理列のデフォルト値
  const managementDefaults = {
    'ステータス': STATUS.UNHANDLED,
    '回収予定日': '',
    '開始時間': '',
    '終了時間': '',
    '回収担当者': '',
    'エリア': '',
    '受付連絡ステータス': '未連絡'
  };
  
  // システム列の値
  const systemValues = {
    '受付ID': uuid,
    'カレンダーイベントID': '',
    'カレンダーID': '', 
    '行URL': '',
    '最終更新者': op,
    '最終更新時刻': Utilities.formatDate(now, CONFIG.tz, 'yyyy-MM-dd HH:mm:ss'),
    '削除日時': '',
    '削除元': '',
    '最終更新元': 'sheet'
  };
  
  const out = [];
  
  for (const h of dstHeaders) {
    // 管理列の場合
    if (h in managementDefaults) { 
      out.push(managementDefaults[h]); 
      continue; 
    }
    
    // システム列の場合
    if (h in systemValues) { 
      out.push(systemValues[h]); 
      continue; 
    }
    
    // 特別処理：住所は二段（ご住所・番地以下）を結合
    if (h === '住所') {
      const searchLabels = [
        'ご住所を入力してください。',
        '番地以下の情報を入力してください。',
        'ご住所を入力してください',
        '番地以下の情報を入力してください'
      ];
      
      const usedIdx = new Set();
      const parts = [];
      
      for (const lab of searchLabels) {
        let idx = srcHeaders.indexOf(lab);
        if (idx === -1) {
          const n = normalizeHeaderKey_(lab || '');
          idx = srcHeaders.findIndex(sh => normalizeHeaderKey_(sh).includes(n));
        }
        
        if (idx !== -1 && idx < srcValues.length && !usedIdx.has(idx)) {
          const v = srcValues[idx];
          if (v != null && String(v).trim() !== '') {
            parts.push(String(v).trim());
            usedIdx.add(idx);
          }
        }
      }
      
      if (parts.length) { 
        out.push(parts.join(' ')); 
        continue; 
      }
    }
    
    // 通常処理：見出し振れを吸収
    const candidates = [h].concat(DEST_ALIASES[h] || []);
    let val = '';
    let found = false;
    
    for (const cand of candidates) {
      let idx = srcHeaders.indexOf(cand);
      
      if (idx === -1) {
        const normCand = normalizeHeaderKey_(cand || '');
        if (srcNormToIndex.hasOwnProperty(normCand)) {
          idx = srcNormToIndex[normCand];
        } else {
          const normH = normalizeHeaderKey_(h);
          const candIdx = srcHeaders.findIndex(sh => {
            const ns = normalizeHeaderKey_(sh);
            return ns.includes(normCand) || normCand.includes(ns) || 
                   ns.includes(normH) || normH.includes(ns);
          });
          if (candIdx !== -1) idx = candIdx;
        }
      }
      
      if (idx !== -1 && idx < srcValues.length) {
        val = srcValues[idx];
        found = true;
        break;
      }
    }
    
    // タイムスタンプ・電話番号の特例
    if (!found && h === 'タイムスタンプ' && srcValues.length > 0) {
      val = srcValues[0];
      found = true;
    }
    
    if (h === '電話番号') {
      val = pickOriginalByAliases_('電話番号', srcHeaders, srcValues);
      val = formatTelForSheet_(val); // 先頭ゼロ保持（'付与でテキスト化）
      out.push(val);
      continue;
    }
    
    out.push(val);
  }
  
  return out;
}

/**
 * 電話番号のシート保存用フォーマット
 * 【説明】先頭'を付けてテキスト化（ゼロ保持）
 * @param {any} v - 元の値
 * @returns {string} フォーマット済み電話番号
 */
function formatTelForSheet_(v){
  let s = toHalfWidth_(String(v ?? '')).trim().replace(/[^\d]/g,'');
  if (!s) return '';
  return "'" + s;
}