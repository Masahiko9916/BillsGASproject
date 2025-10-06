/**
 * SpotDataMapping.gs V4 - スポット回収データマッピング
 * 
 * 【Ver4の変更点】
 * - STATUS定数の使用
 * - エラーハンドリングの強化
 * 
 * 【このファイルの役割】
 * - スポット回収フォーム回答から登録シート形式への変換
 * - 医院用・ディーラー用それぞれのマッピング処理
 * - 住所結合、写真URL処理などの特殊処理
 * 
 * 【依存関係】
 * - 依存元：Config.gs、StringUtils.gs
 * - 依存先：なし
 * - 使用者：Main.gs（フォーム送信処理）
 */

/**
 * スポット回収_医院用のフォーム回答を内部形式に変換
 * @param {Array} dstHeaders - 出力先ヘッダー配列
 * @param {Array} srcHeaders - フォームヘッダー配列
 * @param {Array} srcValues - フォーム値配列
 * @returns {Array} 変換された行データ
 */
function buildSpotClinicRowByMapping_(dstHeaders, srcHeaders, srcValues) {
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
    
    // 特別処理：住所は二段を結合
    if (h === '住所') {
      try {
        const addrParts = [];
        const addr1 = pickValueByAliases_('住所', srcHeaders, srcValues, SPOT_CLINIC_ALIASES, srcNormToIndex);
        if (addr1 && String(addr1).trim()) {
          addrParts.push(String(addr1).trim());
        }
        out.push(addrParts.join(' '));
      } catch (e) {
        console.error('住所結合エラー:', e);
        out.push('');
      }
      continue;
    }
    
    // 特別処理：医院電話番号はテキスト化
    if (h === '医院電話番号') {
      try {
        const val = pickValueByAliases_(h, srcHeaders, srcValues, SPOT_CLINIC_ALIASES, srcNormToIndex);
        out.push(formatTelForSheet_(val));
      } catch (e) {
        console.error('医院電話番号変換エラー:', e);
        out.push('');
      }
      continue;
    }
    
    // 特別処理：代表電話番号もテキスト化
    if (h === '代表電話番号') {
      try {
        const val = pickValueByAliases_(h, srcHeaders, srcValues, SPOT_CLINIC_ALIASES, srcNormToIndex);
        out.push(formatTelForSheet_(val));
      } catch (e) {
        console.error('代表電話番号変換エラー:', e);
        out.push('');
      }
      continue;
    }
    
    // 通常処理：エイリアスを使って値を取得
    try {
      const val = pickValueByAliases_(h, srcHeaders, srcValues, SPOT_CLINIC_ALIASES, srcNormToIndex);
      
      // タイムスタンプの特例
      if (!val && h === 'タイムスタンプ' && srcValues.length > 0) {
        out.push(srcValues[0]);
        continue;
      }
      
      out.push(val);
    } catch (e) {
      console.error(`項目${h}の取得エラー:`, e);
      out.push('');
    }
  }
  
  return out;
}

/**
 * スポット回収_ディーラー用のフォーム回答を内部形式に変換
 * @param {Array} dstHeaders - 出力先ヘッダー配列
 * @param {Array} srcHeaders - フォームヘッダー配列
 * @param {Array} srcValues - フォーム値配列
 * @returns {Array} 変換された行データ
 */
function buildSpotDealerRowByMapping_(dstHeaders, srcHeaders, srcValues) {
  const now = new Date();
  const op = getActiveEmail_();
  const uuid = Utilities.getUuid();
  
  // フォームヘッダーの正規化インデックス作成
  const srcNormToIndex = {};
  for (let i = 0; i < srcHeaders.length; i++) {
    const norm = normalizeHeaderKey_(srcHeaders[i]);
    if (!(norm in srcNormToIndex)) srcNormToIndex[norm] = i;
  }
  
  // 希望日選択から回収予定日の初期値を取得（終了時間を30分後に）
  const hopeDate = pickValueByAliases_('希望日選択', srcHeaders, srcValues, SPOT_DEALER_ALIASES, srcNormToIndex);
  let initialDate = '', initialStart = '', initialEnd = '';
  
  if (hopeDate && hopeDate instanceof Date) {
    initialDate = new Date(hopeDate.getFullYear(), hopeDate.getMonth(), hopeDate.getDate());
    initialStart = hopeDate;
    initialEnd = new Date(hopeDate.getTime() + 30 * 60 * 1000); // 30分後
  } else if (hopeDate) {
    try {
      const d = new Date(hopeDate);
      if (!isNaN(d.getTime())) {
        initialDate = new Date(d.getFullYear(), d.getMonth(), d.getDate());
        initialStart = d;
        initialEnd = new Date(d.getTime() + 30 * 60 * 1000); // 30分後
      }
    } catch (e) {
      console.error('希望日変換エラー:', e);
    }
  }
  
  // 管理列のデフォルト値
  const managementDefaults = {
    'ステータス': STATUS.UNHANDLED,
    '回収予定日': initialDate,
    '開始時間': initialStart,
    '終了時間': initialEnd,
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
    
    // 特別処理：住所は二段を結合
    if (h === '住所') {
      try {
        const addrParts = [];
        const addr1 = pickValueByAliases_('住所', srcHeaders, srcValues, SPOT_DEALER_ALIASES, srcNormToIndex);
        if (addr1 && String(addr1).trim()) {
          addrParts.push(String(addr1).trim());
        }
        out.push(addrParts.join(' '));
      } catch (e) {
        console.error('住所結合エラー:', e);
        out.push('');
      }
      continue;
    }
    
    // 特別処理：医院電話番号はテキスト化
    if (h === '医院電話番号') {
      try {
        const val = pickValueByAliases_(h, srcHeaders, srcValues, SPOT_DEALER_ALIASES, srcNormToIndex);
        out.push(formatTelForSheet_(val));
      } catch (e) {
        console.error('医院電話番号変換エラー:', e);
        out.push('');
      }
      continue;
    }
    
    // 特別処理：ディーラー電話番号もテキスト化
    if (h === 'ディーラー電話番号') {
      try {
        const val = pickValueByAliases_(h, srcHeaders, srcValues, SPOT_DEALER_ALIASES, srcNormToIndex);
        out.push(formatTelForSheet_(val));
      } catch (e) {
        console.error('ディーラー電話番号変換エラー:', e);
        out.push('');
      }
      continue;
    }
    
    // 通常処理：エイリアスを使って値を取得
    try {
      const val = pickValueByAliases_(h, srcHeaders, srcValues, SPOT_DEALER_ALIASES, srcNormToIndex);
      
      // タイムスタンプの特例
      if (!val && h === 'タイムスタンプ' && srcValues.length > 0) {
        out.push(srcValues[0]);
        continue;
      }
      
      out.push(val);
    } catch (e) {
      console.error(`項目${h}の取得エラー:`, e);
      out.push('');
    }
  }
  
  return out;
}

/**
 * エイリアスを使って値を取得
 * @param {string} headerName - 検索するヘッダー名
 * @param {Array} srcHeaders - フォームヘッダー配列
 * @param {Array} srcValues - フォーム値配列
 * @param {Object} aliases - エイリアス定義オブジェクト
 * @param {Object} srcNormToIndex - 正規化済みインデックス
 * @returns {any} 見つかった値（見つからない場合は空文字）
 */
function pickValueByAliases_(headerName, srcHeaders, srcValues, aliases, srcNormToIndex) {
  const candidates = [headerName].concat(aliases[headerName] || []);
  
  for (const cand of candidates) {
    let idx = srcHeaders.indexOf(cand);
    
    // 完全一致で見つからない場合は正規化して検索
    if (idx === -1) {
      const normCand = normalizeHeaderKey_(cand || '');
      if (srcNormToIndex.hasOwnProperty(normCand)) {
        idx = srcNormToIndex[normCand];
      } else {
        idx = srcHeaders.findIndex(sh => normalizeHeaderKey_(sh).includes(normCand));
      }
    }
    
    if (idx !== -1 && idx < srcValues.length) {
      return srcValues[idx];
    }
  }
  
  return '';
}