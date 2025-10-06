/**
 * Notifications.gs V4 - 通知機能
 * 
 * 【Ver4の変更点】
 * - buildErrorNotification_をScheduledTasks.gsから移動（重複解消）
 * - STATUS定数を使用
 * 
 * 【このファイルの役割】
 * - Google Chat Webhookを使った通知送信
 * - 各種メッセージ文の組み立て
 * - 受付・登録・取消等の状況通知
 * 
 * 【依存関係】
 * - 依存元：Config.gs（Webhook URL、メンション設定）、SheetUtils.gs、StringUtils.gs
 * - 依存先：なし
 * - 使用者：Main.gs、RegularCollection.gs、ScheduledTasks.gs等
 */

/**
 * Google Chat 通知送信【Ver4：整合性保持】
 * 【説明】設定されたWebhook URLにメッセージを送信
 * - シート種別に応じて通知先を自動振り分け
 * @param {string} kind - 通知種別（受付、カレンダー登録完了、担当者変更完了等）
 * @param {string} text - 送信するメッセージ本文
 * @param {string} sheetType - シート種別（'regular', 'spot_clinic', 'spot_dealer'）
 */
function postChat_(kind, text, sheetType) {
  // 通知先URLの決定
  let webhookUrl = CONFIG.chatWebhookUrl;
  
  if (sheetType === 'spot_clinic' && SPOT_CONFIG.webhooks.clinic) {
    webhookUrl = SPOT_CONFIG.webhooks.clinic;
  } else if (sheetType === 'spot_dealer' && SPOT_CONFIG.webhooks.dealer) {
    webhookUrl = SPOT_CONFIG.webhooks.dealer;
  }
  
  if (!webhookUrl || webhookUrl.startsWith('<<')) return;
  
  // メンション機能
  const mentionEnabled = Array.isArray(CONFIG.chatMentionOnKinds) && 
                         CONFIG.chatMentionOnKinds.indexOf(kind) !== -1 && 
                         CONFIG.chatMentionUser;
  const mention = mentionEnabled ? `<${CONFIG.chatMentionUser}> ` : '';
  
  const payload = { text: `${mention}【${kind}】\n${text}` };
  
  try {
    UrlFetchApp.fetch(webhookUrl, {
      method: 'post', 
      contentType: 'application/json',
      payload: JSON.stringify(payload), 
      muteHttpExceptions: true
    });
  } catch (e) {
    console.error('Chat通知送信エラー:', e);
  }
}

/**
 * 受付通知メッセージの組み立て（定期回収用）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh - 登録シート
 * @param {number} row - 対象行番号
 * @returns {string} 受付通知メッセージ
 */
function buildReceiptSummary_(sh, row) {
  const clinic = nz_(getCell_(sh, row, '医院名'));
  const person = nz_(getCell_(sh, row, '担当者名'));
  const tel = formatTelForChat_(getCell_(sh, row, '電話番号'));
  const addr = nz_(getCell_(sh, row, '住所'));
  
  const lines = [];
  lines.push('新規申請を受け付けました。');
  lines.push(`・医院名: ${clinic}`);
  lines.push(`・担当者: ${person}`);
  if (tel) lines.push(`・電話: ${tel}`);
  if (addr) lines.push(`・住所: ${addr}`);
  lines.push('');
  lines.push('回収廃棄物');
  
  const wastes = collectWasteLinesForChat_(sh, row);
  if (wastes.length === 0) {
    lines.push('・（数量の入力がありません）');
  } else {
    wastes.forEach(w => lines.push(`・${w}`));
  }
  
  return lines.join('\n');
}

/**
 * スポット回収の受付通知メッセージ組み立て（医院用）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh - 登録シート
 * @param {number} row - 対象行番号
 * @returns {string} 受付通知メッセージ
 */
function buildSpotClinicReceiptSummary_(sh, row) {
  const clinic = nz_(getCell_(sh, row, '医院名'));
  const person = nz_(getCell_(sh, row, '担当者名'));
  const tel = formatTelForChat_(getCell_(sh, row, '医院電話番号'));
  const addr = nz_(getCell_(sh, row, '住所'));
  const requestType = nz_(getCell_(sh, row, '依頼種別'));
  
  const lines = [];
  lines.push('【スポット回収_医院】新規申請を受け付けました。');
  lines.push(`・医院名: ${clinic}`);
  lines.push(`・担当者: ${person}`);
  if (tel) lines.push(`・電話: ${tel}`);
  if (addr) lines.push(`・住所: ${addr}`);
  lines.push(`・依頼種別: ${requestType}`);
  lines.push('');
  lines.push('回収品目');
  
  const items = collectSpotClinicItems_(sh, row);
  if (items.length === 0) {
    lines.push('・（品目の入力がありません）');
  } else {
    items.forEach(item => lines.push(`・${item}`));
  }
  
  return lines.join('\n');
}

/**
 * スポット回収の受付通知メッセージ組み立て（ディーラー用）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh - 登録シート
 * @param {number} row - 対象行番号
 * @returns {string} 受付通知メッセージ
 */
function buildSpotDealerReceiptSummary_(sh, row) {
  const clinic = nz_(getCell_(sh, row, '医院名'));
  const clinicPerson = nz_(getCell_(sh, row, '医院担当者名'));
  const clinicTel = formatTelForChat_(getCell_(sh, row, '医院電話番号'));
  const addr = nz_(getCell_(sh, row, '住所'));
  const requestType = nz_(getCell_(sh, row, '依頼種別'));
  
  const company = nz_(getCell_(sh, row, '会社名'));
  const branch = nz_(getCell_(sh, row, '支店名'));
  const dealerPerson = nz_(getCell_(sh, row, 'ディーラー担当者名'));
  const dealerTel = formatTelForChat_(getCell_(sh, row, 'ディーラー電話番号'));
  
  const hopeDate = getCell_(sh, row, '希望日選択');
  const hopeDateStr = hopeDate ? Utilities.formatDate(new Date(hopeDate), CONFIG.tz, 'yyyy-MM-dd HH:mm') : '';
  
  const lines = [];
  lines.push('【スポット回収_ディーラー】新規申請を受け付けました。');
  lines.push('');
  lines.push('＜医療機関情報＞');
  lines.push(`・医院名: ${clinic}`);
  lines.push(`・担当者: ${clinicPerson}`);
  if (clinicTel) lines.push(`・電話: ${clinicTel}`);
  if (addr) lines.push(`・住所: ${addr}`);
  lines.push('');
  lines.push('＜ディーラー情報＞');
  lines.push(`・会社名: ${company}`);
  if (branch) lines.push(`・支店名: ${branch}`);
  lines.push(`・担当者: ${dealerPerson}`);
  if (dealerTel) lines.push(`・電話: ${dealerTel}`);
  lines.push(`・依頼種別: ${requestType}`);
  if (hopeDateStr) lines.push(`・回収希望日: ${hopeDateStr}`);
  lines.push('');
  lines.push('回収品目');
  
  const items = collectSpotDealerItems_(sh, row);
  if (items.length === 0) {
    lines.push('・（品目の入力がありません）');
  } else {
    items.forEach(item => lines.push(`・${item}`));
  }
  
  return lines.join('\n');
}

/**
 * スポット医院の回収品目を収集
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh - 登録シート
 * @param {number} row - 対象行番号
 * @returns {Array<string>} 品目リスト
 */
function collectSpotClinicItems_(sh, row) {
  const items = [];
  
  if (getCell_(sh, row, '医療機器_有無') || getCell_(sh, row, '医療機器_名称数量')) {
    items.push('✓ 医療機器');
  }
  
  if (getCell_(sh, row, '什器備品_名称数量') || getCell_(sh, row, '家電_名称数量')) {
    items.push('✓ 什器・備品・家電');
  }
  
  if (getCell_(sh, row, '書類_梱包済み') || getCell_(sh, row, '書類_ダンボール数')) {
    items.push('✓ 書類');
  }
  
  if (getCell_(sh, row, '廃液薬品_名称容量数量')) {
    items.push('✓ 廃液・薬品');
  }
  
  if (getCell_(sh, row, 'その他廃棄物')) {
    items.push('✓ その他廃棄物');
  }
  
  return items;
}

/**
 * スポットディーラーの回収品目を収集
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh - 登録シート
 * @param {number} row - 対象行番号
 * @returns {Array<string>} 品目リスト
 */
function collectSpotDealerItems_(sh, row) {
  const items = [];
  
  if (getCell_(sh, row, '医療機器_有無') || getCell_(sh, row, '医療機器_名称数量')) {
    items.push('✓ 医療機器（搬出対応情報あり）');
  }
  
  if (getCell_(sh, row, '什器備品_名称数量') || getCell_(sh, row, '家電_名称数量')) {
    items.push('✓ 什器・備品・家電（搬出対応情報あり）');
  }
  
  if (getCell_(sh, row, '書類_梱包済み') || getCell_(sh, row, '書類_ダンボール数')) {
    items.push('✓ 書類');
  }
  
  if (getCell_(sh, row, '廃液薬品_名称容量数量')) {
    items.push('✓ 廃液・薬品');
  }
  
  if (getCell_(sh, row, 'その他廃棄物')) {
    items.push('✓ その他廃棄物');
  }
  
  return items;
}

/**
 * カレンダー登録通知メッセージの組み立て
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh - 登録シート
 * @param {number} row - 対象行番号
 * @returns {string} 登録通知メッセージ
 */
function buildRegisterNoticeMessage_(sh, row) {
  const clinic = nz_(getCell_(sh, row, '医院名')) || '医院名未設定';
  const assignee = nz_(getCell_(sh, row, '回収担当者')) || '担当未設定';
  const dateCell = getCell_(sh, row, '回収予定日');
  const dateText = dateCell ? formatJaDateWithWeekday_(dateCell) : '日付未設定';
  
  return [
    clinic,
    assignee,
    '',
    'いつもお世話になっております。',
    '廃棄物回収の訪問日',
    dateText,
    '上記の日程でお伺いいたします。',
    'なお、時間帯は分かりかねますので、ご了承ください。',
    'よろしくお願い致します。'
  ].join('\n');
}

/**
 * 日本語曜日付き日付フォーマット
 * @param {Date|any} v - 日付値
 * @returns {string} フォーマット済み日付文字列
 */
function formatJaDateWithWeekday_(v) {
  const d = (v instanceof Date) ? v : new Date(v);
  const md = Utilities.formatDate(d, CONFIG.tz, 'M月d日');
  const yo = ['日','月','火','水','木','金','土'][d.getDay()];
  return md + yo + '曜日';
}

/**
 * カレンダータイトル組み立て（定期回収用）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh - 登録シート
 * @param {number} row - 対象行番号
 * @returns {string} カレンダータイトル
 */
function buildTitle_(sh, row) {
  const clinic = nz_(getCell_(sh, row, '医院名'));
  const area = nz_(getCell_(sh, row, 'エリア'));
  
  if (area && clinic) return `${area}　${clinic}`;
  if (area) return area;
  return clinic || '回収';
}

/**
 * カレンダータイトル組み立て（スポット回収用）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh - 登録シート
 * @param {number} row - 対象行番号
 * @param {string} type - 'clinic' or 'dealer'
 * @returns {string} カレンダータイトル
 */
function buildSpotTitle_(sh, row, type) {
  const clinic = nz_(getCell_(sh, row, '医院名'));
  const area = nz_(getCell_(sh, row, 'エリア'));
  
  if (area && clinic) return `${area}　${clinic}`;
  if (area) return area;
  return clinic || 'スポット回収';
}

/**
 * 回収廃棄物一覧の組み立て（定期回収用）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh - 登録シート
 * @param {number} row - 対象行番号
 * @returns {Array<string>} 廃棄物リスト
 */
function collectWasteLinesForChat_(sh, row) {
  const out = [];
  const qtyMap = [
    {h:'50L容器', label:'50L容器', unit:'個'},
    {h:'30L容器', label:'30L容器', unit:'個'},
    {h:'20L容器', label:'20L容器', unit:'個'},
    {h:'感染性廃棄物', label:'ダンボール箱', unit:'箱'},
    {h:'廃プラ袋', label:'廃プラ袋', unit:'袋'},
    {h:'ガラス袋', label:'ガラス・石膏くず袋', unit:'袋'},
    {h:'印象歯袋', label:'印象歯袋', unit:'袋'},
    {h:'石膏くず袋', label:'石膏くず袋', unit:'袋'},
    {h:'定着液', label:'定着液', unit:'個'},
    {h:'現像液', label:'現像液', unit:'個'},
    {h:'金属くず袋', label:'金属くず袋', unit:'袋'}
  ];
  
  qtyMap.forEach(({h, label, unit}) => {
    const q = toIntQty_(getCell_(sh, row, h));
    if (q > 0) out.push(`${label} ${q}${unit}`);
  });
  
  return out;
}