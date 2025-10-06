/**
 * ScheduledTasks.gs V4.1 - ステータス監視による自動処理
 * - 担当者変更完了後のステータスを「更新完了」に修正
 */

/**
 * メイン処理：全シートの待ち状態を自動処理
 */
function processScheduledTasks() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    console.log('=== 別の処理が実行中のため終了 ===');
    return;
  }
  
  try {
    console.log('=== 自動処理開始 ===');
    const ss = SpreadsheetApp.openById(CONFIG.regSpreadsheetId);
    
    const targets = [
      { name: CONFIG.regSheetName, type: 'regular' },
      { name: SPOT_CONFIG.sheets.clinic, type: 'spot_clinic' },
      { name: SPOT_CONFIG.sheets.dealer, type: 'spot_dealer' }
    ];
    
    targets.forEach(target => {
      const sh = ss.getSheetByName(target.name);
      if (!sh) {
        console.log(`シート「${target.name}」が見つかりません。スキップします。`);
        return;
      }
      
      console.log(`\n--- シート「${target.name}」を処理中 ---`);
      try {
        processSheetTasks_(sh, target.type);
      } catch (e) {
        console.error(`シート「${target.name}」の処理でエラー:`, e);
        postChat_('エラー', `自動処理でエラーが発生しました\nシート: ${target.name}\nエラー: ${e.message}`, target.type);
      }
    });
    
    console.log('\n=== 自動処理完了 ===');
    
  } finally {
    lock.releaseLock();
  }
}

/**
 * 1つのシート内の待ち状態を処理
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh - 対象シート
 * @param {string} sheetType - シート種別
 */
function processSheetTasks_(sh, sheetType) {
  const headers = sheetHeaders_(sh);
  const statusCol = headers.indexOf('ステータス');
  if (statusCol < 0) return;
  
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  
  const statusValues = sh.getRange(2, statusCol + 1, lastRow - 1, 1).getValues();
  let processedCount = 0;
  const maxProcess = CONFIG.AUTO_PROCESS.maxProcessPerRun;
  
  for (let i = 0; i < statusValues.length; i++) {
    if (processedCount >= maxProcess) {
      console.log(`処理上限（${maxProcess}件）に達しました。`);
      break;
    }
    
    const row = 2 + i;
    const status = String(statusValues[i][0] || '').trim();
    
    try {
      if (status === STATUS.CALENDAR_REGISTER) {
        processRegisterTask_(sh, row, sheetType);
        processedCount++;
      } else if (status === STATUS.RESYNC_REGISTER) {
        processUpdateTask_(sh, row, sheetType);
        processedCount++;
      } else if (status === STATUS.CANCEL_REGISTER) {
        processCancelTask_(sh, row, sheetType);
        processedCount++;
      } else if (status === STATUS.ASSIGNEE_CHANGE_REGISTER) {
        handleAssigneeChange_(sh, row);
        // ★★★ 修正箇所 ★★★
        setCell_(sh, row, 'ステータス', STATUS.RESYNC_COMPLETE); // 完了後のステータスを「更新完了」に
        setLastUpdated_(sh, row, 'sheet');
        console.log(`  ${row}行目: 担当者変更完了`);
        processedCount++;
      }
    } catch (err) {
      console.error(`  ${row}行目でエラー:`, err);
      setCell_(sh, row, 'ステータス', STATUS.ERROR);
      setLastUpdated_(sh, row, 'sheet');
      const errorMsg = buildErrorNotification_(sh, row, status, err.message);
      postChat_('エラー', errorMsg, sheetType);
      processedCount++;
    }
  }
}

/**
 * カレンダー登録タスクを処理
 */
function processRegisterTask_(sh, row, sheetType) {
  console.log(`  ${row}行目: ${STATUS.CALENDAR_REGISTER} → 処理開始`);
  
  const eventId = getCell_(sh, row, 'カレンダーイベントID');
  if (eventId) throw new Error('既にカレンダーイベントが作成されています');
  
  if (sheetType === 'regular') {
    createCalendarEventFromRow_(sh, row);
  } else {
    createSpotCalendarEvent_(sh, row, sheetType.replace('spot_', ''));
  }
  
  setCell_(sh, row, 'ステータス', STATUS.CALENDAR_COMPLETE);
  setLastUpdated_(sh, row, 'sheet');
  console.log(`  ${row}行目: カレンダー登録完了`);
}

/**
 * イベント更新タスクを処理
 */
function processUpdateTask_(sh, row, sheetType) {
  console.log(`  ${row}行目: ${STATUS.RESYNC_REGISTER} → 処理開始`);
  
  const eventId = nz_(getCell_(sh, row, 'カレンダーイベントID'));
  if (!eventId) throw new Error('カレンダーイベントIDがありません');
  
  if (sheetType === 'regular') {
    updateCalendarEventTimes_(sh, row, eventId);
  } else {
    updateSpotCalendarEventTimes_(sh, row, eventId, sheetType.replace('spot_', ''));
  }
  
  setCell_(sh, row, 'ステータス', STATUS.RESYNC_COMPLETE);
  setLastUpdated_(sh, row, 'sheet');
  console.log(`  ${row}行目: 再同期完了`);
}

/**
 * イベント削除タスクを処理
 */
function processCancelTask_(sh, row, sheetType) {
  console.log(`  ${row}行目: ${STATUS.CANCEL_REGISTER} → 処理開始`);
  
  const eventId = nz_(getCell_(sh, row, 'カレンダーイベントID'));
  if (eventId) {
    deleteCalendarEvent_(sh, row, eventId);
  } else {
    console.log(`  ${row}行目: 既に削除済み（イベントID無し）`);
  }
  
  setCell_(sh, row, 'カレンダーイベントID', '');
  setCell_(sh, row, 'ステータス', STATUS.CANCEL_COMPLETE);
  setCell_(sh, row, '削除日時', new Date());
  setCell_(sh, row, '削除元', 'sheet');
  setLastUpdated_(sh, row, 'sheet');
  
  let title;
  if (sheetType === 'regular') {
    title = buildTitle_(sh, row);
  } else {
    title = buildSpotTitle_(sh, row, sheetType.replace('spot_', ''));
  }
  
  const rowUrl = getCell_(sh, row, '行URL') || `${sh.getParent().getUrl()}#gid=${sh.getSheetId()}&range=${row}:${row}`;
  postChat_('取消', `${title} の予定を取消しました\n${rowUrl}`, sheetType);
  console.log(`  ${row}行目: 取消完了`);
}

/**
 * エラー通知メッセージの組み立て
 */
function buildErrorNotification_(sh, row, taskType, errorMessage) {
  const clinicName = nz_(getCell_(sh, row, '医院名')) || '（医院名未設定）';
  const assignee = nz_(getCell_(sh, row, '回収担当者')) || '（担当者未設定）';
  const rowUrl = getCell_(sh, row, '行URL') || `${sh.getParent().getUrl()}#gid=${sh.getSheetId()}&range=${row}:${row}`;
  
  const lines = [];
  lines.push('【自動処理エラー】');
  lines.push('');
  lines.push(`種別：${taskType}`);
  lines.push(`医院名：${clinicName}`);
  lines.push(`担当者：${assignee}`);
  lines.push(`エラー内容：${errorMessage}`);
  lines.push('');
  lines.push(`行URL：${rowUrl}`);
  lines.push('');
  lines.push('※ ステータスを「エラー」に変更しました。');
  lines.push('※ 修正後、ステータスを元に戻すと再実行されます。');
  
  return lines.join('\n');
}