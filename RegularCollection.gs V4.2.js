/**
 * RegularCollection.gs - カレンダー機能
 * - 削除済みイベントのIDをクリアするよう修正
 * - カレンダーからの担当者変更検知時のステータスを「担当者変更あり」に設定
 * - deleteCalendarEvent_ のエラーハンドリングを強化
 * - 【改善】担当者変更の入力形式を「」で囲む形式に変更し、堅牢性を向上
 */

/**
 * 登録シートの行からカレンダーイベントを作成（定期回収用）【Ver4：Calendar API使用】
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh - 登録シート
 * @param {number} row - 対象行番号
 */
function createCalendarEventFromRow_(sh, row) {
  const d = getCell_(sh, row, '回収予定日');
  const s = getCell_(sh, row, '開始時間');
  const e = getCell_(sh, row, '終了時間');
  
  if (!d || !s || !e) {
    throw new Error('回収予定日/開始時間/終了時間は必須です');
  }
  
  const area = nz_(getCell_(sh, row, 'エリア'));
  if (!area) {
    throw new Error('エリアが未入力のため登録できません。');
  }
  
  const assigneeName = nz_(getCell_(sh, row, '回収担当者'));
  if (!assigneeName) {
    throw new Error('回収担当者が未選択のため登録できません。');
  }

  if (!validateAvailability_(sh, row, d, s, e)) {
    throw new Error('回収不可能日の曜日に該当しています（登録不可）');
  }

  const email = getAssigneeEmailByName_(assigneeName);
  if (!email) {
    throw new Error('担当者マスタにメールアドレスが見つかりません: ' + assigneeName);
  }

  const start = mergeDateTime_(d, s);
  const end = mergeDateTime_(d, e);
  const title = buildTitle_(sh, row);
  const description = buildDescription_(sh, row, d, s, e);
  const location = buildLocation_(sh, row);

  // Calendar API使用（CalendarAppは使用しない）
  const event = {
    summary: title,
    location: location,
    description: description,
    start: { dateTime: start.toISOString() },
    end: { dateTime: end.toISOString() },
    guestsCanModify: true
  };
  
  const createdEvent = Calendar.Events.insert(event, email);

  setLastUpdated_(sh, row, 'sheet', {
    'カレンダーイベントID': createdEvent.id,
    'カレンダーID': email
  });

  postChat_('カレンダー登録完了', title + '\nイベントID: ' + createdEvent.id + '\n' + getCell_(sh, row, '行URL'), 'regular');
  postChat_('カレンダー登録通知', buildRegisterNoticeMessage_(sh, row), 'regular');
}

/**
 * スポット回収のカレンダーイベント作成【Ver4：Calendar API使用】
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh - 登録シート
 * @param {number} row - 対象行番号
 * @param {string} type - 'clinic' or 'dealer'
 */
function createSpotCalendarEvent_(sh, row, type) {
  const d = getCell_(sh, row, '回収予定日');
  const s = getCell_(sh, row, '開始時間');
  const e = getCell_(sh, row, '終了時間');
  
  if (!d || !s || !e) {
    throw new Error('回収予定日/開始時間/終了時間は必須です');
  }
  
  const area = nz_(getCell_(sh, row, 'エリア'));
  if (!area) {
    throw new Error('エリアが未入力のため登録できません。');
  }
  
  const assigneeName = nz_(getCell_(sh, row, '回収担当者'));
  if (!assigneeName) {
    throw new Error('回収担当者が未選択のため登録できません。');
  }

  const email = getAssigneeEmailByName_(assigneeName);
  if (!email) {
    throw new Error('担当者マスタにメールアドレスが見つかりません: ' + assigneeName);
  }

  const start = mergeDateTime_(d, s);
  const end = mergeDateTime_(d, e);
  const title = buildSpotTitle_(sh, row, type);
  const description = buildSpotDescription_(sh, row, d, s, e, type);
  const location = buildSpotLocation_(sh, row);

  // Calendar API使用
  const event = {
    summary: title,
    location: location,
    description: description,
    start: { dateTime: start.toISOString() },
    end: { dateTime: end.toISOString() },
    guestsCanModify: true
  };

  const createdEvent = Calendar.Events.insert(event, email);

  setLastUpdated_(sh, row, 'sheet', {
    'カレンダーイベントID': createdEvent.id,
    'カレンダーID': email
  });

  const sheetType = (type === 'clinic') ? 'spot_clinic' : 'spot_dealer';
  postChat_('カレンダー登録完了', title + '\nイベントID: ' + createdEvent.id + '\n' + getCell_(sh, row, '行URL'), sheetType);
}

/**
 * カレンダーイベントの時間更新【Ver4：Calendar API使用】
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh - 登録シート
 * @param {number} row - 対象行番号
 * @param {string} eventId - イベントID
 */
function updateCalendarEventTimes_(sh, row, eventId) {
  const d = getCell_(sh, row, '回収予定日');
  const s = getCell_(sh, row, '開始時間');
  const e = getCell_(sh, row, '終了時間');
  
  if (!d || !s || !e) {
    throw new Error('日時が未入力です');
  }
  
  const start = mergeDateTime_(d, s);
  const end = mergeDateTime_(d, e);
  
  const calId = nz_(getCell_(sh, row, 'カレンダーID'));
  if (!calId) {
    throw new Error('カレンダーIDが見つかりません');
  }
  
  try {
    Calendar.Events.patch(
      {
        start: { dateTime: start.toISOString() },
        end: { dateTime: end.toISOString() },
        description: buildDescription_(sh, row, d, s, e),
        location: buildLocation_(sh, row)
      },
      calId,
      normalizeEventId_(eventId)
    );
  } catch (e) {
    postChat_('エラー', `イベント更新に失敗しました（ID: ${eventId}）\nエラー: ${e.message}`, 'regular');
    throw new Error('イベント更新に失敗しました: ' + e.message);
  }
}

/**
 * スポット回収のイベント時間更新【Ver4：Calendar API使用】
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh - 登録シート
 * @param {number} row - 対象行番号
 * @param {string} eventId - イベントID
 * @param {string} type - 'clinic' or 'dealer'
 */
function updateSpotCalendarEventTimes_(sh, row, eventId, type) {
  const d = getCell_(sh, row, '回収予定日');
  const s = getCell_(sh, row, '開始時間');
  const e = getCell_(sh, row, '終了時間');
  
  if (!d || !s || !e) {
    throw new Error('日時が未入力です');
  }
  
  const start = mergeDateTime_(d, s);
  const end = mergeDateTime_(d, e);
  
  const calId = nz_(getCell_(sh, row, 'カレンダーID'));
  if (!calId) {
    throw new Error('カレンダーIDが見つかりません');
  }
  
  const sheetType = (type === 'clinic') ? 'spot_clinic' : 'spot_dealer';
  
  try {
    Calendar.Events.patch(
      {
        start: { dateTime: start.toISOString() },
        end: { dateTime: end.toISOString() },
        description: buildSpotDescription_(sh, row, d, s, e, type),
        location: buildSpotLocation_(sh, row)
      },
      calId,
      normalizeEventId_(eventId)
    );
  } catch (e) {
    postChat_('エラー', `イベント更新に失敗しました（ID: ${eventId}）\nエラー: ${e.message}`, sheetType);
    throw new Error('イベント更新に失敗しました: ' + e.message);
  }
}

/**
 * カレンダーイベントの削除【Ver4：Calendar API使用】
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh - 登録シート
 * @param {number} row - 対象行番号
 * @param {string} eventId - イベントID
 */
function deleteCalendarEvent_(sh, row, eventId) {
  const calId = nz_(getCell_(sh, row, 'カレンダーID'));
  if (!calId) {
    throw new Error('イベントの削除に失敗しました（カレンダーIDが不明）');
  }
  
  try {
    Calendar.Events.remove(calId, normalizeEventId_(eventId));
  } catch (e) {
    if (e.message.includes('Not Found') || e.message.includes('Resource has been deleted')) {
      console.warn(`イベント(ID: ${eventId})は既に削除されているようです。処理を続行します。`);
    } else {
      console.error('Calendar API削除失敗:', e.message);
      throw new Error(`イベントの削除に失敗しました: ${e.message}`);
    }
  }
}

/**
 * 担当者変更時の自動カレンダー移管【Ver4：Calendar API使用】
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh - 登録シート
 * @param {number} row - 対象行番号
 */
function handleAssigneeChange_(sh, row) {
  const eventIdRaw = nz_(getCell_(sh, row, 'カレンダーイベントID'));
  if (!eventIdRaw) return;

  const newAssigneeName = nz_(getCell_(sh, row, '回収担当者'));
  if (!newAssigneeName) return;

  const newCalId = getAssigneeEmailByName_(newAssigneeName);
  if (!newCalId) {
    throw new Error('担当者マスタにメールが見つかりません: ' + newAssigneeName);
  }

  const oldCalId = nz_(getCell_(sh, row, 'カレンダーID'));
  const normEventId = normalizeEventId_(eventIdRaw);

  if (oldCalId === newCalId) return;

  transferEventByCopyAndDelete_(oldCalId, normEventId, newCalId, sh, row);
}

/**
 * Calendar API: イベントをコピー＆削除で移管する【Ver4：エラー解決版】
 * @param {string} fromCalId - 移動元カレンダーID
 * @param {string} eventId - イベントID
 * @param {string} toCalId - 移動先カレンダーID
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh - 登録シート
 * @param {number} row - 対象行番号
 */
function transferEventByCopyAndDelete_(fromCalId, eventId, toCalId, sh, row) {
  try {
    const originalEvent = Calendar.Events.get(fromCalId, eventId);
    
    const newAssigneeName = nz_(getCell_(sh, row, '回収担当者'));

    const updatedDescription = String(originalEvent.description || '').replace(
      /変更先担当者[：:][^\n]+/,
      `変更先担当者：「${newAssigneeName}」 → 処理済み`
    );
    
    const newEventData = {
      summary: originalEvent.summary,
      description: updatedDescription,
      location: originalEvent.location,
      start: originalEvent.start,
      end: originalEvent.end,
      guestsCanModify: true
    };

    const createdEvent = Calendar.Events.insert(newEventData, toCalId);

    Calendar.Events.remove(fromCalId, eventId);
    
    setLastUpdated_(sh, row, 'calendar', {
      'カレンダーID': toCalId,
      'カレンダーイベントID': createdEvent.id
    });

  } catch (e) {
    const msg = (e && e.message) ? e.message : String(e);
    throw new Error('イベント移管（コピー＆削除）に失敗しました: ' + msg);
  }
}


/**
 * 全ての対象シートのカレンダー差分同期（5分おき）【ロック機能追加】
 */
function syncCalendarToSheet() {
  // ★★★ 修正箇所 START ★★★
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    console.log('=== 別の処理（sync or processTasks）が実行中のため同期処理を終了 ===');
    return;
  }
  // ★★★ 修正箇所 END ★★★
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.regSpreadsheetId);

    const targetSheets = [
      CONFIG.regSheetName,
      SPOT_CONFIG.sheets.clinic,
      SPOT_CONFIG.sheets.dealer
    ];

    targetSheets.forEach(sheetName => {
      const sh = ss.getSheetByName(sheetName);
      if (!sh) {
        console.warn(`シート「${sheetName}」が見つからないため、同期をスキップします。`);
        return;
      }
      
      console.log(`--- シート「${sheetName}」の同期を開始 ---`);
      try {
        syncSingleSheet_(sh);
      } catch(e) {
        console.error(`シート「${sheetName}」の同期中にエラーが発生しました: ${e.message}`);
      }
      console.log(`--- シート「${sheetName}」の同期を完了 ---`);
    });
  // ★★★ 修正箇所 START ★★★
  } finally {
    lock.releaseLock();
  }
  // ★★★ 修正箇所 END ★★★
}

/**
 * 1つのシートのカレンダー同期処理（内部関数）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh - 対象シート
 */
function syncSingleSheet_(sh) {
  const headers = sheetHeaders_(sh).map(h => String(h).trim());
  const idCol = headers.indexOf('カレンダーイベントID');
  if (idCol < 0) {
    throw new Error(`シート「${sh.getName()}」のヘッダーに「カレンダーイベントID」がありません。`);
  }
  
  const calIdCol = headers.indexOf('カレンダーID');
  const idMap = new Map();
  const lastRow = sh.getLastRow();
  
  if (lastRow >= 2) {
    const ids = sh.getRange(2, idCol + 1, lastRow - 1, 1).getValues();
    const calIds = calIdCol >= 0 ? sh.getRange(2, calIdCol + 1, lastRow - 1, 1).getValues() : [];
    
    for (let i = 0; i < ids.length; i++) {
      const raw = String(ids[i][0] || '').trim();
      if (!raw) continue;
      
      const norm = normalizeEventId_(raw);
      const cal = calIdCol >= 0 ? String((calIds[i] || [])[0] || '').trim() : '';
      
      idMap.set(`${cal}|${norm}`, 2 + i);
      idMap.set(`${cal}|${norm}@google.com`, 2 + i);
      if (!cal) {
        idMap.set(`|${norm}`, 2 + i);
        idMap.set(`|${norm}@google.com`, 2 + i);
      }
    }
  }
  
  const targetCalIds = collectActiveCalendarIdsFromSheet_(sh, headers);
  const limits = CONFIG.SYNC_LIMITS || { maxCalendarsPerRun: 20, maxPagesPerCalendar: 3, maxEventsPerRun: 500 };
  
  let calendarsProcessed = 0;
  let totalEventsProcessed = 0;
  
  for (const calId of targetCalIds) {
    if (calendarsProcessed >= limits.maxCalendarsPerRun) break;
    calendarsProcessed++;
    
    const props = PropertiesService.getScriptProperties();
    const tokenKey = `${CONFIG.PROP_SYNC_TOKEN_PREFIX}:${calId}`;
    let syncToken = props.getProperty(tokenKey);
    let pageToken = null, wroteToken = false, safetyCounter = 0, pages = 0;
    
    try {
      do {
        if (++safetyCounter > 200) throw new Error('ページングが多すぎます。');
        if (pages >= limits.maxPagesPerCalendar) break;
        if (totalEventsProcessed >= limits.maxEventsPerRun) break;
        
        const params = { maxResults: 250, showDeleted: true, pageToken: pageToken || undefined };
        if (syncToken) {
          params.syncToken = syncToken;
        } else {
          params.updatedMin = new Date(Date.now() - CONFIG.initialLookbackDays * 86400000).toISOString();
        }
        
        const resp = Calendar.Events.list(calId, params);
        const items = resp.items || [];
        
        for (let i = 0; i < items.length; i++) {
          upsertRowFromEventMulti_(sh, headers, idMap, calId, items[i]);
          totalEventsProcessed++;
          if (totalEventsProcessed >= limits.maxEventsPerRun) break;
        }
        
        pageToken = resp.nextPageToken || null;
        if (!pageToken && resp.nextSyncToken) {
          props.setProperty(tokenKey, resp.nextSyncToken);
          wroteToken = true;
        }
        pages++;
      } while (pageToken && totalEventsProcessed < limits.maxEventsPerRun);
      
    } catch (e) {
      const msg = String(e && e.message || e);
      if (msg.includes('Sync token is no longer valid') || msg.includes('410')) {
        props.deleteProperty(tokenKey);
        if (!wroteToken) {
          const resp = Calendar.Events.list(calId, {
            maxResults: 250,
            showDeleted: true,
            updatedMin: new Date(Date.now() - CONFIG.initialLookbackDays * 86400000).toISOString()
          });
          (resp.items || []).forEach(ev => upsertRowFromEventMulti_(sh, headers, idMap, calId, ev));
          if (resp.nextSyncToken) props.setProperty(tokenKey, resp.nextSyncToken);
        }
      } else {
        throw e;
      }
    }
    
    if (totalEventsProcessed >= limits.maxEventsPerRun) break;
  }
}

/**
 * シートからアクティブなカレンダーIDを収集
 */
function collectActiveCalendarIdsFromSheet_(sh, headers){
  const calIdCol = headers.indexOf('カレンダーID');
  const set = new Set();
  
  if (calIdCol >= 0) {
    const lastRow = sh.getLastRow();
    if (lastRow >= 2) {
      const vals = sh.getRange(2, calIdCol+1, lastRow-1, 1).getValues();
      vals.forEach(v => { 
        const s = String(v[0]||'').trim(); 
        if (s) set.add(s); 
      });
    }
  }
  
  if (set.size === 0 && CONFIG.calendarId) {
    set.add(CONFIG.calendarId);
  }
  
  return Array.from(set);
}

/**
 * カレンダーイベントから行を更新/作成
 */
function upsertRowFromEventMulti_(sh, headers, idMap, calId, ev) {
  const id = String(ev.id || '').trim();
  if (!id) return;
  
  const row = idMap.get(`${calId}|${id}`) || 
              idMap.get(`${calId}|${id}@google.com`) || 
              idMap.get(`|${id}`) || 
              idMap.get(`|${id}@google.com`);
              
  const statusCol = headers.indexOf('ステータス');
  const dateCol = headers.indexOf('回収予定日');
  const startCol = headers.indexOf('開始時間');
  const endCol = headers.indexOf('終了時間');
  const updCol = headers.indexOf('最終更新時刻');
  const updSrcCol = headers.indexOf('最終更新元');
  const delAtCol = headers.indexOf('削除日時');
  const delByCol = headers.indexOf('削除元');
  const calIdCol = headers.indexOf('カレンダーID');
  const evIdCol = headers.indexOf('カレンダーイベントID');
  const assigneeCol = headers.indexOf('回収担当者');
  
  if (!row) {
    return;
  }
  
  const isCancelled = ev.status === 'cancelled';
  
  if (!isCancelled && ev.description) {
    try {
      const isAssigneeChanged = processAssigneeChangeFromCalendar_(sh, row, ev, calId);
      if (isAssigneeChanged) {
        return;
      }
    } catch (changeErr) {
      console.error('担当者変更処理エラー:', changeErr);
    }
  }
  
  if (isCancelled) {
    if (statusCol >= 0) sh.getRange(row, statusCol+1).setValue(STATUS.CALENDAR_DELETED);
    if (delAtCol >= 0) sh.getRange(row, delAtCol+1).setValue(ev.updated ? new Date(ev.updated) : new Date());
    if (delByCol >= 0) sh.getRange(row, delByCol+1).setValue('calendar');
    
    if (evIdCol >= 0) sh.getRange(row, evIdCol+1).setValue('');
    if (calIdCol >= 0) sh.getRange(row, calIdCol+1).setValue('');

  } else {
    const times = extractTimesFromEvent_(ev);
    if (dateCol >= 0 && times.date) sh.getRange(row, dateCol+1).setValue(times.date);
    if (startCol >= 0 && times.start) sh.getRange(row, startCol+1).setValue(times.start);
    if (endCol >= 0 && times.end) sh.getRange(row, endCol+1).setValue(times.end);
    
    const currentStatus = getCell_(sh, row, 'ステータス');
    if (currentStatus !== STATUS.ASSIGNEE_CHANGED_FROM_CAL) {
      if (statusCol >= 0) sh.getRange(row, statusCol+1).setValue(STATUS.CALENDAR_COMPLETE);
    }
  }
  
  if (updCol >= 0) sh.getRange(row, updCol+1).setValue(new Date());
  if (updSrcCol >= 0) sh.getRange(row, updSrcCol+1).setValue('calendar');
  
  const name = getAssigneeNameByEmail_(calId);
  if (name && assigneeCol >= 0) sh.getRange(row, assigneeCol+1).setValue(name);
}

/**
 * カレンダーからの担当者変更処理
 * @returns {boolean} 担当者変更が処理された場合は true を返す
 */
function processAssigneeChangeFromCalendar_(sh, row, ev, currentCalId) {
  const description = String(ev.description || '');
  const changeInfo = extractAssigneeChangeFromDescription_(description);
  
  if (!changeInfo || !changeInfo.newAssigneeName || changeInfo.processed) {
    return false;
  }
  
  const newAssigneeName = changeInfo.newAssigneeName.trim();
  if (!newAssigneeName || newAssigneeName.includes('（ここに')) {
    return false;
  }

  const newCalId = getAssigneeEmailByName_(newAssigneeName);
  if (!newCalId) {
    updateDescriptionWithError_(ev, currentCalId, newAssigneeName);
    return false;
  }
  
  if (currentCalId === newCalId) {
    updateDescriptionAsProcessed_(ev, currentCalId, newAssigneeName);
    return false;
  }
  
  try {
    const normEventId = normalizeEventId_(ev.id);
    transferEventByCopyAndDelete_(currentCalId, normEventId, newCalId, sh, row);
    
    setLastUpdated_(sh, row, 'calendar', {
      '回収担当者': newAssigneeName,
      'ステータス': STATUS.ASSIGNEE_CHANGED_FROM_CAL
    });
    
    postChat_('担当者変更完了（カレンダー操作）', 
      `カレンダーから担当者が変更されました\n` +
      `新担当者: ${newAssigneeName}\n` +
      `行URL: ${getCell_(sh, row, '行URL')}`,
      sh.getName().includes('スポット') ? 'spot_clinic' : 'regular'
    );
    
    return true;
    
  } catch (e) {
    console.error('担当者変更エラー:', e);
    updateDescriptionWithError_(ev, currentCalId, newAssigneeName);
    return false;
  }
}


/**
 * 説明文から担当者変更情報を抽出
 * @param {string} description - カレンダーの説明文
 * @returns {Object|null} {newAssigneeName: string, processed: boolean} または null
 */
function extractAssigneeChangeFromDescription_(description) {
  if (!description) return null;
  
  const regex = /変更先担当者[：:]「([^」]+)」/; // ★★★ 修正箇所 ★★★
  const match = description.match(regex);
  
  if (!match) return null;
  
  const newAssigneeName = match[1].trim();
  const processed = description.includes('→ 処理済み') || description.includes('→ エラー');
  
  return {
    newAssigneeName: newAssigneeName,
    processed: processed
  };
}

/**
 * 説明文を「処理済み」に更新【Ver4：Calendar API使用】
 * @param {Object} ev - カレンダーイベントオブジェクト
 * @param {string} calId - カレンダーID
 * @param {string} assigneeName - 担当者名
 */
function updateDescriptionAsProcessed_(ev, calId, assigneeName) {
  try {
    const description = String(ev.description || '');
    const updatedDesc = description.replace(
      /変更先担当者[：:]「[^」]*」/,
      `変更先担当者：「${assigneeName}」 → 処理済み`
    );
    
    Calendar.Events.patch(
      { description: updatedDesc },
      calId,
      normalizeEventId_(ev.id)
    );
  } catch (e) {
    console.error('説明文更新エラー:', e);
  }
}

/**
 * 説明文を「エラー」に更新【Ver4：Calendar API使用】
 * @param {Object} ev - カレンダーイベントオブジェクト
 * @param {string} calId - カレンダーID
 * @param {string} assigneeName - 担当者名
 */
function updateDescriptionWithError_(ev, calId, assigneeName) {
  try {
    const description = String(ev.description || '');
    const updatedDesc = description.replace(
      /変更先担当者[：:]「[^」]*」/,
      `変更先担当者：「${assigneeName}」 → エラー`
    );
    
    Calendar.Events.patch(
      { description: updatedDesc },
      calId,
      normalizeEventId_(ev.id)
    );
  } catch (e) {
    console.error('説明文更新エラー:', e);
  }
}

/**
 * カレンダーイベントから時間情報を抽出
 */
function extractTimesFromEvent_(ev) {
  let start = null, end = null, dateOnly = null;
  
  if (ev.start && ev.start.date) {
    dateOnly = new Date(ev.start.date + 'T00:00:00');
  } else if (ev.start && ev.start.dateTime) {
    start = new Date(ev.start.dateTime);
    if (ev.end && ev.end.dateTime) {
      end = new Date(ev.end.dateTime);
    } else if (ev.end && ev.end.date) {
      end = new Date(ev.end.date + 'T00:00:00');
    }
  }
  
  return {
    date: dateOnly || (start ? new Date(start.getFullYear(), start.getMonth(), start.getDate()) : null),
    start: start || null,
    end: end || null
  };
}

function normalizeEventId_(id) { 
  return String(id || '').replace(/@google\.com$/i, ''); 
}

function validateAvailability_(sh, row, d, s, e) {
  const parts = [
    nz_(getCell_(sh, row, '休診日')),
    nz_(getCell_(sh, row, '午前休診')),
    nz_(getCell_(sh, row, '午後休診'))
  ].filter(Boolean).join('、');
  
  const ngWeekdays = parseWeekdayListStrict_(parts);
  const dow = new Date(d).getDay();
  return !ngWeekdays.has(dow);
}

function parseWeekdayListStrict_(text) {
  const set = new Set();
  if (!text) return set;
  
  let s = toHalfWidth_(String(text)).trim();
  s = s.replace(/[，]/g, ',').replace(/[・！]/g, ',').replace(/　/g, ' ');
       
  const tokens = s.split(/[,\s、]+/).map(t => t.trim()).filter(Boolean);
  const JA = ['日曜日','月曜日','火曜日','水曜日','木曜日','金曜日','土曜日'];
  const MAP = new Map(JA.map((name, i) => [name, i]));
  
  for (const t of tokens) {
    if (MAP.has(t)) set.add(MAP.get(t));
  }
  return set;
}

/**
 * イベント説明文の組み立て（定期回収用）【Ver4：整合性保持】
 */
function buildDescription_(sh, row, d, s, e) {
  const f = (h) => nz_(getCell_(sh, row, h));
  const dt = Utilities.formatDate(new Date(d), CONFIG.tz, 'yyyy-MM-dd');
  const st = Utilities.formatDate(toDate_(s), CONFIG.tz, 'HH:mm');
  const et = Utilities.formatDate(toDate_(e), CONFIG.tz, 'HH:mm');
  
  const L = [];
  L.push('■ 概要');
  L.push('・医院名: ' + f('医院名'));
  
  if (f('郵便番号') || f('住所')) {
    const addrParts = [];
    if (f('郵便番号')) addrParts.push('〒' + f('郵便番号'));
    if (f('住所')) addrParts.push(f('住所'));
    L.push('・住所: ' + addrParts.join(' '));
  }
  
  L.push('・担当者: ' + [f('担当者名'), f('電話番号')].filter(Boolean).join(' / '));
  L.push('・回収予定: ' + `${dt} ${st}—${et}`);
  
  if (f('備考')) L.push('・備考: ' + f('備考'));
  if (f('対応について')) L.push('・対応: ' + f('対応について'));
  if (f('不要品詳細')) L.push('・不要品詳細: ' + f('不要品詳細'));
  
  L.push('');
  L.push('■ 回収廃棄物');
  const wastes = collectWasteLinesForChat_(sh, row);
  if (wastes.length === 0) {
    L.push('・（数量の入力がありません）');
  } else {
    wastes.forEach(w => L.push('・' + w));
  }
  
  const h = f('休診日'), am = f('午前休診'), pm = f('午後休診'), nt = f('回収不可能時間');
  if (h || am || pm || nt) {
    L.push('');
    L.push('■ 回収不可（参考）');
    if (h) L.push('・休診日: ' + h);
    if (am) L.push('・午前休診: ' + am);
    if (pm) L.push('・午後休診: ' + pm);
    if (nt) L.push('・時間: ' + nt);
  }
  
  L.push('');
  L.push('■ 担当者変更');
  L.push('・変更先担当者：「」'); // ★★★ 修正箇所 ★★★
  
  L.push('');
  L.push('■ 管理');
  L.push('・受付ID: ' + f('受付ID'));
  L.push('・シート行: ' + f('行URL'));
  L.push('・回収担当者: ' + f('回収担当者'));
  L.push('・カレンダーID: ' + f('カレンダーID'));
  
  return L.join('\n');
}

/**
 * スポット回収のイベント説明文【Ver4：整合性保持】
 */
function buildSpotDescription_(sh, row, d, s, e, type) {
  const f = (h) => nz_(getCell_(sh, row, h));
  const dt = Utilities.formatDate(new Date(d), CONFIG.tz, 'yyyy-MM-dd');
  const st = Utilities.formatDate(toDate_(s), CONFIG.tz, 'HH:mm');
  const et = Utilities.formatDate(toDate_(e), CONFIG.tz, 'HH:mm');
  
  const L = [];
  L.push('■ 概要');
  L.push('・医院名: ' + f('医院名'));
  
  if (f('郵便番号') || f('住所')) {
    const addrParts = [];
    if (f('郵便番号')) addrParts.push('〒' + f('郵便番号'));
    if (f('住所')) addrParts.push(f('住所'));
    L.push('・住所: ' + addrParts.join(' '));
  }
  
  if (type === 'dealer') {
    L.push('・医院担当者: ' + f('医院担当者名'));
    L.push('・ディーラー: ' + [f('会社名'), f('支店名')].filter(Boolean).join(' '));
    L.push('・ディーラー担当者: ' + f('ディーラー担当者名'));
  } else {
    L.push('・担当者: ' + f('担当者名'));
  }
  
  L.push('・回収予定: ' + `${dt} ${st}—${et}`);
  L.push('・依頼種別: ' + f('依頼種別'));
  
  L.push('');
  L.push('■ 回収品目');
  const items = (type === 'clinic') ? collectSpotClinicItems_(sh, row) : collectSpotDealerItems_(sh, row);
  if (items.length === 0) {
    L.push('・（品目の入力がありません）');
  } else {
    items.forEach(item => L.push('・' + item));
  }
  
  L.push('');
  L.push('■ 担当者変更');
  L.push('・変更先担当者：「」'); // ★★★ 修正箇所 ★★★
  
  L.push('');
  L.push('■ 管理');
  L.push('・受付ID: ' + f('受付ID'));
  L.push('・シート行: ' + f('行URL'));
  L.push('・回収担当者: ' + f('回収担当者'));
  L.push('・カレンダーID: ' + f('カレンダーID'));
  
  return L.join('\n');
}

function buildLocation_(sh, row) {
  const clinic = nz_(getCell_(sh, row, '医院名'));
  const zip = nz_(getCell_(sh, row, '郵便番号'));
  const addr = nz_(getCell_(sh, row, '住所'));
  
  const parts = [];
  if (zip) parts.push(`〒${zip}`);
  if (addr) parts.push(addr);
  
  return parts.length ? `${clinic}（${parts.join(' ')}）` : clinic;
}

function buildSpotLocation_(sh, row) {
  return buildLocation_(sh, row);
}

function mergeDateTime_(dateCell, timeCell){
  const d = (dateCell instanceof Date) ? new Date(dateCell) : new Date(dateCell);
  const t = toDate_(timeCell);
  d.setHours(t.getHours(), t.getMinutes(), 0, 0);
  return d;
}

function toDate_(v){
  if (v instanceof Date) return v;
  if (typeof v === 'number') { 
    const base = new Date(1899, 11, 30); 
    return new Date(base.getTime() + v * 86400000); 
  }
  
  const m = String(v).match(/(\d{1,2}):(\d{2})/);
  if (m) { 
    const d = new Date(); 
    d.setHours(+m[1], +m[2], 0, 0); 
    return d; 
  }
  
  const d2 = new Date(v); 
  return isNaN(d2.getTime()) ? new Date() : d2;
}