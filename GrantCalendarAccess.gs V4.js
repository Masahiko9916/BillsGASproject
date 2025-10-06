/**
 * GrantCalendarAccess.gs V4 - カレンダー権限一括付与
 * 
 * 【Ver4の変更点】
 * - エラー処理の改善
 * - ログ出力の強化
 * - Calendar APIとの整合性向上
 * 
 * 【このファイルの役割】
 * - 管理者が自分のアカウントに全担当者のカレンダー編集権限を一括付与
 * - 担当者マスタから自動的にリストを取得
 * - Google Calendar API（v3）を使用
 * 
 * 【実行条件】
 * - Google Workspace管理者権限で実行すること
 * - Calendar APIが有効化されていること
 * 
 * 【依存関係】
 * - Config.gs（担当者マスタの設定）
 * - AssigneeMaster.gs（担当者情報取得）
 * - SheetUtils.gs（シート操作）
 */

/**
 * メイン処理：全担当者のカレンダーに編集権限を一括付与
 * 【説明】スプレッドシートのメニューから実行する
 */
function grantCalendarAccessToAdmin() {
  const ui = SpreadsheetApp.getUi();
  
  // 管理者メールアドレスの取得
  const adminEmail = Session.getActiveUser().getEmail();
  
  if (!adminEmail || !adminEmail.includes('@')) {
    ui.alert('エラー', '管理者のメールアドレスが取得できませんでした。', ui.ButtonSet.OK);
    return;
  }
  
  // 確認ダイアログ
  const response = ui.alert(
    'カレンダー権限の一括付与',
    `全担当者のカレンダーに、以下のアカウントへの編集権限を付与します：\n\n${adminEmail}\n\n実行してよろしいですか？`,
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    ui.alert('キャンセルしました。');
    return;
  }
  
  // 処理開始
  ui.alert('処理を開始します。完了までしばらくお待ちください。\n（25名の場合、約15秒かかります）');
  
  try {
    const result = grantCalendarAccessBulk_(adminEmail);
    
    // 結果表示
    const message = 
      `=== 処理完了 ===\n\n` +
      `対象者数: ${result.total}名\n` +
      `成功: ${result.success}名\n` +
      `スキップ（既存）: ${result.skipped}名\n` +
      `失敗: ${result.error}名\n\n` +
      (result.errorDetails.length > 0 
        ? `エラー詳細:\n${result.errorDetails.join('\n')}`
        : '');
    
    ui.alert('完了', message, ui.ButtonSet.OK);
    
    // ログにも出力
    console.log('=== カレンダー権限付与完了 ===');
    console.log(message);
    
  } catch (e) {
    ui.alert('エラー', `処理中にエラーが発生しました:\n${e.message}`, ui.ButtonSet.OK);
    console.error('カレンダー権限付与エラー:', e);
  }
}

/**
 * カレンダー権限の一括付与処理（内部関数）
 * @param {string} adminEmail - 管理者のメールアドレス
 * @returns {Object} 処理結果（total, success, skipped, error, errorDetails）
 */
function grantCalendarAccessBulk_(adminEmail) {
  console.log(`=== カレンダー権限付与開始 ===`);
  console.log(`管理者: ${adminEmail}`);
  
  // 担当者マスタから担当者リストを取得
  const assigneeEmails = getAssigneeEmailsForGrant_();
  
  if (assigneeEmails.length === 0) {
    throw new Error('担当者マスタに有効なメールアドレスが見つかりません。');
  }
  
  console.log(`対象者数: ${assigneeEmails.length}名`);
  console.log(`担当者リスト: ${assigneeEmails.join(', ')}`);
  
  const result = {
    total: assigneeEmails.length,
    success: 0,
    skipped: 0,
    error: 0,
    errorDetails: []
  };
  
  // 各担当者に対して処理
  assigneeEmails.forEach((email, index) => {
    console.log(`\n[${index + 1}/${assigneeEmails.length}] ${email} を処理中...`);
    
    try {
      // Calendar APIを使用してACLルールを追加
      const rule = {
        scope: {
          type: 'user',
          value: adminEmail
        },
        role: 'writer' // 編集権限（予定の変更権限）
      };
      
      // 既存のルールを確認（重複防止）
      let alreadyExists = false;
      try {
        const existingRules = Calendar.Acl.list(email);
        if (existingRules && existingRules.items) {
          alreadyExists = existingRules.items.some(item => 
            item.scope && item.scope.value === adminEmail
          );
        }
      } catch (checkErr) {
        // 既存ルールの確認に失敗しても続行
        console.warn(`  既存ルールの確認に失敗（続行します）: ${checkErr.message}`);
      }
      
      if (alreadyExists) {
        console.log(`  ⚠️ 既に権限が付与されています - スキップ`);
        result.skipped++;
        return;
      }
      
      // ACLルールを挿入
      Calendar.Acl.insert(rule, email);
      console.log(`  ✅ 成功`);
      result.success++;
      
    } catch (e) {
      const errorMsg = `${email}: ${e.message}`;
      console.error(`  ❌ 失敗: ${e.message}`);
      result.error++;
      result.errorDetails.push(errorMsg);
      
      // カレンダーが見つからない場合の詳細メッセージ
      if (e.message.includes('Not Found') || e.message.includes('404')) {
        console.error(`  → カレンダーが見つかりません。メールアドレスを確認してください。`);
      } else if (e.message.includes('Forbidden') || e.message.includes('403')) {
        console.error(`  → 権限がありません。管理者権限で実行してください。`);
      }
    }
    
    // API制限対策：少し待機（25名で約12.5秒）
    Utilities.sleep(500);
  });
  
  console.log(`\n=== 処理完了 ===`);
  console.log(`成功: ${result.success}件`);
  console.log(`スキップ: ${result.skipped}件`);
  console.log(`失敗: ${result.error}件`);
  
  return result;
}

/**
 * 担当者マスタから担当者のメールアドレスリストを取得
 * @returns {Array<string>} メールアドレスの配列
 */
function getAssigneeEmailsForGrant_() {
  try {
    const sh = getAssigneeSheet_();
    const headers = sheetHeaders_(sh).map(h => String(h).trim());
    
    console.log('取得したヘッダー:', headers);
    console.log('探しているヘッダー:', CONFIG.assigneeEmailHeader);
    
    const emailCol = headers.indexOf(CONFIG.assigneeEmailHeader);
    
    if (emailCol < 0) {
      throw new Error(`担当者マスタに「${CONFIG.assigneeEmailHeader}」列がありません。`);
    }
    
    console.log(`メールアドレス列: ${emailCol + 1}列目`);
    
    const lastRow = sh.getLastRow();
    if (lastRow < 2) {
      return [];
    }
    
    console.log(`データ行数: ${lastRow - 1}行`);
    
    const values = sh.getRange(2, emailCol + 1, lastRow - 1, 1).getValues();
    
    console.log(`取得した生データ数: ${values.length}`);
    
    // メールアドレスのバリデーション
    const emails = [];
    values.forEach((row, index) => {
      const email = String(row[0] || '').trim();
      
      // 空欄をスキップ
      if (!email) {
        console.log(`${index + 2}行目: 空欄 - スキップ`);
        return;
      }
      
      // @を含むかチェック
      if (!email.includes('@')) {
        console.warn(`${index + 2}行目: 無効なメールアドレス "${email}" - スキップ`);
        return;
      }
      
      console.log(`${index + 2}行目: ✅ "${email}"`);
      emails.push(email);
    });
    
    console.log(`有効なメールアドレス: ${emails.length}件`);
    
    // 重複を削除
    const uniqueEmails = [...new Set(emails)];
    
    if (uniqueEmails.length !== emails.length) {
      const duplicateCount = emails.length - uniqueEmails.length;
      console.warn(`⚠️ 重複するメールアドレスを${duplicateCount}件削除しました（${emails.length} → ${uniqueEmails.length}）`);
      
      // 重複しているメールアドレスを表示
      const emailCounts = {};
      emails.forEach(email => {
        emailCounts[email] = (emailCounts[email] || 0) + 1;
      });
      
      Object.keys(emailCounts).forEach(email => {
        if (emailCounts[email] > 1) {
          console.warn(`  - ${email}: ${emailCounts[email]}回出現`);
        }
      });
    }
    
    console.log(`最終的なメールアドレス数: ${uniqueEmails.length}件`);
    
    return uniqueEmails;
    
  } catch (e) {
    console.error('担当者リスト取得エラー:', e);
    throw new Error(`担当者マスタの読み込みに失敗しました: ${e.message}`);
  }
}

/**
 * テスト実行：1人だけに権限を付与してテスト
 * 【説明】本番実行前のテスト用
 */
function testGrantCalendarAccessSingle() {
  const ui = SpreadsheetApp.getUi();
  const adminEmail = Session.getActiveUser().getEmail();
  
  // テスト対象のメールアドレスを入力
  const response = ui.prompt(
    'テスト実行',
    'テスト対象のメールアドレスを入力してください:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const testEmail = response.getResponseText().trim();
  
  if (!testEmail || !testEmail.includes('@')) {
    ui.alert('無効なメールアドレスです。');
    return;
  }
  
  try {
    console.log(`=== テスト実行: ${testEmail} ===`);
    
    const rule = {
      scope: {
        type: 'user',
        value: adminEmail
      },
      role: 'writer'
    };
    
    // 既存ルールの確認
    try {
      const existingRules = Calendar.Acl.list(testEmail);
      const alreadyExists = existingRules.items.some(item => 
        item.scope.value === adminEmail
      );
      
      if (alreadyExists) {
        ui.alert('既に権限が付与されています。');
        return;
      }
    } catch (e) {
      console.warn('既存ルール確認エラー:', e);
    }
    
    // ACLルール追加
    Calendar.Acl.insert(rule, testEmail);
    
    console.log('✅ テスト成功');
    ui.alert('成功', `${testEmail} のカレンダーに権限を付与しました。`, ui.ButtonSet.OK);
    
  } catch (e) {
    console.error('❌ テスト失敗:', e);
    ui.alert('失敗', `エラー: ${e.message}`, ui.ButtonSet.OK);
  }
}

/**
 * 権限の確認：自分がアクセスできるカレンダーを一覧表示【V4改善版】
 * 【説明】設定が正しく反映されているか確認する
 */
function listAccessibleCalendars() {
  const adminEmail = Session.getActiveUser().getEmail();
  console.log(`=== アクセス可能なカレンダー一覧 ===`);
  console.log(`ユーザー: ${adminEmail}`);
  
  try {
    const assigneeEmails = getAllAssigneeEmails_();
    let accessibleCount = 0;
    let inaccessibleCount = 0;
    const accessibleList = [];
    const inaccessibleList = [];
    
    assigneeEmails.forEach(email => {
      let success = false;
      // ★★★ 修正箇所: リトライ処理を追加 ★★★
      for (let i = 0; i < 3; i++) { // 最大3回試行
        try {
          // Calendar APIで権限確認を試みる
          const acl = Calendar.Acl.list(email, { maxResults: 1 });
          
          // API呼び出しが成功し、有効な応答が得られたかを確認
          if (acl && acl.items) {
            console.log(`✅ ${email} - アクセス可能`);
            accessibleList.push(email);
            accessibleCount++;
            success = true;
            break; // 成功したのでループを抜ける
          } else {
             // aclオブジェクトは存在するが、itemsがない場合も失敗と見なす
             console.warn(`⚠️ ${email} - APIから予期しない応答。リトライします... (${i + 1}/3)`);
          }
        } catch (e) {
          console.warn(`❌ ${email} - API呼び出しエラー。リトライします... (${i + 1}/3): ${e.message}`);
        }
        if (i < 2) Utilities.sleep(1000 * (i + 1)); // 失敗した場合、1秒、2秒と待機時間を増やす
      }
      
      if (!success) {
        console.log(`❌ ${email} - アクセス不可（3回試行後）`);
        inaccessibleList.push(email);
        inaccessibleCount++;
      }
      
      Utilities.sleep(200); // 次のメールアドレスへのAPI制限対策
    });
    
    console.log(`\n=== 結果 ===`);
    console.log(`アクセス可能: ${accessibleCount}件`);
    console.log(`アクセス不可: ${inaccessibleCount}件`);
    
    const detailMessage = 
      `=== 確認結果 ===\n\n` +
      `アクセス可能: ${accessibleCount}件\n` +
      `アクセス不可: ${inaccessibleCount}件\n\n` +
      (inaccessibleCount > 0 
        ? `アクセス不可リスト:\n${inaccessibleList.slice(0, 10).join('\n')}` +
          (inaccessibleList.length > 10 ? `\n他${inaccessibleList.length - 10}件` : '')
        : '全てのカレンダーにアクセス可能です。');
    
    SpreadsheetApp.getUi().alert('確認完了', detailMessage, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (e) {
    console.error('エラー:', e);
    SpreadsheetApp.getUi().alert('エラー', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}