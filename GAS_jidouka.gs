// ── 設定項目 ──
/**
 * 日報を取る対象シート名を動的に取得
 * シート名の先頭が「*」で始まるシートを対象とする
 */
function getTargetSheetNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const targetSheets = [];
  
  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    if (sheetName.startsWith('*')) {
      targetSheets.push(sheetName);
    }
  });
  
  if (targetSheets.length === 0) {
    throw new Error('対象シートが見つかりません。シート名の先頭に「*」を付けてください。');
  }
  
  Logger.log('対象シート: ' + targetSheets.map(name => getDisplaySheetName(name)).join(', '));
  return targetSheets;
}

/**
 * 指定シートからメールアドレス設定を取得
 * F6セル: TO宛先
 * G6セル: 担当者名
 * G9セル: チケット価格
 * F12〜F17セル: CC宛先
 */
function getEmailSettingsForSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    throw new Error(`シート「${sheetName}」が見つかりません`);
  }
  
  // F6セルからTO宛先を取得
  const toAddress = sheet.getRange('F6').getValue();
  if (!toAddress || typeof toAddress !== 'string') {
    throw new Error(`TO宛先メールアドレスが設定されていません（${sheetName}のF6セル）`);
  }
  
  // G6セルから担当者名を取得
  const personName = sheet.getRange('G6').getValue();
  if (!personName || typeof personName !== 'string') {
    throw new Error(`担当者名が設定されていません（${sheetName}のG6セル）`);
  }
  
  // G9セルからチケット価格を取得
  const ticketPrice = sheet.getRange('G9').getValue();
  if (!ticketPrice || typeof ticketPrice !== 'number') {
    throw new Error(`チケット価格が設定されていません（${sheetName}のG9セル）`);
  }
  
  // F12〜F17セルからCC宛先を取得
  const ccRange = sheet.getRange('F12:F17');
  const ccValues = ccRange.getValues();
  const ccAddresses = [];
  
  ccValues.forEach(row => {
    const email = row[0];
    if (email && typeof email === 'string' && email.trim() !== '') {
      ccAddresses.push(email.trim());
    }
  });
  
  const ccString = ccAddresses.join(',');
  
  Logger.log(`[${getDisplaySheetName(sheetName)}] TO宛先: ${toAddress}`);
  Logger.log(`[${getDisplaySheetName(sheetName)}] 担当者名: ${personName}`);
  Logger.log(`[${getDisplaySheetName(sheetName)}] チケット価格: ${ticketPrice}`);
  Logger.log(`[${getDisplaySheetName(sheetName)}] CC宛先: ${ccString}`);
  
  return {
    to: toAddress.trim(),
    cc: ccString,
    personName: personName.trim(),
    ticketPrice: ticketPrice
  };
}

/**
 * 従来の関数（後方互換性のため）
 */
function getEmailSettings() {
  const sheetNames = getTargetSheetNames();
  return getEmailSettingsForSheet(sheetNames[0]);
}

/**
 * シート名から表示用の名前を取得（先頭の「*」を除去）
 */
function getDisplaySheetName(sheetName) {
  return sheetName.startsWith('*') ? sheetName.slice(1) : sheetName;
}


/**
 * 指定シートの日報を生成
 */
function generateDailyReportForSheet(sheetName, date) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    throw new Error(`シート「${sheetName}」が見つかりません`);
  }
  
  const emailSettings = getEmailSettingsForSheet(sheetName);
  const dateStr = Utilities.formatDate(date, tz, 'yyyy/MM/dd');
  
  const rows = sheet.getDataRange().getValues().slice(1);
  let lastDate = null;
  const lines = [];
  
  rows.forEach(r => {
    if (r[0] instanceof Date) lastDate = r[0];
    if (lastDate && Utilities.formatDate(lastDate, tz, 'yyyy/MM/dd') === dateStr) {
      const slot = r[1], web = Number(r[2])||0, onl = Number(r[3])||0;
      if (slot && slot.toString().trim() !== '') {
        const totalTickets = web + onl;
        const totalAmount = totalTickets * emailSettings.ticketPrice;
        lines.push(`・${slot}: ${totalTickets}枚 × ${emailSettings.ticketPrice}円 = ${totalAmount}円`);
      }
    }
  });
  
  if (lines.length === 0) {
    lines.push('（データがありません）');
  }
  
  Logger.log(`[${getDisplaySheetName(sheetName)}] 日報生成完了: ${lines.length}行`);
  return lines;
}

/**
 * 指定シートの週報PDFを生成
 */
function generateWeeklyReportPDF(sheetName, startDate, endDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    throw new Error(`シート「${sheetName}」が見つかりません`);
  }
  
  const rows = sheet.getDataRange().getValues().slice(1);
  let lastDate = null;
  const weekly = {};
  let totalWeb = 0, totalOnl = 0;
  
  rows.forEach(r => {
    if (r[0] instanceof Date) lastDate = r[0];
    if (lastDate >= startDate && lastDate <= endDate) {
      const key = Utilities.formatDate(lastDate, tz, 'yyyy-MM-dd');
      weekly[key] = weekly[key] || { web:0, onsite:0 };
      weekly[key].web    += Number(r[2])||0;
      weekly[key].onsite += Number(r[3])||0;
    }
  });
  
  // 一時シートに週報を出力
  const displayName = getDisplaySheetName(sheetName);
  const tempName = `${displayName}_週報`;
  if (ss.getSheetByName(tempName)) {
    ss.deleteSheet(ss.getSheetByName(tempName));
  }
  const rpt = ss.insertSheet(tempName);
  
  rpt.appendRow(['日付','Webチケット','現地チケット','合計']);
  Object.keys(weekly).sort().forEach(d => {
    const w = weekly[d].web, o = weekly[d].onsite;
    rpt.appendRow([d, w, o, w+o]);
    totalWeb += w; totalOnl += o;
  });
  rpt.appendRow(['合計', totalWeb, totalOnl, totalWeb+totalOnl]);
  
  // スタイル調整
  const lr = rpt.getLastRow(), lc = 4;
  rpt.getRange(1,1,lr,lc).setBorder(true,true,true,true,true,true);
  rpt.getRange(1,1,1,lc).setBackground('#d9ead3');
  for (let i=2; i<=lr; i++) {
    rpt.getRange(i,1,1,lc)
       .setBackground(i%2===0? '#f3f3f3':'#ffffff');
  }
  rpt.getRange(lr,1,1,lc).setFontWeight('bold').setBackground('#d9ead3');
  rpt.setFrozenRows(1);
  
  // PDFエクスポートURL組み立て
  const exportUrl = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?` +
    [
      'format=pdf',
      'portrait=false',
      'size=A4',
      'fitw=true',
      'top_margin=0.5',
      'bottom_margin=0.5',
      'left_margin=0.5',
      'right_margin=0.5',
      'gridlines=true',
      'printbackground=true',
      'fzr=true',
      `gid=${rpt.getSheetId()}`
    ].join('&');
  
  const blob = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
  }).getBlob().setName(
    `${displayName}_週報_` +
    `${Utilities.formatDate(startDate,tz,'yyyyMMdd')}` +
    `_〜${Utilities.formatDate(endDate,tz,'yyyyMMdd')}.pdf`
  );
  
  ss.deleteSheet(rpt);
  Logger.log(`[${getDisplaySheetName(sheetName)}] 週報PDF生成完了`);
  return blob;
}

/**
 * 指定シートの日報・週報を個別送信
 */
function sendReportForSheet(sheetName, date) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const dateStr = Utilities.formatDate(date, tz, 'yyyy/MM/dd');
  
  const emailSettings = getEmailSettingsForSheet(sheetName);
  
  // 日報生成
  const dailyLines = generateDailyReportForSheet(sheetName, date);
  
  // 週報PDF生成（過去7日間）
  const startDate = new Date(date);
  startDate.setDate(date.getDate() - 6);
  const weeklyPDF = generateWeeklyReportPDF(sheetName, startDate, date);
  
  // メール本文作成
  const displayName = getDisplaySheetName(sheetName);
  const bodyLines = [
    `${emailSettings.personName}様`,
    '',
    '',
    'お世話になっております。',
    `上映報告botより『${displayName}』の${dateStr}の日報をお送りします。`,
    '',
    `【${displayName}】 日報`,
    ...dailyLines,
    '',
    '以上、ご確認ください。'
  ];
  
  const body = bodyLines.join('\n');
  const subject = `『${displayName}』上映報告`;
  
  // メール送信
  MailApp.sendEmail({
    to: emailSettings.to,
    cc: emailSettings.cc,
    subject: subject,
    body: body,
    attachments: [weeklyPDF]
  });
  
  Logger.log(`[${getDisplaySheetName(sheetName)}] メール送信完了`);
}

/**
 * 指定シートの完全な日報・週報処理を実行
 * （集計 → PDF生成 → メール送信 → 一時ファイル削除）
 */
function processSheetReport(sheetName, date) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const dateStr = Utilities.formatDate(date, tz, 'yyyy/MM/dd');
  
  Logger.log(`[${getDisplaySheetName(sheetName)}] 処理開始: ${dateStr}`);
  
  try {
    // 1. メール設定取得
    Logger.log(`[${getDisplaySheetName(sheetName)}] メール設定取得中...`);
    const emailSettings = getEmailSettingsForSheet(sheetName);
    
    // 2. 日報データ集計
    Logger.log(`[${getDisplaySheetName(sheetName)}] 日報データ集計中...`);
    const dailyLines = generateDailyReportForSheet(sheetName, date);
    
    // 3. 週報データ集計とPDF生成
    Logger.log(`[${getDisplaySheetName(sheetName)}] 週報PDF生成中...`);
    const startDate = new Date(date);
    startDate.setDate(date.getDate() - 6);
    const weeklyPDF = generateWeeklyReportPDF(sheetName, startDate, date);
    
    // 4. メール本文作成
    Logger.log(`[${getDisplaySheetName(sheetName)}] メール本文作成中...`);
    const displayName = getDisplaySheetName(sheetName);
    const bodyLines = [
      `${emailSettings.personName}様`,
      '',
      '',
      'お世話になっております。',
      `上映報告botより『${displayName}』の${dateStr}の日報をお送りします。`,
      '',
      `【${displayName}】 日報`,
      ...dailyLines,
      '',
      '以上、ご確認ください。'
    ];
    
    const body = bodyLines.join('\n');
    const subject = `『${displayName}』上映報告`;
    
    // 5. メール送信
    Logger.log(`[${getDisplaySheetName(sheetName)}] メール送信中...`);
    MailApp.sendEmail({
      to: emailSettings.to,
      cc: emailSettings.cc,
      subject: subject,
      body: body,
      attachments: [weeklyPDF]
    });
    
    Logger.log(`[${getDisplaySheetName(sheetName)}] 処理完了: メール送信済み`);
    
  } catch (error) {
    Logger.log(`[${getDisplaySheetName(sheetName)}] 処理エラー: ${error.message}`);
    throw error;
  }
}

/**
 * 全シートの日報・週報を個別送信
 */
function sendAllSheetReports(date) {
  const sheetNames = getTargetSheetNames();
  const yesterday = date || new Date(new Date().getTime() - 24 * 60 * 60 * 1000);
  
  sheetNames.forEach(sheetName => {
    try {
      sendReportForSheet(sheetName, yesterday);
    } catch (error) {
      Logger.log(`[${getDisplaySheetName(sheetName)}] 送信エラー: ${error.message}`);
    }
  });
  
  Logger.log('全シート個別送信完了');
}

/**
 * 全シートの日報・週報を順次処理で送信
 * （各シートごとに：集計 → PDF生成 → メール送信 → 次のシートへ）
 */
function sendDailyReportSequential(date) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const targetDate = date || new Date(new Date().getTime() - 24 * 60 * 60 * 1000);
  const dateStr = Utilities.formatDate(targetDate, tz, 'yyyy/MM/dd');
  
  Logger.log(`=== 順次処理開始: ${dateStr} ===`);
  
  try {
    // 対象シート取得
    const sheetNames = getTargetSheetNames();
    Logger.log(`対象シート数: ${sheetNames.length}個`);
    
    const results = {
      total: sheetNames.length,
      success: 0,
      failed: 0,
      errors: []
    };
    
    // 各シートを順次処理
    sheetNames.forEach((sheetName, index) => {
      const displayName = getDisplaySheetName(sheetName);
      Logger.log(`--- ${index + 1}/${sheetNames.length}: ${displayName} (${sheetName}) ---`);
      
      try {
        // シート単位の完全処理実行
        processSheetReport(sheetName, targetDate);
        results.success++;
        
      } catch (error) {
        results.failed++;
        results.errors.push({
          sheetName: sheetName,
          error: error.message
        });
        Logger.log(`[${getDisplaySheetName(sheetName)}] 処理失敗: ${error.message}`);
      }
      
      // 次のシートへ進む前に少し待機（API制限対策）
      if (index < sheetNames.length - 1) {
        Utilities.sleep(1000);
      }
    });
    
    // 処理結果のサマリー
    Logger.log(`=== 順次処理完了: ${dateStr} ===`);
    Logger.log(`成功: ${results.success}/${results.total}シート`);
    Logger.log(`失敗: ${results.failed}/${results.total}シート`);
    
    if (results.errors.length > 0) {
      Logger.log('エラー詳細:');
      results.errors.forEach(err => {
        Logger.log(`  - ${err.sheetName}: ${err.error}`);
      });
    }
    
    return results;
    
  } catch (error) {
    Logger.log(`順次処理全体エラー: ${error.message}`);
    throw error;
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('日報・週報メール')
    .addItem('手動で送信（統合）', 'manualDailyReport')
    .addItem('手動で送信（個別）', 'manualIndividualReport')
    .addToUi();
}

/**
 * 日報テキスト生成＋週報PDF作成＆メール送信（UI 呼び出しなし）
 * 新実装：シート単位で順次処理を実行
 */
function sendDailyReportCore() {
  const today = new Date();
  const yesterday = new Date(today.getFullYear(), today.getMonth(), today.getDate() - 1);
  
  Logger.log('sendDailyReportCore: 順次処理に移行');
  
  try {
    // 新しい順次処理実装を使用
    const results = sendDailyReportSequential(yesterday);
    
    Logger.log(`sendDailyReportCore完了: 成功${results.success}件、失敗${results.failed}件`);
    
    // 失敗があった場合は警告ログを出力
    if (results.failed > 0) {
      Logger.log('一部のシートで処理が失敗しました。詳細は上記ログを確認してください。');
    }
    
  } catch (error) {
    Logger.log(`sendDailyReportCore エラー: ${error.message}`);
    throw error;
  }
}

/** 手動実行用（統合送信） */
function manualDailyReport() {
  const ui = SpreadsheetApp.getUi();
  try {
    sendDailyReportCore();
    
    // ログから処理結果を取得するため、少し待機
    Utilities.sleep(500);
    
    ui.alert('日報・週報PDF（シート別順次処理）をメール送信しました\n詳細はログを確認してください。');
  } catch (e) {
    ui.alert('エラーが発生しました:\n' + e.message);
    Logger.log(e.stack);
  }
}

/** 手動実行用（個別送信） */
function manualIndividualReport() {
  const ui = SpreadsheetApp.getUi();
  try {
    const yesterday = new Date(new Date().getTime() - 24 * 60 * 60 * 1000);
    
    // 新しい順次処理を使用
    const results = sendDailyReportSequential(yesterday);
    
    // 結果に応じてメッセージを変更
    if (results.failed === 0) {
      ui.alert(`日報・週報PDF（シート別個別）をメール送信しました\n成功: ${results.success}/${results.total}シート`);
    } else {
      ui.alert(`日報・週報PDF（シート別個別）の送信が完了しました\n成功: ${results.success}/${results.total}シート\n失敗: ${results.failed}/${results.total}シート\n\n詳細はログを確認してください。`);
    }
  } catch (e) {
    ui.alert('エラーが発生しました:\n' + e.message);
    Logger.log(e.stack);
  }
}

/** トリガー用（UI 呼び出しなし） */
function triggerDailyReport() {
  try {
    sendDailyReportCore();
  } catch (e) {
    Logger.log('自動送信エラー: ' + e.stack);
  }
}
