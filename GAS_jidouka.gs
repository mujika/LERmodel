// ── 設定項目 ──
/** 日報を取る対象シート名 */
const SHEET_NAMES   = ['。', '', ''];

/**
 * ----- 設定項目 -----
 * メール送信先（To と Cc）
 */
const RECIPIENT_TO = '';
const RECIPIENT_CC = ',,';


function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('日報・週報メール')
    .addItem('手動で送信', 'manualDailyReport')
    .addToUi();
}

/**
 * 日報テキスト生成＋週報PDF作成＆メール送信（UI 呼び出しなし）
 */
function sendDailyReportCore() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const tz  = ss.getSpreadsheetTimeZone();
  const today     = new Date();
  const yesterday = new Date(today.getFullYear(), today.getMonth(), today.getDate() - 1);
  const yStr      = Utilities.formatDate(yesterday, tz, 'yyyy/MM/dd');

  // メール本文用の行を集める
  const bodyLines = [
    'お疲れさまです。',
    `${yStr} の日報と週報をお送りします。`,
    ''
  ];
  // PDF blob をためておく配列
  const attachments = [];

  SHEET_NAMES.forEach(sheetName => {
    // --- 日報部分 ---
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(`シート「${sheetName}」が見つかりません`);
    const rows = sheet.getDataRange().getValues().slice(1);
    let lastDate = null;
    const lines = [];
    rows.forEach(r => {
      if (r[0] instanceof Date) lastDate = r[0];
      if (lastDate && Utilities.formatDate(lastDate, tz, 'yyyy/MM/dd') === yStr) {
        const slot = r[1], web = Number(r[2])||0, onl = Number(r[3])||0;
        lines.push(`・${slot}: ${web+onl}枚 (Web:${web}枚, 現地:${onl}枚)`);
      }
    });
    if (lines.length === 0) {
      lines.push('（データがありません）');
    }

    bodyLines.push(`【${sheetName}】 日報`);
    bodyLines.push(...lines);
    bodyLines.push('');

    // --- 週報PDF部分 ---
    // 期間：前日から遡って6日間（計7日）
    const startW = new Date(yesterday);
    startW.setDate(yesterday.getDate() - 6);
    const endW = yesterday;

    lastDate = null;
    const weekly = {};
    let totalWeb = 0, totalOnl = 0;
    rows.forEach(r => {
      if (r[0] instanceof Date) lastDate = r[0];
      if (lastDate >= startW && lastDate <= endW) {
        const key = Utilities.formatDate(lastDate, tz, 'yyyy-MM-dd');
        weekly[key] = weekly[key] || { web:0, onsite:0 };
        weekly[key].web    += Number(r[2])||0;
        weekly[key].onsite += Number(r[3])||0;
      }
    });

    // 一時シートに週報を出力
    const tempName = `${sheetName}_週報`;
    if (ss.getSheetByName(tempName)) ss.deleteSheet(tempName);
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

    // PDFエクスポートURL組み立て (A4横/横幅に合わせる/余白0.5inch)
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
      `${sheetName}_週報_` +
      `${Utilities.formatDate(startW,tz,'yyyyMMdd')}` +
      `_〜${Utilities.formatDate(endW,tz,'yyyyMMdd')}.pdf`
    );

    attachments.push(blob);
    ss.deleteSheet(rpt);
  });

  bodyLines.push('以上、ご確認ください。');
  const body = bodyLines.join('\n');
  const subject = `[日報・週報] ${yStr}`;

  // メール送信（複数PDFを添付）
  MailApp.sendEmail({
    to:          RECIPIENT_TO,
    cc:          RECIPIENT_CC,
    subject:     subject,
    body:        body,
    attachments: attachments
  });
}

/** 手動実行用 */
function manualDailyReport() {
  const ui = SpreadsheetApp.getUi();
  try {
    sendDailyReportCore();
    ui.alert('日報・週報PDF（複数シート分）をメール送信しました');
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
