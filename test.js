function replaceSlashWithNewline() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("投稿管理");
  var range = sheet.getRange("D:D");
  var values = range.getValues();
  for (var i = 0; i < values.length; i++) {
    if (typeof values[i][0] === 'string') {
      values[i][0] = values[i][0].replace(/▽▽/g, "\n\n").replace(/▽/g, "\n");
    }
  }
  range.setValues(values);
}

function cleanPostSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('投稿管理');
  var data = sheet.getDataRange().getValues();
  var deleteRows = [];
  
  for (var i = 1; i < data.length; i++) {
    var status = String(data[i][2] || '').trim();
    if (status === 'posted' || status === 'expired') {
      deleteRows.push(i + 1);
      console.log('削除対象 行' + (i+1) + ': status=' + status + ', text=' + String(data[i][3]).substring(0, 30));
    }
  }
  
  deleteRows.sort(function(a, b) { return b - a; });
  deleteRows.forEach(function(r) { sheet.deleteRow(r); });
  console.log('削除:', deleteRows.length + '件');
}

function cleanHistory() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hist = ss.getSheetByName('履歴');
  var data = hist.getDataRange().getValues();
  var deleteRows = [];
  var seen = {};
  
  for (var i = 1; i < data.length; i++) {
    var text = String(data[i][2] || '').trim();
    var postId = String(data[i][5] || '');
    
    // 空テキスト行は削除
    if (!text && !postId) {
      deleteRows.push(i + 1);
      continue;
    }
    
    // 同じthreadsPostIdの重複を削除（最初の1件だけ残す）
    if (postId) {
      if (seen[postId]) {
        deleteRows.push(i + 1);
      } else {
        seen[postId] = true;
      }
    }
  }
  
  deleteRows.sort(function(a, b) { return b - a; });
  deleteRows.forEach(function(r) { hist.deleteRow(r); });
  console.log('削除:', deleteRows.length + '件。残り:', hist.getLastRow() - 1 + '件');
}
function fullStatusCheck() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var now = new Date();
  console.log('現在時刻:', now.toISOString());
  
  // 1. 投稿管理の状態
  var sheet = ss.getSheetByName('投稿管理');
  var data = sheet.getDataRange().getValues();
  var statusCount = {};
  var dueCount = 0;
  
  for (var i = 1; i < data.length; i++) {
    var status = String(data[i][2] || '').trim();
    if (!status) continue;
    statusCount[status] = (statusCount[status] || 0) + 1;
    
    var scheduled = new Date(data[i][6]);
    if ((status === 'scheduled' || status === '予約済み') && !isNaN(scheduled.getTime()) && scheduled <= now) {
      dueCount++;
      console.log('★投稿対象 行' + (i+1) + ': status=' + status + ', groupId=' + (data[i][9]||'なし') + ', text=' + String(data[i][3]).substring(0, 30));
    }
  }
  console.log('\n=== 投稿管理 ===');
  console.log('総行数:', data.length - 1);
  console.log('ステータス別:', JSON.stringify(statusCount));
  console.log('投稿対象(時刻到来):', dueCount, '件');
  
  // 2. 履歴の重複チェック
  var hist = ss.getSheetByName('履歴');
  var hData = hist.getDataRange().getValues();
  var textCount = {};
  for (var j = 1; j < hData.length; j++) {
    var t = String(hData[j][2] || '').substring(0, 30);
    textCount[t] = (textCount[t] || 0) + 1;
  }
  var duplicates = [];
  for (var key in textCount) {
    if (textCount[key] > 1) duplicates.push(key + ' x' + textCount[key]);
  }
  console.log('\n=== 履歴 ===');
  console.log('総件数:', hData.length - 1);
  console.log('重複:', duplicates.length > 0 ? duplicates.join(', ') : 'なし');
  
  // 3. API接続テスト
  var settings = getSettings(ss);
  try {
    var res = UrlFetchApp.fetch('https://graph.threads.net/v1.0/' + settings.user_id + '?fields=id,username&access_token=' + settings.access_token, {muteHttpExceptions: true});
    console.log('\n=== API ===');
    console.log('ステータス:', res.getResponseCode(), res.getContentText().substring(0, 100));
  } catch(e) {
    console.log('\n=== API エラー ===');
    console.log(e.message);
  }
}



function cleanDuplicateHistory() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hist = ss.getSheetByName('履歴');
  var data = hist.getDataRange().getValues();
  var deleteRows = [];
  var kept = false;
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][2] || '').indexOf('周りに合わせすぎて') !== -1) {
      if (!kept) {
        kept = true; // 1件だけ残す
        console.log('残す: 行' + (i+1));
      } else {
        deleteRows.push(i + 1);
      }
    }
  }
  
  // 下から削除
  deleteRows.sort(function(a, b) { return b - a; });
  deleteRows.forEach(function(r) { hist.deleteRow(r); });
  console.log('削除完了: ' + deleteRows.length + '件削除、1件残し');
}
function showProcessFunction() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('投稿管理');
  var data = sheet.getDataRange().getValues();
  
  // まず error 行の状態を確認
  for (var i = 1; i < data.length; i++) {
    var status = String(data[i][2] || '').trim();
    var id = String(data[i][0] || '');
    if (id === '5d06257f-71f0-4edf-b69e-12a9213465b1') {
      console.log('問題の行' + (i+1) + ': status=' + status + ', error=' + data[i][8] + ', groupId=' + data[i][9] + ', order=' + data[i][10]);
    }
    // night-0219 グループも確認
    if (String(data[i][9] || '') === 'night-0219') {
      console.log('night-0219 行' + (i+1) + ': id=' + id + ', status=' + status + ', order=' + data[i][10] + ', text=' + String(data[i][3] || '').substring(0, 30));
    }
  }
  
  // 重複投稿の履歴確認
  var hist = ss.getSheetByName('履歴');
  var hData = hist.getDataRange().getValues();
  var count = 0;
  for (var j = 1; j < hData.length; j++) {
    if (String(hData[j][2] || '').indexOf('周りに合わせすぎて') !== -1) {
      count++;
    }
  }
  console.log('「周りに合わせすぎて」の履歴件数: ' + count);
}
function fixAndPostRemaining() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('投稿管理');
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    var id = String(data[i][0] || '');
    
    // 行84: processing → posted に変更（既に投稿済みなので）
    if (id === 'tree-1771422277583-1') {
      sheet.getRange(i + 1, 3).setValue('posted');
      console.log('行' + (i+1) + ': posted に変更');
    }
    
    // 行85: ツリーのグループを外して単独投稿として予約し直す
    if (id === 'tree-1771422277583-2') {
      sheet.getRange(i + 1, 10).setValue('');  // groupId をクリア
      sheet.getRange(i + 1, 11).setValue('');  // orderNum をクリア
      sheet.getRange(i + 1, 12).setValue('');  // reply_to_id をクリア
      // 予約時刻を3分後に設定
      var newTime = new Date();
      newTime.setMinutes(newTime.getMinutes() + 3);
      sheet.getRange(i + 1, 7).setValue(newTime);
      sheet.getRange(i + 1, 3).setValue('scheduled');
      console.log('行' + (i+1) + ': 単独投稿として3分後に再予約');
    }
  }
  
  console.log('修復完了。3分後にトリガーが自動投稿します。');
}



function checkCurrentState() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('投稿管理');
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    var groupId = String(data[i][9] || '');
    var id = String(data[i][0] || '');
    if (groupId.indexOf('tree-1771422277583') !== -1 || id.indexOf('tree-1771422277583') !== -1) {
      console.log('行' + (i+1) + ': id=' + id + ', status=' + data[i][2] + ', groupId=' + groupId + ', order=' + data[i][10]);
    }
  }
  
  // 履歴も確認
  var hist = ss.getSheetByName('履歴');
  if (hist) {
    var hData = hist.getDataRange().getValues();
    console.log('--- 履歴の最新5件 ---');
    for (var j = Math.max(1, hData.length - 5); j < hData.length; j++) {
      console.log('履歴行' + (j+1) + ': id=' + hData[j][0] + ', text=' + String(hData[j][2] || '').substring(0, 30) + ', threadPostId=' + hData[j][5]);
    }
  }
}


function retryErrorPost() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('投稿管理');
  
  // 行84のステータスを scheduled に戻す
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === 'tree-1771422277583-1') {
      sheet.getRange(i + 1, 3).setValue('scheduled');  // status
      sheet.getRange(i + 1, 9).setValue('');            // エラーメッセージをクリア
      console.log('行' + (i+1) + ' をscheduledに戻しました');
      break;
    }
  }
  
  // processScheduledPosts を実行
  processScheduledPosts();
}

function testThreadsAPI() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settings = getSettings(ss);
  var token = settings.access_token;
  var userId = settings.user_id;
  
  // 1. プロフィール取得テスト
  try {
    var profileUrl = 'https://graph.threads.net/v1.0/' + userId + '?fields=id,username&access_token=' + token;
    var res = UrlFetchApp.fetch(profileUrl, {muteHttpExceptions: true});
    console.log('プロフィール取得:', res.getResponseCode(), res.getContentText().substring(0, 200));
  } catch(e) {
    console.log('プロフィールエラー:', e.message);
  }
  
  // 2. 投稿テスト（コンテナ作成のみ、公開はしない）
  try {
    var containerUrl = 'https://graph.threads.net/v1.0/' + userId + '/threads';
    var res2 = UrlFetchApp.fetch(containerUrl, {
      method: 'post',
      payload: {
        media_type: 'TEXT',
        text: 'APIテスト（これは公開されません）',
        access_token: token
      },
      muteHttpExceptions: true
    });
    console.log('コンテナ作成:', res2.getResponseCode(), res2.getContentText().substring(0, 300));
  } catch(e) {
    console.log('コンテナエラー:', e.message);
  }
}
function fixBrInPosts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('投稿管理');
  var data = sheet.getDataRange().getValues();
  var fixed = 0;
  for (var i = 1; i < data.length; i++) {
    var text = String(data[i][3] || '');
    if (text.indexOf('<br>') !== -1) {
      var newText = text.replace(/<br\s*\/?>/gi, '\n');
      sheet.getRange(i + 1, 4).setValue(newText);
      fixed++;
      console.log('行' + (i+1) + ' 修正: <br>を改行に置換');
    }
  }
  console.log('修正完了: ' + fixed + '件');
}

function checkTreeGroup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('投稿管理');
  var data = sheet.getDataRange().getValues();
  
  // tree-1771422277583 グループの行を探す
  for (var i = 1; i < data.length; i++) {
    var groupId = String(data[i][9] || '').trim();
    var id = String(data[i][0] || '');
    var status = String(data[i][2] || '').trim();
    var orderNum = data[i][10];
    var text = String(data[i][3] || '').substring(0, 40);
    
    if (groupId.indexOf('tree-1771422277583') !== -1 || id.indexOf('tree-1771422277583') !== -1) {
      console.log('行' + (i+1) + ': id=' + id + ', status=' + status + ', groupId=' + groupId + ', order=' + orderNum + ', text=' + text);
    }
  }
  
  // ついでに error の行も確認
  console.log('--- error行 ---');
  for (var i = 1; i < data.length; i++) {
    var status = String(data[i][2] || '').trim();
    if (status === 'error') {
      console.log('行' + (i+1) + ': id=' + data[i][0] + ', text=' + String(data[i][3] || '').substring(0, 40) + ', errorMsg=' + data[i][8]);
    }
  }
}


function testProcessNow() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('投稿管理');
  var data = sheet.getDataRange().getValues();
  var now = new Date();
  
  console.log('現在時刻:', now.toISOString());
  console.log('行数:', data.length - 1);
  
  for (var i = 1; i < data.length; i++) {
    var status = String(data[i][2]).trim();
    var scheduledTime = data[i][6];
    var text = String(data[i][3]).substring(0, 30);
    
    if (!scheduledTime) continue;
    var scheduled = new Date(scheduledTime);
    var diff = (scheduled.getTime() - now.getTime()) / 1000 / 60; // 分
    
    console.log('行' + (i+1) + ': status=' + status + ', 予約=' + scheduled.toISOString() + ', 差=' + Math.round(diff) + '分, text=' + text);
    
    if (status === 'scheduled' || status === '予約済み') {
      if (scheduled <= now) {
        console.log('  → ★ 投稿対象！');
      } else {
        console.log('  → まだ時刻が来ていません');
      }
    }
  }
}

function EXPORT_ALL_FILES_TO_NEW_SS() {
  // 新しいスプレッドシートを作成
  var newSs = SpreadsheetApp.create('Threads Insight Master - ソースコード_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss'));
  var newSsUrl = newSs.getUrl();
  Logger.log('📄 新規スプレッドシート作成: ' + newSsUrl);

  // デフォルトの空シートを後で削除するために保持
  var defaultSheet = newSs.getSheets()[0];
  var fileCount = 0;

  // ── HTMLファイル ──
  var htmlFiles = [
    'index', 'styles', 'app',
    'screen_welcome', 'screen_dashboard', 'screen_analytics',
    'screen_competitor', 'screen_generate', 'screen_drafts',
    'screen_settings', 'screen_keywords'
  ];

  for (var i = 0; i < htmlFiles.length; i++) {
    try {
      var src = HtmlService.createTemplateFromFile(htmlFiles[i]).getRawContent();
      var sheetName = htmlFiles[i] + '.html';
      var sheet = newSs.insertSheet(sheetName);
      writeSourceToSheet_(sheet, sheetName, 'HTML', src);
      fileCount++;
      Logger.log('✅ ' + sheetName + ' (' + src.split('\n').length + '行)');
    } catch (e) {
      Logger.log('❌ ' + htmlFiles[i] + '.html: ' + e.message);
    }
  }

  // ── GASファイル ──
  var gsFiles = [
    'Code', 'Auth', 'Insights', 'Analytics', 'Gemini',
    'Drafts', 'KeywordSearch', 'Sheets',
    'Config', 'Utils', 'API',
    'test', 'Test', 'TestCompetitor'
  ];
  var found = {};

  for (var i = 0; i < gsFiles.length; i++) {
    var name = gsFiles[i];
    if (found[name]) continue;
    try {
      var src = ScriptApp.getResource(name).getDataAsString();
      found[name] = true;
      var sheetName = name + '.gs';
      var sheet = newSs.insertSheet(sheetName);
      writeSourceToSheet_(sheet, sheetName, 'GS', src);
      fileCount++;
      Logger.log('✅ ' + sheetName + ' (' + src.split('\n').length + '行)');
    } catch (e) {
      // ファイルなし → スキップ
    }
  }

  // デフォルトシートを削除（ファイルが1つ以上あれば）
  if (fileCount > 0) {
    try { newSs.deleteSheet(defaultSheet); } catch (e) {}
  }

  // ── 目次シートを先頭に作成 ──
  var tocSheet = newSs.insertSheet('目次', 0);
  tocSheet.appendRow(['#', 'ファイル名', 'タイプ', '行数']);
  tocSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');

  var allSheets = newSs.getSheets();
  var idx = 1;
  for (var i = 0; i < allSheets.length; i++) {
    var s = allSheets[i];
    if (s.getName() === '目次') continue;
    var name = s.getName();
    var type = name.endsWith('.gs') ? 'GS' : 'HTML';
    // 行数はA2セルに記載済み
    var lineCount = '';
    try { lineCount = s.getRange('B2').getValue(); } catch (e) {}
    tocSheet.appendRow([idx, name, type, lineCount]);
    idx++;
  }

  tocSheet.setColumnWidth(1, 40);
  tocSheet.setColumnWidth(2, 250);
  tocSheet.setColumnWidth(3, 60);
  tocSheet.setColumnWidth(4, 80);

  Logger.log('');
  Logger.log('══════════════════════════════════════');
  Logger.log('✅ エクスポート完了: ' + fileCount + 'ファイル');
  Logger.log('📎 URL: ' + newSsUrl);
  Logger.log('══════════════════════════════════════');

  // URLをダイアログで表示（ブラウザ上で実行時）
  try {
    var htmlOutput = HtmlService
      .createHtmlOutput(
        '<p>エクスポート完了（' + fileCount + 'ファイル）</p>' +
        '<p><a href="' + newSsUrl + '" target="_blank">📎 新しいスプレッドシートを開く</a></p>'
      )
      .setWidth(400)
      .setHeight(120);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'エクスポート完了');
  } catch (e) {
    // UIなし環境ではスキップ
  }
}

/**
 * ソースコードをシートに書き込む（1行1行を別セルに）
 */
function writeSourceToSheet_(sheet, fileName, type, source) {
  var lines = source.split('\n');

  // ヘッダー情報（A1:B1〜B3）
  sheet.getRange('A1').setValue('ファイル名').setFontWeight('bold');
  sheet.getRange('B1').setValue(fileName);
  sheet.getRange('A2').setValue('行数').setFontWeight('bold');
  sheet.getRange('B2').setValue(lines.length);
  sheet.getRange('A3').setValue('タイプ').setFontWeight('bold');
  sheet.getRange('B3').setValue(type);

  // 区切り行
  sheet.getRange('A4').setValue('── ソースコード ──').setFontWeight('bold');
  sheet.getRange(4, 1, 1, 3).setBackground('#f0f0f0');

  // ソースコード（A5〜）- 1行ずつ書き込み
  if (lines.length > 0) {
    var data = lines.map(function(line) { return [line]; });
    sheet.getRange(5, 1, data.length, 1).setValues(data);
  }

  // 書式設定
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 120);

  // コード部分のフォントをmonospaceに
  if (lines.length > 0) {
    sheet.getRange(5, 1, lines.length, 1)
      .setFontFamily('Courier New')
      .setFontSize(10)
      .setWrap(false);
  }

  // A列幅をコード用に広げる
  sheet.setColumnWidth(1, 1200);
}
