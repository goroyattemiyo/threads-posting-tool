// @ts-nocheck
/**
 * Threads投稿ツール - メインAPI
 * 
 * このファイルはWebアプリのバックエンドとして機能します。
 * PWAからのリクエストを処理し、スプレッドシートとThreads APIを連携します。
 */

// ===========================================
// 設定
// ===========================================

var CONFIG = {
  THREADS_API_BASE: 'https://graph.threads.net/v1.0',
  THREADS_AUTH_URL: 'https://threads.net/oauth/authorize',
  THREADS_TOKEN_URL: 'https://graph.threads.net/oauth/access_token',
  SCOPES: 'threads_basic,threads_content_publish,threads_manage_insights,threads_manage_replies'
};

/**
 * URLの末尾スラッシュを除去する
 * @param {string} url
 * @return {string}
 */
function normalizeUrl_(url) {
  return url ? url.replace(/\/+$/, '') : '';
}

// デプロイURLを動的に取得
function getDeploymentUrl() {
  try {
    return ScriptApp.getService().getUrl();
  } catch (e) {
    console.log('デプロイURLの取得に失敗:', e.message);
    return '';
  }
}

// スプレッドシートIDを取得（コンテナバウンドスクリプトの場合）
function getBoundSpreadsheetId() {
  // 1. バインドされたスプレッドシートから取得
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) return ss.getId();
  } catch (e) {
    // スタンドアロンスクリプトの場合
  }
  // 2. PropertiesService から取得（スタンドアロン用）
  try {
    var props = PropertiesService.getUserProperties();
    var savedId = props.getProperty('SHEET_ID');
    if (savedId) {
      console.log('PropertiesService から sheetId 取得:', savedId);
      return savedId;
    }
  } catch (e2) {
    console.log('PropertiesService エラー:', e2.message);
  }
  return null;
}

// sheetId を PropertiesService に保存する関数
function saveSheetIdToProperties(sheetId) {
  if (!sheetId) return { success: false, error: 'sheetId is empty' };
  try {
    PropertiesService.getUserProperties().setProperty('SHEET_ID', sheetId);
    console.log('sheetId を PropertiesService に保存:', sheetId);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}
// ===========================================
// テスト領域
// ===========================================
/**
 * 履歴シートの列ズレデータを検出・修復する
 */
function fixHistoryData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('履歴');
  if (!sheet) {
    console.log('履歴シートが見つかりません');
    return;
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    console.log('履歴データなし');
    return;
  }
  
  var data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
  var fixedCount = 0;
  var deletedRows = [];
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var rowNum = i + 2;
    
    var colA = String(row[0]);  // id
    var colB = String(row[1]);  // account_id
    var colC = String(row[2]);  // text
    var colE = row[4];          // posted_at
    var colF = String(row[5]);  // threads_post_id
    
    // ★検出パターン1: B列にaccount_idではなくテキストが入っている（account_id列が抜けてる）
    // account_idは "acc-" で始まるか "default"
    var isValidAccountId = colB.indexOf('acc-') === 0 || colB === 'default' || colB === '';
    
    if (!isValidAccountId && colB.length > 20) {
      // B列にテキストが入っている → account_idが欠落して1列ズレ
      console.log('列ズレ検出 (行' + rowNum + '): B列にテキストあり → 修復');
      
      // 現在: A=id, B=text, C=media_url, D=posted_at, E=threads_post_id, F=likes, G=replies, H=fetched_at, I=group_id, J=reply_to_id
      // 正しい: A=id, B=account_id, C=text, D=media_url, E=posted_at, F=threads_post_id, G=likes, H=replies, I=fetched_at, J=group_id, K=reply_to_id
      
      var fixedRow = [
        row[0],     // A: id (そのまま)
        '',         // B: account_id (不明なので空)
        row[1],     // C: text (元B列)
        row[2],     // D: media_url (元C列)
        row[3],     // E: posted_at (元D列)
        row[4],     // F: threads_post_id (元E列)
        row[5],     // G: likes (元F列)
        row[6],     // H: replies (元G列)
        row[7],     // I: fetched_at (元H列)
        row[8],     // J: group_id (元I列)
        row[9]      // K: reply_to_id (元J列)
      ];
      
      sheet.getRange(rowNum, 1, 1, 11).setValues([fixedRow]);
      fixedCount++;
      continue;
    }
    
    // ★検出パターン2: C列（text）に "scheduled" や "expired" などステータス値が入っている
    var statusValues = ['scheduled', '予約済み', 'posted', 'error', 'expired', 'processing'];
    if (statusValues.indexOf(colC) !== -1) {
      console.log('ステータス混入検出 (行' + rowNum + '): text="' + colC + '" → 削除対象');
      deletedRows.push(rowNum);
      continue;
    }
    
    // ★検出パターン3: posted_atが空でthreads_post_idも空（未投稿データの混入）
    if (!colE && !colF && colC !== '') {
      console.log('未投稿データ混入 (行' + rowNum + '): threads_post_id空 → 削除対象');
      deletedRows.push(rowNum);
      continue;
    }
  }
  
  // 不正データの行を削除（下から）
  deletedRows.sort(function(a, b) { return b - a; });
  for (var j = 0; j < deletedRows.length; j++) {
    sheet.deleteRow(deletedRows[j]);
    console.log('行削除:', deletedRows[j]);
  }
  
  console.log('修復完了: ' + fixedCount + '件修復、' + deletedRows.length + '件削除');
}


function exportFilesToSheets() {
  // 新しいスプレッドシートを作成
  const ss = SpreadsheetApp.create('GASファイルエクスポート_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd_HHmmss'));
  
  // 各ファイルの内容を取得する関数マップ
  const fileMap = {
    'コード.gs': function() { return ScriptApp.getResource('コード').getDataAsString(); },
    'index.html': function() { return HtmlService.createHtmlOutputFromFile('index').getContent(); },
    'styles.html': function() { return HtmlService.createHtmlOutputFromFile('styles').getContent(); },
    'app.html': function() { return HtmlService.createHtmlOutputFromFile('app').getContent(); },
    'phase1.html': function() { return HtmlService.createHtmlOutputFromFile('phase1').getContent(); },
    'phase2.html': function() { return HtmlService.createHtmlOutputFromFile('phase2').getContent(); },
    'phase3.html': function() { return HtmlService.createHtmlOutputFromFile('phase3').getContent(); },
    'phase4.html': function() { return HtmlService.createHtmlOutputFromFile('phase4').getContent(); }
  };

  const defaultSheet = ss.getSheets()[0];
  let first = true;

  for (const [fileName, getContent] of Object.entries(fileMap)) {
    let sheet;
    if (first) {
      defaultSheet.setName(fileName);
      sheet = defaultSheet;
      first = false;
    } else {
      sheet = ss.insertSheet(fileName);
    }

    let source = '';
    try {
      source = getContent();
    } catch (e) {
      source = '※取得エラー: ' + e.message;
    }

    // ヘッダー
    sheet.getRange('A1').setValue('ファイル名').setFontWeight('bold');
    sheet.getRange('B1').setValue(fileName);

    sheet.getRange('A3').setValue('ソースコード').setFontWeight('bold');

    // ソースを1行1セルで書き込み
    const lines = source.split('\n');
    const outputData = lines.map(function(line) { return [line]; });
    if (outputData.length > 0) {
      sheet.getRange(4, 1, outputData.length, 1).setValues(outputData);
    }

    sheet.setColumnWidth(1, 800);
  }

  const ssUrl = ss.getUrl();
  Logger.log('エクスポート完了: ' + ssUrl);

  try {
    SpreadsheetApp.getUi().alert('エクスポート完了!\n\n' + ssUrl);
  } catch (e) {
    Logger.log(ssUrl);
  }
}

/** 1行=1セルで A列に書き込む（セル5万文字制限に強い） */
function writeSourceToSheetLines_(sheet, fileName, type, source) {
  const lines = source ? source.split('\n') : [''];

  sheet.getRange('A1').setValue('file').setFontWeight('bold');
  sheet.getRange('B1').setValue(fileName);
  sheet.getRange('A2').setValue('type').setFontWeight('bold');
  sheet.getRange('B2').setValue(type);
  sheet.getRange('A3').setValue('lines').setFontWeight('bold');
  sheet.getRange('B3').setValue(lines.length);

  sheet.getRange('A5').setValue('--- source ---').setFontWeight('bold');
  sheet.getRange(5, 1, 1, 2).setBackground('#f0f0f0');

  const values = lines.map(l => [l]);
  sheet.getRange(6, 1, values.length, 1).setValues(values);

  sheet.setColumnWidth(1, 1200);
  sheet.getRange(6, 1, Math.max(1, values.length), 1)
    .setFontFamily('Courier New')
    .setFontSize(10)
    .setWrap(false);
}

function sanitizeSheetName_(name) {
  // 禁止文字: [ ] : * ? / \
  let s = String(name).replace(/[\[\]\:\*\?\/\\]/g, ' ');
  s = s.replace(/\s+/g, ' ').trim();
  if (!s) s = 'sheet';
  if (s.length > 100) s = s.slice(0, 100);
  return s;
}

function uniqueSheetName_(base, used) {
  if (!used.has(base)) return base;
  for (let i = 2; i < 1000; i++) {
    const cand = `${base} (${i})`;
    if (!used.has(cand) && cand.length <= 100) return cand;
  }
  throw new Error('シート名重複が多すぎます: ' + base);
}


function writeSourceToSheet_(sheet, fileName, type, source) {
  var lines = source.split('\n');

  sheet.getRange('A1').setValue('ファイル名').setFontWeight('bold');
  sheet.getRange('B1').setValue(fileName);
  sheet.getRange('A2').setValue('行数').setFontWeight('bold');
  sheet.getRange('B2').setValue(lines.length);
  sheet.getRange('A3').setValue('タイプ').setFontWeight('bold');
  sheet.getRange('B3').setValue(type);

  sheet.getRange('A4').setValue('── ソースコード ──').setFontWeight('bold');
  sheet.getRange(4, 1, 1, 3).setBackground('#f0f0f0');

  if (lines.length > 0) {
    var data = lines.map(function(line) { return [line]; });
    sheet.getRange(5, 1, data.length, 1).setValues(data);
  }

  sheet.setColumnWidth(1, 1200);
  sheet.setColumnWidth(2, 120);

  if (lines.length > 0) {
    sheet.getRange(5, 1, lines.length, 1)
      .setFontFamily('Courier New')
      .setFontSize(10)
      .setWrap(false);
  }
}


// ===========================================
// ここまで
// ===========================================

// ===========================================
// Webアプリ エントリーポイント
// ===========================================
/**
 * アクティブなシートIDを設定（トリガー用）
 */
function setActiveSheetId(sheetId) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('ACTIVE_SHEET_ID', sheetId);
  console.log('ACTIVE_SHEET_ID を設定:', sheetId);
  return { success: true };
}

function updateSpreadsheetId() {
  var newSheetId = '1MrhU_isCTm1AC1saDEYG-e59x0Kn-50L4VfD8yualbI';
  var props = PropertiesService.getScriptProperties();
  props.setProperty('SPREADSHEET_ID', newSheetId);
  console.log('SPREADSHEET_ID を更新しました:', newSheetId);
}

function doGet(e) {
  var DEPLOYMENT_URL = getDeploymentUrl();
  var BOUND_SHEET_ID = getBoundSpreadsheetId();
  
  console.log('=== doGet ===');
  console.log('DEPLOYMENT_URL:', DEPLOYMENT_URL);
  console.log('BOUND_SHEET_ID:', BOUND_SHEET_ID);
  
  var page = (e && e.parameter && e.parameter.page) ? e.parameter.page : '';
  var code = (e && e.parameter && e.parameter.code) ? e.parameter.code : '';
  var state = (e && e.parameter && e.parameter.state) ? e.parameter.state : '';
  
  // sheetId の決定（優先順位：バウンドシート > パラメータ > state）
  var sheetId = '';
  
  if (BOUND_SHEET_ID) {
    sheetId = BOUND_SHEET_ID;
  } else if (e && e.parameter && e.parameter.sheetId) {
    sheetId = e.parameter.sheetId;
  } else if (state) {
    var stateParts = state.split(':::');
    if (stateParts.length > 0 && stateParts[0]) {
      sheetId = stateParts[0];
    }
  }
  
  console.log('使用するsheetId:', sheetId);
  
  // デプロイURLを設定シートに保存（app_idがある場合のみ）
  if (sheetId && DEPLOYMENT_URL) {
    try {
      var ssUrl = SpreadsheetApp.openById(sheetId);
      var settingsUrl = getSettings(ssUrl);
      if (settingsUrl.app_id && (!settingsUrl.app_url || settingsUrl.app_url !== DEPLOYMENT_URL)) {
        saveSettings(ssUrl, { app_url: DEPLOYMENT_URL });
        console.log('app_url を保存しました:', DEPLOYMENT_URL);
      }
    } catch (urlError) {
      console.log('app_url 保存エラー:', urlError.message);
    }
  }
  
  // OAuth コールバック処理（code がある場合）
  var stateSheetId = '';
  if (state) {
    var statePartsOAuth = state.split(':::');
    if (statePartsOAuth.length > 0) {
      stateSheetId = statePartsOAuth[0];
    }
  }
  
  if (code && (stateSheetId || sheetId) && !page) {
    var targetSheetId = stateSheetId || sheetId;
    var tokenResult = { success: false, error: '' };
    
    try {
      var ssToken = SpreadsheetApp.openById(targetSheetId);
      tokenResult = exchangeToken(ssToken, code);
    } catch (tokenErr) {
      console.log('トークン交換エラー:', tokenErr.message);
      tokenResult = { success: false, error: tokenErr.message };
    }
    
    var statusIcon = tokenResult.success ? '✅' : '❌';
    var statusTitle = tokenResult.success ? '認証成功！' : '認証失敗';
    var statusDesc = tokenResult.success 
      ? 'Threadsアカウントとの連携が完了しました。'
      : 'エラー: ' + (tokenResult.error || '不明なエラー');
    
    var callbackHtml = '<!DOCTYPE html><html><head><meta charset="utf-8">' +
      '<meta name="viewport" content="width=device-width,initial-scale=1">' +
      '<title>認証結果</title>' +
      '<style>body{font-family:sans-serif;display:flex;justify-content:center;align-items:center;min-height:100vh;margin:0;background:#f5f5f5;}' +
      '.container{text-align:center;padding:40px;background:white;border-radius:16px;box-shadow:0 2px 10px rgba(0,0,0,0.1);max-width:400px;}' +
      '.icon{font-size:64px;margin-bottom:16px;}' +
      '.title{font-size:24px;font-weight:bold;margin-bottom:8px;}' +
      '.desc{color:#666;margin-bottom:24px;}</style></head>' +
      '<body><div class="container">' +
      '<div class="icon">' + statusIcon + '</div>' +
      '<div class="title">' + statusTitle + '</div>' +
      '<div class="desc">' + statusDesc + '</div>' +
      '<button onclick="window.close()" style="' +
      'padding:12px 32px;font-size:16px;background:linear-gradient(135deg,#00ba7c,#1da1f2);' +
      'color:#fff;border:none;border-radius:8px;cursor:pointer;margin-top:20px;' +
      '">このウィンドウを閉じる</button>' +
      '<p style="margin-top:12px;font-size:13px;color:#888;">' +
      '閉じられない場合は手動でこのタブを閉じ、元の画面に戻ってください' +
      '</p>' +
      '</div>' +
      '<script>localStorage.setItem("threads_tool_sheet_id", "' + targetSheetId + '");</script>' +
      '</body></html>';
    
    return HtmlService.createHtmlOutput(callbackHtml)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  
  // トリガー自動設定（初回のみ）
  if (sheetId) {
    try {
      var ssTrigger = SpreadsheetApp.openById(sheetId);
      ensureScheduleTrigger(ssTrigger);
    } catch (triggerErr) {
      console.log('トリガー設定スキップ:', triggerErr.message);
    }
  }
  
  // 認証開始リクエスト
  if (page === 'auth' && sheetId) {
    try {
      var ssAuth = SpreadsheetApp.openById(sheetId);
      var settingsAuth = getSettings(ssAuth);
      
      if (settingsAuth.app_id && settingsAuth.app_secret) {
        var redirectUri = normalizeUrl_(DEPLOYMENT_URL);
        var stateParam = sheetId + ':::' + Utilities.getUuid();
        
        saveSettings(ssAuth, { oauth_state: stateParam });
        
        var authUrl = 'https://threads.net/oauth/authorize' +
          '?client_id=' + settingsAuth.app_id +
          '&redirect_uri=' + encodeURIComponent(normalizeUrl_(redirectUri)) +
          '&scope=' + encodeURIComponent(CONFIG.SCOPES) +
          '&response_type=code' +
          '&force_authentication=1' +
          '&state=' + encodeURIComponent(stateParam);
        
        return HtmlService.createHtmlOutput(
          '<!DOCTYPE html><html><head><meta name="viewport" content="width=device-width, initial-scale=1">' +
          '<style>body{font-family:-apple-system,sans-serif;text-align:center;padding:40px 20px;background:#f5f5f5;}' +
          '.card{background:white;border-radius:16px;padding:32px 24px;max-width:400px;margin:0 auto;box-shadow:0 4px 12px rgba(0,0,0,0.1);}' +
          '.btn{display:block;width:100%;padding:14px;background:#000;color:white;text-decoration:none;border-radius:12px;font-weight:600;margin-bottom:12px;box-sizing:border-box;border:none;cursor:pointer;}' +
          '.btn-secondary{background:#f0f0f0;color:#333;}</style></head>' +
          '<body><div class="card">' +
          '<div style="font-size:48px;margin-bottom:16px;">🔐</div>' +
          '<h2>Threads認証</h2>' +
          '<p>Threadsアカウントと連携します。</p>' +
          '<a href="' + authUrl + '" class="btn">認証ページを開く</a>' +
          '</div></body></html>'
        ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      }
    } catch (authErr) {
      console.error('Auth redirect error:', authErr);
    }
  }
  
  // 初期データを一括取得
  var initialData = {
    sheetId: sheetId,
    settings: {},
    user: null,
    accounts: [],
    activeAccount: null,
    tokenWarnings: [],
    initialScreen: 'welcome'
  };
  
 if (sheetId) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    
    // 設定を取得
    initialData.settings = getSettings(ss);
    
    // 認証済みの場合
    if (initialData.settings && initialData.settings.access_token) {
      // アカウント一覧（直接配列が返る）
      initialData.accounts = getAccounts(ss) || [];
      
      // アクティブアカウント（直接オブジェクトが返る）
      initialData.activeAccount = getActiveAccount(ss) || null;
      
      // ユーザー情報
      if (initialData.activeAccount) {
        initialData.user = {
          username: initialData.activeAccount.username,
          profilePicUrl: initialData.activeAccount.profilePicUrl
        };
      }
      
      // トークン期限チェック
      var tokenWarnings = checkTokenExpiry(ss);
      initialData.tokenWarnings = Array.isArray(tokenWarnings) ? tokenWarnings : [];
      
      initialData.initialScreen = 'compose';
    } else if (initialData.settings && initialData.settings.app_id) {
      initialData.initialScreen = 'setup-auth';
    } else {
      initialData.initialScreen = 'setup';
    }
  } catch (dataErr) {
    console.error('Initial data error:', dataErr);
  }
}
  
  console.log('初期画面:', initialData.initialScreen);
  
  // HTMLテンプレートを生成
  var template = HtmlService.createTemplateFromFile('index');
  template.serverData = JSON.stringify(initialData);
  template.deploymentUrl = DEPLOYMENT_URL;
  
  return template.evaluate()
    .setTitle('Threads 投稿ツール')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

  

    function createNewSpreadsheet() {
      try {
    var ss = SpreadsheetApp.create('Threads投稿ツール');
    var sheetId = ss.getId();
    var sheetUrl = ss.getUrl();
    var file = DriveApp.getFileById(sheetId);
    
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
    
    var deploymentUrl = '';
    try {
      deploymentUrl = ScriptApp.getService().getUrl();
    } catch (e) {
      console.log('deploymentUrl取得エラー:', e.message);
    }
    
    // 必要なシートを初期化
    initializeSheets(ss);
    
    // デフォルトの「シート1」を削除
    var defaultSheet = ss.getSheetByName('シート1');
    if (defaultSheet && ss.getSheets().length > 1) {
      ss.deleteSheet(defaultSheet);
    }
    
    // 設定シートにURLを書き込む
    var settingsSheet = ss.getSheetByName('設定');
    if (settingsSheet) {
      var data = settingsSheet.getDataRange().getValues();
      
      for (var i = 0; i < data.length; i++) {
        if (data[i][0] === 'spreadsheet_url') {
          settingsSheet.getRange(i + 1, 2).setValue(sheetUrl);
        }
        if (data[i][0] === 'app_url') {
          settingsSheet.getRange(i + 1, 2).setValue(deploymentUrl);
        }
      }
      SpreadsheetApp.flush();
    }
    
    // READMEシートを追加
    var readmeSheet = ss.getSheetByName('README');
    if (!readmeSheet) {
      readmeSheet = ss.insertSheet('README', 0);
    }
    
    var readmeContent = [
      ['Threads投稿ツール'],
      [''],
      ['このスプレッドシートは投稿ツールのデータ保存用です。'],
      [''],
      ['【重要】アプリURL（ブックマークしてください）'],
      [deploymentUrl || '（デプロイ後に設定されます）'],
      [''],
      ['【このスプレッドシートのURL】'],
      [sheetUrl],
      [''],
      ['【シートの説明】'],
      ['・設定：アプリの設定情報（※編集しないでください）'],
      ['・投稿管理：予約投稿のデータ'],
      ['・履歴：投稿履歴'],
      ['・分析：投稿の分析データ'],
      [''],
      ['【別のデバイスからアクセスする場合】'],
      ['1. このスプレッドシートのURLをコピー'],
      ['2. アプリを開いて「既にスプレッドシートがある方」を選択'],
      ['3. URLを貼り付けて接続'],
      [''],
      ['【注意事項】'],
      ['・このスプレッドシートを削除すると設定が消えます'],
      ['・app_secretは絶対に他人に共有しないでください']
    ];
    
    readmeSheet.getRange(1, 1, readmeContent.length, 1).setValues(readmeContent);
    readmeSheet.setColumnWidth(1, 600);
    readmeSheet.getRange(1, 1).setFontSize(18).setFontWeight('bold');
    readmeSheet.getRange(5, 1).setFontWeight('bold').setBackground('#fff3cd');
    readmeSheet.getRange(6, 1).setFontColor('#1a73e8').setFontSize(12);
    readmeSheet.getRange(8, 1).setFontWeight('bold');
    readmeSheet.getRange(9, 1).setFontColor('#1a73e8').setFontSize(11);
    
    ss.setActiveSheet(readmeSheet);
    ss.moveActiveSheet(1);
    
    console.log('Created spreadsheet:', sheetId);
    
    return {
      success: true,
      sheetId: sheetId,
      url: sheetUrl,
      name: ss.getName()
    };
    
  } catch (error) {
    console.error('createNewSpreadsheet error:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const result = processApiRequest(data);
    return createJsonResponse(result);
  } catch (error) {
    console.error('API Error:', error);
    return createJsonResponse({ success: false, error: error.message });
  }
}

function createJsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ===========================================
// スプレッドシート検証・初期化
// ===========================================

function validateSheetId(sheetId) {
  try {
    if (!sheetId) {
      return { valid: false, error: 'シートIDが空です' };
    }
    
    const ss = SpreadsheetApp.openById(sheetId);
    initializeSheets(ss);
    
    return { valid: true, name: ss.getName() };
    
  } catch (e) {
    console.error('validateSheetId error:', e);
    return { 
      valid: false, 
      error: 'スプレッドシートにアクセスできません。URLを確認してください。' 
    };
  }
}

function initializeSheets(ss) {
  if (!ss) {
    console.error('initializeSheets: ss is undefined');
    return;
  }
  
  var requiredSheets = [
    {
      name: '設定',
      headers: null,
      initialData: [
        ['app_id', ''],
        ['app_secret', ''],
        ['access_token', ''],
        ['user_id', ''],
        ['token_expires', ''],
        ['username', ''],
        ['profile_pic_url', ''],
        ['setup_completed', 'FALSE'],
        ['spreadsheet_url', ''],
        ['app_url', ''],
        ['active_account', ''],  // ← 追加
        ['trigger_configured', '']

      ]
    },
    {
      name: 'アカウント',  // ← 新規追加
      headers: ['account_id', 'access_token', 'user_id', 'username', 'profile_pic_url', 'token_expires', 'created_at']
    },
    {
      name: '投稿管理',
      headers: ['id', 'account_id', 'status', 'text', 'media_url', 'media_type', 'scheduled_time', 'created_at', 'updated_at', 'group_id', 'order_num', 'reply_to_id','retry_count']
    },
    {
      name: '履歴',
      headers: ['id', 'account_id', 'text', 'media_url', 'posted_at', 'threads_post_id', 'likes', 'replies', 'fetched_at', 'group_id', 'reply_to_id']  // ← account_id追加
    },
    {
      name: '分析',
      headers: ['post_id', 'date', 'likes', 'replies', 'views', 'fetched_at']
    }
  ];
  
  requiredSheets.forEach(function(sheetDef) {
    var sheet = ss.getSheetByName(sheetDef.name);
    
    if (!sheet) {
      sheet = ss.insertSheet(sheetDef.name);
      
      if (sheetDef.initialData) {
        sheet.getRange(1, 1, sheetDef.initialData.length, sheetDef.initialData[0].length)
          .setValues(sheetDef.initialData);
      } else if (sheetDef.headers) {
        sheet.getRange(1, 1, 1, sheetDef.headers.length)
          .setValues([sheetDef.headers]);
        sheet.getRange(1, 1, 1, sheetDef.headers.length)
          .setFontWeight('bold');
      }
    } else {
      // 既存シートの場合、必要なキー/列を確認して追加
      if (sheetDef.name === '設定') {
        ensureSettingsKeys(sheet, sheetDef.initialData);
      } else if (sheetDef.headers) {
        ensureSheetHeaders(sheet, sheetDef.headers);
      }
    }
  });
}

/**
 * シートのヘッダーを確認し、不足している列を追加
 */
function ensureSheetHeaders(sheet, requiredHeaders) {
  var lastCol = sheet.getLastColumn();
  
  if (lastCol === 0) {
    // 空のシート → ヘッダーを設定
    sheet.getRange(1, 1, 1, requiredHeaders.length)
      .setValues([requiredHeaders]);
    sheet.getRange(1, 1, 1, requiredHeaders.length)
      .setFontWeight('bold');
    return;
  }
  
  var currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  
  // 各必要なヘッダーが存在するか確認
  requiredHeaders.forEach(function(header, index) {
    if (currentHeaders.indexOf(header) === -1) {
      // ヘッダーが存在しない場合、適切な位置に列を挿入
      // account_id は2列目に挿入する特別処理
      if (header === 'account_id' && index === 1) {
        sheet.insertColumnAfter(1);
        sheet.getRange(1, 2).setValue(header).setFontWeight('bold');
        console.log('列を挿入しました: ' + header + ' (B列)');
      }
    }
  });
}


function ensureSettingsKeys(sheet, requiredKeys) {
  const data = sheet.getDataRange().getValues();
  const existingKeys = data.map(row => row[0]);
  
  requiredKeys.forEach(keyValue => {
    const key = keyValue[0];
    if (!existingKeys.includes(key)) {
      sheet.appendRow(keyValue);
    }
  });
}

// ===========================================
// 設定管理
// ===========================================

function getSettings(ss) {
  console.log('=== getSettings 開始 ===');
  
  var sheet = ss.getSheetByName('設定');
  if (!sheet) {
    console.log('設定シートが見つかりません');
    return {};
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow === 0) {
    console.log('設定シートが空です');
    return {};
  }
  
  var data = sheet.getRange(1, 1, lastRow, 2).getValues();
  var settings = {};
  
  for (var i = 0; i < data.length; i++) {
    var key = data[i][0];
    var value = data[i][1];
    if (key && key !== '') {
      settings[key] = value;
    }
  }
  
  console.log('取得した設定:', JSON.stringify(settings));
  return settings;
}

function saveSettings(ss, params) {
  console.log('=== saveSettings 開始 ===');
  console.log('params:', JSON.stringify(params));
  
  // spreadsheet_url が設定される場合、このシートIDを登録リストに追加
  if (params.spreadsheet_url || params.setup_completed === 'TRUE') {
    var sheetId = ss.getId();
    registerSheetId(sheetId);
  }
  
  var sheet = ss.getSheetByName('設定');
  
  var sheet = ss.getSheetByName('設定');
  if (!sheet) {
    console.log('設定シートが見つかりません。作成します。');
    sheet = ss.insertSheet('設定');
  }
  
  var settingsMap = {
    'app_id': params.app_id,
    'app_secret': params.app_secret,
    'access_token': params.access_token,
    'user_id': params.user_id,
    'token_expires': params.token_expires,
    'username': params.username,
    'profile_pic_url': params.profile_pic_url,
    'setup_completed': params.setup_completed,
    'oauth_state': params.oauth_state,
    'spreadsheet_url': params.spreadsheet_url,
    'app_url': params.app_url,
    'active_account': params.active_account,
    'trigger_configured': params.trigger_configured
  };
  
  var lastRow = sheet.getLastRow();
  var data = lastRow > 0 ? sheet.getRange(1, 1, lastRow, 2).getValues() : [];
  
  for (var key in settingsMap) {
    var value = settingsMap[key];
    if (value === undefined || value === null) {
      continue;
    }
    
    var found = false;
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(value);
        console.log('更新:', key, '=', value);
        found = true;
        break;
      }
    }
    
    if (!found) {
      var newRow = sheet.getLastRow() + 1;
      sheet.getRange(newRow, 1).setValue(key);
      sheet.getRange(newRow, 2).setValue(value);
      console.log('追加:', key, '=', value);
    }
  }
  
  SpreadsheetApp.flush();
  
  console.log('=== saveSettings 完了 ===');
  return { success: true };
}


// ===========================================
// OAuth認証
// ===========================================

function getAuthUrl(ss) {
  var settings = getSettings(ss);
  var appId = settings.app_id;
  
  // 設定シートの app_url を使用（ハードコードしない）
  var redirectUri = settings.app_url;
  
  if (!redirectUri) {
    // app_url が未設定の場合は動的に取得
    redirectUri = getDeploymentUrl();
  }
  
  if (!appId) {
    return { success: false, error: 'App IDを設定してください' };
  }
  
  if (!redirectUri) {
    return { success: false, error: 'アプリURLが取得できません' };
  }
  
  // stateにsheetIdを含める
  var sheetId = ss.getId();
  var state = sheetId + ':::' + Utilities.getUuid();
  
  redirectUri = normalizeUrl_(redirectUri);

  console.log('=== getAuthUrl ===');
  console.log('app_id:', appId);
  console.log('redirect_uri:', redirectUri);
  console.log('sheetId:', sheetId);
  console.log('state:', state);
  
  saveSettings(ss, { oauth_state: state });
  
  var authUrl = CONFIG.THREADS_AUTH_URL +
  '?client_id=' + appId +
  '&redirect_uri=' + encodeURIComponent(normalizeUrl_(redirectUri)) +
  '&scope=' + CONFIG.SCOPES +
  '&response_type=code' +
  '&force_authentication=1' +
  '&state=' + state;
  
  return {
    success: true,
    url: authUrl,
    state: state
  };
}


function exchangeToken(ss, code) {
  var settings = getSettings(ss);
  var appId = settings.app_id;
  var appSecret = settings.app_secret;
  
  // 設定シートの app_url を使用（ハードコードしない）
  var redirectUri = settings.app_url;
  
  if (!redirectUri) {
    redirectUri = getDeploymentUrl();
  }
  redirectUri = normalizeUrl_(redirectUri);
  
  // 短期トークンを取得
  var tokenUrl = 'https://graph.threads.net/oauth/access_token';
  
  var payload = {
    client_id: appId,
    client_secret: appSecret,
    grant_type: 'authorization_code',
    redirect_uri: normalizeUrl_(redirectUri),
    code: code
  };
  
  var tokenResponse = UrlFetchApp.fetch(tokenUrl, {
    method: 'post',
    payload: payload,
    muteHttpExceptions: true
  });
  
  var responseText = tokenResponse.getContentText();
  var tokenData = JSON.parse(responseText);
  
  if (tokenData.error) {
    // エラー詳細を含めて返す
    throw new Error('Threads API Error: ' + responseText);
  }
  
  var shortLivedToken = tokenData.access_token;
  var userId = String(tokenData.user_id);
  
  // 長期トークンに交換
  var longTokenUrl = 'https://graph.threads.net/access_token' +
    '?grant_type=th_exchange_token' +
    '&client_secret=' + appSecret +
    '&access_token=' + shortLivedToken;
  
  var longTokenResponse = UrlFetchApp.fetch(longTokenUrl, {
    muteHttpExceptions: true
  });
  
  var longTokenData = JSON.parse(longTokenResponse.getContentText());
  
  if (longTokenData.error) {
    throw new Error('Long token error: ' + longTokenResponse.getContentText());
  }
  
  var accessToken = longTokenData.access_token;
  var expiresIn = longTokenData.expires_in;
  var expiresAt = new Date(Date.now() + expiresIn * 1000).toISOString();
  
  // ユーザー情報を取得
  var username = '';
  var profilePicUrl = '';
  
  try {
    var userInfoUrl = 'https://graph.threads.net/' + userId + 
      '?fields=id,username,threads_profile_picture_url&access_token=' + accessToken;
    
    var userInfoResponse = UrlFetchApp.fetch(userInfoUrl, {
      muteHttpExceptions: true
    });
    
    var userInfo = JSON.parse(userInfoResponse.getContentText());
    
    if (userInfo.username) {
      username = userInfo.username;
    }
    if (userInfo.threads_profile_picture_url) {
      profilePicUrl = userInfo.threads_profile_picture_url;
    }
  } catch (e) {
    // ユーザー情報取得エラーは無視
  }
  
  // 設定を保存（後方互換性のため）
  saveSettings(ss, {
    access_token: accessToken,
    user_id: userId,
    token_expires: expiresAt,
    username: username,
    profile_pic_url: profilePicUrl,
    setup_completed: 'TRUE'
  });
  
  // アカウントシートにも追加
  var accountResult = addAccount(ss, {
    userId: userId,
    accessToken: accessToken,
    username: username,
    profilePicUrl: profilePicUrl,
    tokenExpires: expiresAt
  });
  
  console.log('アカウント追加結果:', accountResult);
  
  return {
    success: true,
    user_id: userId,
    username: username,
    expires_at: expiresAt,
    account_id: accountResult.accountId,
    is_new_account: accountResult.isNew
  };
}

// ===========================================
// ユーザー情報
// ===========================================

function getUserProfile(ss) {
  // まずアクティブアカウントを確認
  var activeAccount = getActiveAccount(ss);
  
  if (activeAccount && activeAccount.accessToken) {
    // アクティブアカウントの情報を返す
    return {
      success: true,
      user: {
        username: activeAccount.username || '',
        profilePicUrl: activeAccount.profilePicUrl || '',
        userId: activeAccount.userId || ''
      }
    };
  }
  
  // 後方互換性：設定シートから取得
  var settings = getSettings(ss);
  
  if (!settings.access_token) {
    throw new Error('認証が必要です');
  }
  
  var username = settings.username || '';
  var profilePicUrl = settings.profile_pic_url || '';
  var userId = String(settings.user_id);
  
  if (!username && userId && settings.access_token) {
    try {
      var userInfoUrl = CONFIG.THREADS_API_BASE + '/' + userId + 
        '?fields=id,username,threads_profile_picture_url&access_token=' + settings.access_token;
      
      var response = UrlFetchApp.fetch(userInfoUrl, {
        muteHttpExceptions: true
      });
      
      var userInfo = JSON.parse(response.getContentText());
      
      if (userInfo.username) {
        username = userInfo.username;
        profilePicUrl = userInfo.threads_profile_picture_url || '';
        
        saveSettings(ss, {
          username: username,
          profile_pic_url: profilePicUrl
        });
      }
    } catch (e) {
      console.error('getUserProfile fetch error:', e.message);
    }
  }
  
  return {
    success: true,
    user: {
      username: username,
      profilePicUrl: profilePicUrl,
      userId: userId
    }
  };
}


function uploadImage(ss, base64Data, fileName) {
  try {
    console.log('=== uploadImage (Catbox) ===');
    console.log('fileName:', fileName);
    
    // MIMEタイプを判定
    var mimeType = 'image/jpeg';
    if (base64Data.indexOf('data:image/png') === 0) {
      mimeType = 'image/png';
    } else if (base64Data.indexOf('data:image/gif') === 0) {
      mimeType = 'image/gif';
    } else if (base64Data.indexOf('data:image/webp') === 0) {
      mimeType = 'image/webp';
    }
    
    // data:image/xxx;base64, の部分を除去
    var base64Content = base64Data.replace(/^data:image\/\w+;base64,/, '');
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Content), mimeType, fileName);
    
    // ファイルサイズチェック（8MB以下 - Threads API制限）
    if (blob.getBytes().length > 8 * 1024 * 1024) {
      return {
        success: false,
        error: '画像は8MB以下にしてください'
      };
    }
    
    // Catbox APIにアップロード
    var formData = {
      'reqtype': 'fileupload',
      'fileToUpload': blob
    };
    
    var options = {
      'method': 'post',
      'payload': formData,
      'muteHttpExceptions': true
    };
    
    var response = UrlFetchApp.fetch('https://catbox.moe/user/api.php', options);
    var responseText = response.getContentText();  
    
    console.log('Catbox response:', responseText);
    
    // Catboxは成功時にURLを直接返す
    if (responseText && responseText.indexOf('https://files.catbox.moe/') === 0) {
      var publicUrl = responseText.trim();
      
      console.log('Uploaded to Catbox:', publicUrl);
      
      return {
        success: true,
        url: publicUrl,
        name: fileName,
        mimeType: mimeType
      };
    } else {
      throw new Error('Catboxアップロード失敗: ' + responseText);
    }
    
  } catch (error) {
    console.error('uploadImage error:', error);
    return {
      success: false,
      error: error.message
    };
  }
}


function uploadVideo(ss, base64Data, fileName) {
  try {
    console.log('=== uploadVideo (Catbox) ===');
    console.log('fileName:', fileName);
    
    // MIMEタイプを判定
    var mimeType = 'video/mp4';
    if (base64Data.indexOf('data:video/quicktime') === 0) {
      mimeType = 'video/quicktime';
    } else if (base64Data.indexOf('data:video/webm') === 0) {
      mimeType = 'video/webm';
    } else if (base64Data.indexOf('data:video/mov') === 0) {
      mimeType = 'video/quicktime';
    }
    
    // data:video/xxx;base64, の部分を除去
    var base64Content = base64Data.replace(/^data:video\/\w+;base64,/, '');
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Content), mimeType, fileName);
    
    // ファイルサイズチェック（100MB以下）
    if (blob.getBytes().length > 100 * 1024 * 1024) {
      return {
        success: false,
        error: '動画は100MB以下にしてください'
      };
    }
    
    // Catbox APIにアップロード
    var formData = {
      'reqtype': 'fileupload',
      'fileToUpload': blob
    };
    
    var options = {
      'method': 'post',
      'payload': formData,
      'muteHttpExceptions': true
    };
    
    var response = UrlFetchApp.fetch('https://catbox.moe/user/api.php', options);
    var responseText = response.getContentText();  
    
    console.log('Catbox response:', responseText);
    
    // Catboxは成功時にURLを直接返す
    if (responseText && responseText.indexOf('https://files.catbox.moe/') === 0) {
      var publicUrl = responseText.trim();
      
      console.log('Uploaded to Catbox:', publicUrl);
      
      return {
        success: true,
        url: publicUrl,
        name: fileName,
        mimeType: mimeType
      };
    } else {
      throw new Error('Catboxアップロード失敗: ' + responseText);
    }
    
  } catch (error) {
    console.error('uploadVideo error:', error);
    return {
      success: false,
      error: error.message
    };
  }
}
// ===========================================
// 投稿機能
// ===========================================
/**
 * ツリー投稿を即時実行
 */
function createTreePost(posts, sheetId) {
  if (!sheetId) {
    // フォールバック: スクリプトプロパティから取得
    var props = PropertiesService.getScriptProperties();
    sheetId = props.getProperty('ACTIVE_SHEET_ID');
  }
  
  if (!sheetId) {
    return { success: false, error: 'シートが設定されていません' };
  }
  
  var ss = SpreadsheetApp.openById(sheetId);
  var settings = getSettings(ss);
  
  if (!settings.access_token) {
    return { success: false, error: '認証が必要です' };
  }
  
  var historySheet = ss.getSheetByName('履歴');
  var groupId = 'tree-' + Date.now();
  var lastPostId = null;
  var postedCount = 0;
  
  console.log('ツリー投稿開始:', posts.length + '件');
  
  for (var i = 0; i < posts.length; i++) {
    var post = posts[i];
    var replyToId = lastPostId;
    
    console.log('投稿 ' + (i + 1) + '/' + posts.length + ':', post.text.substring(0, 30));
    
    try {
      var result;
      
      if (i === 0) {
        // 親投稿
        result = createPost(ss, post.text, post.mediaUrl, post.mediaType);
      } else {
        // 返信投稿
        result = createReplyPost(ss, post.text, post.mediaUrl, post.mediaType, replyToId);
      }
      
      if (result && result.success) {
        console.log('投稿成功:', result.postId);
        lastPostId = result.postId;
        postedCount++;
        
        // 履歴に追加
        if (historySheet) {
          var timestamp = new Date().toISOString();
          historySheet.appendRow([
            'tree-' + Date.now() + '-' + i,
            post.text,
            post.mediaUrl || '',
            timestamp,
            result.postId,
            0,
            0,
            timestamp,
            groupId,
            replyToId || ''
          ]);
        }
        
        // API制限対策
        if (i < posts.length - 1) {
          Utilities.sleep(3000);
        }
        
      } else {
        var errorMsg = result ? result.error : 'Unknown error';
        console.error('投稿失敗:', errorMsg);
        return { 
          success: false, 
          error: '投稿 ' + (i + 1) + ' で失敗: ' + errorMsg,
          postedCount: postedCount 
        };
      }
      
    } catch (e) {
      console.error('投稿エラー:', e.message);
      return { 
        success: false, 
        error: '投稿 ' + (i + 1) + ' でエラー: ' + e.message,
        postedCount: postedCount 
      };
    }
  }
  
  console.log('ツリー投稿完了:', postedCount + '件');
  
  return { success: true, postedCount: postedCount, groupId: groupId };
}

/**
 * ツリー投稿を予約
 */
function scheduleTreePost(posts, scheduledTime, sheetId) {
  var props = PropertiesService.getScriptProperties();
  var activeSheetId = props.getProperty('ACTIVE_SHEET_ID');
  
  if (!activeSheetId) {
    return { success: false, error: 'シートが設定されていません' };
  }
  
  var ss = SpreadsheetApp.openById(activeSheetId);
  var sheet = ss.getSheetByName('投稿管理');
  
  if (!sheet) {
    return { success: false, error: '投稿管理シートが見つかりません' };
  }
  
  // アクティブアカウントを取得
  var activeAccount = getActiveAccount(ss);
  var accountId = activeAccount ? activeAccount.accountId : '';
  
  var treeGroupId = 'tree-' + Date.now();
  var scheduledDate = new Date(scheduledTime);
  var now = new Date();
  
  // フォーマット: YYYY/MM/DD HH:MM:SS
  var formattedTime = Utilities.formatDate(scheduledDate, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
  var formattedNow = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
  
  console.log('ツリー投稿予約:', posts.length + '件', 'グループID:', treeGroupId, '予定時刻:', formattedTime);
  
  for (var i = 0; i < posts.length; i++) {
    var post = posts[i];
    var postId = treeGroupId + '-' + (i + 1);
    
    // ヘッダー順: A:id, B:account_id, C:status, D:text, E:media_url, F:media_type, G:scheduled_time, H:created_at, I:updated_at, J:group_id, K:order_num, L:reply_to_id
    sheet.appendRow([
      postId,              // A: id
      accountId,           // B: account_id
      'scheduled',         // C: status
      post.text,           // D: text
      post.mediaUrl || '', // E: media_url
      post.mediaType || '',// F: media_type
      formattedTime,       // G: scheduled_time
      formattedNow,        // H: created_at
      '',                  // I: updated_at
      treeGroupId,         // J: group_id
      i + 1,               // K: order_num
      ''                   // L: reply_to_id
    ]);
    
    console.log('予約追加:', postId, 'order:', i + 1);
  }
  
  return { 
    success: true, 
    message: posts.length + '件のツリー投稿を予約しました',
    groupId: treeGroupId,
    scheduledTime: formattedTime
  };
}

/**
 * 返信投稿を作成
 */
function createReplyPost(ss, text, mediaUrl, mediaType, replyToId) {
  var auth = getActiveAccountAuth(ss);
  
  if (!auth || !auth.accessToken) {
    var settings = getSettings(ss);
    if (!settings.access_token) {
      return { success: false, error: '認証が必要です' };
    }
    auth = {
      accessToken: settings.access_token,
      userId: String(settings.user_id)
    };
  }
  
  var userId = auth.userId;
  var accessToken = auth.accessToken;
  
  try {
    // コンテナ作成パラメータ
    var containerParams = {
      text: text,
      reply_to_id: replyToId,
      access_token: accessToken  // ← 修正
    };
    
    // メディアがある場合
    if (mediaUrl && mediaType) {
      if (mediaType === 'IMAGE') {
        containerParams.media_type = 'IMAGE';
        containerParams.image_url = mediaUrl;
      } else if (mediaType === 'VIDEO') {
        containerParams.media_type = 'VIDEO';
        containerParams.video_url = mediaUrl;
      }
    } else {
      // テキストのみの場合
      containerParams.media_type = 'TEXT';
    }
    
    console.log('返信コンテナパラメータ:', JSON.stringify(containerParams));
    
    // コンテナ作成
    var containerUrl = CONFIG.THREADS_API_BASE + '/' + userId + '/threads';
    var containerResponse = UrlFetchApp.fetch(containerUrl, {
      method: 'post',
      payload: containerParams,
      muteHttpExceptions: true
    });
    
    var containerResult = JSON.parse(containerResponse.getContentText());
    console.log('返信コンテナ結果:', JSON.stringify(containerResult));
    
    if (containerResult.error) {
      return { success: false, error: containerResult.error.message };
    }
    
    var containerId = containerResult.id;
    
    // コンテナの準備状況を確認（最大30秒待機）
    var maxAttempts = 6;
    var waitTime = 5000;
    var isReady = false;
    
    for (var attempt = 0; attempt < maxAttempts; attempt++) {
      Utilities.sleep(waitTime);
      
      // ステータス確認
      var statusUrl = CONFIG.THREADS_API_BASE + '/' + containerId + '?fields=status&access_token=' + accessToken;  // ← 修正
      var statusResponse = UrlFetchApp.fetch(statusUrl, { muteHttpExceptions: true });
      var statusResult = JSON.parse(statusResponse.getContentText());
      
      console.log('コンテナステータス確認 ' + (attempt + 1) + ':', JSON.stringify(statusResult));
      
      if (statusResult.status === 'FINISHED') {
        isReady = true;
        break;
      } else if (statusResult.status === 'ERROR') {
        return { success: false, error: 'コンテナ作成エラー' };
      }
    }
    
    if (!isReady) {
      console.log('コンテナ準備タイムアウト、公開を試行します');
    }
    
    // 公開
    var publishUrl = CONFIG.THREADS_API_BASE + '/' + userId + '/threads_publish';
    var publishResponse = UrlFetchApp.fetch(publishUrl, {
      method: 'post',
      payload: {
        creation_id: containerId,
        access_token: accessToken  // ← 修正
      },
      muteHttpExceptions: true
    });
    
    var publishResult = JSON.parse(publishResponse.getContentText());
    console.log('返信公開結果:', JSON.stringify(publishResult));
    
    if (publishResult.error) {
      console.log('返信公開失敗:', publishResult.error.message);
      return { success: false, error: publishResult.error.message };
    }
    
    return { success: true, postId: publishResult.id };
    
  } catch (e) {
    console.error('返信投稿エラー:', e.message);
    return { success: false, error: e.message };
  }
}


/**
 * ツリー投稿を処理（グループ単位）
 */
function processTreePosts(ss, groupId, now) {
  var TREE_POST_LIMIT = 10;
  var sheet = ss.getSheetByName('投稿管理');
  if (!sheet) return { success: false, error: '投稿管理シートが見つかりません' };
  
  var data = sheet.getDataRange().getValues();
  
  var colIndex = {
    id: 0,
    account_id: 1,
    status: 2,
    text: 3,
    media_url: 4,
    media_type: 5,
    scheduled_time: 6,
    group_id: 9,
    order_num: 10,
    reply_to_id: 11
  };
  
  var treePosts = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[colIndex.group_id] === groupId && row[colIndex.status] !== 'posted' && row[colIndex.status] !== 'processing') {
      treePosts.push({
        rowIndex: i,
        id: row[colIndex.id],
        accountId: row[colIndex.account_id],  // ★追加
        text: row[colIndex.text],
        mediaUrl: row[colIndex.media_url],
        mediaType: row[colIndex.media_type],
        orderNum: Number(row[colIndex.order_num]) || 1,
        scheduledTime: new Date(row[colIndex.scheduled_time])
      });
    }
  }
  
  treePosts.sort(function(a, b) {
    return a.orderNum - b.orderNum;
  });

  if (treePosts.length > TREE_POST_LIMIT) {
    return { success: false, error: 'ツリー投稿は最大' + TREE_POST_LIMIT + '件までです' };
  }

  if (treePosts.length === 0) {
    return { success: false, error: 'ツリー投稿が見つかりません' };
  }
  
  console.log('ツリー投稿を処理:', groupId, treePosts.length + '件');
  
  var postedRows = [];
  var lastPostId = null;
  
  for (var j = 0; j < treePosts.length; j++) {
    var post = treePosts[j];
    var rowNum = post.rowIndex + 1;
    var replyToId = lastPostId;
    
    sheet.getRange(rowNum, colIndex.status + 1).setValue('processing');
    SpreadsheetApp.flush();
    
    try {
      var result;
      
      if (j === 0) {
        // 親投稿（createPost内部でaddToHistoryが呼ばれる）
        console.log('親投稿:', post.text.substring(0, 30));
        result = createPost(ss, post.text, post.mediaUrl, post.mediaType);
      } else {
        // 返信投稿（createReplyPostは履歴追加しないので手動で追加）
        console.log('返信投稿:', post.text.substring(0, 30), '-> reply_to:', replyToId);
        result = createReplyPost(ss, post.text, post.mediaUrl, post.mediaType, replyToId);
      }
      
      if (result && result.success) {
        console.log('投稿成功:', result.postId);
        lastPostId = result.postId;
        
        sheet.getRange(rowNum, colIndex.reply_to_id + 1).setValue(replyToId || '');
        
        // 返信投稿のみ履歴追加（親投稿はcreatePost内で追加済み）
        if (j > 0) {
          addToHistory(ss, post.text, post.mediaUrl, result.postId, post.accountId || '');
        }
        
        sheet.getRange(rowNum, colIndex.status + 1).setValue('posted');
        postedRows.push(rowNum);
        
      } else {
        var errorMsg = result ? result.error : 'Unknown error';
        console.error('投稿失敗:', errorMsg);
        sheet.getRange(rowNum, colIndex.status + 1).setValue('error');
        sheet.getRange(rowNum, 9).setValue(errorMsg);
        return { success: false, error: errorMsg, postedCount: j };
      }
      
      if (j < treePosts.length - 1) {
        Utilities.sleep(3000);
      }
      
    } catch (e) {
      console.error('投稿エラー:', e.message);
      sheet.getRange(rowNum, colIndex.status + 1).setValue('error');
      sheet.getRange(rowNum, 9).setValue(e.message);
      return { success: false, error: e.message, postedCount: j };
    }
  }
  
  postedRows.sort(function(a, b) { return b - a; });
  for (var k = 0; k < postedRows.length; k++) {
    sheet.deleteRow(postedRows[k]);
  }
  
  console.log('ツリー投稿完了:', groupId, postedRows.length + '件投稿');
  
  return { success: true, postedCount: postedRows.length };
}

function createPost(ss, text, mediaUrl, mediaType) {
  // アクティブアカウントの認証情報を取得
  var auth = getActiveAccountAuth(ss);
  
  if (!auth || !auth.accessToken) {
    // 後方互換性：設定シートから取得
    var settings = getSettings(ss);
    if (!settings.access_token) {
      throw new Error('認証が必要です');
    }
    auth = {
      accessToken: settings.access_token,
      userId: String(settings.user_id)
    };
  }
  
  var accessToken = auth.accessToken;
  var userId = String(auth.userId);
  
  console.log('=== createPost ===');
  console.log('userId:', userId);
  
  var containerParams = {
    text: text,
    access_token: accessToken
  };
  
  if (mediaUrl && mediaUrl.trim() !== '') {
    containerParams.media_type = mediaType || 'IMAGE';
    if (mediaType === 'VIDEO') {
      containerParams.video_url = mediaUrl;
    } else {
      containerParams.image_url = mediaUrl;
    }
  } else {
    containerParams.media_type = 'TEXT';
  }
  
  console.log('Creating container with params:', JSON.stringify(containerParams));
  
  // コンテナ作成
  var containerResponse = UrlFetchApp.fetch(
    CONFIG.THREADS_API_BASE + '/' + userId + '/threads',
    {
      method: 'POST',
      payload: containerParams,
      muteHttpExceptions: true
    }
  );
  
  var containerText = containerResponse.getContentText();
  console.log('Container response:', containerText);
  
  var containerData = JSON.parse(containerText);
  
  if (containerData.error) {
    throw new Error(containerData.error.message);
  }
  
  var containerId = containerData.id;
  
  // 処理待ち
  Utilities.sleep(3000);
  
  // 公開
  var publishResponse = UrlFetchApp.fetch(
    CONFIG.THREADS_API_BASE + '/' + userId + '/threads_publish',
    {
      method: 'POST',
      payload: {
        creation_id: containerId,
        access_token: accessToken
      },
      muteHttpExceptions: true
    }
  );
  
  var publishText = publishResponse.getContentText();
  console.log('Publish response:', publishText);
  
  var publishData = JSON.parse(publishText);
  
  if (publishData.error) {
    throw new Error(publishData.error.message);
  }
  
  // 履歴に追加（account_id付き）
  var activeAccount = getActiveAccount(ss);
  var accountId = activeAccount ? activeAccount.accountId : 'default';
  addToHistory(ss, text, mediaUrl, publishData.id, accountId);
  
  return { 
    success: true, 
    postId: publishData.id 
  };
}


function waitForMediaProcessing(accessToken, containerId) {
  const maxAttempts = 30;
  const waitTime = 2000;
  
  for (let i = 0; i < maxAttempts; i++) {
    const statusResponse = UrlFetchApp.fetch(
      `${CONFIG.THREADS_API_BASE}/${containerId}?fields=status&access_token=${accessToken}`
    );
    
    const statusData = JSON.parse(statusResponse.getContentText());
    
    if (statusData.status === 'FINISHED') {
      return true;
    }
    
    if (statusData.status === 'ERROR') {
      throw new Error('メディアの処理に失敗しました');
    }
    
    Utilities.sleep(waitTime);
  }
  
  throw new Error('メディアの処理がタイムアウトしました');
}

function schedulePost(ss, text, mediaUrl, mediaType, scheduledTime) {
  var sheet = ss.getSheetByName('投稿管理');
  
  var id = Utilities.getUuid();
  var now = new Date();
  
  // アクティブアカウントのIDを取得
  var activeAccount = getActiveAccount(ss);
  var accountId = activeAccount ? activeAccount.accountId : 'default';
  
  // 新しい列順序: id, account_id, status, text, media_url, media_type, scheduled_time, created_at, updated_at, group_id, order_num, reply_to_id
  sheet.appendRow([
    id,
    accountId,
    '予約済み',
    text,
    mediaUrl || '',
    mediaType || '',
    new Date(scheduledTime),
    now,
    now,
    '',  // group_id
    '',  // order_num
    ''   // reply_to_id
  ]);
  
  return { success: true, postId: id };
}

/**
 * 予約投稿用トリガーの自動設定（初回のみ）
 * 修正版: ACTIVE_SHEET_IDも確実に設定
 */
function ensureScheduleTrigger(ss) {
  try {
    var settings = getSettings(ss);
    
    // ★修正: ACTIVE_SHEET_IDを毎回保存（バウンドでない場合の保険）
    var props = PropertiesService.getScriptProperties();
    var currentActiveId = props.getProperty('ACTIVE_SHEET_ID');
    var ssId = ss.getId();
    
    if (!currentActiveId || currentActiveId !== ssId) {
      props.setProperty('ACTIVE_SHEET_ID', ssId);
      console.log('ACTIVE_SHEET_ID を更新:', ssId);
    }
    
    // 既に設定済みならスキップ
    if (settings.trigger_configured) {
      return;
    }
    
    // トリガーを確認・設定
    var triggers = ScriptApp.getProjectTriggers();
    var hasScheduleTrigger = false;
    
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'processScheduledPosts') {
        hasScheduleTrigger = true;
        break;
      }
    }
    
    if (!hasScheduleTrigger) {
      ScriptApp.newTrigger('processScheduledPosts')
        .timeBased()
        .everyMinutes(1)
        .create();
      console.log('予約投稿トリガーを自動設定しました');
    }
    
    // フラグを保存
    saveSettings(ss, { trigger_configured: 'TRUE' });
    
  } catch (e) {
    console.log('トリガー設定エラー:', e.message);
  }
}

function getScheduledPosts(ss, showAllAccounts) {
  var sheet = ss.getSheetByName('投稿管理');
  if (!sheet) {
    console.log('投稿管理シートが見つかりません');
    return [];
  }
  
  var lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) {
    return [];
  }
  
  var data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
  
  // アクティブアカウントを取得
  var activeAccount = getActiveAccount(ss);
  var activeAccountId = activeAccount ? activeAccount.accountId : 'default';
  
  // アカウント一覧を取得（ユーザー名表示用）
  var accounts = getAccounts(ss);
  var accountMap = {};
  accounts.forEach(function(acc) {
    accountMap[acc.accountId] = acc.username;
  });
  
  var posts = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var id = row[0];
    var accountId = row[1] || 'default';
    var status = String(row[2] || '').trim();
    var text = row[3];
    
    if (!id && !text) {
      continue;
    }
    
    // アカウントフィルタリング
    if (!showAllAccounts && accountId !== activeAccountId) {
      continue;
    }
    
    // posted以外を表示
    if (status === 'posted' || status === '投稿完了') {
      continue;
    }
    
    // ★修正: 日付の安全な変換
    var scheduledTimeStr = '';
    if (row[6]) {
      try {
        var d = new Date(row[6]);
        if (!isNaN(d.getTime())) {
          scheduledTimeStr = d.toISOString();
        }
      } catch (e) {
        console.log('日付変換エラー (行' + (i + 2) + '):', row[6], e.message);
        scheduledTimeStr = '';
      }
    }
    
    var createdAtStr = '';
    if (row[7]) {
      try {
        var d2 = new Date(row[7]);
        if (!isNaN(d2.getTime())) {
          createdAtStr = d2.toISOString();
        }
      } catch (e) {
        createdAtStr = '';
      }
    }
    
    var updatedAtStr = '';
    if (row[8]) {
      try {
        var d3 = new Date(row[8]);
        if (!isNaN(d3.getTime())) {
          updatedAtStr = d3.toISOString();
        }
      } catch (e) {
        updatedAtStr = '';
      }
    }
    
    posts.push({
      id: id || ('row-' + i),
      accountId: accountId,
      accountUsername: accountMap[accountId] || accountId,
      status: status || '予約済み',
      text: text || '',
      mediaUrl: row[4] || '',
      mediaType: row[5] || '',
      scheduledTime: scheduledTimeStr,
      createdAt: createdAtStr,
      updatedAt: updatedAtStr,
      groupId: row[9] || '',
      orderNum: row[10] || '',
      retryCount: row[12] || 0,
      errorMessage: ''
    });
  }
  
  posts.sort(function(a, b) {
    if (!a.scheduledTime) return 1;
    if (!b.scheduledTime) return -1;
    return new Date(a.scheduledTime) - new Date(b.scheduledTime);
  });
  
  return posts;
}

/**
 * 失敗した予約投稿を再試行
 */
function retryScheduledPost(ss, postId) {
  try {
    var sheet = ss.getSheetByName('投稿管理');
    if (!sheet) {
      return { success: false, error: '投稿管理シートが見つかりません' };
    }
    
    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    var postData = null;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === postId) {
        rowIndex = i + 1;
        postData = {
          text: data[i][2],
          mediaUrl: data[i][3],
          mediaType: data[i][4]
        };
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, error: '投稿が見つかりません' };
    }
    
    // ステータスを「再試行中」に更新
    sheet.getRange(rowIndex, 2).setValue('retrying');
    sheet.getRange(rowIndex, 9).setValue(''); // エラーメッセージをクリア
    
    // 投稿を実行
    var result = createPost(ss, postData.text, postData.mediaUrl, postData.mediaType);
    
    if (result && result.success) {
      // 履歴シートに追加
      var historySheet = ss.getSheetByName('履歴');
     
      
      // 予約投稿シートから削除
      sheet.deleteRow(rowIndex);
      
      return { success: true, postId: result.postId };
    } else {
      var errorMsg = result ? result.error : 'Unknown error';
      sheet.getRange(rowIndex, 2).setValue('error');
      sheet.getRange(rowIndex, 9).setValue(errorMsg);
      return { success: false, error: errorMsg };
    }
    
  } catch (error) {
    console.error('retryScheduledPost error:', error);
    return { success: false, error: error.message };
  }
}


function deleteScheduledPost(ss, postId) {
  const sheet = ss.getSheetByName('投稿管理');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === postId) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  
  throw new Error('投稿が見つかりません');
}

function addToHistory(ss, text, mediaUrl, threadsPostId, accountId) {
  var sheet = ss.getSheetByName('履歴');
  var id = Utilities.getUuid();
  var now = new Date();
  
  // accountId が指定されていない場合はアクティブアカウントを使用
  if (!accountId) {
    var activeAccount = getActiveAccount(ss);
    accountId = activeAccount ? activeAccount.accountId : 'default';
  }
  
  // 新しい列順序: id, account_id, text, media_url, posted_at, threads_post_id, likes, replies, fetched_at, group_id, reply_to_id
  sheet.appendRow([
    id,
    accountId,
    text,
    mediaUrl || '',
    now,
    threadsPostId,
    0,
    0,
    now,
    '',  // group_id
    ''   // reply_to_id
  ]);
}


function getHistory(ss, showAllAccounts) {
  try {
    var sheet = ss.getSheetByName('履歴');
    if (!sheet) {
      console.log('履歴シートが見つかりません');
      return [];
    }
    
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      console.log('履歴データなし');
      return [];
    }
    
    var data = sheet.getDataRange().getValues();
    
    // アクティブアカウントを取得
    var activeAccount = getActiveAccount(ss);
    var activeAccountId = activeAccount ? activeAccount.accountId : 'default';
    
    // アカウント一覧を取得（ユーザー名表示用）
    var accounts = getAccounts(ss);
    var accountMap = {};
    accounts.forEach(function(acc) {
      accountMap[acc.accountId] = acc.username;
    });
    
    var history = [];
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) {
        var accountId = data[i][1] || 'default';
        
        // アカウントフィルタリング
        if (!showAllAccounts && accountId !== activeAccountId) {
          continue;
        }
        
        var mediaUrl = data[i][3] ? String(data[i][3]) : '';
        var mediaType = '';
        if (mediaUrl && mediaUrl.match(/\.(mp4|mov|webm)$/i)) {
          mediaType = 'VIDEO';
        } else if (mediaUrl) {
          mediaType = 'IMAGE';
        }
        
        // ★修正: 日付の安全な変換
        var postedAtStr = '';
        if (data[i][4]) {
          try {
            var d = new Date(data[i][4]);
            if (!isNaN(d.getTime())) {
              postedAtStr = d.toISOString();
            }
          } catch (e) {
            postedAtStr = '';
          }
        }
        
        history.push({
          id: String(data[i][0] || ''),
          accountId: accountId,
          accountUsername: accountMap[accountId] || accountId,
          text: String(data[i][2] || ''),
          mediaUrl: mediaUrl,
          mediaType: mediaType,
          postedAt: postedAtStr,
          threadsPostId: String(data[i][5] || ''),
          likes: Number(data[i][6]) || 0,
          replies: Number(data[i][7]) || 0,
          groupId: data[i][9] || '',
          replyToId: data[i][10] || ''
        });
      }
    }
    
    history.sort(function(a, b) {
      if (!a.postedAt) return 1;
      if (!b.postedAt) return -1;
      return new Date(b.postedAt) - new Date(a.postedAt);
    });
    
    console.log('履歴件数:', history.length);
    return history;
    
  } catch (error) {
    console.error('getHistory error:', error);
    return [];
  }
}


// ===========================================
// 分析機能
// ===========================================

/**
 * ユーザーインサイトを取得（過去7日間）
 */
function getUserInsights(ss) {
  var auth = getActiveAccountAuth(ss);
  
  if (!auth || !auth.accessToken) {
    console.log('getUserInsights: 認証情報がありません');
    return { success: false, error: '認証が必要です' };
  }
  
  try {
    // 過去7日間の期間を設定
    var now = new Date();
    var until = Math.floor(now.getTime() / 1000);
    var since = Math.floor((now.getTime() - 7 * 24 * 60 * 60 * 1000) / 1000);
    
    var url = CONFIG.THREADS_API_BASE + '/' + auth.userId + '/threads_insights' +
      '?metric=views,followers_count' +
      '&period=day' +
      '&since=' + since +
      '&until=' + until +
      '&access_token=' + auth.accessToken;
    
    console.log('getUserInsights URL:', url.replace(auth.accessToken, '***'));
    
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var responseText = response.getContentText();
    
    console.log('getUserInsights response:', responseText);
    
    var data = JSON.parse(responseText);
    
    if (data.error) {
      console.error('getUserInsights error:', data.error);
      return { success: false, error: data.error.message };
    }
    
    // レスポンスを整形
    var insights = {
      views: 0,
      followersCount: 0,
      period: {
        since: new Date(since * 1000).toISOString(),
        until: new Date(until * 1000).toISOString()
      }
    };

    if (data.data) {
        data.data.forEach(function(metric) {
    if (metric.name === 'views') {
      if (metric.values && metric.values.length > 0) {
        var total = 0;
        metric.values.forEach(function(v) { total += v.value || 0; });
        insights.views = total;
      } else if (metric.total_value) {
        insights.views = metric.total_value.value || 0;
      }
    }
    if (metric.name === 'followers_count' && metric.total_value) {
      insights.followersCount = metric.total_value.value || 0;
     }
   });
    }
    
    return { success: true, data: insights };
    
  } catch (error) {
    console.error('getUserInsights exception:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 投稿のインサイトを取得
 */
function getPostInsights(ss, threadsPostId) {
  var auth = getActiveAccountAuth(ss);
  
  if (!auth || !auth.accessToken) {
    return { success: false, error: '認証が必要です' };
  }
  
  try {
    var url = CONFIG.THREADS_API_BASE + '/' + threadsPostId + '/insights' +
      '?metric=views,likes,replies,reposts,quotes' +
      '&access_token=' + auth.accessToken;
    
    console.log('getPostInsights URL:', url.replace(auth.accessToken, '***'));
    
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var responseText = response.getContentText();
    
    var data = JSON.parse(responseText);
    
    if (data.error) {
      return { success: false, error: data.error.message };
    }
    
    var insights = {
      views: 0,
      likes: 0,
      replies: 0,
      reposts: 0,
      quotes: 0
    };
    
    if (data.data) {
      data.data.forEach(function(metric) {
        if (metric.values && metric.values[0]) {
          insights[metric.name] = metric.values[0].value || 0;
        }
      });
    }
    
    return { success: true, data: insights };
    
  } catch (error) {
    console.error('getPostInsights exception:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 履歴の投稿すべてのインサイトを更新
 */
function updateAllPostInsights(ss) {
  var historySheet = ss.getSheetByName('履歴');
  if (!historySheet) {
    return { success: false, error: '履歴シートが見つかりません' };
  }
  
  var lastRow = historySheet.getLastRow();
  if (lastRow <= 1) {
    return { success: true, updated: 0 };
  }
  
  var data = historySheet.getRange(2, 1, lastRow - 1, 11).getValues();
  var updatedCount = 0;
  
  for (var i = 0; i < data.length; i++) {
    var threadsPostId = data[i][5]; // F列: threads_post_id
    
    if (threadsPostId) {
      var result = getPostInsights(ss, threadsPostId);
      
      if (result.success) {
        var rowNum = i + 2;
        historySheet.getRange(rowNum, 7).setValue(result.data.likes);    // G列: likes
        historySheet.getRange(rowNum, 8).setValue(result.data.replies);  // H列: replies
        historySheet.getRange(rowNum, 9).setValue(new Date());           // I列: fetched_at
        updatedCount++;
        
        // API制限対策
        Utilities.sleep(500);
      }
    }
  }
  
  console.log('インサイト更新完了:', updatedCount, '件');
  return { success: true, updated: updatedCount };
}

/**
 * 分析データを取得（UI用）
 */
function getInsights(ss) {
  var result = {
    user: null,
    recentPosts: [],
    analyticsData: []
  };
  
  // ユーザーインサイト
  var userInsights = getUserInsights(ss);
  if (userInsights.success) {
    result.user = userInsights.data;
  }
  
  // 分析シートからデータを取得
  var analyticsSheet = ss.getSheetByName('分析');
  if (analyticsSheet) {
    var lastRow = analyticsSheet.getLastRow();
    if (lastRow > 1) {
      var data = analyticsSheet.getRange(2, 1, lastRow - 1, 6).getValues();
      
      for (var i = 0; i < data.length; i++) {
        var row = data[i];
        if (row[0]) {  // post_id がある場合
          result.analyticsData.push({
            postId: row[0],
            date: row[1],
            likes: row[2] || 0,
            replies: row[3] || 0,
            views: row[4] || 0,
            fetchedAt: row[5]
          });
        }
      }
      
      // 日付の新しい順にソート
      result.analyticsData.sort(function(a, b) {
        return new Date(b.date) - new Date(a.date);
      });
    }

  }
  
  // 履歴から最近の投稿を取得
var history = getHistory(ss, false);

// analyticsData に投稿本文を追加
result.analyticsData = result.analyticsData.map(function(analytics) {
  var post = history.find(function(h) {
    return h.threadsPostId === analytics.postId;
  });
  
  if (post) {
    analytics.text = post.text;
  }
  
  return analytics;
});

// recentPosts も設定
result.recentPosts = history.slice(0, 10).map(function(post) {
  var analytics = result.analyticsData.find(function(a) {
    return a.postId === post.threadsPostId;
  });
  
  if (analytics) {
    post.views = analytics.views;
    post.likes = analytics.likes;
    post.replies = analytics.replies;
  }
  
  return post;
});

// ★Phase3-2: 分析シートにデータを蓄積
  try {
    var analyticsSheet = ss.getSheetByName('分析');
    if (analyticsSheet && result.recentPosts && result.recentPosts.length > 0) {
      var existingPostIds = {};
      var existingData = analyticsSheet.getDataRange().getValues();
      for (var ei = 1; ei < existingData.length; ei++) {
        existingPostIds[existingData[ei][0]] = true;
      }
      
      var now = new Date();
      result.recentPosts.forEach(function(post) {
        if (post.threadsPostId && !existingPostIds[post.threadsPostId]) {
          // 新しい投稿のインサイトを取得して保存
          var postInsight = getPostInsights(ss, post.threadsPostId);
          if (postInsight.success) {
            analyticsSheet.appendRow([
              post.threadsPostId,
              post.postedAt || now.toISOString(),
              postInsight.data.likes || 0,
              postInsight.data.replies || 0,
              postInsight.data.views || 0,
              now.toISOString()
            ]);
            
            // result にも反映
            post.views = postInsight.data.views || 0;
            post.likes = postInsight.data.likes || 0;
            post.replies = postInsight.data.replies || 0;
          }
          Utilities.sleep(300);
        }
      });
    }
  } catch (analyticsError) {
    console.log('分析データ蓄積エラー:', analyticsError.message);
  }
return result;
}
// ===========================================
// Phase 3-1: トークン自動更新
// ===========================================

// processScheduledPosts 定義確認メモ:
// - L2448: バウンドスプレッドシート優先 + トークン更新 + 予約処理（採用）
// - L2530: ACTIVE_SHEET_IDのみ参照の旧定義（削除）
/**
 * 予約投稿を処理（トリガーで1分ごとに実行）
 */
function processScheduledPosts() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    console.log('別のプロセスが実行中。スキップします。');
    return;
  }
  
  try {
    console.log('=== processScheduledPosts 開始 ===');
    
    // バウンドスプレッドシートを優先
    var sheetId = null;
    try {
      sheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
    } catch (e) {}
    
    if (!sheetId) {
      var props = PropertiesService.getScriptProperties();
      sheetId = props.getProperty('ACTIVE_SHEET_ID');
    }
    
    if (!sheetId) {
      console.log('シートIDが未設定。処理をスキップします。');
      return;
    }
    
    console.log('処理対象シートID:', sheetId);
    var ss = SpreadsheetApp.openById(sheetId);
    var now = new Date();
    
    // トークン自動更新
    try {
      refreshExpiringTokens(ss);
    } catch (tokenErr) {
      console.log('トークン更新スキップ:', tokenErr.message);
    }
    
    // 予約投稿を処理
    processSheetScheduledPosts(sheetId, now);
    
    console.log('=== processScheduledPosts 完了 ===');
    
  } catch (e) {
    console.error('processScheduledPosts エラー:', e.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * トークン期限警告データを取得（フロント表示用）
 * 改善版: 残日数と状態を返す
 */
function getTokenWarnings(ss) {
  var accounts = getAccounts(ss);
  var warnings = [];
  var now = new Date();
  
  for (var i = 0; i < accounts.length; i++) {
    var account = accounts[i];
    if (!account.tokenExpires) continue;
    
    var expiryDate = new Date(account.tokenExpires);
    var daysLeft = Math.ceil((expiryDate.getTime() - now.getTime()) / (1000 * 60 * 60 * 24));
    
    if (daysLeft <= 0) {
      warnings.push({
        accountId: account.accountId,
        username: account.username,
        daysLeft: 0,
        status: 'expired',
        message: '@' + account.username + ' のトークンが期限切れです。再認証してください。'
      });
    } else if (daysLeft <= 5) {
      warnings.push({
        accountId: account.accountId,
        username: account.username,
        daysLeft: daysLeft,
        status: 'critical',
        message: '@' + account.username + ' のトークンがあと ' + daysLeft + ' 日で期限切れです。'
      });
    } else if (daysLeft <= 10) {
      warnings.push({
        accountId: account.accountId,
        username: account.username,
        daysLeft: daysLeft,
        status: 'warning',
        message: '@' + account.username + ' のトークンがあと ' + daysLeft + ' 日で期限切れです（自動更新予定）。'
      });
    }
  }
  
  return warnings;
}

// ===========================================
// 予約投稿の自動実行
// ===========================================

/**
 * 長期トークンをリフレッシュ（有効期限が10日以内のアカウント）
 */
function refreshExpiringTokens(ss) {
  try {
    var accounts = getAccounts(ss);
    var settings = getSettings(ss);
    var appSecret = settings.app_secret;
    
    if (!appSecret) {
      console.log('app_secret が未設定。トークン更新スキップ');
      return;
    }
    
    var now = new Date();
    var REFRESH_THRESHOLD_DAYS = 10;
    
    var accountSheet = ss.getSheetByName('アカウント');
    if (!accountSheet || accountSheet.getLastRow() <= 1) return;
    
    var accountData = accountSheet.getRange(2, 1, accountSheet.getLastRow() - 1, 7).getValues();
    
    for (var i = 0; i < accountData.length; i++) {
      var row = accountData[i];
      var accountId = row[0];
      var accessToken = row[1];
      var tokenExpires = row[5];
      
      if (!accessToken || !tokenExpires) continue;
      
      var expiryDate = new Date(tokenExpires);
      if (isNaN(expiryDate.getTime())) continue;
      
      var daysLeft = Math.ceil((expiryDate.getTime() - now.getTime()) / (1000 * 60 * 60 * 24));
      
      if (daysLeft <= REFRESH_THRESHOLD_DAYS && daysLeft > 0) {
        console.log('トークンをリフレッシュします: ' + accountId + ' (残り' + daysLeft + '日)');
        
        try {
          var refreshUrl = 'https://graph.threads.net/refresh_access_token' +
            '?grant_type=th_refresh_token' +
            '&access_token=' + accessToken;
          
          var response = UrlFetchApp.fetch(refreshUrl, { muteHttpExceptions: true });
          var data = JSON.parse(response.getContentText());
          
          if (data.access_token && data.expires_in) {
            var newExpiresAt = new Date(Date.now() + data.expires_in * 1000).toISOString();
            var rowNum = i + 2;
            accountSheet.getRange(rowNum, 2).setValue(data.access_token);
            accountSheet.getRange(rowNum, 6).setValue(newExpiresAt);
            
            var activeAccount = getActiveAccount(ss);
            if (activeAccount && activeAccount.accountId === accountId) {
              saveSettings(ss, {
                access_token: data.access_token,
                token_expires: newExpiresAt
              });
            }
            
            console.log('トークン更新成功: ' + accountId);
          } else if (data.error) {
            console.error('トークン更新失敗: ' + data.error.message);
          }
        } catch (refreshError) {
          console.error('リフレッシュエラー:', refreshError.message);
        }
        
        Utilities.sleep(1000);
      }
    }
  } catch (e) {
    console.error('refreshExpiringTokens エラー:', e.message);
  }
}


/**
 * 登録済みシートID一覧を取得
 */
function getRegisteredSheetIds() {
  var props = PropertiesService.getScriptProperties();
  var idsJson = props.getProperty('REGISTERED_SHEET_IDS');
  
  if (!idsJson) {
    return [];
  }
  
  try {
    return JSON.parse(idsJson);
  } catch (e) {
    return [];
  }
}

/**
 * スプレッドシートIDを登録リストに追加
 */
function registerSheetId(sheetId) {
  var props = PropertiesService.getScriptProperties();
  var sheetIds = getRegisteredSheetIds();
  
  if (sheetIds.indexOf(sheetId) === -1) {
    sheetIds.push(sheetId);
    props.setProperty('REGISTERED_SHEET_IDS', JSON.stringify(sheetIds));
    console.log('シートIDを登録しました:', sheetId);
  }
  
  return sheetIds;
}

/**
 * 予約投稿を処理（個別シート）
 */
function processSheetScheduledPosts(sheetId, now) {
  var ss = SpreadsheetApp.openById(sheetId);
  if (!ss) {
    console.log('スプレッドシートを開けません:', sheetId);
    return;
  }
  
  var settings = getSettings(ss);
  if (!settings.access_token) {
    console.log('認証されていません:', sheetId);
    return;
  }
  
  console.log('認証済みユーザー:', settings.username || settings.user_id);
  
  var sheet = ss.getSheetByName('投稿管理');
  if (!sheet) {
    console.log('投稿管理シートが見つかりません:', sheetId);
    return;
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    console.log('予約投稿なし:', sheetId);
    return;
  }
  
  var data = sheet.getDataRange().getValues();
  var postedRows = [];
  var postedCount = 0;
  var processedGroups = {};
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var postId = row[0];
    var accountId = row[1];
    var status = String(row[2]).trim();
    var text = row[3];
    var mediaUrl = row[4];
    var mediaType = row[5];
    var scheduledTime = row[6];
    var groupId = row[9];
    
    if (!postId && !text) {
      continue;
    }
    
    var skipStatuses = ['posted', '投稿完了', 'processing', 'error', 'エラー', 'expired'];
    if (skipStatuses.indexOf(status) !== -1) {
      continue;
    }
    
    var validStatuses = ['scheduled', '予約済み', 'retrying', '再試行中'];
    if (validStatuses.indexOf(status) === -1) {
      console.log('不明なステータスをスキップ:', status, 'postId:', postId);
      continue;
    }
    
    if (!postId) {
      postId = Utilities.getUuid();
      sheet.getRange(i + 1, 1).setValue(postId);
    }
    
    if (!scheduledTime) {
      console.log('予約時刻が未設定。スキップ:', postId);
      continue;
    }
    
    var scheduled = new Date(scheduledTime);
    
    if (isNaN(scheduled.getTime())) {
      console.log('無効な予約時刻。スキップ:', postId, 'scheduledTime:', scheduledTime);
      continue;
    }
    
    if (scheduled <= now) {
      // 予約時刻から2時間以上経過したらスキップ
      var MAX_DELAY_MS = 2 * 60 * 60 * 1000;
      if (now.getTime() - scheduled.getTime() > MAX_DELAY_MS) {
        console.log('予約時刻から2時間以上経過。スキップ:', postId);
        sheet.getRange(i + 1, 3).setValue('expired');
        continue;
      }
      
      // ツリー投稿
      if (groupId && !processedGroups[groupId]) {
        console.log('ツリー投稿を処理:', groupId);
        processedGroups[groupId] = true;
        
        if (accountId) {
          try { setActiveAccount(ss, accountId); } catch (accErr) {
            console.log('アカウント切り替えエラー:', accErr.message);
          }
        }
        
        var treeResult = processTreePosts(ss, groupId, now);
        if (treeResult.success) {
          postedCount += treeResult.postedCount;
        }
        
        data = sheet.getDataRange().getValues();
        i = 0;
        continue;
      }
      
      // 単発投稿
      if (!groupId) {
        console.log('投稿実行:', postId, 'テキスト:', text.substring(0, 30), 'アカウント:', accountId);
        
        if (accountId) {
          try { setActiveAccount(ss, accountId); } catch (accErr) {
            console.log('アカウント切り替えエラー:', accErr.message);
          }
        }
        
        sheet.getRange(i + 1, 3).setValue('processing');
        SpreadsheetApp.flush();
        
        try {
          // ★ createPost内部でaddToHistoryが呼ばれるので、ここでは履歴追加しない
          var result = createPost(ss, text, mediaUrl, mediaType);
          
          if (result && result.success) {
            console.log('投稿成功', result.postId);
            sheet.getRange(i + 1, 3).setValue('posted');
            postedRows.push(i + 1);
            postedCount++;
          } else {
            var err = result ? result.error : 'Unknown error';
            console.error('投稿失敗', err);
            sheet.getRange(i + 1, 3).setValue('error');
            sheet.getRange(i + 1, 9).setValue(err);
          }
          
        } catch (postError) {
          console.error('投稿エラー:', postError.message);
          var currentRetry = Number(sheet.getRange(i + 1, 13).getValue()) || 0;
          var MAX_RETRIES = 3;
          
          if (currentRetry < MAX_RETRIES) {
            var retryTime = new Date(now.getTime() + 5 * 60 * 1000);
            sheet.getRange(i + 1, 3).setValue('scheduled');
            sheet.getRange(i + 1, 7).setValue(retryTime);
            sheet.getRange(i + 1, 13).setValue(currentRetry + 1);
            console.log('リトライ予定 (' + (currentRetry + 1) + '/' + MAX_RETRIES + '): ' + postId);
          } else {
            sheet.getRange(i + 1, 3).setValue('error');
            sheet.getRange(i + 1, 9).setValue('リトライ上限(' + MAX_RETRIES + '回)超過: ' + postError.message);
            console.error('リトライ上限超過:', postId);
          }
        }
        
        Utilities.sleep(3000);
      }
    }
  }
  
  // 投稿済み行を削除（下から）
  postedRows.sort(function(a, b) { return b - a; });
  for (var j = 0; j < postedRows.length; j++) {
    sheet.deleteRow(postedRows[j]);
  }
  
  console.log('シート処理完了:', sheetId, postedCount, '件投稿');
}

function checkScheduledPosts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('投稿管理');
  const settings = getSettings(ss);
  
  if (!settings.access_token) {
    console.log('認証が必要です');
    return;
  }
  
  const now = new Date();
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const status = data[i][1];
    const scheduledTime = new Date(data[i][5]);
    
    if (status === '予約済み' && scheduledTime <= now) {
      const text = data[i][2];
      const mediaUrl = data[i][3];
      const mediaType = data[i][4];
      
      try {
        const result = createPost(ss, text, mediaUrl, mediaType);
        
        sheet.getRange(i + 1, 2).setValue('投稿完了');
        sheet.getRange(i + 1, 8).setValue(new Date());
        
        console.log(`投稿完了: ${text.substring(0, 30)}...`);
        
      } catch (error) {
        sheet.getRange(i + 1, 2).setValue('エラー');
        console.error(`投稿エラー: ${error.message}`);
      }
    }
  }
}

function updateScheduledPost(ss, postId, text, mediaUrl, mediaType, scheduledTime) {
  try {
    var sheet = ss.getSheetByName('投稿管理');
    if (!sheet) {
      return { success: false, error: '投稿管理シートが見つかりません' };
    }
    
    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === postId) {
        rowIndex = i + 1; // シートの行番号（1始まり）
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, error: '投稿が見つかりません' };
    }
    
    // 更新（日付はDateオブジェクトに変換）
      sheet.getRange(rowIndex, 4).setValue(text);              // D列
      sheet.getRange(rowIndex, 5).setValue(mediaUrl || '');    // E列
      sheet.getRange(rowIndex, 6).setValue(mediaType || '');   // F列
      sheet.getRange(rowIndex, 7).setValue(new Date(scheduledTime)); // G列
      sheet.getRange(rowIndex, 9).setValue(new Date());        // I列
    var currentStatus = sheet.getRange(rowIndex, 3).getValue(); // C列
    if (currentStatus === 'error') {
      sheet.getRange(rowIndex, 3).setValue('scheduled');     // C列
      sheet.getRange(rowIndex, 10).setValue('');             // エラーメッセージ列
    }
    
    console.log('Updated post:', postId);
    
    return { success: true };
    
  } catch (error) {
    console.error('updateScheduledPost error:', error);
    return { success: false, error: error.message };
  }
}



// ===========================================
// APIリクエスト処理
// ===========================================

function processApiRequest(params) {
  var action = params.action;
  var sheetId = params.sheetId;
  
  console.log('=== processApiRequest ===');
  console.log('params:', JSON.stringify(params));
  console.log('action:', action);
  console.log('sheetId:', sheetId);
  
  // ACTIVE_SHEET_ID を更新（予約投稿トリガー用）
  if (sheetId) {
    var props = PropertiesService.getScriptProperties();
    props.setProperty('ACTIVE_SHEET_ID', sheetId);
  }
  
  try {
    // createSpreadsheetはsheetId不要
    if (action === 'createSpreadsheet') {
      return { success: true, data: createNewSpreadsheet() };
    }
    
    if (!sheetId) {
      console.log('ERROR: sheetId is empty');
      return { success: false, error: 'シートIDが指定されていません' };
    }
    
    let ss = null;
    try {
      ss = SpreadsheetApp.openById(sheetId);
      console.log('Spreadsheet opened:', ss.getName());
    } catch (e) {
      console.log('ERROR: Cannot open spreadsheet:', e.message);
      return { success: false, error: 'スプレッドシートを開けません: ' + e.message };
    }
    
    let result;
    switch (action) {
      case 'getTokenWarnings':
        result = getTokenWarnings(ss);
        break;
      case 'refreshToken':
        result = refreshExpiringTokens(ss);
        break;
      case 'getSettings':
        result = getSettings(ss);
        break;
      case 'saveSettings':
        result = saveSettings(ss, params);
        break;
      case 'getAuthUrl':
        result = getAuthUrl(ss);
        break;
      case 'exchangeToken':
        result = exchangeToken(ss, params.code);
        break;
      case 'getUserProfile':
        result = getUserProfile(ss);
        break;
      case 'createPost':
        result = createPost(ss, params.text, params.mediaUrl, params.mediaType);
        break;
      case 'schedulePost':
        result = schedulePost(ss, params.text, params.mediaUrl, params.mediaType, params.scheduledTime);
        break;
      case 'getScheduledPosts':
        result = getScheduledPosts(ss, params.showAllAccounts);
        break;
      case 'deletePost':
        result = deleteScheduledPost(ss, params.postId);
        break;
      case 'getHistory':
        result = getHistory(ss, params.showAllAccounts);
        break;
      case 'getInsights':
        console.log('=== getInsights 開始 ===');
        try {
          var rawResult = getInsights(ss);
        result = JSON.parse(JSON.stringify(rawResult));
          console.log('=== getInsights 成功 ===');
        } catch (e) {
          console.log('=== getInsights エラー ===', e.message);
        result = { user: { views: 0, followersCount: 0 }, recentPosts: [], analyticsData: [] };
        }
        break;
      case 'getUserInsights':
        result = getUserInsights(ss);
        break;
      case 'getPostInsights':
        result = getPostInsights(ss, params.postId);
        break;
      case 'updateAllPostInsights':
        result = updateAllPostInsights(ss);
        break;
      case 'validateSheetId':
        result = validateSheetId(sheetId);
        break;
      case 'updateScheduledPost':
        result = updateScheduledPost(ss, params.postId, params.text, params.mediaUrl, params.mediaType, params.scheduledTime);
        break;
      case 'uploadImage':
        result = uploadImage(ss, params.base64Data, params.fileName);
        break;
      case 'uploadVideo':
        result = uploadVideo(ss, params.base64Data, params.fileName);
        break;
      case 'getImageFolderInfo':
        result = getImageFolderInfo(ss);
        break;
      case 'retryScheduledPost':
        result = retryScheduledPost(ss, params.postId);
        break;
      case 'createTreePost':
        result = createTreePost(params.posts, sheetId);
        break;
      case 'scheduleTreePost':
        result = scheduleTreePost(params.posts, params.scheduledTime, sheetId);
        break;
      case 'getAccounts':
        result = getAccounts(ss);
        break;
      case 'getActiveAccount':
        result = getActiveAccount(ss);
        break;
      case 'setActiveAccount':
        result = setActiveAccount(ss, params.accountId);
        break;
      case 'addAccount':
        result = addAccount(ss, params.accountData);
        break;
      case 'removeAccount':
        result = removeAccount(ss, params.accountId);
        break;
      case 'importAccountFromSheet':
        result = importAccountFromSheet(ss, params.sourceSheetId);
        break;
      case 'checkTokenExpiry':
        result = checkTokenExpiry(ss);
        break;
      default:
        return { success: false, error: 'Unknown action: ' + action };
    }
    
    console.log('result:', JSON.stringify(result));
    return { success: true, data: result };
    
  } catch (error) {
    console.error('API Error:', error);
    return { success: false, error: error.message || 'Unknown error' };
  }
}

// ===========================================
// 複数アカウント管理
// ===========================================

/**
 * アカウント一覧を取得
 */
function getAccounts(ss) {
  var sheet = ss.getSheetByName('アカウント');
  if (!sheet) {
    console.log('アカウントシートが見つかりません');
    return [];
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return [];
  }
  
  var data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  var accounts = [];
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (row[0]) {  // account_id が存在する場合
      accounts.push({
        accountId: row[0],
        accessToken: row[1],
        userId: row[2],
        username: row[3],
        profilePicUrl: row[4],
        tokenExpires: row[5],
        createdAt: row[6]
      });
    }
  }
  
  console.log('アカウント数:', accounts.length);
  return accounts;
}

/**
 * アクティブアカウントを取得
 */
function getActiveAccount(ss) {
  var settings = getSettings(ss);
  var activeAccountId = settings.active_account;
  
  // アクティブアカウントが未設定の場合
  if (!activeAccountId) {
    // アカウント一覧から最初のアカウントを取得
    var accounts = getAccounts(ss);
    if (accounts.length > 0) {
      // 最初のアカウントをアクティブに設定
      setActiveAccount(ss, accounts[0].accountId);
      return accounts[0];
    }
    
    // アカウントシートにデータがない場合、設定シートから取得（後方互換性）
    if (settings.access_token && settings.user_id) {
      return {
        accountId: 'default',
        accessToken: settings.access_token,
        userId: settings.user_id,
        username: settings.username || '',
        profilePicUrl: settings.profile_pic_url || '',
        tokenExpires: settings.token_expires || ''
      };
    }
    
    return null;
  }
  
  // アクティブアカウントIDでアカウントを検索
  var accounts = getAccounts(ss);
  for (var i = 0; i < accounts.length; i++) {
    if (accounts[i].accountId === activeAccountId) {
      return accounts[i];
    }
  }
  
  // 見つからない場合、最初のアカウントを返す
  if (accounts.length > 0) {
    setActiveAccount(ss, accounts[0].accountId);
    return accounts[0];
  }
  
  return null;
}

/**
 * アクティブアカウントを切り替え
 */
function setActiveAccount(ss, accountId) {
  console.log('アクティブアカウントを設定:', accountId);
  
  saveSettings(ss, { active_account: accountId });
  
  return { success: true, accountId: accountId };
}

// 別のスプレッドシートからアカウントをインポート
function importAccountFromSheet(ss, sourceSheetId) {
  try {
    // ソースのスプレッドシートを開く
    var sourceSs;
    try {
      sourceSs = SpreadsheetApp.openById(sourceSheetId);
    } catch (e) {
      return { success: false, error: 'スプレッドシートを開けません。IDを確認してください。' };
    }
    
    // ソースの設定シートを読み取る
    var sourceSettingsSheet = sourceSs.getSheetByName('設定');
    var sourceAccountsSheet = sourceSs.getSheetByName('アカウント');
    
    if (!sourceSettingsSheet) {
      return { success: false, error: '設定シートが見つかりません。正しいスプシIDか確認してください。' };
    }
    
    // ソースの設定を取得
    var sourceSettings = {};
    var sourceData = sourceSettingsSheet.getDataRange().getValues();
    for (var i = 0; i < sourceData.length; i++) {
      if (sourceData[i][0]) {
        sourceSettings[sourceData[i][0]] = sourceData[i][1];
      }
    }
    
    // 必要な情報があるかチェック
    if (!sourceSettings.access_token || !sourceSettings.user_id) {
      return { success: false, error: 'インポート元のスプシで認証が完了していません。' };
    }
    
    // 現在のスプシのアカウントシートを取得
    var accountsSheet = ss.getSheetByName('アカウント');
    if (!accountsSheet) {
      return { success: false, error: 'アカウントシートが見つかりません。' };
    }
    
    // 既存のアカウントをチェック（重複防止）
    var existingData = accountsSheet.getDataRange().getValues();
    for (var i = 1; i < existingData.length; i++) {
      if (existingData[i][1] === sourceSettings.user_id) {
        return { success: false, error: 'このアカウントは既に追加されています。' };
      }
    }
    
    // 新しいアカウントIDを生成
    var newAccountId = 'account_' + Date.now();
    
    // アカウントを追加
    var newRow = [
      newAccountId,
      sourceSettings.access_token,
      sourceSettings.user_id,
      sourceSettings.username || '',
      sourceSettings.profile_pic_url || '',
      sourceSettings.token_expires || '',
      new Date().toISOString()
    ];
    
    accountsSheet.appendRow(newRow);
    
    return { 
      success: true, 
      data: {
        accountId: newAccountId,
        username: sourceSettings.username || sourceSettings.user_id,
        message: '@' + (sourceSettings.username || sourceSettings.user_id) + ' をインポートしました'
      }
    };
    
  } catch (e) {
    console.error('importAccountFromSheet error:', e);
    return { success: false, error: 'インポートエラー: ' + e.message };
  }
}

// トークン期限をチェック
function checkTokenExpiry(ss) {
  var accounts = getAccounts(ss);
  var warnings = [];
  var now = new Date();
  var warningDays = 5;
  
  accounts.forEach(function(account) {
    if (account.tokenExpires) {
      var expiryDate = new Date(account.tokenExpires);
      var daysLeft = Math.ceil((expiryDate - now) / (1000 * 60 * 60 * 24));
      
      if (daysLeft <= warningDays && daysLeft > 0) {
        warnings.push({
          accountId: account.accountId,
          username: account.username,
          daysLeft: daysLeft
        });
      } else if (daysLeft <= 0) {
        warnings.push({
          accountId: account.accountId,
          username: account.username,
          daysLeft: 0,
          expired: true
        });
      }
    }
  });
  
  return { success: true, data: warnings };
}

/**
 * 新規アカウントを追加
 */
function addAccount(ss, accountData) {
  var sheet = ss.getSheetByName('アカウント');
  if (!sheet) {
    throw new Error('アカウントシートが見つかりません');
  }
  
  // 既存アカウントをチェック（同じuser_idがあれば更新）
  var accounts = getAccounts(ss);
  var existingIndex = -1;
  
  for (var i = 0; i < accounts.length; i++) {
    if (accounts[i].userId === accountData.userId) {
      existingIndex = i;
      break;
    }
  }
  
  var accountId = accountData.accountId || 'acc-' + accountData.userId;
  var now = new Date().toISOString();
  
  var rowData = [
    accountId,
    accountData.accessToken,
    accountData.userId,
    accountData.username || '',
    accountData.profilePicUrl || '',
    accountData.tokenExpires || '',
    now
  ];
  
  if (existingIndex >= 0) {
    // 既存アカウントを更新（ヘッダー行 + インデックス + 1）
    var rowNum = existingIndex + 2;
    sheet.getRange(rowNum, 1, 1, 7).setValues([rowData]);
    console.log('アカウントを更新しました:', accountId);
  } else {
    // 新規追加
    sheet.appendRow(rowData);
    console.log('アカウントを追加しました:', accountId);
  }
  
  // 追加したアカウントをアクティブに設定
  setActiveAccount(ss, accountId);
  
  return { 
    success: true, 
    accountId: accountId,
    isNew: existingIndex < 0
  };
}

/**
 * アカウントを削除
 */
function removeAccount(ss, accountId) {
  var sheet = ss.getSheetByName('アカウント');
  if (!sheet) {
    throw new Error('アカウントシートが見つかりません');
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    throw new Error('削除するアカウントがありません');
  }
  
  var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var rowToDelete = -1;
  
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === accountId) {
      rowToDelete = i + 2;  // ヘッダー行 + インデックス + 1
      break;
    }
  }
  
  if (rowToDelete === -1) {
    throw new Error('アカウントが見つかりません: ' + accountId);
  }
  
  sheet.deleteRow(rowToDelete);
  console.log('アカウントを削除しました:', accountId);
  
  // 削除後、別のアカウントをアクティブに設定
  var remainingAccounts = getAccounts(ss);
  if (remainingAccounts.length > 0) {
    setActiveAccount(ss, remainingAccounts[0].accountId);
  } else {
    saveSettings(ss, { active_account: '' });
  }
  
  return { success: true };
}

/**
 * アクティブアカウントの認証情報を取得（内部用）
 */
function getActiveAccountAuth(ss) {
  var account = getActiveAccount(ss);
  
  if (!account) {
    // 後方互換性：設定シートから取得
    var settings = getSettings(ss);
    if (settings.access_token) {
      return {
        accessToken: settings.access_token,
        userId: settings.user_id
      };
    }
    return null;
  }
  
  return {
    accessToken: account.accessToken,
    userId: account.userId
  };
}
