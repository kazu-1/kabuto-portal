/**
 * 「室温管理」シートにデータを追記する関数
 */
/**
 * SORACOM APIトークンの管理と差分データ取得を行う最適化版
 */
function getTempMinamiise() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const simId = scriptProperties.getProperty('SIM_ID_4');
  const sheetId = scriptProperties.getProperty('SHEET_ID_4');
  const ss = sheetId ? SpreadsheetApp.openById(sheetId) : SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('室温管理');

  if (!sheet) {
    Logger.log('「室温管理」シートが見つかりません。');
    return;
  }

  // --- 1. 取得開始時刻（from）の決定 ---
  let from;
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    // A列（timestamp）の最終行から時刻を取得し、1秒(1000ms)加算して重複を回避
    const lastTimestamp = sheet.getRange(lastRow, 1).getValue();
    from = new Date(lastTimestamp).getTime() + 1000;
  } else {
    // 初回やシートが空の場合は1時間前から
    from = new Date().getTime() - (60 * 60 * 1000);
  }
  const to = new Date().getTime();

  // --- 2. 認証トークンの取得（再利用ロジック） ---
  let apiKey = scriptProperties.getProperty('SORACOM_API_KEY');
  let token = scriptProperties.getProperty('SORACOM_TOKEN');
  
  // トークンがない、または古い（今回は簡易的にリクエストしてエラーなら再取得する手法を採用）
  if (!apiKey || !token) {
    const authData = refreshSoracomToken(scriptProperties);
    apiKey = authData.apiKey;
    token = authData.token;
  }

  // --- 3. データ取得 ---
  let allData = [];
  try {
    allData = fetchSoracomData(simId, apiKey, token, from, to);
  } catch (e) {
    // トークン切れの可能性があるため、一度だけ再認証してリトライ
    if (e.message.indexOf('401') !== -1) {
      const authData = refreshSoracomToken(scriptProperties);
      allData = fetchSoracomData(simId, authData.apiKey, authData.token, from, to);
    } else {
      throw e;
    }
  }

  if (allData.length === 0) {
    Logger.log('新規データはありませんでした。');
    return;
  }

  // --- 4. データ整形 ---
  const rowsToAppend = allData.map(item => {
    try {
      const timestamp = new Date(item.time);
      const contentObj = JSON.parse(item.content);
      const decoded = Utilities.newBlob(Utilities.base64Decode(contentObj.payload)).getDataAsString('utf-8');
      const p = JSON.parse(decoded);
      return [timestamp, p.temperature || '', p.humidity || ''];
    } catch (e) {
      return null;
    }
  }).filter(row => row !== null);

  // --- 5. 追記（古い順に並び替え） ---
  rowsToAppend.reverse();
  if (rowsToAppend.length > 0) {
    sheet.getRange(lastRow + 1, 1, rowsToAppend.length, 3).setValues(rowsToAppend);
    Logger.log(`${rowsToAppend.length} 件の新規データを追記しました。`);
  }
}

/**
 * SORACOM認証を実行しプロパティを更新する補助関数
 */
function refreshSoracomToken(props) {
  Logger.log('SORACOMトークンを更新します...');
  const email = props.getProperty('SORACOM_EMAIL_MINAMIISE');
  const password = props.getProperty('SORACOM_PASSWORD_MINAMIISE');
  
  const response = UrlFetchApp.fetch('https://g.api.soracom.io/v1/auth', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ email, password })
  });
  
  const data = JSON.parse(response.getContentText());
  props.setProperty('SORACOM_API_KEY', data.apiKey);
  props.setProperty('SORACOM_TOKEN', data.token);
  return data;
}

/**
 * データ取得リクエストを行う補助関数
 */
function fetchSoracomData(simId, apiKey, token, from, to) {
  let results = [];
  let lastKey = null;
  while (true) {
    let url = `https://api.soracom.io/v1/sims/${simId}/data?from=${from}&to=${to}&sort=desc&limit=1000`;
    if (lastKey) url += `&last_evaluated_key=${encodeURIComponent(JSON.stringify(lastKey))}`;
    
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { 'X-Soracom-API-Key': apiKey, 'X-Soracom-Token': token },
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() !== 200) {
      throw new Error('API Error: ' + response.getResponseCode());
    }

    const data = JSON.parse(response.getContentText());
    results = results.concat(data);
    if (!data.lastEvaluatedKey) break;
    lastKey = data.lastEvaluatedKey;
  }
  return results;
}

/**
 * 1時間おきに実行するトリガーを設定する関数
 * 初回に一度だけ手動で実行してください
 */
function create10MinTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  // 既存の同名トリガーがあれば削除（二重登録防止）
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'getTempMinamiise') ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger('getTempMinamiise')
    .timeBased()
    // .everyHours(1)
    .everyMinutes(10)
    .create();
    
  Logger.log('10分ごとごとのトリガーを設定しました。');
}