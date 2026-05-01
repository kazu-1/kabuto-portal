/**
 * Webページを表示する
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('飼育管理ポータル')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 室温データの取得（1時間おきに間引く版）
 */
function getSheetData(startDateStr, endDateStr) {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const sheetId = scriptProperties.getProperty('SHEET_ID_4');
    const ss = sheetId ? SpreadsheetApp.openById(sheetId) : SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('室温管理');
    
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    let start = startDateStr ? new Date(startDateStr + " 00:00:00").getTime() : null;
    let end = endDateStr ? new Date(endDateStr + " 23:59:59").getTime() : null;

    const results = [];
    const processedHours = new Set(); // すでに取得した「日+時間」を記録する用

    // データの最後（最新）から遡ってチェック
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      if (!(row[0] instanceof Date)) continue;
      
      const rowDate = row[0];
      const rowTime = rowDate.getTime();
      
      // 期間フィルター
      if (start && rowTime < start) continue;
      if (end && rowTime > end) continue;

      // 「2026-01-13-07」のような一意のキーを作る
      const hourKey = rowDate.getFullYear() + "-" + (rowDate.getMonth()+1) + "-" + rowDate.getDate() + "-" + rowDate.getHours();
      
      // その時間のデータがまだ未登録なら採用（最新の10分データを優先）
      if (!processedHours.has(hourKey)) {
        results.push({
          timestamp: rowTime,
          temp: parseFloat(row[1]) || 0,
          humi: parseFloat(row[2]) || 0
        });
        processedHours.add(hourKey);
      }
    }
    
    // グラフ表示のために古い順に並び替えて返す
    return results.sort((a, b) => a.timestamp - b.timestamp);
  } catch (e) {
    console.error(e.message);
    return [];
  }
}

/**
 * A列から特定のテキスト（前方一致）を含む行番号を返す補助関数
 */
function findRowByText(sheet, text) {
  const data = sheet.getRange("A:A").getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] && String(data[i][0]).indexOf(text) === 0) {
      return i + 1; // 行番号（1始まり）を返す
    }
  }
  return null;
}

/**
 * 指定した開始行から、A列が空欄になるまでの連続したデータを取得する補助関数
 */
function getDataUntilBlank(sheet, startRow, numColumns) {
  const lastRow = sheet.getLastRow();
  if (startRow > lastRow) return [];
  
  // 開始行から最終行までのA列を取得して、空欄を探す
  const columnA = sheet.getRange(startRow, 1, lastRow - startRow + 1, 1).getValues();
  let rowCount = 0;
  for (let i = 0; i < columnA.length; i++) {
    if (columnA[i][0] === "") break; // 空白行で終了
    rowCount++;
  }
  
  if (rowCount === 0) return [];
  return sheet.getRange(startRow, 1, rowCount, numColumns).getValues();
}

/**
 * A. 採卵・孵化月毎の集計 データを取得
 */
function getEggSummaryData() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const sheetId = scriptProperties.getProperty('SHEET_ID_4');
    const ss = sheetId ? SpreadsheetApp.openById(sheetId) : SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('卵の集計');
    
    const startRowHeader = findRowByText(sheet, "A.");
    if (!startRowHeader) return [];

    // 見出しの2行下から、A列が空白になるまで取得（A列〜E列の5列分）
    const values = getDataUntilBlank(sheet, startRowHeader + 2, 5);
    
    return values.map(row => {
      return {
        label: row[0] + "年" + row[1] + "月",
        eggCount: parseFloat(row[2]) || 0,
        hatchCount: parseFloat(row[3]) || 0,
        larvaCount: parseFloat(row[4]) || 0
      };
    });
  } catch (e) {
    console.error(e.message);
    return [];
  }
}

/**
 * C. カブトムシごとの集計 データを取得
 */
function getBeetleSummaryData() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const sheetId = scriptProperties.getProperty('SHEET_ID_4');
    const ss = sheetId ? SpreadsheetApp.openById(sheetId) : SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('卵の集計');
    
    const startRowHeader = findRowByText(sheet, "C.");
    if (!startRowHeader) return [];

    // 見出しの2行下から、A列が空白になるまで取得（A列〜K列の11列分）
    const values = getDataUntilBlank(sheet, startRowHeader + 2, 14);
    
    return values.map(row => {
      return {
        year: row[0],
        month: row[1],
        label: row[0] + "年" + row[1] + "月",
        caseId: row[2],
        fullCaseName: row[0] + "年" + row[1] + "月-" + row[2],
        total: row[11],     // L列
        hatched: row[12],   // M列
        rate: row[13]      // N列
      };
    });
  } catch (e) {
    console.error(e.message);
    return [];
  }
}