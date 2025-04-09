
function listSheetIndex() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const baseUrl = ss.getUrl();
  
  // indexシートを取得、なければ作成
  let sheetToWrite = ss.getSheetByName("index");
  if (!sheetToWrite) {
    sheetToWrite = ss.insertSheet("index");
  } else {
    // 既存データを消す（No.、シート名、行数、列数のみ）
    const lastRow = sheetToWrite.getLastRow();
    
    // A〜D列（No.、シート名、行数、列数）だけをクリア
    if (lastRow > 2) { // 2行目以降にデータがある場合
      sheetToWrite.getRange(3, 1, lastRow - 2, 4).clearContent(); // A〜D列だけクリア
    }
  }

  const sheets = ss.getSheets();
  
  // A1にタイトル
  sheetToWrite.getRange("A1").setValue("index");

  // ヘッダー行（再設定）
  const headers = ['No.', 'シート名（リンク付き）', '行数', '列数', '説明'];
  sheetToWrite.getRange(2, 1, 1, headers.length).setValues([headers]);

  const data = [];

  for (let i = 0; i < sheets.length; i++) {
    const sheet = sheets[i];
    const name = sheet.getName();
    const sheetId = sheet.getSheetId();
    const linkFormula = `=HYPERLINK("${baseUrl}#gid=${sheetId}", "${name}")`;
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    data.push([
      i + 1,
      linkFormula,
      lastRow,
      lastCol,
    ]);
  }

  // データ出力（3行目以降）
  const startRow = 3;

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowIndex = startRow + i;
    sheetToWrite.getRange(rowIndex, 1).setValue(row[0]);        // No.
    sheetToWrite.getRange(rowIndex, 2).setFormula(row[1]);      // リンク付きシート名
    sheetToWrite.getRange(rowIndex, 3).setValue(row[2]);        // 行数
    sheetToWrite.getRange(rowIndex, 4).setValue(row[3]);        // 列数
  }

  // 書式設定（列幅や日付など必要なら追加可能）
}
