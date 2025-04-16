function exportSheetsAsSpreadsheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const listSheet = ss.getSheetByName("index");
  const data = listSheet.getDataRange().getValues(); // 全データ取得

  const folder = DriveApp.createFolder("書き出しスプレッドシート_" + new Date().toISOString());

  for (let i = 2; i < data.length; i++) { // 3行目から開始（配列のインデックスは0から始まるため2）
    const sheetName = data[i][1]; // B列（インデックスは1）
    const fileName = data[i][4];  // E列（インデックスは4）

    if (!sheetName || !fileName) continue; // 空のセルはスキップ

    const originalSheet = ss.getSheetByName(sheetName);
    if (!originalSheet) {
      Logger.log(`シート「${sheetName}」が見つかりません`);
      continue;
    }

    // 既存のスプレッドシートを探す
    let newSpreadsheet;
    const files = DriveApp.getFilesByName(fileName);
    if (files.hasNext()) {
      // 既存のスプレッドシートが見つかった場合
      newSpreadsheet = SpreadsheetApp.open(files.next());
      Logger.log(`既存のスプレッドシート「${fileName}」にシートを追加します。`);
    } else {
      // 新しいスプレッドシートを作成
      newSpreadsheet = SpreadsheetApp.create(fileName);
      Logger.log(`新しいスプレッドシート「${fileName}」を作成しました。`);
    }

    // 新しいシートをコピーして追加
    let newSheet = originalSheet.copyTo(newSpreadsheet);
    
    // シート名の重複を避けるための処理
    let newSheetName = sheetName;
    let counter = 1;
    while (newSpreadsheet.getSheetByName(newSheetName)) {
      newSheetName = `${sheetName}_${counter}`;
      counter++;
    }
    newSheet.setName(newSheetName);

    // デフォルトで作られる「Sheet1」などの空シートを削除
    const defaultSheet = newSpreadsheet.getSheets().find(s => s.getSheetName() === "Sheet1");
    if (defaultSheet) {
      newSpreadsheet.deleteSheet(defaultSheet);
    }

    // 作成したスプレッドシートを指定フォルダに移動
    const newFile = DriveApp.getFileById(newSpreadsheet.getId());
    folder.addFile(newFile);
    DriveApp.getRootFolder().removeFile(newFile); // 「マイドライブ」から削除（元の場所から移動）
  }
}
