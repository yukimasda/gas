function createIndex() {
  const folderId = '109Z_9ppKDpp3ygfApkhPIUfn8r0OwvTq';  // 対象のフォルダID
  const outputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('INDEX');
  
  if (!outputSheet) {
    throw new Error('「INDEX」シートが見つかりません。事前に作成してください。');
  }

  outputSheet.clearContents();
  // ヘッダ行追加（2行目）
  const boldTextStyle = SpreadsheetApp.newTextStyle().setBold(true).build();
  outputSheet.getRange(1, 1).setValue('INDEX').setTextStyle(boldTextStyle); // A1セルにタイトル
 
  
  let row = 3;  // データの開始行

  const rootFolder = DriveApp.getFolderById(folderId);

  function getAllFiles(folder, parentPath = '') {
    let result = [];

    const currentPath = parentPath ? `${parentPath}/${folder.getName()}` : folder.getName();

    const filesArray = [];
    const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    while (files.hasNext()) {
      filesArray.push(files.next());
    }
    filesArray.sort((a, b) => a.getName().localeCompare(b.getName(), 'ja'));

    for (const file of filesArray) {
      result.push({
        folderName: folder.getName(),
        folderUrl: folder.getUrl(),
        file: file
      });
    }

    const subFolders = [];
    const subIter = folder.getFolders();
    while (subIter.hasNext()) {
      subFolders.push(subIter.next());
    }
    subFolders.sort((a, b) => a.getName().localeCompare(b.getName(), 'ja'));

    for (const subFolder of subFolders) {
      result = result.concat(getAllFiles(subFolder, currentPath));
    }

    return result;
  }

  const files = getAllFiles(rootFolder);

  let currentFolderName = '';

  for (const item of files) {
    const file = item.file;
    const folderName = item.folderName;
    const folderUrl = item.folderUrl;
    const fileName = file.getName();
    const fileId = file.getId();

    const spreadsheet = SpreadsheetApp.openById(fileId);
    const sheets = spreadsheet.getSheets();

    // フォルダ名が変わったらB列にリンクを追加
    if (folderName !== currentFolderName) {
      currentFolderName = folderName;
      outputSheet.getRange(row, 2).setFormula(`=HYPERLINK("${folderUrl}", "${folderName}")`);
      row++;
    }

    // C列: ファイル名（リンクなし）
    outputSheet.getRange(row, 3).setValue(fileName);

    // D列: シート名をリンク付きで追加
    for (let i = 0; i < sheets.length; i++) {
      const sheet = sheets[i];
      const sheetLink = `https://docs.google.com/spreadsheets/d/${fileId}#gid=${sheet.getSheetId()}`;
      
      // 最初のシートのときだけ空行を挿入
      if (i === 0) {
        row++;
      }
      outputSheet.getRange(row, 4).setFormula(`=HYPERLINK("${sheetLink}", "${sheet.getName()}")`);
      row++;
    }

    row++; // ファイルごとに空行を挿入
  }
}
