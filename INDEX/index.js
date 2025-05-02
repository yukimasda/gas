function createIndex() {
  const folderId = '109Z_9ppKDpp3ygfApkhPIUfn8r0OwvTq';  // 対象のフォルダID
  const outputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('INDEX');
  
  if (!outputSheet) {
    throw new Error('「INDEX」シートが見つかりません。事前に作成してください。');
  }

  // 3行目以降のデータをクリア
  outputSheet.getRange(3, 1, outputSheet.getLastRow() - 2, outputSheet.getLastColumn()).clear();
  
  // 処理開始時にD1セルにメッセージを表示
  outputSheet.getRange(1, 4).setValue('INDEX更新中です。しばらくお待ちください。');
  SpreadsheetApp.flush();

  // ヘッダ行追加
  const boldTextStyle = SpreadsheetApp.newTextStyle().setBold(true).build();
  outputSheet.getRange(1, 1).setValue('INDEX').setTextStyle(boldTextStyle);
  
  // 2行目にヘッダーを追加
  const headerStyle = SpreadsheetApp.newTextStyle().setBold(true).build();
  outputSheet.getRange(2, 2).setValue('フォルダ').setTextStyle(headerStyle);
  outputSheet.getRange(2, 3).setValue('ファイル').setTextStyle(headerStyle);
  outputSheet.getRange(2, 4).setValue('シート').setTextStyle(headerStyle);
  
  let row = 3;
  const rootFolder = DriveApp.getFolderById(folderId);
  const batchSize = 10; // バッチサイズを10に変更
  let batchData = [];
  let currentFolderName = '';

  function getAllFiles(folder, parentPath = '') {
    let result = [];
    const currentPath = parentPath ? `${parentPath}/${folder.getName()}` : folder.getName();

    // 00_INDEXフォルダを除外
    if (folder.getName() === '00_INDEX') {
      return result;
    }

    // ファイルの取得を最適化
    const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    const filesArray = [];
    while (files.hasNext()) {
      filesArray.push(files.next());
    }
    filesArray.sort((a, b) => a.getName().localeCompare(b.getName(), 'ja'));

    // サブフォルダの取得を最適化
    const subFolders = [];
    const subIter = folder.getFolders();
    while (subIter.hasNext()) {
      subFolders.push(subIter.next());
    }
    subFolders.sort((a, b) => a.getName().localeCompare(b.getName(), 'ja'));

    // ファイル情報を一括で追加
    for (const file of filesArray) {
      result.push({
        folderName: folder.getName(),
        folderUrl: folder.getUrl(),
        file: file
      });
    }

    // サブフォルダを再帰的に処理
    for (const subFolder of subFolders) {
      result = result.concat(getAllFiles(subFolder, currentPath));
    }

    return result;
  }

  const files = getAllFiles(rootFolder);

  // バッチ処理用の関数
  function processBatch() {
    if (batchData.length > 0) {
      // まず基本データを書き込む
      const range = outputSheet.getRange(row, 1, batchData.length, 4);
      range.setValues(batchData);

      // リンクを設定
      for (let i = 0; i < batchData.length; i++) {
        const data = batchData[i];
        const currentRow = row + i;
        
        // 背景色の設定
        if (data[1] && data[1].url) { // フォルダ行
          outputSheet.getRange(currentRow, 2, 1, 3).setBackground('#e6f3ff'); // 薄い青
        } else { // ファイル行
          if (data[2]) { // ファイル行の場合のみ
            outputSheet.getRange(currentRow, 3, 1, 2).setBackground('#f0f0f0'); // 薄いグレー
          }
        }

        if (data[1] && data[1].url) { // フォルダリンク
          const cell = outputSheet.getRange(currentRow, 2);
          const richText = SpreadsheetApp.newRichTextValue()
            .setText(data[1].text)
            .setLinkUrl(data[1].url)
            .build();
          cell.setRichTextValue(richText);
        }
        if (data[3] && data[3].url) { // シートリンク
          const cell = outputSheet.getRange(currentRow, 4);
          const richText = SpreadsheetApp.newRichTextValue()
            .setText(data[3].text)
            .setLinkUrl(data[3].url)
            .build();
          cell.setRichTextValue(richText);
        }
      }

      row += batchData.length;
      batchData = [];
    }
  }

  for (const item of files) {
    const file = item.file;
    const folderName = item.folderName;
    const folderUrl = item.folderUrl;
    const fileName = file.getName();
    const fileId = file.getId();

    // フォルダ名が変わった場合の処理
    if (folderName !== currentFolderName) {
      processBatch(); // バッチを処理
      currentFolderName = folderName;
      batchData.push(['', { text: folderName, url: folderUrl }, '', '']);
    }

    const spreadsheet = SpreadsheetApp.openById(fileId);
    const sheets = spreadsheet.getSheets();

    // ファイル情報をバッチに追加
    batchData.push(['', '', fileName, '']);

    // シート情報をバッチに追加
    for (let i = 0; i < sheets.length; i++) {
      const sheet = sheets[i];
      const sheetLink = `https://docs.google.com/spreadsheets/d/${fileId}#gid=${sheet.getSheetId()}`;
      batchData.push(['', '', '', { text: sheet.getName(), url: sheetLink }]);
    }

    // バッチサイズに達したら処理
    if (batchData.length >= batchSize) {
      processBatch();
    }
  }

  // 残りのバッチを処理
  processBatch();

  // 罫線の設定
  const lastRow = outputSheet.getLastRow();
  const borderStyle = SpreadsheetApp.BorderStyle.SOLID;
  const borderColor = '#cccccc';
  
  // 外枠
  outputSheet.getRange(2, 2, lastRow - 1, 3).setBorder(true, true, true, true, null, null, borderColor, borderStyle);
  
  // 内部の横線
  outputSheet.getRange(2, 2, lastRow - 1, 3).setBorder(null, null, null, null, true, null, borderColor, borderStyle);

  // 処理完了時に更新日時を表示
  const now = new Date();
  const formattedDate = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年MM月dd日 HH:mm');
  outputSheet.getRange(1, 4).setValue(`更新日時: ${formattedDate}`);
}
