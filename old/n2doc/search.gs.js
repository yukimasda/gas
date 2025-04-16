function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('🔍検索ツール')
      .addItem('横断検索', 'showSettingsDialog')
      .addToUi();
}

function showSettingsDialog() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
      'サブフォルダを検索しますか？',
      ui.ButtonSet.YES_NO
  );

  // ユーザーが「はい」を選択した場合
  if (response == ui.Button.YES) {
    Logger.log('サブフォルダも検索します');
    executeSearchWithOption(true);
  } else if (response == ui.Button.NO) {
    Logger.log('サブフォルダは検索しません');
    executeSearchWithOption(false);
  }
}

function executeSearchWithOption(searchSubfolders) {
  const searchSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("横断検索");
  const keywordValues = searchSheet.getRange("B2:B" + searchSheet.getLastRow()).getValues();
  const keywords = keywordValues.map(row => row[0]).filter(k => k && k.toString().trim() !== "");

  if (keywords.length === 0) {
    searchSheet.getRange("C2").setValue("❗ キーワードが空です");
    return;
  }

  searchSheet.getRange("C2:H" + searchSheet.getLastRow()).clearContent();  // 結果初期化 (Hまでクリア)
  searchSheet.getRange("C1:G1").setValues([["ヒットワード", "フォルダ名", "ファイル名", "シート名", "ヒットしたセルの値"]]); // ヘッダー変更
  searchSheet.getRange("C2").setValue("🔄 検索中...");
  SpreadsheetApp.flush();

  // フォルダIDを取得
  const folderValues = searchSheet.getRange("A2:A" + searchSheet.getLastRow()).getValues();
  const folderIds = folderValues
    .map(row => {
      const raw = row[0];
      if (!raw) return null;
      const str = raw.toString().trim();
      const match = str.match(/[-\w]{25,}/);
      return match ? match[0] : null;
    })
    .filter(id => id);

  if (folderIds.length === 0) {
    searchSheet.getRange("C2").setValue("❗ フォルダIDが空です");
    return;
  }

  let results = [];
  let processedFiles = 0;
  let processedSheets = 0;
  let totalFiles = 0;
  let totalSheets = 0;

  // フォルダ内のファイル数とシート数を取得
  for (const folderId of folderIds) {
    try {
      const folder = DriveApp.getFolderById(folderId);
      const allFiles = getAllFilesWithParent(folder, searchSubfolders); // 親フォルダ情報も取得

      totalFiles += allFiles.length;
      allFiles.forEach(item => {
        const ss = SpreadsheetApp.openById(item.file.getId());
        totalSheets += ss.getSheets().length;
      });
    } catch (e) {
      Logger.log(`❗ フォルダID ${folderId} の検索中にエラー: ${e}`);
    }
  }

  // 実際の検索処理
  for (const folderId of folderIds) {
    try {
      const folder = DriveApp.getFolderById(folderId);
      const allFilesWithParent = getAllFilesWithParent(folder, searchSubfolders); // 親フォルダ情報も取得

      for (const item of allFilesWithParent) {
        const file = item.file;
        const parentFolder = item.parent;
        const fileId = file.getId();
        const fileUrl = `https://docs.google.com/spreadsheets/d/${fileId}/edit`;
        const ss = SpreadsheetApp.openById(fileId);
        const sheets = ss.getSheets();
        const folderLink = parentFolder ? `=HYPERLINK("${parentFolder.getUrl()}", "${parentFolder.getName()}")` : "";

        sheets.forEach(sheet => {
          const sheetName = sheet.getName();
          const sheetId = sheet.getSheetId();
          const values = sheet.getDataRange().getValues();

          values.forEach((row, r) => {
            row.forEach((cell, c) => {
              const cellValue = cell.toString();
              const matched = keywords.filter(k => new RegExp(k, 'i').test(cellValue));
              if (matched.length > 0) {
                const cellA1 = sheet.getRange(r + 1, c + 1).getA1Notation();
                const sheetLink = `=HYPERLINK("${fileUrl}#gid=${sheetId}&range=${cellA1}", "${sheetName}")`;
                results.push([matched.join(", "), folderLink, file.getName(), sheetLink, cellValue]); // フォルダ名を追加
              }
            });
          });

          // 進捗表示: 10シートごとに更新
          processedSheets++;
          if (processedSheets % 10 === 0) {
            searchSheet.getRange("C2").setValue(`🔄 ファイル ${processedFiles}/${totalFiles} | シート ${processedSheets}/${totalSheets} 検索中...`);
            SpreadsheetApp.flush();  // 進捗を反映させる
          }
        });
        processedFiles++;
      }
    } catch (e) {
      Logger.log(`❗ フォルダID ${folderId} の検索中にエラー: ${e}`);
    }
  }

  if (results.length > 0) {
    searchSheet.getRange(2, 3, results.length, 5).setValues(results); // 幅を5に変更
    searchSheet.getRange("H1").setValue(`✅ ${results.length} 件ヒット`); // G1からH1に変更
  } else {
    searchSheet.getRange("C2").setValue("🛑 一致するデータは見つかりませんでした");
  }
}

// 再帰的にファイル収集（Sheetsのみ）と親フォルダ情報を取得
function getAllFilesWithParent(folder, includeSubfolders) {
  const items = [];
  const fileIterator = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (fileIterator.hasNext()) {
    items.push({ file: fileIterator.next(), parent: folder });
  }

  if (includeSubfolders) {
    const subFolders = folder.getFolders();
    while (subFolders.hasNext()) {
      const sub = subFolders.next();
      items.push(...getAllFilesWithParent(sub, true));
    }
  }

  return items;
}

