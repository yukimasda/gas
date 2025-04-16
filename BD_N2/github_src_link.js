function getAllFiles() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GitHub Src Links');
  if (!sheet) throw new Error('「GitHub Src Links」シートが見つかりません');
  
  // シートをクリア
  sheet.clear();
  
  // ヘッダー設定
  sheet.getRange('A1').setValue('GitHub repo list');
  sheet.getRange('A1:C1').merge().setFontWeight('bold').setHorizontalAlignment('center');
  
  // カラム名の設定
  sheet.getRange('A2').setValue('ディレクトリ');
  sheet.getRange('B2').setValue('ファイル名');
  sheet.getRange('C2').setValue('説明');
  sheet.getRange('A2:C2').setFontWeight('bold').setBackground('#f3f3f3');
  
  // カラム幅の設定
  sheet.setColumnWidth(1, 250); // A列：ディレクトリ
  sheet.setColumnWidth(2, 200); // B列：ファイル名
  sheet.setColumnWidth(3, 400); // C列：説明
  
  // リポジトリのルートディレクトリ内のファイル一覧を取得
  const rootUrl = `https://api.github.com/repos/${repo}/contents/`;
  fetchFiles(rootUrl, sheet, 3, "", ""); // 3行目から表示開始、初期パスは空、初期ディレクトリは空
  
  // 罫線を追加
  const lastRow = sheet.getLastRow();
  if (lastRow > 2) {
    sheet.getRange(2, 1, lastRow - 1, 3).setBorder(true, true, true, true, true, true);
  }
}

/**
 * 特定のディレクトリやファイルを除外すべきかどうかを判断する
 */
function shouldExclude(path, name) {
  // 除外するディレクトリやファイル名のリスト
  const excludeDirs = ['dist', 'node_modules', '.git'];
  const excludeExtensions = ['.min.js', '.map'];
  
  // パス全体またはディレクトリ名が除外リストに含まれているかチェック
  for (const dir of excludeDirs) {
    if (path.includes(`/${dir}/`) || path === dir || name === dir) {
      return true;
    }
  }
  
  // 除外する拡張子をチェック
  for (const ext of excludeExtensions) {
    if (name.endsWith(ext)) {
      return true;
    }
  }
  
  return false;
}

function fetchFiles(url, sheet, startRow, currentPath, lastDir) {
  // URLにdistディレクトリが含まれている場合はスキップ
  if (shouldExclude(currentPath, '')) {
    return startRow; // 現在の行を変更せずに返す
  }

  const options = {
    headers: {
      Authorization: `token ${token}`,
      Accept: "application/vnd.github.v3.raw"
    },
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());

  if (Array.isArray(data)) {
    let row = startRow;
    let currentDirDisplayed = false; // 現在のディレクトリが表示されたかどうかを追跡
    
    // ファイルとディレクトリを分離して、ファイルを先に処理
    const files = data.filter(item => item.type === 'file' && !shouldExclude(currentPath, item.name));
    const dirs = data.filter(item => item.type === 'dir' && !shouldExclude(currentPath, item.name));
    
    // ファイルをファイル名でソート
    files.sort((a, b) => a.name.localeCompare(b.name));
    
    // ファイルの処理
    files.forEach(item => {
      const fileName = item.name;
      const fileLink = item.html_url;
      const filePath = currentPath ? `${currentPath}` : "";

      // ディレクトリが変わった場合のみディレクトリ名を表示
      if (lastDir !== currentPath) {
        lastDir = currentPath;
        currentDirDisplayed = false;
      }
      
      // A列にディレクトリ（同じディレクトリでは最初だけ表示）
      if (!currentDirDisplayed) {
        sheet.getRange(row, 1).setValue(filePath);
        currentDirDisplayed = true;
      }
      
      // B列にファイル名（リンク付き）
      sheet.getRange(row, 2).setFormula(`=HYPERLINK("${fileLink}", "${fileName}")`);
      
      // C列は空にしておく（説明処理を削除）
      
      row++;
    });
    
    // ディレクトリをソート
    dirs.sort((a, b) => a.name.localeCompare(b.name));
    
    // ディレクトリの処理（再帰的に）
    for (const dir of dirs) {
      const subDirUrl = dir.url;
      const subDirPath = currentPath ? `${currentPath}/${dir.name}` : dir.name;
      row = fetchFiles(subDirUrl, sheet, row, subDirPath, lastDir);
    }
    
    return row;
  }
  return startRow;
}
