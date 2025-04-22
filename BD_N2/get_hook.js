// このファイルでは token、repo、branch の定義は省略します。
// これらの変数は他のソースで定義されていると想定します。

// メイン関数を非同期に修正
async function fetchHooksFromGitHub() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Hook List');
  if (!sheet) {
    Browser.msgBox("エラー", "「Hook List」シートが見つかりません。シートを作成してください。", Browser.Buttons.OK);
    return;
  }

  const [owner, repoName] = repo.split('/');
  const folder = 'inc';
  
  sheet.clear();
  sheet.appendRow(['ファイル名', 'クラス名', 'フック名', 'コールバック関数名', '種別', '行番号', '推定される役割']);

  let totalHooks = 0;
  let rowBuffer = [];  // 結果を一時保存するバッファ
  let fileCount = 0;  // ファイル処理カウンター
  let remaining = null;
  let resetTimeJST = null;
  
  try {
    const files = listPhpFiles(owner, repoName, folder);
    
    for (const file of files) {
      fileCount++;
      let content;
      
      // 10ファイルごとにAPI制限をチェック
      if (fileCount % 10 === 1) {
        const response = await fetchFileContent(owner, repoName, file.path);
        content = response.content;
        remaining = response.remaining;
        resetTimeJST = response.resetTimeJST;
      } else {
        content = (await fetchFileContent(owner, repoName, file.path)).content;
      }

      const lines = content.split('\n');
      let currentClass = '';
      
      for (let index = 0; index < lines.length; index++) {
        const line = lines[index];
        const classMatch = line.match(/class\s+(\w+)/);
        if (classMatch) {
          currentClass = classMatch[1];
        }
        
        if (line.includes('add_action(') || line.includes('add_filter(')) {
          const type = line.includes('add_action(') ? 'action' : 'filter';
          const fileUrl = `https://github.com/${repo}/blob/master/${file.path}`;
          const lineLink = `=HYPERLINK("${fileUrl}#L${index + 1}", "${index + 1}")`;
          
          let hookName = '';
          const hookMatch = line.match(/['"]([^'"]+)['"]/);
          if (hookMatch) {
            hookName = hookMatch[1];
          }

          let callback = '';
          const callbackMatch = line.match(/,\s*['"]([^'"]+)['"]/);
          if (callbackMatch) {
            callback = callbackMatch[1];
          }

          // バッファに追加
          rowBuffer.push([file.path, currentClass, hookName, callback, type, lineLink, '']);
          totalHooks++;
          
          // 10行たまったらまとめて書き込み
          if (rowBuffer.length >= 10) {
            const startRow = totalHooks - rowBuffer.length + 2;
            sheet.getRange(startRow, 1, rowBuffer.length, 7).setValues(rowBuffer);
            rowBuffer = [];  // バッファをクリア
            
            // API制限情報を含めて表示を更新
            const statusMessage = remaining 
              ? `検索中... ${totalHooks}件のフックが見つかりました (残りAPI制限: ${remaining}, ${resetTimeJST}にリセット)`
              : `検索中... ${totalHooks}件のフックが見つかりました`;
            sheet.getRange("A1").setValue(statusMessage);
            SpreadsheetApp.flush();
          }
        }
      }
    }
    
    // 残りのバッファを書き込み
    if (rowBuffer.length > 0) {
      const startRow = totalHooks - rowBuffer.length + 2;
      sheet.getRange(startRow, 1, rowBuffer.length, 7).setValues(rowBuffer);
    }
    
    sheet.getRange("A1").setValue(`検索完了: ${totalHooks}件のフックが見つかりました`);
    
    // 結果をソート
    if (totalHooks > 0) {
      const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7);
      dataRange.sort([{column: 1, ascending: true}, {column: 6, ascending: true}]);
    }
    
  } catch (error) {
    Logger.log(`検索中にエラー: ${error}`);
    sheet.getRange("A1").setValue(`エラーが発生しました: ${error}`);
  }
}

function listPhpFiles(owner, repoName, path) {
  const url = `https://api.github.com/repos/${owner}/${repoName}/contents/${path}`;
  const res = UrlFetchApp.fetch(url, {
    headers: { Authorization: `token ${token}` }
  });
  const json = JSON.parse(res.getContentText());
  let files = [];
  json.forEach(item => {
    if (item.type === 'file' && item.name.endsWith('.php')) {
      files.push(item);
    } else if (item.type === 'dir') {
      files = files.concat(listPhpFiles(owner, repoName, item.path));
    }
  });
  return files;
}

async function fetchFileContent(owner, repoName, path) {
  const url = `https://api.github.com/repos/${owner}/${repoName}/contents/${path}`;
  const res = UrlFetchApp.fetch(url, {
    headers: { Authorization: `token ${token}` }
  });
  const json = JSON.parse(res.getContentText());
  
  // API制限情報を取得
  const remaining = res.getHeaders()['x-ratelimit-remaining'];
  const resetTime = new Date(res.getHeaders()['x-ratelimit-reset'] * 1000);
  const resetTimeJST = Utilities.formatDate(resetTime, 'Asia/Tokyo', 'HH:mm:ss');

  return {
    content: Utilities.newBlob(Utilities.base64Decode(json.content)).getDataAsString(),
    remaining: remaining,
    resetTimeJST: resetTimeJST
  };
}
