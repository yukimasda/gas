async function searchHookUsages() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Hook List');
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  // H列をクリア
  sheet.getRange(2, 8, lastRow - 1, 1).clearContent();
  sheet.getRange("H1").setValue("フック使用箇所");

  const hookNames = sheet.getRange(2, 3, lastRow - 1, 1).getValues().flat().filter(k => k);
  if (hookNames.length === 0) return;

  if (!token) {
    sheet.getRange("H1").setValue("⚠️ GitHub トークンが設定されていません");
    return;
  }

  const [owner, repoName] = "steamships/neo-neng".split('/');
  
  try {
    // まずsrcディレクトリ配下のファイル一覧を再帰的に取得
    async function listPhpFiles(path = 'src') {
      const url = `https://api.github.com/repos/${owner}/${repoName}/contents/${path}`;
      const response = await UrlFetchApp.fetch(url, {
        headers: { 
          'Authorization': `token ${token}`,
          'Accept': 'application/vnd.github.v3+json',
          'User-Agent': 'Google Apps Script'
        }
      });
      
      const items = JSON.parse(response.getContentText());
      let files = [];
      
      for (const item of items) {
        if (item.type === 'file' && (item.name.endsWith('.ts') || item.name.endsWith('.js') || item.name.endsWith('.yml'))) {
          files.push(item);
        } else if (item.type === 'dir') {
          files = files.concat(await listPhpFiles(item.path));
        }
      }
      
      return files;
    }

    // ファイル一覧を取得
    const files = await listPhpFiles();
    let fileCount = 0;
    let processedCount = 0;
    let updateBuffer = [];  // 更新用バッファ
    const matchCounts = new Map(); // 各hookNameのマッチ数を追跡

    // I列以降をクリア
    sheet.getRange(2, 8, lastRow - 1, 5).clearContent(); // H列からL列までクリア

    for (const file of files) {
      fileCount++;
      sheet.getRange("H1").setValue(`ファイル検索中... (${fileCount}/${files.length})`);
      SpreadsheetApp.flush();
      
      // blobからファイル内容を取得
      const blobUrl = `https://api.github.com/repos/${owner}/${repoName}/git/blobs/${file.sha}`;
      const blobResponse = await UrlFetchApp.fetch(blobUrl, {
        headers: {
          'Authorization': `token ${token}`,
          'Accept': 'application/vnd.github.v3+json',
          'User-Agent': 'Google Apps Script'
        }
      });

      const blobData = JSON.parse(blobResponse.getContentText());
      const content = Utilities.newBlob(Utilities.base64Decode(blobData.content)).getDataAsString();

      // API制限の表示
      const remaining = blobResponse.getHeaders()['x-ratelimit-remaining'];
      const resetTime = new Date(blobResponse.getHeaders()['x-ratelimit-reset'] * 1000);
      const resetTimeJST = Utilities.formatDate(resetTime, 'Asia/Tokyo', 'HH:mm:ss');
      
      // 各フックの検索
      for (const hookName of hookNames) {
        if (content.includes(hookName)) {
          const fileUrl = `https://github.com/${owner}/${repoName}/blob/master/${file.path}`;
          const hyperlink = `=HYPERLINK("${fileUrl}", "${file.path}")`;
          
          // マッチ数をカウント
          if (!matchCounts.has(hookName)) {
            matchCounts.set(hookName, 0);
          }
          const matchCount = matchCounts.get(hookName);
          matchCounts.set(hookName, matchCount + 1);
          
          // 更新バッファに追加（列を計算）
          updateBuffer.push({
            row: hookNames.indexOf(hookName) + 2,
            col: 8 + matchCount, // H列(8)から開始
            value: hyperlink
          });
          processedCount++;
          
          // 10件たまったら一括更新
          if (updateBuffer.length >= 10) {
            sheet.getRange("H1").setValue(`バッチ更新中... (${processedCount}件のマッチ) - Core API制限: ${remaining}/5000`);
            SpreadsheetApp.flush();
            
            // バッファ内の各行を個別に更新
            for (const update of updateBuffer) {
              sheet.getRange(update.row, update.col).setFormula(update.value);
              SpreadsheetApp.flush();
              Utilities.sleep(100);
            }
            
            updateBuffer = [];
            
            sheet.getRange("H1").setValue(`検索中... (${processedCount}件のマッチを発見) - 次のファイルへ`);
            SpreadsheetApp.flush();
          }
        }
      }

      await Utilities.sleep(1000);
    }

    // 残りのバッファを処理
    if (updateBuffer.length > 0) {
      sheet.getRange("H1").setValue(`最終バッチ更新中... (${processedCount}件のマッチ)`);
      SpreadsheetApp.flush();
      
      for (const update of updateBuffer) {
        sheet.getRange(update.row, update.col).setFormula(update.value);
        SpreadsheetApp.flush();
        Utilities.sleep(100);
      }
    }

    sheet.getRange("H1").setValue(`検索完了: ${processedCount}件のマッチが見つかりました`);
    SpreadsheetApp.flush();

  } catch (error) {
    Logger.log(`検索エラー: ${error.message}`);
    sheet.getRange("H1").setValue(`エラーが発生しました: ${error.message}`);
    SpreadsheetApp.flush();
  }
}