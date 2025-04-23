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
        if (item.type === 'file' && (item.name.endsWith('.ts') || item.name.endsWith('.js') || item.name.endsWith('.vue'))) {
          files.push(item);
        } else if (item.type === 'dir') {
          files = files.concat(await listPhpFiles(item.path));
        }
      }
      
      return files;
    }

    // 最初に全ファイルの内容を一括取得
    const files = await listPhpFiles();
    let processedCount = 0;
    let apiRemaining = 5000;
    let apiResetTime = '';
    
    // ファイルの内容をキャッシュとして保持
    const fileContents = new Map();
    
    // ファイルを5件ずつバッチ処理で取得
    for (let i = 0; i < files.length; i += 5) {
      const fileBatch = files.slice(i, i + 5);
      sheet.getRange("H1").setValue(`ファイル内容を取得中... (${i + 1}-${Math.min(i + 5, files.length)}/${files.length})`);
      SpreadsheetApp.flush();
      
      const contentPromises = fileBatch.map(async file => {
        // blobsの代わりにcontentsエンドポイントを使用
        const contentUrl = `https://api.github.com/repos/${owner}/${repoName}/contents/${file.path}`;
        const contentResponse = await UrlFetchApp.fetch(contentUrl, {
          headers: {
            'Authorization': `token ${token}`,
            'Accept': 'application/vnd.github.v3+json',
            'User-Agent': 'Google Apps Script'
          }
        });
        
        // API制限情報を更新
        apiRemaining = contentResponse.getHeaders()['x-ratelimit-remaining'];
        const resetTime = new Date(contentResponse.getHeaders()['x-ratelimit-reset'] * 1000);
        apiResetTime = Utilities.formatDate(resetTime, 'Asia/Tokyo', 'HH:mm:ss');
        
        const contentData = JSON.parse(contentResponse.getContentText());
        return {
          file: file,
          content: Utilities.newBlob(Utilities.base64Decode(contentData.content)).getDataAsString()
        };
      });

      const contents = await Promise.all(contentPromises);
      
      // キャッシュに保存
      for (const {file, content} of contents) {
        fileContents.set(file.path, content);
      }
      
      await Utilities.sleep(2000); // API制限を考慮した待機
    }
    
    // 全ファイルの内容を取得後、各hookNameで検索
    for (const hookName of hookNames) {
      const row = hookNames.indexOf(hookName) + 2;
      let matches = [];
      
      sheet.getRange("H1").setValue(`"${hookName}" を検索中... (${hookNames.indexOf(hookName) + 1}/${hookNames.length})`);
      SpreadsheetApp.flush();
      
      // キャッシュされた全ファイルの内容から検索
      for (const [filePath, content] of fileContents.entries()) {
        if (content.includes(hookName)) {
          const fileUrl = `https://github.com/${owner}/${repoName}/blob/master/${filePath}`;
          matches.push({
            url: fileUrl,
            displayText: filePath
          });
          processedCount++;
        }
      }
      
      // 検索結果をセルに書き込み
      if (matches.length > 0) {
        const richText = SpreadsheetApp.newRichTextValue();
        const texts = matches.map(m => m.displayText);
        richText.setText(texts.join('\n'));
        
        let currentPos = 0;
        for (const match of matches) {
          richText.setLinkUrl(currentPos, currentPos + match.displayText.length, match.url);
          currentPos += match.displayText.length + 1;
        }
        
        sheet.getRange(row, 8).setRichTextValue(richText.build());
        SpreadsheetApp.flush();
      }
    }

    sheet.getRange("H1").setValue(`検索完了: ${processedCount}件のマッチが見つかりました`);

  } catch (error) {
    Logger.log(`検索エラー: ${error.message}`);
    sheet.getRange("H1").setValue(`エラーが発生しました: ${error.message}`);
  }
}