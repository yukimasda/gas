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
    let updateBuffer = new Map(); // hookNameごとのバッファを管理
    
    // H列のみクリア
    sheet.getRange(2, 8, lastRow - 1, 1).clearContent();

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
          const row = hookNames.indexOf(hookName) + 2;
          
          // 既存のセルの内容を取得
          if (!updateBuffer.has(row)) {
            const existingCell = sheet.getRange(row, 8);
            const existingRichText = existingCell.getRichTextValue();
            updateBuffer.set(row, []);
            
            // 既存のリンクがある場合はバッファに追加
            if (existingRichText && existingRichText.getText()) {
              const text = existingRichText.getText();
              const urls = existingRichText.getLinkUrls();
              const lines = text.split('\n');
              
              lines.forEach((line, index) => {
                if (urls[index]) {
                  updateBuffer.get(row).push({
                    url: urls[index],
                    displayText: line
                  });
                }
              });
            }
          }
          
          // 新しいリンクをバッファに追加
          updateBuffer.get(row).push({
            url: fileUrl,
            displayText: file.path
          });
          processedCount++;
          
          // 10件たまったら一括更新
          if (processedCount % 10 === 0) {
            sheet.getRange("H1").setValue(`バッチ更新中... (${processedCount}件のマッチ) - Core API制限: ${remaining}/5000`);
            SpreadsheetApp.flush();
            
            // バッファ内の各行を更新
            for (const [row, links] of updateBuffer.entries()) {
              const richText = SpreadsheetApp.newRichTextValue();
              const texts = links.map(l => l.displayText);
              richText.setText(texts.join('\n'));
              
              // 各リンクのURLを設定
              let currentPos = 0;
              for (const link of links) {
                richText.setLinkUrl(currentPos, currentPos + link.displayText.length, link.url);
                currentPos += link.displayText.length + 1; // +1 for newline
              }
              
              sheet.getRange(row, 8).setRichTextValue(richText.build());
              SpreadsheetApp.flush();
              Utilities.sleep(100);
            }
            
            // バッファをクリアせずに維持
            sheet.getRange("H1").setValue(`検索中... (${processedCount}件のマッチを発見) - 次のファイルへ`);
            SpreadsheetApp.flush();
          }
        }
      }

      await Utilities.sleep(1000);
    }

    // 残りのバッファを処理
    if (updateBuffer.size > 0) {
      sheet.getRange("H1").setValue(`最終バッチ更新中... (${processedCount}件のマッチ)`);
      SpreadsheetApp.flush();
      
      for (const [row, links] of updateBuffer.entries()) {
        const richText = SpreadsheetApp.newRichTextValue();
        const texts = links.map(l => l.displayText);
        richText.setText(texts.join('\n'));
        
        let currentPos = 0;
        for (const link of links) {
          richText.setLinkUrl(currentPos, currentPos + link.displayText.length, link.url);
          currentPos += link.displayText.length + 1;
        }
        
        sheet.getRange(row, 8).setRichTextValue(richText.build());
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