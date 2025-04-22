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
  let processedCount = 0;

  for (let row = 2; row <= lastRow; row++) {
    const hookName = sheet.getRange(row, 3).getValue();
    if (!hookName) continue;

    processedCount++;
    sheet.getRange("H1").setValue(`検索中... (${processedCount}/${hookNames.length})`);
    SpreadsheetApp.flush();

    try {
      const query = `"${hookName}" path:src/ repo:${owner}/${repoName}`;
      const url = `https://api.github.com/search/code?q=${encodeURIComponent(query)}`;
      
      const options = {
        headers: {
          Authorization: `token ${token}`,
          Accept: "application/vnd.github.v3+json"
        },
        muteHttpExceptions: true
      };

      const response = UrlFetchApp.fetch(url, options);
      
      // レート制限の確認と表示
      const remaining = response.getHeaders()['x-ratelimit-remaining'];
      const resetTime = new Date(response.getHeaders()['x-ratelimit-reset'] * 1000);
      const resetTimeJST = Utilities.formatDate(resetTime, 'Asia/Tokyo', 'HH:mm:ss');
      Logger.log(`残りAPI制限: ${remaining}, リセット時間: ${resetTimeJST}`);
      sheet.getRange("H1").setValue(`検索中... (${processedCount}/${hookNames.length}) - 残りAPI制限: ${remaining} (${resetTimeJST}にリセット)`);

      const data = JSON.parse(response.getContentText());

      if (data.message) {
        if (data.message.includes("API rate limit exceeded")) {
          sheet.getRange("H1").setValue(`⚠️ GitHub APIレート制限超過 (${resetTimeJST}にリセット)`);
          return;
        }
        continue;
      }

      if (data.items && data.items.length > 0) {
        const results = data.items.map(item => {
          const fileUrl = `https://github.com/${owner}/${repoName}/blob/master/${item.path}`;
          return `${item.path} (${fileUrl})`;
        });
        
        sheet.getRange(row, 8).setValue(results.join('\n'));
      }

      await Utilities.sleep(2000);

    } catch (error) {
      Logger.log(`検索エラー (${hookName}): ${error.message}`);
      continue;
    }
  }

  sheet.getRange("H1").setValue(`検索完了: ${processedCount}件のフックを検索しました`);
}