async function searchHookUsagesWP() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Hook List');
    if (!sheet) return;
  
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return;
  
    // I列をクリア
    sheet.getRange(2, 9, lastRow - 1, 1).clearContent();
    sheet.getRange("I1").setValue("WordPress使用箇所");
  
    const hookNames = sheet.getRange(2, 3, lastRow - 1, 1).getValues().flat().filter(k => k);
    if (hookNames.length === 0) return;
  
    // GitHubトークンの確認
    const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
    if (!token) {
      sheet.getRange("I1").setValue("⚠️ GitHub トークンが設定されていません");
      return;
    }
  
    let allResults = new Map();
    let processedCount = 0;
  
    // 各フックに対して検索を実行
    for (const hookName of hookNames) {
      processedCount++;
      sheet.getRange("I1").setValue(`検索中... (${processedCount}/${hookNames.length})`);
      SpreadsheetApp.flush();
  
      try {
        // 検索クエリの構築
        const query = `"${hookName}" repo:WordPress/WordPress`;
        const encodedQuery = encodeURIComponent(query);
        const url = `https://api.github.com/search/code?q=${encodedQuery}&per_page=100`;
        
        const options = {
          headers: {
            Authorization: `token ${token}`,
            Accept: "application/vnd.github.v3+json"
          },
          muteHttpExceptions: true
        };
  
        const response = UrlFetchApp.fetch(url, options);
        const data = JSON.parse(response.getContentText());
  
        if (data.message) {
          if (data.message.includes("API rate limit exceeded")) {
            sheet.getRange("I1").setValue("⚠️ GitHub APIレート制限超過。しばらく待ってから再試行してください。");
            return;
          }
          Logger.log(`GitHub APIエラー: ${data.message}`);
          continue;
        }
  
        if (data.items && data.items.length > 0) {
          const results = data.items.map(item => `${item.path}`);
          allResults.set(hookName, results);
        }
  
        // API制限を考慮して待機
        await Utilities.sleep(2000);
  
      } catch (error) {
        Logger.log(`検索エラー (${hookName}): ${error.message}`);
        continue;
      }
    }
  
    // 結果をシートに書き込み
    let row = 2;
    for (const [hookName, results] of allResults) {
      if (results && results.length > 0) {
        sheet.getRange(row, 9).setValue(results.join('\n'));
      } else {
        sheet.getRange(row, 9).setValue('検索結果なし');
      }
      row++;
    }
  
    sheet.getRange("I1").setValue(`検索完了: ${processedCount}件のフックを検索しました`);
}

async function searchInWordPress(keyword) {
    Logger.log(`WordPress検索開始: ${keyword}`);
    const results = [];
    
    try {
      const searchResults = await searchInRepo(
        'WordPress',
        'WordPress',
        keyword,
        ''  // WordPressの場合、全体を検索
      );
      Logger.log(`WordPress検索結果: ${searchResults.length}件`);
      results.push(...searchResults);
    } catch (error) {
      Logger.log(`WordPress検索エラー: ${error}`);
      throw error;
    }
    
    Logger.log(`WordPress検索完了: ${results.length}件の結果`);
    return results;
}

async function searchInRepo(owner, repo, keyword, path) {
    const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
    const url = `https://api.github.com/search/code?q=${encodeURIComponent(keyword)}+in:file${path ? '+path:' + encodeURIComponent(path) : ''}+repo:${owner}/${repo}`;
    
    Logger.log(`検索URL: ${url}`);
  
    const options = {
      method: 'get',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Accept': 'application/vnd.github.v3+json'
      },
      muteHttpExceptions: true
    };
  
    try {
      const response = await UrlFetchApp.fetch(url, options);
      const contentText = response.getContentText();
      const json = JSON.parse(contentText);
      
      Logger.log(`レスポンス本文: ${contentText}`);
      
      if (json.message && json.message.includes("API rate limit exceeded")) {
        throw new Error("API_RATE_LIMIT_EXCEEDED");
      }
      
      if (json.items) {
        Logger.log(`検索結果件数: ${json.items.length}`);
        return json.items.map(item => `${item.path}#L${item.line || '1'}`);
      }
      return [];
    } catch (error) {
      Logger.log(`API呼び出しエラー: ${error}`);
      throw error;
    }
}
