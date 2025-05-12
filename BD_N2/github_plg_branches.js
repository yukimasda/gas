function getAllBranches() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GitHub Plg Links');
  if (!sheet) throw new Error('「GitHub Plg Links」シートが見つかりません');
  
  // シートをクリア
  sheet.clear();
  
  // ヘッダー設定
  sheet.getRange('A1').setValue('GitHub Plugin Branches');
  sheet.getRange('A1:E1').merge().setFontWeight('bold').setHorizontalAlignment('center');
  
  // カラム名の設定
  sheet.getRange('A2').setValue('ブランチ名');
  sheet.getRange('B2').setValue('最終コミット日時');
  sheet.getRange('C2').setValue('コミッター');
  sheet.getRange('D2').setValue('Plugin Name');
  sheet.getRange('E2').setValue('Description');
  sheet.getRange('A2:E2').setFontWeight('bold').setBackground('#f3f3f3');
  
  // カラム幅の設定
  sheet.setColumnWidth(1, 200); // A列：ブランチ名
  sheet.setColumnWidth(2, 200); // B列：最終コミット日時
  sheet.setColumnWidth(3, 150); // C列：コミッター
  sheet.setColumnWidth(4, 300); // D列：Plugin Name
  sheet.setColumnWidth(5, 400); // E列：Description
  
  let page = 1;
  let allBranches = [];
  
  try {
    // 全てのブランチを取得（ページネーション対応）
    while (true) {
      const branchesUrl = `https://api.github.com/repos/steamships/n2-plugins/branches?per_page=100&page=${page}`;
      const options = {
        headers: {
          Authorization: `token ${token}`,
          Accept: "application/vnd.github.v3+json"
        },
        muteHttpExceptions: true
      };
      
      const response = UrlFetchApp.fetch(branchesUrl, options);
      const branches = JSON.parse(response.getContentText());
      
      if (branches.length === 0) break;
      
      allBranches = allBranches.concat(branches);
      page++;
      
      // API制限を考慮して少し待機
      Utilities.sleep(1000);
    }
    
    let row = 3;
    for (const branch of allBranches) {
      // 進捗状況を表示
      sheet.getRange('A1').setValue(`GitHub Plugin Branches (処理中... ${row-2}/${allBranches.length})`);
      
      // index.phpの内容を取得
      const indexPhpUrl = `https://api.github.com/repos/steamships/n2-plugins/contents/index.php?ref=${encodeURIComponent(branch.name)}`;
      let pluginName = '-';
      let description = '-';
      
      try {
        const indexResponse = UrlFetchApp.fetch(indexPhpUrl, {
          headers: {
            Authorization: `token ${token}`,
            Accept: "application/vnd.github.v3.raw"
          },
          muteHttpExceptions: true
        });
        
        if (indexResponse.getResponseCode() === 200) {
          const content = indexResponse.getContentText();
          
          // Plugin Name を抽出
          const pluginNameMatch = content.match(/Plugin Name:\s*(.+)/);
          if (pluginNameMatch) {
            pluginName = pluginNameMatch[1].trim();
          }
          
          // Description を抽出
          const descriptionMatch = content.match(/Description:\s*(.+)/);
          if (descriptionMatch) {
            description = descriptionMatch[1].trim();
          }
        }
      } catch (error) {
        Logger.log(`index.php取得エラー (${branch.name}): ${error.message}`);
      }
      
      // コミット情報の取得（既存のコード）
      const commitUrl = branch.commit.url;
      const commitResponse = UrlFetchApp.fetch(commitUrl, {
        headers: {
          Authorization: `token ${token}`,
          Accept: "application/vnd.github.v3+json"
        },
        muteHttpExceptions: true
      });
      const commitData = JSON.parse(commitResponse.getContentText());
      
      // 日時をJST（日本時間）に変換
      const commitDate = new Date(commitData.commit.author.date);
      const jstDate = Utilities.formatDate(commitDate, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
      
      // GitHubのブランチURLを作成
      const branchUrl = `https://github.com/steamships/n2-plugins/tree/${encodeURIComponent(branch.name)}`;
      
      // リッチテキストでリンクを設定
      const richTextValue = SpreadsheetApp.newRichTextValue()
        .setText(branch.name)
        .setLinkUrl(branchUrl)
        .build();
      
      // データを設定
      sheet.getRange(row, 1).setRichTextValue(richTextValue);
      sheet.getRange(row, 2).setValue(jstDate);
      sheet.getRange(row, 3).setValue(commitData.commit.author.name);
      sheet.getRange(row, 4).setValue(pluginName);
      sheet.getRange(row, 5).setValue(description);
      
      row++;
      
      // API制限を考慮して少し待機
      Utilities.sleep(1000);
    }
    
    // タイトルを最終的な状態に更新
    sheet.getRange('A1').setValue('GitHub Plugin Branches');
    
    // 罫線を追加
    const lastRow = sheet.getLastRow();
    if (lastRow > 2) {
      sheet.getRange(2, 1, lastRow - 1, 5).setBorder(true, true, true, true, true, true);
    }
    
  } catch (error) {
    Logger.log(`エラーが発生しました: ${error.message}`);
    throw new Error(`ブランチ情報の取得に失敗しました: ${error.message}`);
  }
} 