function getAllBranches() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GitHub Plg Links');
  if (!sheet) throw new Error('「GitHub Plg Links」シートが見つかりません');
  
  // シートをクリア
  sheet.clear();
  SpreadsheetApp.flush(); // クリア処理を即時反映

  // 初期メッセージを表示
  sheet.getRange('A1').setValue('GitHub Plugin Branches (準備中...)');
  SpreadsheetApp.flush();

  // ヘッダー設定を6列に拡張
  sheet.getRange('A1').setValue('GitHub Plugin Branches');
  sheet.getRange('A1:F1').merge().setFontWeight('bold').setHorizontalAlignment('center');
  
  // カラム名の設定に分類を追加
  sheet.getRange('F2').setValue('分類');
  sheet.getRange('A2').setValue('ブランチ名');
  sheet.getRange('B2').setValue('最終コミット日時');
  sheet.getRange('C2').setValue('コミッター');
  sheet.getRange('D2').setValue('Plugin Name');
  sheet.getRange('E2').setValue('Description');
  sheet.getRange('A2:F2').setFontWeight('bold').setBackground('#f3f3f3');
  
  // カラム幅の設定に分類列を追加
  sheet.setColumnWidth(6, 100); // F列：分類
  sheet.setColumnWidth(1, 200); // A列：ブランチ名
  sheet.setColumnWidth(2, 200); // B列：最終コミット日時
  sheet.setColumnWidth(3, 150); // C列：コミッター
  sheet.setColumnWidth(4, 300); // D列：Plugin Name
  sheet.setColumnWidth(5, 400); // E列：Description
  
  let page = 1;
  let allBranches = [];
  
  try {
    // ブランチ取得開始メッセージ
    sheet.getRange('A1').setValue('GitHub Plugin Branches (ブランチ一覧取得中...)');
    SpreadsheetApp.flush();

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
    
    // データ収集の一括処理用の配列
    const promises = allBranches.map((branch, index) => {
      return {
        branch: branch,
        indexPhpUrl: `https://api.github.com/repos/steamships/n2-plugins/contents/index.php?ref=${encodeURIComponent(branch.name)}`,
        commitUrl: branch.commit.url,
        index: index
      };
    });

    // 10件ずつの一括処理用にチャンク分割
    const chunkSize = 10;
    const chunks = [];
    for (let i = 0; i < promises.length; i += chunkSize) {
      chunks.push(promises.slice(i, i + chunkSize));
    }

    let allData = [];
    
    // チャンクごとに一括処理
    for (let i = 0; i < chunks.length; i++) {
      const chunk = chunks[i];
      const processedCount = i * chunkSize;
      
      // 進捗状況を表示
      sheet.getRange('A1').setValue(`GitHub Plugin Branches (データ収集中... ${processedCount}/${allBranches.length})`);
      SpreadsheetApp.flush();

      // 並列でリクエストを実行
      const responses = chunk.map(item => {
        const options = {
          headers: {
            Authorization: `token ${token}`,
            Accept: "application/vnd.github.v3.raw"
          },
          muteHttpExceptions: true
        };

        return {
          index: item.index,
          branch: item.branch,
          indexPhp: UrlFetchApp.fetch(item.indexPhpUrl, options),
          commit: UrlFetchApp.fetch(item.commitUrl, {
            ...options,
            headers: {
              ...options.headers,
              Accept: "application/vnd.github.v3+json"
            }
          })
        };
      });

      // レスポンスを処理
      for (const response of responses) {
        const branch = response.branch;
        let pluginName = '-';
        let description = '-';

        // index.phpの処理
        if (response.indexPhp.getResponseCode() === 200) {
          const content = response.indexPhp.getContentText();
          const pluginNameMatch = content.match(/Plugin Name:\s*(.+)/);
          const descriptionMatch = content.match(/Description:\s*(.+)/);
          
          if (pluginNameMatch) pluginName = pluginNameMatch[1].trim();
          if (descriptionMatch) description = descriptionMatch[1].trim();
        }

        // コミット情報の処理
        const commitData = JSON.parse(response.commit.getContentText());
        const commitDate = new Date(commitData.commit.author.date);
        const jstDate = Utilities.formatDate(commitDate, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

        // 分類を判定
        let category = 'その他';
        if (description !== '-') {
          if (/[都道府県市町村]/.test(description)) {
            category = '自治体用';
          } else if (description.includes('ポータル拡張')) {
            category = 'ポータル用';
          }
        }

        // データを配列に追加
        allData.push({
          branch: branch.name,
          branchUrl: `https://github.com/steamships/n2-plugins/tree/${encodeURIComponent(branch.name)}`,
          date: jstDate,
          committer: commitData.commit.author.name,
          pluginName: pluginName,
          description: description,
          category: category
        });
      }

      // API制限を考慮して少し待機
      Utilities.sleep(1000);
    }

    // ソート前の状態表示
    sheet.getRange('A1').setValue('GitHub Plugin Branches (データ整理中...)');
    SpreadsheetApp.flush();

    // カテゴリの優先順位を定義
    const categoryOrder = {
      'ポータル用': 1,
      '自治体用': 2,
      'その他': 3
    };

    // ソート
    allData.sort((a, b) => categoryOrder[a.category] - categoryOrder[b.category]);

    // 書き込みフェーズ
    sheet.getRange('A1').setValue('GitHub Plugin Branches (データ書き込み中...)');
    SpreadsheetApp.flush();
    
    let currentRow = 3;
    for (const data of allData) {
      // リッチテキストでリンクを設定
      const richTextValue = SpreadsheetApp.newRichTextValue()
        .setText(data.branch)
        .setLinkUrl(data.branchUrl)
        .build();

      // データを設定
      const range = sheet.getRange(currentRow, 1, 1, 6);
      range.setValues([[
        data.branch, // A列のテキストは後でリッチテキストで上書き
        data.date,
        data.committer,
        data.pluginName,
        data.description,
        data.category
      ]]);

      // ブランチ名をリッチテキストで設定
      sheet.getRange(currentRow, 1).setRichTextValue(richTextValue);

      // カテゴリに応じて背景色を設定
      const backgroundColor = data.category === 'ポータル用' ? '#e6e6fa' :  // 薄い紫 (Lavender)
                            data.category === '自治体用' ? '#fff2cc' :      // 薄い黄色 (指定色)
                            '#e6f3ff';                                      // 薄い青
      range.setBackground(backgroundColor);

      currentRow++;
    }

    // 最終更新
    sheet.getRange('A1').setValue('GitHub Plugin Branches');
    SpreadsheetApp.flush();
    
    // 罫線を追加（6列に拡張）
    const lastRow = sheet.getLastRow();
    if (lastRow > 2) {
      sheet.getRange(2, 1, lastRow - 1, 6).setBorder(true, true, true, true, true, true);
    }
    
  } catch (error) {
    sheet.getRange('A1').setValue('GitHub Plugin Branches (エラーが発生しました)');
    SpreadsheetApp.flush();
    Logger.log(`エラーが発生しました: ${error.message}`);
    throw new Error(`ブランチ情報の取得に失敗しました: ${error.message}`);
  }
} 