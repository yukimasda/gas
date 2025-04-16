// Google Apps Script では .env ファイルは使用できませんが、
// スクリプトプロパティを使用してトークンを安全に管理できます
const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
// スクリプトエディタで以下のように設定してください:
// 1. [ファイル] > [プロジェクトのプロパティ] を開く
// 2. [スクリプトのプロパティ] タブを選択
// 3. [行を追加] をクリック
// 4. プロパティ名を 'GITHUB_TOKEN' に、値をGitHubトークンに設定
const repo = "steamships/neo-neng";

function gitHub_Search() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GitHub Search');
  if (!sheet) throw new Error('「GitHub Search」シートが見つかりません');
  
  // GitHubトークンが設定されているか確認
  if (!token) {
    // msgBoxだけでエラーを表示
    Browser.msgBox("GitHub トークンが設定されていません", 
                   "スクリプトエディタで [ファイル] > [プロジェクトのプロパティ] > [スクリプトのプロパティ] から 'GITHUB_TOKEN' を設定するか、\n「GitHub」メニューから「APIトークンを設定」を選択してください。", 
                   Browser.Buttons.OK);
    
    // 処理を終了
    return;
  }

  // キーワード一覧（A列の3行目以降）
  const keywords = sheet.getRange("A3:A").getValues().flat().filter(k => k);
  if (keywords.length === 0) return;
  
  // パス一覧（B列の3行目以降）- すべてのパスを使用
  const paths = sheet.getRange("B3:B").getValues().flat().filter(p => p).map(p => p.trim());
  
  // ヘッダにパス検索説明を追加
  sheet.getRange("B2").setValue("検索パス（複数指定可）").setFontWeight("bold");
  
  // 検索開始を通知
  sheet.getRange("B1").setValue("検索準備中...").setFontWeight("bold").setHorizontalAlignment("left");
  SpreadsheetApp.flush();

  const maxPages = 10;
  let allResults = new Map(); // ファイルパスをキーとして結果を保存
  
  // 各キーワードごとに個別に検索を実行
  for (let keywordIndex = 0; keywordIndex < keywords.length; keywordIndex++) {
    const keyword = keywords[keywordIndex];
    
    // キーワード検索の進捗を表示
    const keywordProgress = `キーワード検索中: "${keyword}" (${keywordIndex + 1}/${keywords.length})`;
    sheet.getRange("B1").setValue(keywordProgress).setFontWeight("bold").setHorizontalAlignment("left");
    SpreadsheetApp.flush();
    
    // 検索条件を構築 - 単一キーワード検索
    let pathQueries = [];
    if (paths.length > 0) {
      // パスが指定されている場合、各パスと組み合わせて検索
      for (const path of paths) {
        pathQueries.push(`"${keyword}" path:${path}`);
      }
    } else {
      // パスが指定されていない場合は単純にキーワードのみ
      pathQueries.push(`"${keyword}"`);
    }
    
    // 各パスクエリに対して検索を実行
    for (let queryIndex = 0; queryIndex < pathQueries.length; queryIndex++) {
      const pathQuery = pathQueries[queryIndex];
      
      // クエリの進捗表示
      if (paths.length > 0) {
        const pathInfo = paths[queryIndex] || "不明";
        const queryProgress = `"${keyword}" でパス "${pathInfo}" を検索中 (${queryIndex + 1}/${pathQueries.length})`;
        sheet.getRange("B1").setValue(queryProgress).setFontWeight("bold").setHorizontalAlignment("left");
        SpreadsheetApp.flush();
      }
      
      // リポジトリ指定を追加
      const finalQuery = `${pathQuery} repo:${repo}`;
      
      // 生のGitHub検索クエリをログに表示（デバッグ用）
      Logger.log(`検索クエリ: ${finalQuery}`);
      
      // APIリクエスト用にはクエリをエンコード
      const encodedQuery = encodeURIComponent(finalQuery);
      
      // ステップ1: 検索APIを使ってマッチするファイルを見つける
      for (let page = 1; page <= maxPages; page++) {
        const url = `https://api.github.com/search/code?q=${encodedQuery}&per_page=100&page=${page}`;
        const options = {
          headers: {
            Authorization: `token ${token}`,
            Accept: "application/vnd.github.text-match+json"
          },
          muteHttpExceptions: true,
        };

        try {
          const response = UrlFetchApp.fetch(url, options);
          const data = JSON.parse(response.getContentText());
          
          // レート制限などの問題があれば記録
          if (data.message) {
            Logger.log(`GitHub APIエラー: ${data.message}`);
            if (page === 1) {
              // 最初のページでエラーが発生した場合はユーザーに通知
              Browser.msgBox("GitHub API エラー", `検索中にエラーが発生しました: ${data.message}`, Browser.Buttons.OK);
              return;
            }
            break;
          }

          // APIレスポンスからファイル名だけを抽出してログに表示
          if (data.items && data.items.length > 0) {
            const fileNames = data.items.map(item => item.path);
            Logger.log(`ページ ${page} の結果: ${data.items.length} 件のファイルがヒット`);
            Logger.log(`ヒットしたファイル例: ${fileNames.slice(0, 5).join(", ")}${fileNames.length > 5 ? '...(他 ' + (fileNames.length - 5) + ' 件)' : ''}`);
          } else {
            Logger.log(`ページ ${page}: 検索結果なし`);
          }

          if (!data.items || data.items.length === 0) break;

          // 検索結果を保存（重複を排除するためにマップを使用）
          data.items.forEach(item => {
            if (!allResults.has(item.path)) {
              // 新しいファイルの場合はマップに追加
              item.keywords = [keyword]; // このファイルがヒットしたキーワードを記録
              allResults.set(item.path, item);
            } else {
              // 既存のファイルの場合はキーワードを追加
              const existingItem = allResults.get(item.path);
              if (!existingItem.keywords.includes(keyword)) {
                existingItem.keywords.push(keyword);
              }
            }
          });
          
          if (data.items.length < 100) break;
        } catch (error) {
          Logger.log(`API呼び出しエラー (ページ ${page}): ${error.message}`);
          break;
        }
      }
    }
  }
  
  // マップから配列に変換
  const uniqueResults = Array.from(allResults.values());
  
  // 検索結果が多すぎる場合は警告
  if (uniqueResults.length > 100) {
    const message = `検索結果が多すぎます (${uniqueResults.length} ファイル)。検索条件を絞り込んでください。処理を中止します。`;
    Browser.msgBox("検索結果多すぎエラー", message, Browser.Buttons.OK);
    Logger.log(`処理中止: ${message}`);
    
    // エラーメッセージをB1セルに表示
    sheet.getRange("B1").setValue(`❌ ${message}`).setFontWeight("bold").setHorizontalAlignment("left");
    return;
  }
  
  const filesToProcess = uniqueResults;

  // 検索結果がない場合
  if (filesToProcess.length === 0) {
    sheet.getRange("B1").setValue("検索結果: 0件").setFontWeight("bold").setHorizontalAlignment("left");
    return;
  }

  // 結果表示の準備
  sheet.getRange(3, 3, sheet.getLastRow() - 2, 6).clearContent(); // 古い結果をクリア
  
  let row = 3;
  let totalMatches = 0;
  let processedCount = 0;
  
  // ステップ2: 各ファイルの内容を取得して詳細に検索
  filesToProcess.forEach(item => {
    processedCount++;
    
    // 進捗表示を更新（B1セルに表示、左揃え）
    const progressMessage = `処理中... (${processedCount}/${filesToProcess.length} ファイル)`;
    sheet.getRange("B1").setValue(progressMessage).setHorizontalAlignment("left");
    SpreadsheetApp.flush();
    
    const filePath = item.path;
    const fileUrl = item.html_url;
    const fileType = item.path.split('.').pop().toLowerCase();
    
    // このファイルがヒットしたキーワード（複数の場合あり）
    const matchedKeywords = item.keywords || [];
    
    // テキストファイルのみ処理（バイナリファイルはスキップ）
    const textFileExtensions = ['js', 'jsx', 'ts', 'tsx', 'vue', 'json', 'html', 'css', 'scss', 'less', 'md', 'txt', 'xml', 'yml', 'yaml', 'sh', 'py', 'rb', 'php', 'java', 'c', 'cpp', 'h', 'cs'];
    if (!textFileExtensions.includes(fileType)) {
      Logger.log(`スキップ: ${filePath} (テキストファイルではない可能性があります)`);
      return;
    }
    
    try {
      // ファイルの内容を取得
      const contentUrl = `https://api.github.com/repos/${repo}/contents/${filePath}`;
      const options = {
        headers: {
          Authorization: `token ${token}`,
        },
        muteHttpExceptions: true,
      };
      
      const contentResponse = UrlFetchApp.fetch(contentUrl, options);
      const contentData = JSON.parse(contentResponse.getContentText());
      
      if (contentData.message) {
        Logger.log(`ファイル取得エラー (${filePath}): ${contentData.message}`);
        return;
      }
      
      // Base64エンコードされたコンテンツをデコード
      const decodedContent = Utilities.base64Decode(contentData.content);
      let fileContent = '';
      try {
        fileContent = Utilities.newBlob(decodedContent).getDataAsString();
      } catch (e) {
        Logger.log(`文字コード変換エラー (${filePath}): ${e.message}`);
        return;
      }
      
      // 改行でファイルを分割
      const lines = fileContent.split('\n');
      
      // 各キーワードについて検索
      for (const keyword of keywords) {
        // このキーワードがこのファイルに関連していない場合はスキップ
        if (!matchedKeywords.includes(keyword)) continue;
        
        // 全ての行をチェック
        for (let i = 0; i < lines.length; i++) {
          const line = lines[i];
          
          if (line.includes(keyword)) {
            // 前後のコンテキスト（最大2行）を取得
            const startLine = Math.max(0, i - 2);
            const endLine = Math.min(lines.length - 1, i + 2);
            const contextLines = lines.slice(startLine, endLine + 1);
            const snippet = contextLines.join('\n');
            
            // 結果を表に追加
            sheet.getRange(row, 3).setValue(keyword);       // C: ヒットワード
            sheet.getRange(row, 4).setValue(filePath);      // D: ファイルパス
            sheet.getRange(row, 5).setFormula(`=HYPERLINK("${fileUrl}#L${i+1}", "L${i+1}")`); // E: 行番号付きリンク
            sheet.getRange(row, 6).setValue(snippet);      // F: コードスニペット
            
            row++;
            totalMatches++;
            
            // 1000行を超える場合は処理を中断（シート制限の考慮）
            if (totalMatches >= 1000) {
              Logger.log('警告: 最大表示行数（1000行）に達しました。');
              break;
            }
          }
        }
        
        // 1000行を超える場合は処理を中断（シート制限の考慮）
        if (totalMatches >= 1000) {
          break;
        }
      }
      
    } catch (error) {
      Logger.log(`ファイル処理エラー (${filePath}): ${error.message}`);
    }
  });
  
  // 最終的なマッチ数を更新（B1セルに表示、左揃え）
  const finalMessage = `ヒット数: ${totalMatches} マッチ (${filesToProcess.length} ファイル内)`;
  sheet.getRange("B1").setValue(finalMessage).setFontWeight("bold").setHorizontalAlignment("left");
  
  Logger.log(`検索結果: 合計 ${totalMatches} 箇所のマッチを表示しました (${filesToProcess.length} ファイル内)`);
}

