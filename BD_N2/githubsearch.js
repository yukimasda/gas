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
  
  // パス一覧（B列の3行目以降）- 最初の1つだけを使用
  const paths = sheet.getRange("B3:B").getValues().flat().filter(p => p);
  const path = paths.length > 0 ? paths[0].trim() : '';
  
  // ヘッダにパス検索説明を追加
  sheet.getRange("B2").setValue("検索パス（オプション）").setFontWeight("bold");

  // キーワードをORで結んだ検索クエリを作成
  const keywordQuery = keywords.map(k => `"${k}"`).join(" OR ");
  
  // 検索クエリの構築 - シンプルに
  let query = '';
  
  if (path) {
    // パスが指定されている場合はpath:パラメータを追加
    query = `${keywordQuery} path:${path}`;
    Logger.log(`検索条件：キーワード [${keywords.join(", ")}] / パス [${path}]`);
  } else {
    // パスが指定されていない場合はキーワードのみ
    query = keywordQuery;
    Logger.log(`検索条件：キーワード [${keywords.join(", ")}] / パス指定なし（リポジトリ全体）`);
  }
  
  // リポジトリ指定を追加
  const finalQuery = `${query} repo:${repo}`;
  
  // 生のGitHub検索クエリをログに表示（デバッグ用）
  Logger.log(`最終GitHubクエリ: ${finalQuery}`);
  
  // APIリクエスト用にはクエリをエンコード
  const encodedQuery = encodeURIComponent(finalQuery);
  
  const maxPages = 10;
  let allResults = [];

  // 検索開始ログ
  Logger.log(`検索実行: リポジトリ [${repo}] で検索開始`);

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

      allResults = allResults.concat(data.items);
      if (data.items.length < 100) break;
    } catch (error) {
      Logger.log(`検索中にエラーが発生しました: ${error.message}`);
      if (page === 1) {
        // 最初のページでエラーが発生した場合はユーザーに通知
        Browser.msgBox("エラー", `検索中にエラーが発生しました: ${error.message}`, Browser.Buttons.OK);
        return;
      }
      break;
    }
  }

  // ヒット数をログに表示
  Logger.log(`検索完了: 合計 ${allResults.length} 件のファイルがヒットしました`);

  // 表クリア（C～F列の3行目以降）
  sheet.getRange("C3:F" + sheet.getLastRow()).clearContent();

  // ヘッダとタイトル
  sheet.getRange("A1").setValue("GitHub Search").setFontWeight("bold");
  // 進捗メッセージをB1セルに表示（D1は使用しない）
  sheet.getRange("B1").setValue("検索中...").setFontWeight("bold");
  sheet.getRange("A2").setValue("検索キーワード").setFontWeight("bold");
  sheet.getRange("C2").setValue("ヒットワード").setFontWeight("bold");
  sheet.getRange("D2").setValue("ファイルパス").setFontWeight("bold");
  sheet.getRange("E2").setValue("リンク").setFontWeight("bold");
  sheet.getRange("F2").setValue("コードスニペット").setFontWeight("bold");
  SpreadsheetApp.flush();  // 画面更新

  // ファイル処理の上限（レート制限対策）- 100ファイルに拡大
  const maxFilesToProcess = 100;
  const filesToProcess = allResults.slice(0, maxFilesToProcess);
  
  if (allResults.length > maxFilesToProcess) {
    Logger.log(`注意: ${allResults.length} ファイル中、最初の ${maxFilesToProcess} ファイルのみを処理します（レート制限対策）`);
  }

  let row = 3;
  let totalMatches = 0;
  let processedCount = 0;
  
  // ステップ2: 各ファイルの内容を取得して詳細に検索
  filesToProcess.forEach(item => {
    processedCount++;
    
    // 進捗表示を更新（B1セルに表示）
    const progressMessage = `処理中... (${processedCount}/${filesToProcess.length} ファイル)`;
    sheet.getRange("B1").setValue(progressMessage);
    SpreadsheetApp.flush();
    
    const filePath = item.path;
    const fileUrl = item.html_url;
    const fileType = item.path.split('.').pop().toLowerCase();
    
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
  
  // 最終的なマッチ数を更新（B1セルに表示）
  const finalMessage = `ヒット数: ${totalMatches} マッチ (${processedCount}/${allResults.length} ファイル内)`;
  sheet.getRange("B1").setValue(finalMessage).setFontWeight("bold");
  
  Logger.log(`検索結果: 合計 ${totalMatches} 箇所のマッチを表示しました (${processedCount}/${allResults.length} ファイル内)`);
}

