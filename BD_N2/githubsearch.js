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
  sheet.getRange("B1").setValue("例: src/ または src/components/").setFontStyle("italic");

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
  sheet.getRange("C1").setValue(`ヒット数: ${allResults.length}`).setFontWeight("bold").setHorizontalAlignment("right");
  sheet.getRange("A2").setValue("検索キーワード").setFontWeight("bold");
  sheet.getRange("C2").setValue("ヒットワード").setFontWeight("bold");
  sheet.getRange("D2").setValue("ファイルパス").setFontWeight("bold");
  sheet.getRange("E2").setValue("リンク").setFontWeight("bold");
  sheet.getRange("F2").setValue("コードスニペット").setFontWeight("bold");

  let row = 3;
  allResults.forEach(item => {
    const path = item.path;
    const link = item.html_url;
    const matches = item.text_matches || [];
    let matchedWord = 'N/A';
    let snippet = '';

    for (let match of matches) {
      const fragment = match.fragment || '';
      snippet = fragment;
      const found = keywords.find(kw => fragment.includes(kw));
      if (found) {
        matchedWord = found;
        break;
      }
    }

    sheet.getRange(row, 3).setValue(matchedWord); // C: ヒットワード
    sheet.getRange(row, 4).setValue(path);        // D: ファイルパス
    sheet.getRange(row, 5).setFormula(`=HYPERLINK("${link}", "Link")`); // E: リンク
    sheet.getRange(row, 6).setValue(snippet);     // F: コードスニペット
    row++;
  });
}

