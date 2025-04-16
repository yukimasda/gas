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

  // キーワード一覧（A列の3行目以降）
  const keywords = sheet.getRange("A3:A").getValues().flat().filter(k => k);
  if (keywords.length === 0) return;

  // キーワードをORで結んだ検索クエリを作成
  const query = keywords.map(k => `"${k}"`).join(" OR ");
  const encodedQuery = encodeURIComponent(`${query} repo:${repo}`);
  
  // 実際のURLをC1セルに表示
  const apiUrl = `https://api.github.com/search/code?q=${encodedQuery}&per_page=100&page=1`; // 1ページ目のみ
  sheet.getRange("C1").setValue(`検索URL: ${apiUrl}`).setFontWeight("bold");

  const maxPages = 10;
  let allResults = [];

  // 最初の検索クエリをログに表示
  Logger.log(`検索クエリ: ${apiUrl}`);

  for (let page = 1; page <= maxPages; page++) {
    const url = `https://api.github.com/search/code?q=${encodedQuery}&per_page=100&page=${page}`;
    const options = {
      headers: {
        Authorization: `token ${token}`,
        Accept: "application/vnd.github.text-match+json"
      },
      muteHttpExceptions: true,
    };

    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());

    // APIレスポンスからファイル名だけを抽出してログに表示
    if (data.items && data.items.length > 0) {
      const fileNames = data.items.map(item => item.path);
      Logger.log(`APIレスポンス（ページ ${page}）: ${fileNames.join(", ")}`);
    }

    if (!data.items || data.items.length === 0) break;

    allResults = allResults.concat(data.items);
    if (data.items.length < 100) break;
  }

  // ヒット数をログに表示
  Logger.log(`検索結果のヒット数: ${allResults.length}`);

  // 表クリア（B～E列の3行目以降）
  sheet.getRange("B3:E" + sheet.getLastRow()).clearContent();

  // ヘッダとタイトル
  sheet.getRange("A1").setValue("GitHub Search").setFontWeight("bold");
  sheet.getRange("B1").setValue(`ヒット数: ${allResults.length}`).setFontWeight("bold").setHorizontalAlignment("right");
  sheet.getRange("A2").setValue("検索キーワード").setFontWeight("bold");
  sheet.getRange("B2").setValue("ヒットワード").setFontWeight("bold");
  sheet.getRange("C2").setValue("ファイルパス").setFontWeight("bold");
  sheet.getRange("D2").setValue("リンク").setFontWeight("bold");
  sheet.getRange("E2").setValue("コードスニペット").setFontWeight("bold");

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

    sheet.getRange(row, 2).setValue(matchedWord); // B: ヒットワード
    sheet.getRange(row, 3).setValue(path);        // C: ファイルパス
    sheet.getRange(row, 4).setFormula(`=HYPERLINK("${link}", "Link")`); // D: リンク
    sheet.getRange(row, 5).setValue(snippet);     // E: コードスニペット
    row++;
  });
}
