// グローバル変数の定義（apiKeyのみ）
const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');

/**
 * 既存の解析結果をクリア
 */
function clearExistingData(sheet, headers) {
  if (sheet.getLastRow() > 5) {
    sheet.getRange(6, 2, sheet.getLastRow() - 5, headers.length).clearContent();
  }
}

/**
 * GitHubからソースコードを取得
 */
async function fetchGitHubContent(sourcePath) {
  const contentUrl = `https://api.github.com/repos/${repo}/contents/${sourcePath}`;
  const githubWebUrl = `https://github.com/${repo}/blob/main/${sourcePath}`;
  
  const contentResponse = await UrlFetchApp.fetch(contentUrl, {
    headers: {
      'Authorization': `token ${token}`,
      'Accept': 'application/vnd.github.v3+json',
      'User-Agent': 'Google Apps Script'
    },
    muteHttpExceptions: true
  });

  if (contentResponse.getResponseCode() !== 200) {
    throw new Error(`GitHub API Error: ${JSON.parse(contentResponse.getContentText()).message}`);
  }

  const sourceCode = Utilities.newBlob(
    Utilities.base64Decode(JSON.parse(contentResponse.getContentText()).content)
  ).getDataAsString();

  return {
    sourceCode,
    githubWebUrl
  };
}

/**
 * OpenAI APIを呼び出して解析を実行
 */
async function callOpenAI(sourcePath, sourceCode, headers) {
  const systemPrompt = `あなたはソースコードを解析して仕様書を作成する専門家です。
  ファイルの種類に応じて適切な解析を行い、指定された項目の情報を抽出してください。

  解析の基本方針：
  1. ファイルの構造を上から順に解析
  2. 指定された項目に関する情報を優先的に抽出
  3. 階層構造や依存関係を考慮
  4. コメントや関連情報も参考に
  5. 数式やロジックは簡潔でわかりやすい日本語に変換
  
  数式の説明方針：
  - 複雑な計算式は「〜を計算」のような平易な表現に
  - 条件分岐は「〜の場合」のような形で説明
  - 技術的な用語は一般的な言葉に置き換える
  - 具体例を用いて説明

  出力形式：
  - 行区切り: |||
  - 列区切り: ###
  - 必ず上から順に出力
  - 専門用語は避け、誰でも理解できる表現を使用`;

  const userPrompt = `以下のファイルを解析し、仕様書を作成してください。

  ファイル情報：
  - パス: ${sourcePath}
  - 種類: ${getFileType(sourcePath)}

  抽出する項目と基準：
  ${headers.map(header => `${header}: ${header}に関する情報を抽出`).join('\n')}

  解析のポイント：
  1. ${getFileType(sourcePath)}の特徴を考慮した解析
  2. 上記の項目を優先的に抽出
  3. コードの文脈を理解し、適切な情報を抽出


  出力形式：
  ${headers.join('###')}|||
  （ファイルの構造に従って上から順にデータを出力）

  解析対象のソースコード:
  ${sourceCode}`;

  const response = await UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
    method: 'post',
    headers: {
      'Authorization': `Bearer ${apiKey}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify({
      model: "gpt-4",
      messages: [
        { role: "system", content: systemPrompt },
        { role: "user", content: userPrompt }
      ],
      temperature: 0
    }),
    muteHttpExceptions: true
  });

  const content = JSON.parse(response.getContentText()).choices[0].message.content;
  
  // 解析結果をパース
  return content.split('|||')
    .filter(row => row.trim() && !row.includes(headers.join('###')))
    .map(row => {
      const columns = row.split('###').map(col => col.trim());
      return headers.map((_, index) => columns[index] || "未定義");
    });
}

/**
 * ファイルタイプを判定する補助関数
 */
function getFileType(sourcePath) {
  const extension = sourcePath.split('.').pop().toLowerCase();
  const fileTypes = {
    'yml': 'YAML設定ファイル',
    'yaml': 'YAML設定ファイル',
    'vue': 'Vueコンポーネント',
    'js': 'JavaScriptファイル',
    'php': 'PHPファイル',
    'json': 'JSONファイル'
  };
  return fileTypes[extension] || '不明なファイル形式';
}

/**
 * A列に指定されたファイルを順次解析して仕様書を作成
 */
async function analyzeSourcesWithAI() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // トークンのチェック
  if (!token || !apiKey) {
    sheet.getRange("B2").setValue("");
    sheet.getRange("B5").setValue("⚠️ GitHubトークンまたはOpenAI APIキーが設定されていません");
    return;
  }

  // repoからowner, repoNameを取得
  const [owner, repoName] = repo.split('/');
  if (!owner || !repoName) {
    sheet.getRange("B5").setValue("⚠️ リポジトリの形式が正しくありません");
    return;
  }

  // A列のファイル一覧を取得
  const lastRow = sheet.getLastRow();
  const fileRange = sheet.getRange(2, 1, lastRow - 1, 1);
  const files = fileRange.getValues();
  
  // ヘッダー行を取得して検証
  const headerRange = sheet.getRange(5, 2, 1, sheet.getLastColumn() - 1);
  const headers = headerRange.getValues()[0].filter(header => header !== '');
  if (headers.length === 0) {
    sheet.getRange("B5").setValue("⚠️ 5行目にヘッダーが設定されていません");
    return;
  }

  // 既存のデータをクリア
  clearExistingData(sheet, headers);

  let currentRow = 6;  // 結果の書き込み開始行

  // ファイルごとの処理
  for (let i = 0; i < files.length; i++) {
    const sourcePath = files[i][0];
    if (!sourcePath) continue; // 空の行はスキップ

    try {
      // 進捗状況を更新
      sheet.getRange("B2").setValue(`GPT解析中... (${i + 1}/${files.length})`);

      // GitHubからソースコード取得
      const { sourceCode, githubWebUrl } = await fetchGitHubContent(sourcePath);

      // ファイル名をリンク付きで出力
      const richText = SpreadsheetApp.newRichTextValue()
        .setText(sourcePath)
        .setLinkUrl(githubWebUrl)
        .build();
      sheet.getRange(currentRow, 2).setRichTextValue(richText);
      currentRow++;

      // AIによる解析
      const aiResponse = await callOpenAI(sourcePath, sourceCode, headers);
      
      // 解析結果を書き込み
      if (aiResponse.length > 0) {
        sheet.getRange(currentRow, 2, aiResponse.length, headers.length).setValues(aiResponse);
        currentRow += aiResponse.length;
      }

      // 区切りの空行を追加
      sheet.getRange(currentRow, 2).setValue("");
      currentRow++;

      // API制限を考慮して待機
      await Utilities.sleep(2000);

    } catch (error) {
      Logger.log(`ファイル ${sourcePath} の解析中にエラー: ${error.message}`);
      sheet.getRange(currentRow, 2).setValue(`⚠️ エラー: ${error.message}`);
      currentRow += 2;
      continue;
    }
  }

  sheet.getRange("B2").setValue("全ファイルの解析完了");
}
