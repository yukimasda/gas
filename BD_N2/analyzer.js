// グローバル変数の定義（apiKeyのみ）
const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');

async function analyzeSourceWithAI() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // 解析開始時にステータスを表示
  sheet.getRange("B2").setValue("GPT解析中...");
  
  // トークンのチェック
  if (!token || !apiKey) {
    sheet.getRange("B5").setValue("⚠️ GitHubトークンまたはOpenAI APIキーが設定されていません");
    sheet.getRange("B2").setValue(""); // エラー時はB2をクリア
    return;
  }

  const sourcePath = sheet.getRange("A2").getValue();
  if (!sourcePath) {
    sheet.getRange("B5").setValue("⚠️ 解析ソースが指定されていません");
    return;
  }

  // repoからowner, repoNameを取得
  const [owner, repoName] = repo.split('/');
  if (!owner || !repoName) {
    sheet.getRange("B5").setValue("⚠️ リポジトリの形式が正しくありません");
    return;
  }

  try {
    // GitHubからソースコード取得
    const contentUrl = `https://api.github.com/repos/${repo}/contents/${sourcePath}`;
    
    // GitHubのWebUIのURL
    const githubWebUrl = `https://github.com/${repo}/blob/main/${sourcePath}`;

    const contentResponse = await UrlFetchApp.fetch(contentUrl, {
      headers: {
        'Authorization': `token ${token}`,
        'Accept': 'application/vnd.github.v3+json',
        'User-Agent': 'Google Apps Script'
      },
      muteHttpExceptions: true
    });

    const responseCode = contentResponse.getResponseCode();
    if (responseCode !== 200) {
      const errorData = JSON.parse(contentResponse.getContentText());
      sheet.getRange("B2").setValue(""); // エラー時はB2をクリア
      throw new Error(`GitHub API Error (${responseCode}): ${errorData.message}`);
    }

    const contentData = JSON.parse(contentResponse.getContentText());
    const sourceCode = Utilities.newBlob(Utilities.base64Decode(contentData.content)).getDataAsString();

    // A2セルにGitHubリンクを設定
    const richText = SpreadsheetApp.newRichTextValue()
      .setText(sourcePath)
      .setLinkUrl(githubWebUrl)
      .build();
    sheet.getRange("A2").setRichTextValue(richText);

    // ファイルタイプを判定する補助関数
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

    // ヘッダー行を取得して検証
    const headerRange = sheet.getRange(5, 2, 1, sheet.getLastColumn() - 1);
    const headers = headerRange.getValues()[0].filter(header => header !== '');

    if (headers.length === 0) {
      throw new Error('5行目にヘッダーが設定されていません。');
    }

    // システムプロンプト
    const systemPrompt = `あなたはソースコードを解析して仕様書を作成する専門家です。
    ファイルの種類に応じて適切な解析を行い、指定された項目の情報を抽出してください。

    解析の基本方針：
    1. ファイルの構造を上から順に解析
    2. 指定された項目に関する情報を優先的に抽出
    3. 階層構造や依存関係を考慮
    4. コメントや関連情報も参考に

    出力形式：
    - 行区切り: |||
    - 列区切り: ###
    - 必ず上から順に出力`;

    // ユーザープロンプト
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
    4. コードをわかりやすい日本語かつ簡潔に表現してください,
    5. 数式は避ける。


    出力形式：
    ${headers.join('###')}|||
    （ファイルの構造に従って上から順にデータを出力）

    解析対象のソースコード:
    ${sourceCode}`;

    // OpenAI APIを呼び出し
    const openaiResponse = await UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${apiKey}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        model: "gpt-4o",
        messages: [
          {
            role: "system",
            content: systemPrompt
          },
          {
            role: "user",
            content: userPrompt
          }
        ],
        temperature: 0
      }),
      muteHttpExceptions: true
    });

    try {
      const response = JSON.parse(openaiResponse.getContentText());
      
      if (response.error) {
        Logger.log(`API Error: ${response.error.message}`);
        throw new Error(`OpenAI API Error: ${response.error.message}`);
      }

      const content = response.choices[0].message.content;
      // ヘッダー行を含まない形でデータ行のみを取得
      const rows = content.split('|||')
        .filter(row => row.trim())  // 空行を除外
        .filter(row => !row.includes(headers.join('###'))); // ヘッダー行を除外

      // データを配列に変換
      const values = rows.map(row => {
        const columns = row.split('###').map(col => col.trim());
        return headers.map((header, index) => {
          return columns[index] || "未定義";
        });
      });

      // 既存のデータをクリア（6行目以降）
      const lastRow = sheet.getLastRow();
      if (lastRow > 5) {
        sheet.getRange(6, 2, lastRow - 5, headers.length).clearContent();  // headersの長さを使用
      }

      // バッチ処理で書き込み（6行目から開始）
      if (values.length > 0) {
        sheet.getRange(6, 2, values.length, headers.length).setValues(values);  // headersの長さを使用
      }

      sheet.getRange("B2").setValue("解析完了");

    } catch (error) {
      Logger.log(`Error: ${error.message}`);
      Logger.log(`Response: ${openaiResponse.getContentText()}`);
      sheet.getRange("B2").setValue("");
      throw new Error(`解析エラー: ${error.message}`);
    }
  } catch (error) {
    Logger.log(`エラー発生: ${error.message}`);
    sheet.getRange("B5").setValue(`⚠️ ${error.message}`);
    throw error;
  }
}
