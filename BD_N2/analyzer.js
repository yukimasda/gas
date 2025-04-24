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

    // OpenAI APIを呼び出すための設定
    const systemPrompt = `あなたはYAMLファイルを解析して仕様書を作成する専門家です。
与えられたYAMLの構造を理解し、各フィールドの詳細を抽出してください。
出力は表形式で、各行を「|||」で区切り、各列を「###」で区切って返してください。`;

    const userPrompt = `以下のYAMLファイルを解析し、フォーム仕様を作成してください。

解析のポイント：
1. インデントから階層構造を理解する
2. 各フィールドのtype属性を確認する
3. v-if属性から表示条件を抽出する
4. value属性からデフォルト値を取得する
5. スタイル指定やイベントハンドラも考慮する

出力形式：
カテゴリ###項目名###タイプ###初期値###必須###表示条件|||
（以下、データ行）

解析対象のYAML:
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
      const rows = content.split('|||').filter(row => row.trim());

      // 既存のデータをクリア（5行目以降）
      const lastRow = sheet.getLastRow();
      if (lastRow > 4) {
        sheet.getRange(5, 2, lastRow - 4, 6).clearContent();
      }

      // データを配列に変換
      const values = rows.map(row => {
        const columns = row.split('###').map(col => col.trim());
        return [
          columns[0] || "未定義",  // カテゴリ
          columns[1] || "未定義",  // 項目名
          columns[2] || "未定義",  // タイプ
          columns[3] || "未定義",  // 初期値
          columns[4] || "任意",    // 必須
          columns[5] || "常時表示" // 表示条件
        ];
      });

      // バッチ処理で書き込み（5行目から開始）
      if (values.length > 0) {
        sheet.getRange(5, 2, values.length, 6).setValues(values);
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
