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

  // ヘッダーの設定
  const headers = ["カテゴリ", "項目名", "タイプ", "初期値", "必須", "表示条件"];
  for (let i = 0; i < headers.length; i++) {
    sheet.getRange(5, i + 2).setValue(headers[i]);
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

    // ソースコードを4000文字に制限
    const truncatedCode = sourceCode.length > 4000 
      ? sourceCode.substring(0, 4000) + "\n... (長さの制限のため省略されました)"
      : sourceCode;

    // OpenAIのプロンプト設定
    const promptInstruction = `以下のソースコードを解析して、機能仕様を作成してください。
出力は必ず以下のJSON形式で返してください：

{
  "specs": [
    {
      "カテゴリ": "機能のカテゴリ（例：認証、データ処理）",
      "項目名": "具体的な機能名",
      "タイプ": "入力形式（テキスト、ラジオボタン、チェックボックスなど）",
      "初期値": "デフォルト値（未設定の場合は「未」）",
      "必須": "必須/任意/なし",
      "表示条件": "項目が表示される条件（条件がない場合は「常時表示」）"
    }
  ]
}`;

    const prompt = `${promptInstruction}\n\nソースコード:\n${truncatedCode}`;

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
            content: "あなたは優秀なプログラマーで、ソースコードを解析して機能仕様を作成することができます。出力は必ずJSON形式で行い、各項目を詳細に分析します。"
          },
          {role: "user", content: prompt}
        ],
        temperature: 0,
        max_tokens: 4096,
        response_format: { type: "json_object" }
      }),
      muteHttpExceptions: true
    });

    const aiResponse = JSON.parse(openaiResponse.getContentText());
    if (aiResponse.error) {
      throw new Error(`OpenAI API Error: ${aiResponse.error.message}`);
    }

    // JSONをパースして表形式で出力
    const specifications = JSON.parse(aiResponse.choices[0].message.content).specs;
    specifications.forEach((spec, index) => {
      const row = index + 6; // ヘッダーの次の行から開始
      sheet.getRange(row, 2).setValue(spec.カテゴリ);
      sheet.getRange(row, 3).setValue(spec.項目名);
      sheet.getRange(row, 4).setValue(spec.タイプ);
      sheet.getRange(row, 5).setValue(spec.初期値);
      sheet.getRange(row, 6).setValue(spec.必須);
      sheet.getRange(row, 7).setValue(spec.表示条件);
    });

    // 解析完了後、プロンプトをB2に表示
    sheet.getRange("B2").setValue(promptInstruction);

  } catch (error) {
    Logger.log(`エラー発生: ${error.message}`);
    sheet.getRange("B2").setValue(""); // エラー時はB2をクリア
    sheet.getRange("B5").setValue(`⚠️ エラーが発生しました: ${error.message}`);
  }
}
