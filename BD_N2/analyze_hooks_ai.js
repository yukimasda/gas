async function analyzeHooksWithAI() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Hook List');
  if (!sheet) {
    Browser.msgBox("エラー", "「Hook List」シートが見つかりません。先にフックを検索してください。", Browser.Buttons.OK);
    return;
  }

  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) {
    Browser.msgBox(
      "設定エラー",
      "OpenAI APIキーが設定されていません。\n\n" +
      "1. スクリプトエディタを開く\n" +
      "2. プロジェクトの設定を開く\n" +
      "3. スクリプトプロパティに'OPENAI_API_KEY'を追加\n" +
      "4. APIキーを入力して保存",
      Browser.Buttons.OK
    );
    return;
  }

  // GitHubの設定を取得
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) {
    Browser.msgBox("エラー", "GitHub トークンが設定されていません。先にGitHubトークンを設定してください。", Browser.Buttons.OK);
    return;
  }

  const [owner, repoName] = repo.split('/');
  if (!owner || !repoName) {
    Browser.msgBox("エラー", "リポジトリの形式が正しくありません。'オーナー名/リポジトリ名'の形式で指定してください。", Browser.Buttons.OK);
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    Browser.msgBox("エラー", "フックが見つかりません。先にフックを検索してください。", Browser.Buttons.OK);
    return;
  }

  // 分析結果列（G-J列）をクリア
  if (lastRow > 1) {
    sheet.getRange(2, 7, lastRow - 1, 4).clearContent();
  }

  // ヘッダー行を設定
  sheet.getRange("G1").setValue("AIによる分析中...");
  SpreadsheetApp.flush();

  try {
    const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    let analyzed = 0;

    for (const row of data) {
      const [filePath, className, hookName, callback, type] = row;
      
      // GitHubからソースコード取得
      const content = await fetchFileContent(owner, repoName, filePath);
      
      // 関連コードの抽出
      const relevantCode = extractRelevantCode(content, callback, hookName);
      
      // AIによる分析
      const response = await predictHookRole(hookName, type, { path: filePath }, owner, repoName, callback, relevantCode);
      
      // 分析結果を書き込み
      const currentRow = analyzed + 2;
      sheet.getRange(currentRow, 7).setValue(response);
      
      analyzed++;
      if (analyzed % 5 === 0) {
        sheet.getRange("G1").setValue(`分析中... ${analyzed}/${lastRow - 1}`);
        SpreadsheetApp.flush();
      }
      
      // API制限を考慮して待機
      await Utilities.sleep(1000);
    }

    sheet.getRange("G1").setValue(`分析完了: ${analyzed}件のフックを分析しました`);
    
  } catch (error) {
    Logger.log(`AI分析エラー: ${error}`);
    sheet.getRange("G1").setValue(`エラーが発生しました: ${error}`);
  }
}

async function predictHookRole(hookName, type, file, owner, repoName, callback, relevantCode) {
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
    if (!apiKey) {
      return type === 'filter' ? 'データ加工フィルター' : 'アクション実行';
    }

    const prompt = `
WordPressのフックの役割を分析してください：

フック種別: ${type}
フック名: ${hookName}
コールバック関数: ${callback}

関連コード:
${relevantCode}

分析ポイント（各項目50文字以内、合計200文字以内）：
1. このフックが何を実現しているのか
2. どのようなデータを処理しているのか
3. どのような条件で実行されるのか
4. WordPressのコアのフックとの関連性
`;

    const response = await callChatGPT(prompt, apiKey);
    return response.trim();
  } catch (error) {
    Logger.log(`役割予測エラー: ${error}`);
    return type === 'filter' ? 'データ加工フィルター' : 'アクション実行';
  }
}

async function callChatGPT(prompt, apiKey) {
  const url = 'https://api.openai.com/v1/chat/completions';
  const options = {
    method: 'post',
    headers: {
      'Authorization': `Bearer ${apiKey}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify({
      model: "gpt-3.5-turbo",
      messages: [
        {
          role: "system",
          content: "あなたはWordPressの開発者アシスタントです。フックの役割を簡潔に説明してください。各項目は50文字以内、合計で200文字以内でまとめてください。"
        },
        {
          role: "user",
          content: prompt
        }
      ],
      temperature: 0.3,
      max_tokens: 500,  // 約200文字の日本語に対応
      response_format: { "type": "text" }
    })
  };

  const response = await UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response.getContentText());
  return json.choices[0].message.content;
}

function extractRelevantCode(content, callback, hookName) {
  const lines = content.split('\n');
  let relevantCode = '';
  
  // コールバック関数の定義を探す
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (line.includes(`function ${callback}`) || line.includes(`${callback} =`)) {
      // 関数定義の前後10行を取得
      const start = Math.max(0, i - 5);
      const end = Math.min(lines.length, i + 15);
      relevantCode = lines.slice(start, end).join('\n');
      break;
    }
  }
  
  // フック登録部分も含める
  const hookRegistration = lines.find(line => 
    line.includes(hookName) && (line.includes('add_action') || line.includes('add_filter'))
  );
  
  return `// フック登録
${hookRegistration || '// 登録コード不明'}

// コールバック関数
${relevantCode || '// 関数定義不明'}`;
} 