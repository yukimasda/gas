// このファイルでは token、repo、branch の定義は省略します。
// これらの変数は他のソースで定義されていると想定します。

async function predictHookRole(hookName, type, file, owner, repoName, callback) {
  try {
    // APIキーの存在チェック
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
      return type === 'filter' ? 'データ加工フィルター' : 'アクション実行';
    }

    // GitHubからファイルの内容を取得
    const content = await fetchFileContent(owner, repoName, file.path);
    
    // AIに送信するプロンプトを構築
    const prompt = `
以下のWordPressのフックとそのコンテキストから、フックの具体的な役割を100文字程度で説明してください：

フック種別: ${type}
フック名: ${hookName}
コールバック関数: ${callback}

ソースコード:
${content}

特に以下の点に注目して分析してください：
1. このフックが何を実現しているのか
2. どのようなデータを処理しているのか
3. どのような条件で実行されるのか
4. 他のフックやシステムとの関連性
`;

    // ChatGPT APIを呼び出し
    const response = await callChatGPT(prompt, apiKey);
    
    return response.trim();
  } catch (error) {
    Logger.log(`役割予測エラー: ${error}`);
    return type === 'filter' ? 'データ加工フィルター' : 'アクション実行';
  }
}

function callChatGPT(prompt, apiKey) {
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
          content: "あなたはWordPressの開発者アシスタントです。フックの役割を簡潔に説明してください。"
        },
        {
          role: "user",
          content: prompt
        }
      ],
      temperature: 0.3,
      max_tokens: 150
    })
  };

  const response = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response.getContentText());
  return json.choices[0].message.content;
}

// メイン関数を非同期に修正
async function fetchHooksFromGitHub() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Hook List');
  if (!sheet) {
    Browser.msgBox("エラー", "「Hook List」シートが見つかりません。シートを作成してください。", Browser.Buttons.OK);
    return;
  }

  const [owner, repoName] = repo.split('/');
  const folder = 'inc';
  
  sheet.clear();
  sheet.appendRow(['ファイル名', 'クラス名', 'フック名', 'コールバック関数名', '種別', '行番号', '推定される役割']);

  let totalHooks = 0;
  let rowBuffer = [];
  
  sheet.getRange("A1").setValue(`${folder} フォルダを検索中...`);
  SpreadsheetApp.flush();
  
  try {
    const files = listPhpFiles(owner, repoName, folder);
    
    for (const file of files) {
      const content = await fetchFileContent(owner, repoName, file.path);
      const lines = content.split('\n');
      let currentClass = '';
      
      for (let index = 0; index < lines.length; index++) {
        const line = lines[index];
        const classMatch = line.match(/class\s+(\w+)/);
        if (classMatch) {
          currentClass = classMatch[1];
        }
        
        if (line.includes('add_action(') || line.includes('add_filter(')) {
          const type = line.includes('add_action(') ? 'action' : 'filter';
          const fileUrl = `https://github.com/${repo}/blob/master/${file.path}`;
          const lineLink = `=HYPERLINK("${fileUrl}#L${index + 1}", "${index + 1}")`;
          
          let hookName = '';
          const hookMatch = line.match(/['"]([^'"]+)['"]/);
          if (hookMatch) {
            hookName = hookMatch[1];
          }

          let callback = '';
          const callbackMatch = line.match(/,\s*['"]([^'"]+)['"]/);
          if (callbackMatch) {
            callback = callbackMatch[1];
          }

          // AIによる役割の予測
          const role = await predictHookRole(hookName, type, file, owner, repoName, callback);
          
          rowBuffer.push([file.path, currentClass, hookName, callback, type, lineLink, role]);
          totalHooks++;
          
          if (rowBuffer.length >= 10) {
            const startRow = totalHooks - rowBuffer.length + 2;
            sheet.getRange(startRow, 1, rowBuffer.length, 7).setValues(rowBuffer);
            rowBuffer = [];
          }
        }
      }
    }
    
    // 残りのバッファを書き込み
    if (rowBuffer.length > 0) {
      const startRow = totalHooks - rowBuffer.length + 2;
      sheet.getRange(startRow, 1, rowBuffer.length, 7).setValues(rowBuffer);
    }
    
    sheet.getRange("A1").setValue(`検索完了: ${totalHooks}件のフックが見つかりました`);
    
    // 結果をソート
    if (totalHooks > 0) {
      const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7);
      dataRange.sort([{column: 1, ascending: true}, {column: 6, ascending: true}]);
    }
    
  } catch (error) {
    Logger.log(`検索中にエラー: ${error}`);
    sheet.getRange("A1").setValue(`エラーが発生しました: ${error}`);
  }
}

function listPhpFiles(owner, repoName, path) {
  const url = `https://api.github.com/repos/${owner}/${repoName}/contents/${path}`;
  const res = UrlFetchApp.fetch(url, {
    headers: { Authorization: `token ${token}` }
  });
  const json = JSON.parse(res.getContentText());
  let files = [];
  json.forEach(item => {
    if (item.type === 'file' && item.name.endsWith('.php')) {
      files.push(item);
    } else if (item.type === 'dir') {
      files = files.concat(listPhpFiles(owner, repoName, item.path));
    }
  });
  return files;
}

function fetchFileContent(owner, repoName, path) {
  const url = `https://api.github.com/repos/${owner}/${repoName}/contents/${path}`;
  const res = UrlFetchApp.fetch(url, {
    headers: { Authorization: `token ${token}` }
  });
  const json = JSON.parse(res.getContentText());
  return Utilities.newBlob(Utilities.base64Decode(json.content)).getDataAsString();
}

function analyzeHooksWithAI() {
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

  // 以下、既存の処理を続ける
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    Browser.msgBox("エラー", "フックが見つかりません。先にフックを検索してください。", Browser.Buttons.OK);
    return;
  }

  sheet.getRange("A1").setValue("AIによる分析中...");
  SpreadsheetApp.flush();

  try {
    const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    let analyzed = 0;

    data.forEach(function(row) {
      const [filePath, className, hookName, callback, type] = row;
      
      // GitHubからソースコード取得（ownerとrepoNameを使用）
      const content = fetchFileContent(owner, repoName, filePath);
      
      // AIによる分析
      const prompt = `
WordPressのフックの役割を分析してください：

フック種別: ${type}
フック名: ${hookName}
クラス名: ${className}
コールバック: ${callback}

ソースコード:
${content}

分析ポイント：
1. フックの主な目的
2. 処理内容
3. 実行タイミング
4. 関連する機能
`;

      const response = callChatGPT(prompt, apiKey);
      
      // 分析結果を書き込み
      const currentRow = analyzed + 2;
      sheet.getRange(currentRow, 7).setValue(response);
      
      analyzed++;
      if (analyzed % 5 === 0) {
        sheet.getRange("A1").setValue(`分析中... ${analyzed}/${lastRow - 1}`);
        SpreadsheetApp.flush();
      }
    });

    sheet.getRange("A1").setValue(`分析完了: ${analyzed}件のフックを分析しました`);
    
  } catch (error) {
    Logger.log(`AI分析エラー: ${error}`);
    sheet.getRange("A1").setValue(`エラーが発生しました: ${error}`);
  }
}

// 元のfetchHooksFromGitHub関数は基本的な役割予測のみを行う
function predictHookRole(hookName, type) {
  if (hookName.includes('save_post_')) {
    return `${hookName.replace('save_post_', '')}の保存処理`;
  }
  if (hookName.includes('pre_get_posts')) {
    return 'クエリ制御';
  }
  if (hookName.includes('rest_api_init')) {
    return 'REST API登録';
  }
  // 他の基本的なパターン
  return type === 'filter' ? 'データ加工' : 'アクション実行';
}
