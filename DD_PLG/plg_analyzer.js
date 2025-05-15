// グローバル変数の定義
const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');

// モデル名と最大トークン数を定義
const modelName = "chatgpt-4o-latest";
const maxTokens = 15000; // GPT-4のmaxtokenの最大トークン数

/**
 * スプレッドシートの初期設定
 */
function initializeSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // A1, A2にタイトルを設定
  sheet.getRange("A1").setValue("repoURL");
  sheet.getRange("A2").setValue("branch");
  
  // タイトル行のスタイル設定
  const titleRange = sheet.getRange("A1:A2");
  titleRange.setFontWeight('bold');
  titleRange.setBackground('#f3f3f3');

  // ログと解析結果のタイトルを設定
  sheet.getRange("B9").setValue("解析ログ").setFontWeight('bold').setBackground('#f3f3f3');
  sheet.getRange("B11").setValue("解析結果").setFontWeight('bold').setBackground('#f3f3f3');
}

/**
 * リポジトリとブランチ情報を取得
 */
function getRepoInfo() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const repo = sheet.getRange("B1").getValue(); //steamships/neo-neng
  const branch = sheet.getRange("B2").getValue(); //v1
  
  if (!repo || !branch) {
    throw new Error("リポジトリURLまたはブランチが設定されていません。B1セルにリポジトリURL、B2セルにブランチ名を入力してください。");
  }
  
  return { repo, branch };
}

/**
 * 既存の解析結果をクリア
 */
function clearExistingData(sheet, headers) {
  if (sheet.getLastRow() > 11) {  // 8から11に変更
    // A列以外(B列以降)の12行目以降をクリア  // 9から12に変更
    const range = sheet.getRange(12, 2, sheet.getLastRow() - 11, sheet.getLastColumn() - 1);
    range.clear(); // 書式設定を含めて全てクリア
  }
}

/**
 * GitHubからソースコードを取得
 */
async function fetchGitHubContent(sourcePath) {
  const { repo, branch } = getRepoInfo();
  const contentUrl = `https://api.github.com/repos/${repo}/contents/${sourcePath}?ref=${branch}`;
  const githubWebUrl = `https://github.com/${repo}/blob/${branch}/${sourcePath}`;
  
  try {
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

    const responseContent = JSON.parse(contentResponse.getContentText());
    
    // レスポンスが配列の場合（ディレクトリの場合）またはcontentプロパティがない場合
    if (Array.isArray(responseContent) || !responseContent.content) {
      return {
        sourceCode: "このファイルはディレクトリまたは特殊なファイル形式のため、内容を表示できません。",
        githubWebUrl
      };
    }

    const sourceCode = Utilities.newBlob(
      Utilities.base64Decode(responseContent.content)
    ).getDataAsString();

    return {
      sourceCode,
      githubWebUrl
    };
  } catch (error) {
    Logger.log(`Error in fetchGitHubContent for ${sourcePath}: ${error.message}`);
    
    // オリジナルのエラーをスローする前にログを取得
    if (error.message.includes("Cannot read properties of undefined")) {
      Logger.log("Undefined property error detected. Response might not contain expected fields.");
    }
    
    throw error;
  }
}

/**
 * ステータス表示を更新する関数
 */
function updateStatus(sheet, message) {
  const currentStatus = sheet.getRange("B10").getValue();
  const timestamp = new Date().toLocaleTimeString();
  const newStatus = currentStatus ? 
    `${currentStatus}\n${timestamp}: ${message}` : 
    `${timestamp}: ${message}`;
  sheet.getRange("B10").setValue(newStatus);
}

/**
 * OpenAI APIを呼び出して解析を実行
 */
async function callOpenAI(sourcePath, sourceCode, headers) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // B4セルから追加の解析ポイントを取得
  const additionalPoints = sheet.getRange("B4").getValue().trim();
  
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
  ${additionalPoints ? `4. ${additionalPoints}` : ''}

  出力形式：
  ${headers.join('###')}|||
  （ファイルの構造に従って上から順にデータを出力）

  解析対象のソースコード:
  ${sourceCode}`;

  // 推定トークン数を計算して表示
  const estimatedTokens = Math.ceil((systemPrompt.length + userPrompt.length) / 4);
  updateStatus(sheet, `推定トークン数: ${estimatedTokens}`);

  try {
    // モデル名を表示
    updateStatus(sheet, `使用モデル: ${modelName}`);

    // OpenAI API呼び出しでモデル名と最大トークン数を使用
    const response = await UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${apiKey}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        model: modelName, // モデル名を使用
        messages: [
          { role: "system", content: systemPrompt },
          { role: "user", content: userPrompt }
        ],
        temperature: 0,
        max_tokens: maxTokens // 最大トークン数を指定
      }),
      muteHttpExceptions: true
    });

    const responseText = response.getContentText();
    const responseData = JSON.parse(responseText);

    // 使用したトークン数を表示
    const usedTokens = responseData.usage ? responseData.usage.total_tokens : 0;
    updateStatus(sheet, `使用トークン数: ${usedTokens}`);

    // トークン数が制限を超えた場合の処理
    //if (usedTokens > maxTokens) {
    //  Logger.log("トークン数が制限を超えました。");
    //  throw new Error("トークン数が制限を超えました。");
    //}

    const content = responseData.choices[0].message.content;

    // 解析結果をパース
    return content.split('|||')
      .filter(row => row.trim() && !row.includes(headers.join('###')))
      .map(row => {
        const columns = row.split('###').map(col => col.trim());
        return headers.map((_, index) => columns[index] || "未定義");
      });

  } catch (error) {
    Logger.log("APIリクエスト中にエラーが発生しました:", error);
    throw error; // エラーを再スローして、呼び出し元で処理
  }
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
async function analyzePlg() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // 初期化：ステータス表示をクリア
  sheet.getRange("B10").setValue("");
  
  // タイトルの設定
  sheet.getRange("B9").setValue("解析ログ").setFontWeight('bold').setBackground('#f3f3f3');
  sheet.getRange("B11").setValue("解析結果").setFontWeight('bold').setBackground('#f3f3f3');

  // トークンのチェック
  if (!token || !apiKey) {
    sheet.getRange("B4").setValue("");
    updateStatus(sheet, "⚠️ GitHubトークンまたはOpenAI APIキーが設定されていません");
    return;
  }

  try {
    // リポジトリ情報を取得
    const { repo } = getRepoInfo();
    
    // repoからowner, repoNameを取得
    const [owner, repoName] = repo.split('/');
    if (!owner || !repoName) {
      updateStatus(sheet, "⚠️ リポジトリの形式が正しくありません");
      return;
    }

    // A列のファイル一覧を取得（A4セルから開始）
    const startRow = 4;
    const lastRow = sheet.getLastRow();
    const fileRange = sheet.getRange(startRow, 1, lastRow - startRow + 1, 1);
    let files = fileRange.getValues();

    // 空の行をフィルタリング
    files = files.filter(file => file[0]);

    // ヘッダー行を取得して検証
    const headerRange = sheet.getRange(7, 2, 1, sheet.getLastColumn() - 1);
    const headers = headerRange.getValues()[0].filter(header => header !== '');
    if (headers.length === 0) {
      updateStatus(sheet, "⚠️ 7行目にヘッダーが設定されていません");
      return;
    }

    // 既存のデータをクリア
    clearExistingData(sheet, headers);

    let currentRow = 12;  // 結果の書き込み開始行を12に変更

    // ファイルごとの処理
    for (let i = 0; i < files.length; i++) {
      const sourcePath = files[i][0];
      if (!sourcePath) continue;

      try {
        // 進捗状況を更新
        updateStatus(sheet, `GPT解析中... (${i + 1}/${files.length})`);

        // GitHubからソースコード取得
        const { sourceCode, githubWebUrl } = await fetchGitHubContent(sourcePath);

        // B列にファイル名とリンクを設定
        const richText = SpreadsheetApp.newRichTextValue()
          .setText(sourcePath)
          .setLinkUrl(githubWebUrl)
          .build();
        sheet.getRange(currentRow, 2).setRichTextValue(richText);
        currentRow++;

        // ヘッダーを出力
        sheet.getRange(currentRow, 2, 1, headers.length).setValues([headers]);

        // ヘッダーのスタイルを設定
        const headerRange = sheet.getRange(currentRow, 2, 1, headers.length);
        headerRange.setFontWeight('bold');
        headerRange.setBackground('#f3f3f3');

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
        updateStatus(sheet, `⚠️ エラー: ${error.message}`);
        currentRow += 2;
        continue;
      }
    }

    // 解析完了メッセージを表示
    updateStatus(sheet, "全ファイルの解析完了");

  } catch (error) {
    Logger.log("解析中にエラーが発生しました:", error);
    updateStatus(sheet, `⚠️ エラー: ${error.message}`);
  }
}
