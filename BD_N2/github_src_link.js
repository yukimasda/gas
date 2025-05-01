function getAllFiles() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GitHub Src Links');
  if (!sheet) throw new Error('「GitHub Src Links」シートが見つかりません');
  
  // シートをクリア
  sheet.clear();
  
  // ヘッダー設定
  sheet.getRange('A1').setValue('GitHub repo list');
  sheet.getRange('A1:C1').merge().setFontWeight('bold').setHorizontalAlignment('center');
  
  // カラム名の設定
  sheet.getRange('A2').setValue('ディレクトリ');
  sheet.getRange('B2').setValue('ファイル名');
  sheet.getRange('C2').setValue('説明');
  sheet.getRange('A2:C2').setFontWeight('bold').setBackground('#f3f3f3');
  
  // カラム幅の設定
  sheet.setColumnWidth(1, 250); // A列：ディレクトリ
  sheet.setColumnWidth(2, 200); // B列：ファイル名
  sheet.setColumnWidth(3, 400); // C列：説明
  
  // リポジトリのルートディレクトリ内のファイル一覧を取得
  const rootUrl = `https://api.github.com/repos/${repo}/contents/`;
  fetchFiles(rootUrl, sheet, 3, "", ""); // 3行目から表示開始、初期パスは空、初期ディレクトリは空
  
  // 罫線を追加
  const lastRow = sheet.getLastRow();
  if (lastRow > 2) {
    sheet.getRange(2, 1, lastRow - 1, 3).setBorder(true, true, true, true, true, true);
  }
}

/**
 * 特定のディレクトリやファイルを除外すべきかどうかを判断する
 */
function shouldExclude(path, name) {
  // 除外するディレクトリやファイル名のリスト
  const excludeDirs = ['dist', 'node_modules', '.git'];
  const excludeExtensions = ['.min.js', '.map'];
  
  // パス全体またはディレクトリ名が除外リストに含まれているかチェック
  for (const dir of excludeDirs) {
    if (path.includes(`/${dir}/`) || path === dir || name === dir) {
      return true;
    }
  }
  
  // 除外する拡張子をチェック
  for (const ext of excludeExtensions) {
    if (name.endsWith(ext)) {
      return true;
    }
  }
  
  return false;
}

function fetchFiles(url, sheet, startRow, currentPath, lastDir) {
  // URLにdistディレクトリが含まれている場合はスキップ
  if (shouldExclude(currentPath, '')) {
    return startRow; // 現在の行を変更せずに返す
  }

  const options = {
    headers: {
      Authorization: `token ${token}`,
      Accept: "application/vnd.github.v3.raw"
    },
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());

  if (Array.isArray(data)) {
    let row = startRow;
    let currentDirDisplayed = false; // 現在のディレクトリが表示されたかどうかを追跡
    
    // ファイルとディレクトリを分離して、ファイルを先に処理
    const files = data.filter(item => item.type === 'file' && !shouldExclude(currentPath, item.name));
    const dirs = data.filter(item => item.type === 'dir' && !shouldExclude(currentPath, item.name));
    
    // ファイルをファイル名でソート
    files.sort((a, b) => a.name.localeCompare(b.name));
    
    // ファイルの処理
    files.forEach(item => {
      const fileName = item.name;
      const fileLink = item.html_url;
      const filePath = currentPath ? `${currentPath}` : "";

      // ディレクトリが変わった場合のみディレクトリ名を表示
      if (lastDir !== currentPath) {
        lastDir = currentPath;
        currentDirDisplayed = false;
      }
      
      // A列にディレクトリ（同じディレクトリでは最初だけ表示）
      if (!currentDirDisplayed) {
        sheet.getRange(row, 1).setValue(filePath);
        currentDirDisplayed = true;
      }
      
      // B列にファイル名（リンク付き）
      sheet.getRange(row, 2).setFormula(`=HYPERLINK("${fileLink}", "${fileName}")`);
      
      // C列は空にしておく（説明処理を削除）
      
      row++;
    });
    
    // ディレクトリをソート
    dirs.sort((a, b) => a.name.localeCompare(b.name));
    
    // ディレクトリの処理（再帰的に）
    for (const dir of dirs) {
      const subDirUrl = dir.url;
      const subDirPath = currentPath ? `${currentPath}/${dir.name}` : dir.name;
      row = fetchFiles(subDirUrl, sheet, row, subDirPath, lastDir);
    }
    
    return row;
  }
  return startRow;
}

/**
 * GitHub Src Linksシートの各ファイルをAIで要約し、C列に表示する
 */
function summarizeFilesWithAI() {
  // APIキーとトークンの取得
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!token || !apiKey) {
    SpreadsheetApp.getUi().alert('GitHubトークンまたはOpenAI APIキーが設定されていません。');
    return;
  }

  // シートの取得
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GitHub Src Links');
  if (!sheet) {
    SpreadsheetApp.getUi().alert('「GitHub Src Links」シートが見つかりません');
    return;
  }

  // 処理対象データの範囲を取得
  const lastRow = sheet.getLastRow();
  if (lastRow <= 2) {
    SpreadsheetApp.getUi().alert('要約するファイルがありません。先に「GitHub Src Links」を実行してください。');
    return;
  }

  // 処理状況表示用のセルを用意
  sheet.getRange('D1').setValue('AI要約処理状況');
  
  // モデル名と最大トークン数を定義
  const modelName = "chatgpt-4o-latest";
  const maxTokens = 8000;

  // 各行を処理
  for (let row = 3; row <= lastRow; row++) {
    // 進捗状況を更新
    sheet.getRange('D1').setValue(`AI要約処理中... (${row-2}/${lastRow-2})`);
    
    // ファイル情報を取得
    const dirPath = sheet.getRange(row, 1).getValue();
    const fileCell = sheet.getRange(row, 2);
    const fileFormula = fileCell.getFormula();
    
    // ファイルがない行はスキップ
    if (!fileFormula) continue;
    
    try {
      // HYPERLINKからファイル名とURLを抽出
      const matches = fileFormula.match(/=HYPERLINK\("([^"]+)", "([^"]+)"\)/);
      if (!matches) continue;
      
      const fileUrl = matches[1];
      const fileName = matches[2];
      
      // ファイルパスを構築
      let filePath = '';
      if (dirPath) {
        filePath = `${dirPath}/${fileName}`;
      } else {
        filePath = fileName;
      }
      
      // GitHub APIのURLに変換
      const branch = "v1"; // analyzer.jsと同じブランチを使用
      const contentUrl = `https://api.github.com/repos/${repo}/contents/${filePath}?ref=${branch}`;
      
      // ファイル内容の取得
      const contentResponse = UrlFetchApp.fetch(contentUrl, {
        headers: {
          'Authorization': `token ${token}`,
          'Accept': 'application/vnd.github.v3+json'
        },
        muteHttpExceptions: true
      });
      
      if (contentResponse.getResponseCode() !== 200) {
        sheet.getRange(row, 3).setValue('取得エラー');
        continue;
      }
      
      const responseContent = JSON.parse(contentResponse.getContentText());
      
      // ディレクトリの場合はスキップ
      if (Array.isArray(responseContent) || !responseContent.content) {
        sheet.getRange(row, 3).setValue('ディレクトリ');
        continue;
      }
      
      // ファイル内容をデコード
      const sourceCode = Utilities.newBlob(
        Utilities.base64Decode(responseContent.content)
      ).getDataAsString();
      
      // AI要約の実行
      const systemPrompt = `あなたはソースコードを一言で要約する専門家です。
      ファイルの主な目的や機能を最大40文字程度の日本語で簡潔に説明してください。`;
      
      const userPrompt = `以下のファイルの機能や役割を一言で要約してください：
      ファイル名: ${fileName}
      パス: ${filePath}
      
      ソースコード:
      ${sourceCode}`;
      
      // OpenAI API呼び出し
      const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
        method: 'post',
        headers: {
          'Authorization': `Bearer ${apiKey}`,
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify({
          model: modelName,
          messages: [
            { role: "system", content: systemPrompt },
            { role: "user", content: userPrompt }
          ],
          temperature: 0.3,
          max_tokens: 100 // 短い要約なので少なめのトークン数
        }),
        muteHttpExceptions: true
      });
      
      const aiResponse = JSON.parse(response.getContentText());
      const summary = aiResponse.choices[0].message.content.trim();
      
      // C列に要約を設定
      sheet.getRange(row, 3).setValue(summary);
      
      // API制限を考慮して待機
      Utilities.sleep(1000);
      
    } catch (error) {
      Logger.log(`エラー（行 ${row}）: ${error.message}`);
      sheet.getRange(row, 3).setValue('処理エラー');
    }
  }
  
  // 処理完了メッセージ
  sheet.getRange('D1').setValue('AI要約完了！');
}
