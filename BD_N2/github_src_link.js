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
 * 10ファイルごとにまとめて処理し、結果は個別のセルに出力
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

  // バッチ処理のための変数
  const batchSize = 10; // 10ファイルごとに処理
  let fileInfos = [];   // ファイル情報を格納する配列
  let batchCount = 0;   // 処理したバッチの数
  
  // 直近の有効なディレクトリパスを保持する変数
  let lastValidDirPath = "";
  
  // 各行を処理
  for (let row = 3; row <= lastRow; row++) {
    // 進捗状況を更新
    sheet.getRange('D1').setValue(`AI要約準備中... (${row-2}/${lastRow-2})`);
    
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
      
      // ファイル拡張子を確認
      const fileExtension = fileName.split('.').pop().toLowerCase();
      
      // ディレクトリパスを更新：現在のセルが空でない場合は値を保存
      if (dirPath && dirPath.trim() !== '') {
        lastValidDirPath = dirPath;
      }
      
      // ファイルパスを構築（A列が空の場合は直近の有効なディレクトリパスを使用）
      let filePath;
      if (lastValidDirPath) {
        filePath = `${lastValidDirPath}/${fileName}`;
      } else {
        filePath = fileName;
      }
      
      // ファイル情報を配列に追加
      fileInfos.push({
        row: row,
        dirPath: lastValidDirPath, // 直近の有効なディレクトリパスを使用
        fileName: fileName,
        filePath: filePath,
        extension: fileExtension
      });
      
      // バッチサイズに達したらまとめて処理
      if (fileInfos.length >= batchSize || row === lastRow) {
        batchCount++;
        
        // バッチ処理を実行
        processBatch(sheet, fileInfos, batchCount, modelName, maxTokens, apiKey);
        
        // 配列をクリア
        fileInfos = [];
        
        // API制限を考慮して待機
        Utilities.sleep(2000);
      }
    } catch (error) {
      Logger.log(`エラー（行 ${row}）: ${error.message}`);
      sheet.getRange(row, 3).setValue('処理エラー');
    }
  }
  
  // 処理完了メッセージ
  sheet.getRange('D1').setValue('AI要約完了！');
}

/**
 * バッチ単位でファイルを処理
 */
function processBatch(sheet, fileInfos, batchCount, modelName, maxTokens, apiKey) {
  // 進捗状況を更新
  sheet.getRange('D1').setValue(`AI要約処理中... バッチ ${batchCount}`);
  
  // このバッチのファイル情報をまとめる
  let batchFiles = [];
  let mediaOrErrorFiles = [];  // メディアファイルまたはエラーのあるファイル
  
  // 各ファイルを処理
  for (const fileInfo of fileInfos) {
    try {
      // メディアファイルは固定の説明を設定して即時出力
      if (isMediaFile(fileInfo.extension)) {
        sheet.getRange(fileInfo.row, 3).setValue(getFileTypeDescription(fileInfo.extension));
        mediaOrErrorFiles.push(fileInfo);
        continue;
      }
      
      // GitHub APIのURLに変換
      const branch = "v1"; // analyzer.jsと同じブランチを使用
      const contentUrl = `https://api.github.com/repos/${repo}/contents/${fileInfo.filePath}?ref=${branch}`;
      
      // ファイル内容の取得とエラーハンドリングの強化
      Logger.log(`リクエストURL: ${contentUrl}`); // デバッグ用
      
      const contentResponse = UrlFetchApp.fetch(contentUrl, {
        headers: {
          'Authorization': `token ${token}`,
          'Accept': 'application/vnd.github.v3+json'
        },
        muteHttpExceptions: true
      });
      
      if (contentResponse.getResponseCode() !== 200) {
        Logger.log(`API応答エラー: ${contentResponse.getResponseCode()} - ${contentResponse.getContentText()}`);
        sheet.getRange(fileInfo.row, 3).setValue(`取得エラー(${contentResponse.getResponseCode()})`);
        mediaOrErrorFiles.push(fileInfo);
        continue;
      }
      
      const responseContent = JSON.parse(contentResponse.getContentText());
      
      // ディレクトリの場合はスキップ
      if (Array.isArray(responseContent) || !responseContent.content) {
        sheet.getRange(fileInfo.row, 3).setValue('ディレクトリ');
        mediaOrErrorFiles.push(fileInfo);
        continue;
      }
      
      // ファイル内容をデコード
      let sourceCode;
      try {
        sourceCode = Utilities.newBlob(
          Utilities.base64Decode(responseContent.content)
        ).getDataAsString();
      } catch (e) {
        // デコードに失敗した場合は、ファイルタイプの説明を設定
        sheet.getRange(fileInfo.row, 3).setValue(getFileTypeDescription(fileInfo.extension));
        mediaOrErrorFiles.push(fileInfo);
        continue;
      }
      
      // AI解析用にファイル情報を追加
      batchFiles.push({
        ...fileInfo,
        content: sourceCode.length > 4000 ? sourceCode.substring(0, 4000) + "...(省略)..." : sourceCode
      });
      
    } catch (error) {
      Logger.log(`エラー（ファイル ${fileInfo.fileName}）: ${error.message}`);
      sheet.getRange(fileInfo.row, 3).setValue('処理エラー');
      mediaOrErrorFiles.push(fileInfo);
    }
  }
  
  // メディアファイルやエラーファイル以外のファイルがあればAIで一括解析
  if (batchFiles.length > 0) {
    // AI要約の実行
    const systemPrompt = `あなたはソースコードを一言で要約する専門家です。
    複数のファイルの主な目的や機能をそれぞれ最大40文字程度の日本語で簡潔に説明してください。
    出力形式は「ファイル名: 説明」の形式で、ファイルごとに1行ずつ出力してください。`;
    
    let filesContent = "";
    for (const file of batchFiles) {
      filesContent += `\n--- ファイル: ${file.fileName} (${file.filePath}) ---\n${file.content}\n\n`;
    }
    
    const userPrompt = `以下の複数のファイルの機能や役割をそれぞれ一言で要約してください：
    
    ${filesContent}
    
    各ファイルについて、以下の形式で1行ずつ出力してください：
    ファイル名: 説明`;
    
    try {
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
          max_tokens: 800 // バッチ処理のため多めのトークン数
        }),
        muteHttpExceptions: true
      });
      
      const aiResponse = JSON.parse(response.getContentText());
      const summaryText = aiResponse.choices[0].message.content.trim();
      
      // 解析結果を行ごとに分割
      const summaryLines = summaryText.split('\n').filter(line => line.trim() !== '');
      
      // 各ファイルの解析結果を対応するセルに出力
      const fileSummaries = {};
      
      for (const line of summaryLines) {
        const match = line.match(/^([^:]+):\s*(.*)/);
        if (match) {
          const fileName = match[1].trim();
          const summary = match[2].trim();
          fileSummaries[fileName] = summary;
        }
      }
      
      // 結果をそれぞれの行のC列に出力
      for (const file of batchFiles) {
        const summary = fileSummaries[file.fileName] || getFileTypeDescription(file.extension);
        sheet.getRange(file.row, 3).setValue(summary);
      }
      
    } catch (error) {
      Logger.log(`AI解析エラー: ${error.message}`);
      // 解析に失敗した場合、ファイル拡張子に基づく説明を使用
      for (const file of batchFiles) {
        sheet.getRange(file.row, 3).setValue(getFileTypeDescription(file.extension));
      }
    }
  }
}

/**
 * メディアファイルかどうかを判定
 */
function isMediaFile(extension) {
  const mediaExtensions = ['mp3', 'jpg', 'jpeg', 'png', 'svg', 'gif', 'webp', 'ico'];
  return mediaExtensions.includes(extension);
}

/**
 * ファイルタイプに応じた説明を取得
 */
function getFileTypeDescription(extension) {
  const descriptions = {
    // コード関連
    'js': 'JavaScriptソースファイル',
    'ts': 'TypeScriptソースファイル',
    'scss': 'SCSSスタイルシート',
    'sass': 'Sassスタイルシート',
    'php': 'PHPスクリプトファイル',
    
    // マークアップ/設定
    'md': 'Markdownドキュメント',
    'yml': 'YAML設定ファイル',
    'yaml': 'YAML設定ファイル',
    'xml': 'XMLデータファイル',
    'json': 'JSON設定/データファイル',
    'po': '翻訳ファイル（Gettext PO）',
    
    // メディア
    'mp3': '音声ファイル（MP3）',
    'jpg': '画像ファイル（JPG）',
    'jpeg': '画像ファイル（JPEG）',
    'png': '画像ファイル（PNG）',
    'svg': 'ベクターグラフィック（SVG）',
    'gif': 'アニメーション画像（GIF）',
    'webp': '画像ファイル（WebP）',
    'ico': 'アイコンファイル'
  };
  
  return descriptions[extension] || `${extension.toUpperCase()}ファイル`;
}
