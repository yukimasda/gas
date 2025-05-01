function getAllFiles() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GitHub Src Links');
  if (!sheet) throw new Error('「GitHub Src Links」シートが見つかりません');
  
  // シートをクリア
  sheet.clear();
  
  // ヘッダー設定
  sheet.getRange('A1').setValue('GitHub repo list');
  sheet.getRange('A1:G1').merge().setFontWeight('bold').setHorizontalAlignment('center');
  
  // カラム名の設定
  sheet.getRange('A2').setValue('ディレクトリ');
  sheet.getRange('B2').setValue('ファイル名');
  sheet.getRange('C2').setValue('行数');
  sheet.getRange('D2').setValue('文字数');
  sheet.getRange('E2').setValue('概要説明');
  sheet.getRange('F2').setValue('画面名');
  sheet.getRange('G2').setValue('バリデーション内容');
  sheet.getRange('A2:G2').setFontWeight('bold').setBackground('#f3f3f3');
  
  // カラム幅の設定
  sheet.setColumnWidth(1, 250); // A列：ディレクトリ
  sheet.setColumnWidth(2, 200); // B列：ファイル名
  sheet.setColumnWidth(3, 80);  // C列：行数
  sheet.setColumnWidth(4, 100); // D列：文字数
  sheet.setColumnWidth(5, 400); // E列：概要説明
  sheet.setColumnWidth(6, 200); // F列：画面名
  sheet.setColumnWidth(7, 300); // G列：バリデーション内容
  
  // リポジトリのルートディレクトリ内のファイル一覧を取得
  const rootUrl = `https://api.github.com/repos/${repo}/contents/`;
  fetchFiles(rootUrl, sheet, 3, "", ""); // 3行目から表示開始、初期パスは空、初期ディレクトリは空
  
  // 罫線を追加
  const lastRow = sheet.getLastRow();
  if (lastRow > 2) {
    sheet.getRange(2, 1, lastRow - 1, 7).setBorder(true, true, true, true, true, true);
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
    for (const item of files) {
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
      
      // C列に行数を表示
      try {
        // ファイルの内容を取得して行数カウント
        const fileFullPath = currentPath ? `${currentPath}/${fileName}` : fileName;
        const contentUrl = `https://api.github.com/repos/${repo}/contents/${fileFullPath}`;
        const contentResponse = UrlFetchApp.fetch(contentUrl, options);
        
        if (contentResponse.getResponseCode() === 200) {
          // バイナリファイルの場合はbase64エンコードされているため、テキストファイルかどうか判断
          if (isTextFile(fileName)) {
            const content = contentResponse.getContentText();
            // 末尾の改行を含まない内容に調整
            const trimmedContent = content.endsWith('\n') ? content.slice(0, -1) : content;
            
            let lines = trimmedContent.split('\n');
            // ファイル末尾の空行を削除（末尾の改行による余分なカウントを防ぐ）
            if (lines[lines.length - 1] === '') {
              lines.pop();
            }
            const lineCount = lines.length;
            const charCount = trimmedContent.length;
            
            sheet.getRange(row, 3).setValue(lineCount);
            sheet.getRange(row, 4).setValue(charCount);
          } else {
            sheet.getRange(row, 3).setValue('-');
            sheet.getRange(row, 4).setValue('-');
          }
        } else {
          sheet.getRange(row, 3).setValue('-');
          sheet.getRange(row, 4).setValue('-');
        }
      } catch (error) {
        // エラーが発生した場合
        Logger.log(`ファイル情報取得エラー (${fileName}): ${error.message}`);
        sheet.getRange(row, 3).setValue('-');
        sheet.getRange(row, 4).setValue('-');
      }
      
      row++;
    }
    
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
 * ファイルがテキストファイルかどうかを判断する
 */
function isTextFile(filename) {
  const textExtensions = [
    '.js', '.ts', '.jsx', '.tsx', '.html', '.htm', '.css', '.scss', '.sass', '.less',
    '.txt', '.md', '.markdown', '.json', '.xml', '.yml', '.yaml', '.csv', '.tsv',
    '.php', '.py', '.rb', '.java', '.c', '.cpp', '.h', '.cs', '.go', '.swift',
    '.pl', '.pm', '.sh', '.bash', '.ini', '.cfg', '.config', '.sql', '.vue',
    '.jsx', '.tsx', '.handlebars', '.hbs', '.ejs', '.pug', '.jade'
  ];
  
  const ext = '.' + filename.split('.').pop().toLowerCase();
  return textExtensions.includes(ext);
}

/**
 * GitHub Src Linksシートの各ファイルをAIで要約し、G列に表示する
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
  sheet.getRange('G1').setValue('AI要約処理状況');
  
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
    sheet.getRange('G1').setValue(`AI要約準備中... (${row-2}/${lastRow-2})`);
    
    // ファイル情報を取得
    const dirPath = sheet.getRange(row, 1).getValue();
    const fileCell = sheet.getRange(row, 2);
    const fileFormula = fileCell.getFormula();
    const lineCount = sheet.getRange(row, 3).getValue(); // 行数（すでに表示済み）
    
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
        dirPath: lastValidDirPath,
        fileName: fileName,
        filePath: filePath,
        extension: fileExtension,
        lineCount: lineCount // 行数情報も含める
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
      Logger.log(`ファイル情報の取得エラー: ${error.message}`);
      continue;
    }
  }
  
  // 処理完了メッセージ
  sheet.getRange('G1').setValue('AI要約完了！');
}

function processBatch(sheet, fileInfos, batchCount, modelName, maxTokens, apiKey) {
  // 進捗状況を更新
  sheet.getRange('G1').setValue(`AI要約処理中... バッチ ${batchCount}`);
  
  // このバッチのファイル情報をまとめる
  let batchFiles = [];
  let mediaOrErrorFiles = [];  // メディアファイルまたはエラーのあるファイル
  
  // 各ファイルを処理
  for (const fileInfo of fileInfos) {
    try {
      // メディアファイルは固定の説明を設定して即時出力
      if (isMediaFile(fileInfo.extension)) {
        sheet.getRange(fileInfo.row, 5).setValue(getFileTypeDescription(fileInfo.extension));
        sheet.getRange(fileInfo.row, 6).setValue('-');
        sheet.getRange(fileInfo.row, 7).setValue('-');
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
        sheet.getRange(fileInfo.row, 5).setValue(`取得エラー(${contentResponse.getResponseCode()})`);
        sheet.getRange(fileInfo.row, 6).setValue('-');
        sheet.getRange(fileInfo.row, 7).setValue('-');
        mediaOrErrorFiles.push(fileInfo);
        continue;
      }
      
      // ファイル内容の取得
      const content = contentResponse.getContentText();
      const contentJson = JSON.parse(content);
      
      // Base64エンコードされたコンテンツをデコード
      let sourceCode = '';
      try {
        sourceCode = Utilities.base64Decode(contentJson.content);
        sourceCode = Utilities.newBlob(sourceCode).getDataAsString();
      } catch (e) {
        Logger.log(`デコードエラー ${fileInfo.fileName}: ${e.message}`);
        sheet.getRange(fileInfo.row, 5).setValue('デコードエラー');
        sheet.getRange(fileInfo.row, 6).setValue('-');
        sheet.getRange(fileInfo.row, 7).setValue('-');
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
      sheet.getRange(fileInfo.row, 5).setValue('処理エラー');
      sheet.getRange(fileInfo.row, 6).setValue('-');
      sheet.getRange(fileInfo.row, 7).setValue('-');
      mediaOrErrorFiles.push(fileInfo);
    }
  }
  
  // メディアファイルやエラーファイル以外のファイルがあればAIで一括解析
  if (batchFiles.length > 0) {
    // 画面名のリスト
    const screenNames = [
      "ログイン画面", "ダッシュボード", "返礼品一覧", "お知らせ", "ユーザー一覧", 
      "N2設定", "寄附金額・送料設定", "LH設定", "注意書き設定", "ふるさとチョイス設定", 
      "楽天設定", "楽天SFTP", "SFTPログ（キャビアップ）", "SFTPログ（RMS連携機能）", 
      "SFTPログ", "エラーログ", "キャビネット", "キャビ蓮舫", "ブクマURL提供", "N2SYNC", 
      "立替金精算書DL", "寄付金シミュレータ", "エクスポート"
    ].join("、");

    // AI要約の実行
    const systemPrompt = `あなたはソースコードを分析して詳細な情報を抽出する専門家です。
    複数のファイルに対して、以下の3つの分析を行ってください：
    
    1. 各ファイルの主な目的や機能を最大40文字程度の日本語で簡潔に説明
    2. 各ファイルが関連する画面名の特定（指定されたリストから選択）
    3. バリデーション機能の詳細な内容（正規表現、桁数制限、入力チェックなど具体的に）
    
    画面名は以下のリストから選択してください：
    ${screenNames}
    
    これらに当てはまらない場合は「-」と表示してください。
    
    出力形式は以下のとおりです：
    ファイル名: [説明] | [画面名] | [バリデーション内容]
    
    ファイルごとに1行ずつ出力してください。`;
    
    let filesContent = "";
    for (const file of batchFiles) {
      filesContent += `\n--- ファイル: ${file.fileName} (${file.filePath}) ---\n${file.content}\n\n`;
    }
    
    const userPrompt = `以下の複数のファイルを分析してください：
    
    ${filesContent}
    
    各ファイルについて、以下の形式で1行ずつ出力してください：
    ファイル名: [ファイルの説明] | [画面名] | [バリデーション内容]
    
    画面名は以下のリストから選択してください。該当するものがない場合は「-」と表示してください：
    ${screenNames}
    
    バリデーション内容は、具体的な内容（正規表現の内容、桁数制限の詳細、入力チェックの条件など）を記述してください。
    バリデーションがない場合は「-」と表示してください。`;
    
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
          temperature: 0,
          max_tokens: 1500
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
        // "ファイル名: [説明] | [画面名] | [バリデーション内容]" の形式をパース
        const match = line.match(/^([^:]+):\s*(.*?)\s*\|\s*(.*?)\s*\|\s*(.*)/);
        if (match) {
          const fileName = match[1].trim();
          const summary = match[2].trim();
          const screenName = match[3].trim();
          const validation = match[4].trim();
          fileSummaries[fileName] = { summary, screenName, validation };
        } else {
          // フォールバック：従来の形式のパース
          const basicMatch = line.match(/^([^:]+):\s*(.*)/);
          if (basicMatch) {
            const fileName = basicMatch[1].trim();
            const summary = basicMatch[2].trim();
            fileSummaries[fileName] = { 
              summary, 
              screenName: '-',
              validation: '-'
            };
          }
        }
      }
      
      // 結果をそれぞれの行に出力 - 出力先の列を修正
      for (const file of batchFiles) {
        const fileInfo = fileSummaries[file.fileName] || { 
          summary: getFileTypeDescription(file.extension), 
          screenName: '-',
          validation: '-'
        };
        
        sheet.getRange(file.row, 5).setValue(fileInfo.summary);    // E列：概要説明
        sheet.getRange(file.row, 6).setValue(fileInfo.screenName); // F列：画面名
        sheet.getRange(file.row, 7).setValue(fileInfo.validation); // G列：バリデーション内容
      }
      
    } catch (error) {
      Logger.log(`AI解析エラー: ${error.message}`);
      // 解析に失敗した場合のフォールバック処理も修正
      for (const file of batchFiles) {
        sheet.getRange(file.row, 5).setValue(getFileTypeDescription(file.extension)); // E列：概要説明
        sheet.getRange(file.row, 6).setValue('-'); // F列：画面名
        sheet.getRange(file.row, 7).setValue('-'); // G列：バリデーション内容
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
