function getAllBranches() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GitHub Plg Links');
  if (!sheet) throw new Error('「GitHub Plg Links」シートが見つかりません');
  
  // シートをクリア
  sheet.clear();
  SpreadsheetApp.flush(); // クリア処理を即時反映

  // 初期メッセージを表示
  sheet.getRange('A1').setValue('GitHub Plugin Branches (準備中...)');
  SpreadsheetApp.flush();

  // ヘッダー設定を5列に修正
  sheet.getRange('A1:E1').merge().setFontWeight('bold').setHorizontalAlignment('center');
  
  // カラム名の設定を修正（カラム順序変更）
  sheet.getRange('A2').setValue('ブランチ名');
  sheet.getRange('B2').setValue('Plugin Name');
  sheet.getRange('C2').setValue('分類');
  sheet.getRange('D2').setValue('処理概要');
  sheet.getRange('E2').setValue('プラグイン固有フック');
  sheet.getRange('F2').setValue('namespace');
  sheet.getRange('G2').setValue('コールバック関数');
  sheet.getRange('A2:G2').setFontWeight('bold').setBackground('#f3f3f3');
  
  // カラム幅の設定を修正
  sheet.setColumnWidth(1, 200); // A列：ブランチ名
  sheet.setColumnWidth(2, 300); // B列：Plugin Name
  sheet.setColumnWidth(3, 100); // C列：分類
  sheet.setColumnWidth(4, 400); // D列：処理概要
  sheet.setColumnWidth(5, 300); // E列：プラグイン固有フック
  sheet.setColumnWidth(6, 150); // F列：namespace
  sheet.setColumnWidth(7, 300); // G列：コールバック関数
  
  let page = 1;
  let allBranches = [];
  
  try {
    // ブランチ取得開始メッセージ
    sheet.getRange('A1').setValue('GitHub Plugin Branches (ブランチ一覧取得中...)');
    SpreadsheetApp.flush();

    // 全てのブランチを取得（ページネーション対応）
    while (true) {
      const branchesUrl = `https://api.github.com/repos/steamships/n2-plugins/branches?per_page=100&page=${page}`;
      const options = {
        headers: {
          Authorization: `token ${token}`,
          Accept: "application/vnd.github.v3+json"
        },
        muteHttpExceptions: true
      };
      
      const response = UrlFetchApp.fetch(branchesUrl, options);
      const branches = JSON.parse(response.getContentText());
      
      if (branches.length === 0) break;
      
      allBranches = allBranches.concat(branches);
      page++;
      
      // API制限を考慮して少し待機
      Utilities.sleep(1000);
    }
    
    // データ収集の一括処理用の配列
    const promises = allBranches.map((branch, index) => {
      return {
        branch: branch,
        indexPhpUrl: `https://api.github.com/repos/steamships/n2-plugins/contents/index.php?ref=${encodeURIComponent(branch.name)}`,
        commitUrl: branch.commit.url,
        index: index
      };
    });

    // 10件ずつの一括処理用にチャンク分割
    const chunkSize = 10;
    const chunks = [];
    for (let i = 0; i < promises.length; i += chunkSize) {
      chunks.push(promises.slice(i, i + chunkSize));
    }

    let allData = [];
    
    // チャンクごとに一括処理
    for (let i = 0; i < chunks.length; i++) {
      const chunk = chunks[i];
      const processedCount = i * chunkSize;
      
      // 進捗状況を表示
      sheet.getRange('A1').setValue(`GitHub Plugin Branches (データ収集中... ${processedCount}/${allBranches.length})`);
      SpreadsheetApp.flush();

      // 並列でリクエストを実行
      const responses = chunk.map(item => {
        const options = {
          headers: {
            Authorization: `token ${token}`,
            Accept: "application/vnd.github.v3.raw"
          },
          muteHttpExceptions: true
        };

        // index.phpのみ取得
        const indexPhpUrl = `https://api.github.com/repos/steamships/n2-plugins/contents/index.php?ref=${encodeURIComponent(item.branch.name)}`;
        return {
          branch: item.branch,
          indexPhp: UrlFetchApp.fetch(indexPhpUrl, options)
        };
      });

      // レスポンスを処理
      for (const response of responses) {
        const branch = response.branch;
        let pluginName = '-';
        let description = '-';
        let hooks = [];
        let uniqueHooks = [];
        let aiResponse = '';
        let hookInfo = { hooks: [], namespaces: [], callbacks: [] };

        // index.phpの処理
        if (response.indexPhp.getResponseCode() === 200) {
          const content = response.indexPhp.getContentText();
          const pluginNameMatch = content.match(/Plugin Name:\s*(.+)/);
          const descriptionMatch = content.match(/Description:\s*(.+)/);
          
          if (pluginNameMatch) pluginName = pluginNameMatch[1].trim();
          if (descriptionMatch) description = descriptionMatch[1].trim();

          // フック名とコールバック関数を抽出
          const hookTypes = {
            'add_action': 'アクション追加',
            'add_filter': 'フィルター追加',
            'do_action': 'アクション実行',
            'apply_filters': 'フィルター適用'
          };

          let allHooks = [];
          let hookCallbacks = new Map(); // フックとコールバックの対応を保存

          for (const [type, japaneseType] of Object.entries(hookTypes)) {
            const pattern = new RegExp(`${type}\\s*\\(\\s*['"]([^'"]+)['"]\\s*,\\s*['"]?([^'",)]+)['"]?`, 'g');
            let match;
            while ((match = pattern.exec(content)) !== null) {
              const hookName = `${japaneseType}: ${match[1]}`;
              const fullCallback = match[2];
              
              // コールバック関数とネームスペースを分離
              let namespace = '-';
              let callback = fullCallback;
              
              if (fullCallback.includes('\\')) {
                const parts = fullCallback.split('\\');
                namespace = parts[0];
                callback = parts[parts.length - 1];
              }
              
              allHooks.push(hookName);
              hookCallbacks.set(hookName, { namespace, callback });
            }
          }
          uniqueHooks = [...new Set(allHooks)];

          // コールバック関数の実装を探す
          let callbackImplementations = new Map();
          for (const callback of hookCallbacks.values()) {
            try {
              // エスケープが必要な文字をエスケープ
              const escapedCallback = callback.callback.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
              
              // 正規表現を修正
              const functionPattern = new RegExp(
                `function\\s+${escapedCallback}\\s*\\([^)]*\\)\\s*{([^}]+)}`,
                'gs'
              );
              
              const matches = content.match(functionPattern);
              if (matches && matches.length > 0) {
                // 最初のマッチを使用
                const implementation = matches[0]
                  .replace(new RegExp(`function\\s+${escapedCallback}\\s*\\([^)]*\\)\\s*{`), '')
                  .replace(/}\s*$/, '')
                  .trim();
                
                callbackImplementations.set(callback.callback, implementation);
              }
            } catch (e) {
              Logger.log(`関数 ${callback.callback} の解析中にエラーが発生: ${e.message}`);
              continue;
            }
          }

          // すべての関数を抽出（G列のコールバック関数以外も含む）
          const allFunctions = new Map();
          
          // G列のコールバック関数を先に処理
          for (const callback of hookCallbacks.values()) {
            if (!allFunctions.has(callback.callback)) {
              allFunctions.set(callback.callback, {
                isCallback: true,
                implementation: callbackImplementations.get(callback.callback) || ''
              });
            }
          }

          // その他の関数を抽出
          const functionPattern = /function\s+([a-zA-Z0-9_]+)\s*\([^)]*\)\s*{([^}]+)}/gs;
          let match;
          while ((match = functionPattern.exec(content)) !== null) {
            const funcName = match[1];
            const implementation = match[2].trim();
            if (!allFunctions.has(funcName)) {
              allFunctions.set(funcName, {
                isCallback: false,
                implementation: implementation
              });
            }
          }

          // AIによるコード解析を修正
          const aiPrompt = `あなたはWordPressプラグインの専門家として、以下のコードの処理内容を解析してください。
厳密にJSON形式で返してください。それ以外の追加テキストは含めないでください。

解析対象のコード：
${[...allFunctions.entries()].map(([funcName, data]) => 
  `■ ${funcName}:\n${data.implementation}`
).join('\n\n')}

出力形式：
{
  "functions": [
    {
      "name": "関数名",
      "isCallback": true/false,
      "summary": "処理概要（100文字以内）"
    }
  ]
}

※関数が存在しない場合は、以下の形式で出力してください：
{
  "functions": [],
  "summary": "ソース全体の処理概要（100文字以内）"
}

※処理概要は技術的な観点で具体的に記載してください
※必ず有効なJSON形式で出力してください`;

          // AIの応答をJSONとしてパース
          let functionAnalysis;
          let jsonResponse;
          try {
            jsonResponse = analyzeWithAI(aiPrompt);
            Logger.log('AI Response:', jsonResponse);
            
            // レスポンスから余分な文字を除去
            const cleanedResponse = jsonResponse.replace(/^[^{]*/, '').replace(/[^}]*$/, '');
            
            functionAnalysis = JSON.parse(cleanedResponse);
            
            if (!functionAnalysis || !functionAnalysis.functions) {
              throw new Error('Invalid response format');
            }
            
            // 解析結果を指定された形式に整形
            if (functionAnalysis.functions.length > 0) {
              aiResponse = functionAnalysis.functions.map(func => 
                `関数名:${func.name}\n（${func.summary}）`
              ).join('\n\n');
            } else if (functionAnalysis.summary) {
              aiResponse = `処理概要:\n（${functionAnalysis.summary}）`;
            } else {
              throw new Error('No functions or summary found');
            }
            
          } catch (e) {
            Logger.log(`AI応答のパースに失敗: ${e.message}`);
            Logger.log('Raw AI Response:', jsonResponse);
            aiResponse = '※ 関数の解析に失敗しました。管理者に確認してください。';
          }

          // フックとコールバック関数の情報を更新
          hookInfo = {
            hooks: uniqueHooks,
            namespaces: [...hookCallbacks.values()].map(v => v.namespace),
            callbacks: [...hookCallbacks.values()].map(v => v.callback)
          };
        }

        // 分類を判定
        let category = 'その他';
        if (description !== '-') {
          if (/[都道府県市町村]/.test(description)) {
            category = '自治体用';
          } else if (description.includes('ポータル拡張')) {
            category = 'ポータル用';
          }
        }

        // 必要なデータのみ追加（branchUrlは残す）
        allData.push({
          branch: branch.name,
          branchUrl: `https://github.com/steamships/n2-plugins/tree/${encodeURIComponent(branch.name)}`,
          pluginName: pluginName,
          category: category,
          hooks: uniqueHooks.join('\n'),
          namespaces: hookInfo.namespaces.join('\n'),
          callbacks: hookInfo.callbacks.join('\n'),
          analysis: aiResponse
        });
      }

      // API制限を考慮して少し待機
      Utilities.sleep(1000);
    }

    // ソート前の状態表示
    sheet.getRange('A1').setValue('GitHub Plugin Branches (データ整理中...)');
    SpreadsheetApp.flush();

    // カテゴリの優先順位を定義
    const categoryOrder = {
      'ポータル用': 1,
      '自治体用': 2,
      'その他': 3
    };

    // ソート
    allData.sort((a, b) => categoryOrder[a.category] - categoryOrder[b.category]);

    // 書き込みフェーズ
    sheet.getRange('A1').setValue('GitHub Plugin Branches (データ書き込み中...)');
    SpreadsheetApp.flush();
    
    let currentRow = 3;
    for (const data of allData) {
      // リッチテキストでリンクを設定
      const richTextValue = SpreadsheetApp.newRichTextValue()
        .setText(data.branch)
        .setLinkUrl(data.branchUrl)
        .build();

      // データを設定（7列に修正）
      const range = sheet.getRange(currentRow, 1, 1, 7);
      
      // フックとコールバック関数を整形
      let formattedHooks = '';
      let formattedCallbacks = '';
      if (data.hooks) {
        formattedHooks = data.hooks.split('\n').map(hook => `- ${hook}`).join('\n');
        formattedCallbacks = data.callbacks;
      } else {
        formattedHooks = '- なし';
        formattedCallbacks = '- なし';
      }

      range.setValues([[
        data.branch,
        data.pluginName,
        data.category,
        data.analysis,
        formattedHooks,
        data.namespaces || '-',
        data.callbacks || '-'
      ]]);

      // セルの書式設定も修正
      sheet.getRange(currentRow, 4, 1, 4).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
      sheet.getRange(currentRow, 4, 1, 4).setVerticalAlignment('top');

      // ブランチ名をリッチテキストで設定
      sheet.getRange(currentRow, 1).setRichTextValue(richTextValue);

      // カテゴリに応じて背景色を設定
      const backgroundColor = data.category === 'ポータル用' ? '#e6e6fa' :  // 薄い紫
                            data.category === '自治体用' ? '#fff2cc' :      // 薄い黄色
                            '#e6f3ff';                                      // 薄い青
      range.setBackground(backgroundColor);

      currentRow++;
    }

    // 最終更新
    sheet.getRange('A1').setValue('GitHub Plugin Branches');
    SpreadsheetApp.flush();
    
    // 罫線を追加（7列に修正）
    const lastRow = sheet.getLastRow();
    if (lastRow > 2) {
      sheet.getRange(2, 1, lastRow - 1, 7).setBorder(
        true, // top
        true, // left
        true, // bottom
        true, // right
        true, // vertical
        true  // horizontal
      );
    }
    
  } catch (error) {
    sheet.getRange('A1').setValue('GitHub Plugin Branches (エラーが発生しました)');
    SpreadsheetApp.flush();
    Logger.log(`エラーが発生しました: ${error.message}`);
    throw new Error(`ブランチ情報の取得に失敗しました: ${error.message}`);
  }
}

// AIによる解析を行う関数
function analyzeWithAI(prompt) {
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
    if (!apiKey) return '※ AI解析にはAPIキーが必要です';

    const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${apiKey}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        model: "chatgpt-4o-latest",
        messages: [
          {
            role: "system",
            content: "あなたはWordPressプラグインの専門家です。プラグインコードを解析する際は、WordPress固有の機能（フック、オプション、投稿タイプ、データベース操作など）に注目し、技術的な観点から重要な処理を簡潔に説明してください。"
          },
          {
            role: "user",
            content: prompt
          }
        ],
        temperature: 0,
        max_tokens: 15000
      })
    });

    const result = JSON.parse(response.getContentText());
    return result.choices[0].message.content.trim();
  } catch (error) {
    Logger.log(`AI解析エラー: ${error.message}`);
    return '※ AI解析エラー';
  }
} 