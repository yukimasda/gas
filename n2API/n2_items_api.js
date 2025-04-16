/**返礼品取得API 
 * getallがtrueのとき、全返礼品取得します
 * getallがfalseのとき、指定された返礼品コードだけ取得します。
 */
function n2_items_api(getAll = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('n2_items') || ss.insertSheet('n2_items');
  
  // targetシートから自治体コードと返礼品コードを取得
  const targetSheet = ss.getSheetByName('target');
  if (!targetSheet) {
    SpreadsheetApp.getUi().alert("エラー: targetシートが見つかりません");
    return;
  }

  // 自治体コードを取得（A2セル）
  const localGovCode = targetSheet.getRange("A2").getValue();
  if (!localGovCode || localGovCode.toString().trim() === "") {
    SpreadsheetApp.getUi().alert("エラー: 自治体コードが入力されていません");
    return;
  }

  // 返礼品コードを取得（B2セル以降）
  const lastRow = targetSheet.getRange("B:B").getValues().filter(String).length;
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("エラー: 返礼品コードが入力されていません");
    return;
  }

  const skuValues = targetSheet.getRange("B2:B" + lastRow).getValues().flat();
  const validSkuValues = skuValues.filter(sku => sku && sku.toString().trim() !== "");

  if (validSkuValues.length === 0) {
    SpreadsheetApp.getUi().alert("エラー: 有効な返礼品コードが入力されていません");
    return;
  }

  // 認証情報
  const user = PropertiesService.getScriptProperties().getProperty('BASIC_USER');
  const pass = PropertiesService.getScriptProperties().getProperty('BASIC_PASS');
  const headersAuth = {
    "Authorization": "Basic " + Utilities.base64Encode(user + ":" + pass)
  };

  let json;
  if (getAll) {
    const url = `https://n2.steamship.co.jp/${localGovCode}/wp-admin/admin-ajax.php?action=n2_items_api&mode=json`;
    const response = UrlFetchApp.fetch(url, {
      "method": "get",
      "headers": headersAuth,
      "muteHttpExceptions": true
    });

    if (response.getResponseCode() !== 200) {
      SpreadsheetApp.getUi().alert("エラー: データの取得に失敗しました");
      return;
    }

    json = JSON.parse(response.getContentText());
  } else {
    // 各返礼品コードに対してAPIリクエストを実行
    let allItems = [];
    for (const sku of validSkuValues) {
      const url = `https://n2.steamship.co.jp/${localGovCode}/wp-admin/admin-ajax.php?action=n2_items_api&mode=json&code=${sku}`;
      
      Logger.log(`APIリクエスト: ${url}`); // URLをログに出力
      
      const response = UrlFetchApp.fetch(url, {
        "method": "get",
        "headers": headersAuth,
        "muteHttpExceptions": true
      });

      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      Logger.log(`レスポンスコード: ${responseCode}`);
      Logger.log(`レスポンス内容: ${responseText}`);

      if (responseCode === 200) {
        try {
          const responseJson = JSON.parse(responseText);
          Logger.log(`パース後のJSONオブジェクト: ${JSON.stringify(responseJson)}`);
          
          if (responseJson && responseJson.items) {
            const items = Array.isArray(responseJson.items) ? responseJson.items : [responseJson.items];
            Logger.log(`取得したアイテム数: ${items.length}`);
            allItems = allItems.concat(items);
          } else {
            Logger.log(`警告: レスポンスにitemsプロパティがありません: ${JSON.stringify(responseJson)}`);
          }
        } catch (e) {
          Logger.log(`エラー: JSONのパースに失敗: ${e.message}`);
        }
      } else {
        Logger.log(`エラー: APIリクエスト失敗（コード: ${responseCode}）`);
      }
      
      Utilities.sleep(1000);
    }

    json = {
      items: allItems
    };
    
    Logger.log(`全取得アイテム数: ${allItems.length}`);
  }

  if (!json) {
    SpreadsheetApp.getUi().alert("エラー: レスポンスがnullまたはundefinedです");
    return;
  }

  // itemsが存在するか確認
  if (!json.items) {
    SpreadsheetApp.getUi().alert("エラー: itemsプロパティが存在しません");
    return;
  }

  // itemsが配列であるか確認
  if (!Array.isArray(json.items)) {
    SpreadsheetApp.getUi().alert("エラー: itemsプロパティは配列ではありません。実際の値: " + JSON.stringify(json.items));
    return;
  }

  // 返礼品コードを指定した場合の処理
  if (!getAll) {
    // レスポンスの構造を確認
    if (typeof json.items === 'object' && !Array.isArray(json.items)) {
      // オブジェクトの場合は配列に変換
      json.items = [json.items];
    } else if (typeof json.items === 'string') {
      // 文字列の場合はJSONとしてパースを試みる
      try {
        const parsedItem = JSON.parse(json.items);
        json.items = [parsedItem];
      } catch (e) {
        json.items = [{ code: json.items }];
      }
    }
  }

  // itemsが配列であるか再度確認
  if (!Array.isArray(json.items)) {
    SpreadsheetApp.getUi().alert("エラー: itemsプロパティは配列ではありません。実際の値: " + JSON.stringify(json.items));
    return;
  }

  // アイテムが空かチェック
  if (json.items.length === 0) {
    SpreadsheetApp.getUi().alert("エラー: 検索結果が0件です。返礼品コードを確認してください。");
    return;
  }

  // 返礼品コードをログに出力
  const giftCodes = json.items.map(item => item.code); // 返礼品コードを抽出
  Logger.log(giftCodes); // 返礼品コードをログに出力

  const data = json.items;
  
  // データの内容をログに出力（デバッグ用）
  Logger.log("データの内容: " + JSON.stringify(data[0]));
  
  // 最初のアイテムにデータがあるか確認
  if (!data[0] || typeof data[0] !== 'object') {
    SpreadsheetApp.getUi().alert("エラー: 返礼品データの形式が不正です。");
    return;
  }
  
  const headers = Object.keys(data[0]);

  // ヘッダーをA列から書き込む
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // データをA列から書き込む
  const values = data.map(item => headers.map(header => item[header] || ""));
  sheet.getRange(2, 1, values.length, values[0].length).setValues(values);

  SpreadsheetApp.getUi().alert("データの取得が完了しました！");
}

function getAllGifts() {
  n2_items_api(true);
}