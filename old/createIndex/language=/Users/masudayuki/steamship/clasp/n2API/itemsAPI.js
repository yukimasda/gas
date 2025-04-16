function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu(' N2API')
    .addItem('⬇️get', 'n2_output_gift_api')
    .addItem('⬇️items API', 'n2_output_items_api')
    .addToUi();
}

function n2_output_items_api() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const headers = ["返礼品コード", "タイトル", "事業者名", "寄附金額", "説明文", "内容量・規格等", "発送方法"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const lastRow = sheet.getLastRow();
  const skuValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();

  const 自治体コード = 'f015814-atsuma';
  const user = PropertiesService.getScriptProperties().getProperty('BASIC_USER');
  const pass = PropertiesService.getScriptProperties().getProperty('BASIC_PASS');
  const headersAuth = {
    "Authorization": "Basic " + Utilities.base64Encode(user + ":" + pass)
  };

  skuValues.forEach((sku, index) => {
    if (!sku) return;

    const url = `https://n2.steamship.co.jp/${自治体コード}/wp-admin/admin-ajax.php?action=n2_items_api&mode=json&code=${sku}`;
    const options = {
      "method": "get",
      "headers": headersAuth,
      "muteHttpExceptions": true
    };

    try {
      const response = UrlFetchApp.fetch(url, options);
      const code = response.getResponseCode();

      if (code !== 200) {
        sheet.getRange(index + 2, 2).setValue(`エラー: HTTP ${code}`);
        return;
      }

      const json = JSON.parse(response.getContentText());
      const data = json.data;

      const values = [
        sku,
        data.title || "",
        data["事業者名"] || "",
        data["寄附金額"] || "",
        data["説明文"] || "",
        data["内容量・規格等"] || "",
        data["発送方法"] || ""
      ];

      sheet.getRange(index + 2, 1, 1, values.length).setValues([values]);

    } catch (e) {
      sheet.getRange(index + 2, 2).setValue(`エラー: ${e.message}`);
    }
  });

  SpreadsheetApp.getUi().alert("全てのデータ取得が完了しました！");
} 