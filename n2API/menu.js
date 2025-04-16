function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu(' N2API')
    .addItem('返礼品コード指定して取得', 'n2_items_api')
    .addItem('全ての返礼品を取得', 'getAllGifts')
    .addToUi();
}
