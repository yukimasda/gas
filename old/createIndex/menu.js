function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('🔍検索ツール')
      .addItem('横断検索', 'showSettingsDialog')
      .addItem('INDEX作成', 'createIndex')
      .addToUi();
}
