function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // 🔍検索ツール メニュー
  ui.createMenu('🔍Sheet')
    .addItem('横断検索', 'showSettingsDialog')
    .addItem('INDEX作成', 'createIndex')
    .addToUi();

  // 🌐外部ツール メニュー（Github用）
  ui.createMenu('🌲GitHub')
    .addItem('APIトークンを設定', 'setGitHubToken')
    .addItem('GitHub Search', 'gitHub_Search')
    .addItem('GitHub Src Links', 'getAllFiles')
    .addToUi();

  // 🔍フック解析 メニュー
  ui.createMenu('🧠AIフック解析')
    .addItem('フックを検索', 'fetchHooksFromGitHub')
    .addItem('AIで役割を分析', 'analyzeHooksWithAI')
    .addToUi();
}
