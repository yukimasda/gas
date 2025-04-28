function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('🔍Gdrive')
    .addItem('横断検索', 'showSettingsDialog')
    .addItem('INDEX作成', 'createIndex')
    .addToUi();

  ui.createMenu('🌲GitHub')
    .addItem('APIトークンを設定', 'setGitHubToken')
    .addItem('GitHub Search', 'gitHub_Search')
    .addItem('GitHub Src Links', 'getAllFiles')
    .addToUi();

  ui.createMenu('📎フック検索')
    .addItem('フックを検索', 'fetchHooksFromGitHub')
    .addItem('フック使用箇所を検索', 'searchHookUsages')
    .addItem('フック使用箇所を検索(WordPress)', 'searchHookUsagesWP')
    .addToUi();

  ui.createMenu('🧠AIフック解析')
    .addItem('AIで役割を分析', 'analyzeHooksWithAI')
    .addItem('AIで仕様書を作成', 'analyzeSourceWithAI')
    .addItem('複数ファイルを解析', 'analyzeMultipleFiles')
    .addToUi();
}
