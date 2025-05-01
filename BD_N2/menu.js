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
    .addItem('ソース概要AI要約', 'summarizeFilesWithAI')
    .addToUi();

  ui.createMenu('📎hook list')
    .addItem('フックを検索', 'fetchHooksFromGitHub')
    .addItem('フック使用箇所を検索', 'searchHookUsages')
    .addItem('フック使用箇所を検索(WordPress)', 'searchHookUsagesWP')
    .addItem('AIで役割を分析', 'analyzeHooksWithAI')
    .addToUi();

  ui.createMenu('🧠AI解析')
    .addItem('ファイル一覧からList作成', 'analyzeSourcesWithAI')
    .addToUi();
}
