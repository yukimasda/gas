function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('🌲GitHub')
    .addItem('APIトークンを設定', 'setGitHubToken')
    .addItem('GitHub Plugin Branches', 'getAllBranches')
    .addToUi();
}
