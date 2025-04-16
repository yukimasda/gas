function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('GitHub Issue')
    .addItem('GitHub Issue 作成', 'createGitHubIssues')
    .addItem('GitHub Issue 取得', 'fetchGitHubIssues') 
    .addToUi();
}
