/**
 * GitHub APIトークンを設定するための関数
 */
function setGitHubToken() {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
      'GitHub APIトークンの設定',
      'GitHubの個人用アクセストークンを入力してください:',
      ui.ButtonSet.OK_CANCEL
    );
  
    if (response.getSelectedButton() == ui.Button.OK) {
      const token = response.getResponseText().trim();
      PropertiesService.getScriptProperties().setProperty('GITHUB_TOKEN', token);
      ui.alert('GitHubトークンが正常に保存されました。');
    }
  }