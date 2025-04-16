const GITHUB_TOKEN = '';
const GITHUB_REPO = 'steamships/neo-neng'; // ユーザー名/リポジトリ名


function createGitHubIssues() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) { // 1行目はヘッダー
    const [shouldCreate, title, assigneesStr, labelsStr, milestoneStr] = data[i];

    if (shouldCreate === true) {
      const assignees = assigneesStr ? assigneesStr.split(',').map(s => s.trim()) : [];
      const labels = labelsStr ? labelsStr.split(',').map(s => s.trim()) : [];
      const milestone = milestoneStr || null; // マイルストーンはnullまたは指定されたもの

      const payload = {
        title: title,
        assignees: assignees,
        labels: labels,
        milestone: milestone ? milestone : undefined // マイルストーンがあれば設定
      };

      const options = {
        method: 'POST',
        contentType: 'application/json',
        headers: {
          Authorization: 'token ' + GITHUB_TOKEN
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };

      const url = `https://api.github.com/repos/${GITHUB_REPO}/issues`;

      try {
        const response = UrlFetchApp.fetch(url, options);
        const responseCode = response.getResponseCode();
        const json = JSON.parse(response.getContentText());

        if (responseCode === 201) {
          const issueNumber = json.number;
          const issueUrl = json.html_url;

          // ステータス（F列）とリンク付きIssue No（G列）をセット
          sheet.getRange(i + 1, 6).setValue('作成済み'); // F列: ステータス
          sheet.getRange(i + 1, 7).setFormula(`=HYPERLINK("${issueUrl}", "#${issueNumber}")`); // G列: リンク付きIssue No.

        } else {
          sheet.getRange(i + 1, 6).setValue('作成失敗');
          Logger.log('GitHub API Error: ' + JSON.stringify(json)); // エラーレスポンスをログに出力
        }
      } catch (error) {
        sheet.getRange(i + 1, 6).setValue('作成失敗');
        Logger.log('Error: ' + error.message); // エラーメッセージをログに出力
      }

      // チェックボックスをFALSEに戻す
      sheet.getRange(i + 1, 1).setValue(false);
    }
  }
}
