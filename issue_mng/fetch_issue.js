function fetchGitHubIssues() {
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("fetch");

  const url = `https://api.github.com/repos/${GITHUB_REPO}/issues?state=all&per_page=100`;
  
  const options = {
    method: 'GET',
    headers: {
      Authorization: 'token ' + GITHUB_TOKEN,
      Accept: 'application/vnd.github+json'
    },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response.getContentText());

  // ヘッダーを変更して、マイルストーンをラベルの後に追加
  const headers = ['番号', 'タイトル', 'アサイン', 'ラベル', 'マイルストーン', 'ステータス', '作成日時'];
  sheet.clear();
  sheet.appendRow(headers);

  json.forEach(issue => {
    if (issue.pull_request) return;

    const number = issue.number;
    const title = issue.title;
    const assignees = (issue.assignees || []).map(a => a.login).join(', ');
    const labels = (issue.labels || []).map(l => l.name).join(', ');
    const state = issue.state;
    const createdAt = issue.created_at;
    const issueUrl = `https://github.com/${GITHUB_REPO}/issues/${number}`;

    const milestone = issue.milestone ? issue.milestone.title : 'なし';

    const issueLink = `=HYPERLINK("${issueUrl}", "#${number}")`;

    // マイルストーンをラベルの後に追加
    sheet.appendRow([
      issueLink,
      title,
      assignees,
      labels,
      milestone,
      state,
      createdAt
    ]);
  });
}
