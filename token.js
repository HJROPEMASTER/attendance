// Configure these with your GitHub details
const GITHUB_USERNAME = 'hjropemaster'; // or organization name if it's under an org
const REPO = 'attendance';
const BRANCH = 'main';
const FILE_PATH = 'data/sheet-data.csv'; // path in your GitHub repo (prefer subfolder)

function exportSheetToGitHub() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();

  // Filter out any rows that may contain secrets like GitHub tokens
  const filteredData = data.filter(row =>
    !row.some(cell => typeof cell === 'string' && cell.includes('ghp_'))
  );
  const csv = filteredData.map(row => row.join(',')).join('\n');

  const url = `https://api.github.com/repos/${GITHUB_USERNAME}/${REPO}/contents/${FILE_PATH}`;

  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  Logger.log("GitHub Token: " + token); // Debug: check if token is being loaded

  // First, get the current file SHA
  const getOptions = {
    method: 'get',
    headers: {
      Authorization: 'token ' + token
    },
    muteHttpExceptions: true
  };

  const getRes = UrlFetchApp.fetch(url, getOptions);
  const getResCode = getRes.getResponseCode();

  let sha = null;
  if (getResCode === 200) {
    const content = JSON.parse(getRes.getContentText());
    sha = content.sha;
  } else if (getResCode !== 404) {
    Logger.log("Error fetching file: " + getRes.getContentText());
    return;
  }

  // Now prepare the PUT payload
  const payload = {
    message: '📤 Update sheet CSV from Google Sheets',
    content: Utilities.base64Encode(csv),
    branch: BRANCH
  };
  if (sha) payload.sha = sha;

  const putOptions = {
    method: 'put',
    contentType: 'application/json',
    headers: {
      Authorization: 'token ' + token
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const res = UrlFetchApp.fetch(url, putOptions);
  Logger.log(res.getContentText());
}
