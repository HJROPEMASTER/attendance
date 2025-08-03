function exportSheetToGitHub() {
  const props = PropertiesService.getScriptProperties();
  const TOKEN = props.getProperty('GITHUB_TOKEN');
  if (!TOKEN) throw new Error("❗ GitHub token not set. See step below.");

  const REPO = "hjropemaster/sheet-export"; // change to your repo
  const BRANCH = "main"; // or master
  const FILE_PATH = "data.csv"; // file name in GitHub

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const csv = data.map(row =>
    row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(',')
  ).join('\n');
  const encodedContent = Utilities.base64Encode(csv);

  const headers = {
    Authorization: `Bearer ${TOKEN}`,
    Accept: "application/vnd.github.v3+json"
  };

  const url = `https://api.github.com/repos/${REPO}/contents/${FILE_PATH}`;

  // Get file SHA if it already exists
  let sha = null;
  try {
    const getResp = UrlFetchApp.fetch(`${url}?ref=${BRANCH}`, { headers });
    sha = JSON.parse(getResp).sha;
  } catch (_) { /* ignore 404 */ }

  const payload = {
    message: `Updated from Google Sheets at ${new Date().toISOString()}`,
    content: encodedContent,
    branch: BRANCH,
    ...(sha && { sha })
  };

  const response = UrlFetchApp.fetch(url, {
    method: "put",
    headers,
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  const result = response.getContentText();
  if (code >= 200 && code < 300) {
    Logger.log("✅ Success:\n" + result);
  } else {
    throw new Error("❌ GitHub error:\n" + result);
  }
}
