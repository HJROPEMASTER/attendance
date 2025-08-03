// === CONFIGURATION ===
const GITHUB_USERNAME = 'hjropemaster';
const REPO_NAME = 'attendance';
const BRANCH = 'main';
const EXPORT_FOLDER = 'exports'; // GitHub folder path
const FILE_PREFIX = 'sheet_';   // Optional: file name prefix

// === MENU SETUP ===
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("📤 GitHub Tools")
    .addItem("Export Sheet to GitHub", "exportSheetToGitHub")
    .addItem("🔍 Run Diagnostics", "runFullDiagnostic")
    .addToUi();
}

// === EXPORT FUNCTION ===
function exportSheetToGitHub() {
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) return SpreadsheetApp.getUi().alert('❌ No GITHUB_TOKEN set in script properties.');

  const dateStr = new Date().toISOString().slice(0, 10);
  const fileName = `${EXPORT_FOLDER}/${FILE_PREFIX}${dateStr}.csv`;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const csv = data.map(row => row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(',')).join('\n');
  const encoded = Utilities.base64Encode(csv);

  const url = `https://api.github.com/repos/${GITHUB_USERNAME}/${REPO_NAME}/contents/${fileName}`;
  const headers = {
    Authorization: `Bearer ${token}`,
    Accept: "application/vnd.github.v3+json",
    'User-Agent': 'Google-Apps-Script'
  };

  let sha = null;
  try {
    const existing = UrlFetchApp.fetch(`${url}?ref=${BRANCH}`, { headers, muteHttpExceptions: true });
    if (existing.getResponseCode() === 200) {
      sha = JSON.parse(existing.getContentText()).sha;
    }
  } catch (_) {}

  const payload = {
    message: `📤 Exported from Google Sheets on ${new Date().toISOString()}`,
    content: encoded,
    branch: BRANCH,
    ...(sha && { sha })
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'put',
    contentType: 'application/json',
    headers,
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  const resText = response.getContentText();
  if (code >= 200 && code < 300) {
    SpreadsheetApp.getUi().alert(`✅ Exported to GitHub:\n📁 ${fileName}`);
  } else {
    SpreadsheetApp.getUi().alert(`❌ Export failed (${code}):\n${resText}`);
  }
}

// === DIAGNOSTIC FUNCTIONS ===

function runFullDiagnostic() {
  console.log('🚀 Running GitHub Diagnostic...');
  if (!testGitHubTokenBasic()) return;
  testRepositoryAccess();
  listMyRepositories();
  testExactAPICall();
}

function testGitHubTokenBasic() {
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) {
    console.error('❌ No GitHub token set');
    return false;
  }

  const options = {
    method: 'GET',
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/vnd.github.v3+json",
      'User-Agent': 'Google-Apps-Script'
    },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch("https://api.github.com/user", options);
  const code = response.getResponseCode();
  const text = response.getContentText();

  if (code === 200) {
    const user = JSON.parse(text);
    console.log(`✅ Token valid. Logged in as: ${user.login}`);
    return true;
  } else {
    console.error(`❌ Invalid token (${code}): ${text}`);
    return false;
  }
}

function testRepositoryAccess() {
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  const repoURL = `https://api.github.com/repos/${GITHUB_USERNAME}/${REPO_NAME}`;

  const options = {
    method: 'GET',
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/vnd.github.v3+json",
      'User-Agent': 'Google-Apps-Script'
    },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(repoURL, options);
  const code = response.getResponseCode();
  const body = response.getContentText();

  if (code === 200) {
    console.log(`✅ Repo found: ${GITHUB_USERNAME}/${REPO_NAME}`);
    testWriteAccess(GITHUB_USERNAME, REPO_NAME, token);
  } else if (code === 404) {
    console.log(`❌ Repo not found: ${GITHUB_USERNAME}/${REPO_NAME}`);
  } else {
    console.log(`⚠️ Repo access error (${code}): ${body}`);
  }
}

function testWriteAccess(username, repo, token) {
  const url = `https://api.github.com/repos/${username}/${repo}/contents/test-write-${Date.now()}.txt`;
  const options = {
    method: 'GET',
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/vnd.github.v3+json",
      'User-Agent': 'Google-Apps-Script'
    },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  if (code === 404) {
    console.log('✅ Write access test passed (expected 404)');
  } else if (code === 403) {
    console.log('❌ No write access (403 Forbidden)');
  } else {
    console.log(`⚠️ Unexpected response: ${code}`);
  }
}

function listMyRepositories() {
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  const url = 'https://api.github.com/user/repos?per_page=100';

  const options = {
    method: 'GET',
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/vnd.github.v3+json",
      'User-Agent': 'Google-Apps-Script'
    },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  if (code !== 200) {
    console.error(`❌ Failed to list repos (${code})`);
    return;
  }

  const repos = JSON.parse(response.getContentText());
  console.log(`✅ Found ${repos.length} repos`);
  repos.forEach(r => {
    console.log(`- ${r.full_name} (${r.private ? 'private' : 'public'})`);
  });
}

function testExactAPICall() {
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  const filePath = `${EXPORT_FOLDER}/test-dummy.csv`;
  const url = `https://api.github.com/repos/${GITHUB_USERNAME}/${REPO_NAME}/contents/${filePath}`;

  const options = {
    method: 'GET',
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/vnd.github.v3+json",
      'User-Agent': 'Google-Apps-Script'
    },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  const text = response.getContentText();

  console.log(`📡 API call to ${url}`);
  console.log(`📄 Response Code: ${code}`);
  console.log(`📄 Body: ${text}`);
}
