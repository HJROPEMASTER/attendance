/**
 * GitHub Access Diagnostic Script
 * 
 * This script helps diagnose why you're getting "Repository not found" errors
 * Run these functions one by one to identify the issue
 */

// Test different repository configurations
const TEST_CONFIGS = [
  { username: 'hjropemaster', repo: 'attendance' },
  { username: 'HJROPEMASTER', repo: 'attendance' },
  { username: 'hjropemaster', repo: 'ATTENDANCE' },
  { username: 'HJROPEMASTER', repo: 'ATTENDANCE' }
];

/**
 * Step 1: Test if your GitHub token is valid at all
 */
function testGitHubTokenBasic() {
  console.log('🔍 Testing basic GitHub token validity...');
  
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  
  if (!token) {
    console.error('❌ No GitHub token found in Script Properties');
    return false;
  }
  
  console.log(`🔑 Token found, length: ${token.length} characters`);
  console.log(`🔑 Token starts with: ${token.substring(0, 10)}...`);
  
  // Test basic GitHub API access
  const url = 'https://api.github.com/user';
  
  const options = {
    method: 'GET',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Accept': 'application/vnd.github.v3+json',
      'User-Agent': 'Google-Apps-Script-Diagnostic'
    },
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    console.log(`📡 GitHub API Response Code: ${responseCode}`);
    
    if (responseCode === 200) {
      const user = JSON.parse(responseText);
      console.log('✅ Token is valid!');
      console.log(`👤 Authenticated as: ${user.login}`);
      console.log(`📧 Email: ${user.email || 'Not public'}`);
      console.log(`🔒 Token scopes: Check your token permissions`);
      return true;
    } else if (responseCode === 401) {
      console.error('❌ Token is invalid or expired');
      console.error('Response:', responseText);
      return false;
    } else {
      console.error(`❌ Unexpected response: ${responseCode}`);
      console.error('Response:', responseText);
      return false;
    }
    
  } catch (error) {
    console.error('❌ Error testing token:', error.message);
    return false;
  }
}

/**
 * Step 2: Test access to your specific repository with different name variations
 */
function testRepositoryAccess() {
  console.log('🔍 Testing repository access with different name variations...');
  
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  
  if (!token) {
    console.error('❌ No GitHub token found');
    return;
  }
  
  for (const config of TEST_CONFIGS) {
    console.log(`\n🧪 Testing: ${config.username}/${config.repo}`);
    
    const url = `https://api.github.com/repos/${config.username}/${config.repo}`;
    
    const options = {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Accept': 'application/vnd.github.v3+json',
        'User-Agent': 'Google-Apps-Script-Diagnostic'
      },
      muteHttpExceptions: true
    };
    
    try {
      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      
      if (responseCode === 200) {
        const repo = JSON.parse(response.getContentText());
        console.log(`✅ SUCCESS! Repository found: ${repo.full_name}`);
        console.log(`🔒 Private: ${repo.private}`);
        console.log(`📝 Description: ${repo.description || 'No description'}`);
        console.log(`🌟 Stars: ${repo.stargazers_count}`);
        console.log(`🔧 Default branch: ${repo.default_branch}`);
        
        // Test write access
        testWriteAccess(config.username, config.repo, token);
        return config;
        
      } else if (responseCode === 404) {
        console.log(`❌ Repository not found: ${config.username}/${config.repo}`);
      } else if (responseCode === 403) {
        console.log(`🔒 Access forbidden: ${config.username}/${config.repo} (private repo or insufficient permissions)`);
      } else {
        console.log(`⚠️ Unexpected response ${responseCode} for: ${config.username}/${config.repo}`);
      }
      
    } catch (error) {
      console.error(`❌ Error testing ${config.username}/${config.repo}:`, error.message);
    }
  }
  
  console.log('\n❌ No accessible repository found with any name variation');
}

/**
 * Step 3: Test write access to the repository
 */
function testWriteAccess(username, repo, token) {
  console.log(`\n🔍 Testing write access to ${username}/${repo}...`);
  
  // Try to get a file that definitely doesn't exist to test write permissions
  const testUrl = `https://api.github.com/repos/${username}/${repo}/contents/test-write-access-${Date.now()}.txt`;
  
  const options = {
    method: 'GET',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Accept': 'application/vnd.github.v3+json',
      'User-Agent': 'Google-Apps-Script-Diagnostic'
    },
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(testUrl, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode === 404) {
      console.log('✅ Write access test passed (404 expected for non-existent file)');
    } else if (responseCode === 403) {
      console.log('❌ Write access denied (403 Forbidden)');
      console.log('🔧 Your token may not have "repo" scope or write permissions');
    } else {
      console.log(`⚠️ Unexpected response for write test: ${responseCode}`);
    }
    
  } catch (error) {
    console.error('❌ Error testing write access:', error.message);
  }
}

/**
 * Step 4: List your accessible repositories
 */
function listMyRepositories() {
  console.log('🔍 Listing repositories accessible with your token...');
  
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  
  if (!token) {
    console.error('❌ No GitHub token found');
    return;
  }
  
  const url = 'https://api.github.com/user/repos?per_page=100&sort=updated';
  
  const options = {
    method: 'GET',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Accept': 'application/vnd.github.v3+json',
      'User-Agent': 'Google-Apps-Script-Diagnostic'
    },
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      const repos = JSON.parse(response.getContentText());
      console.log(`✅ Found ${repos.length} accessible repositories:`);
      
      repos.forEach((repo, index) => {
        console.log(`${index + 1}. ${repo.full_name} (${repo.private ? 'private' : 'public'})`);
        
        // Check if this might be the attendance repo
        if (repo.name.toLowerCase().includes('attendance') || 
            repo.full_name.toLowerCase().includes('attendance')) {
          console.log(`   🎯 POTENTIAL MATCH: ${repo.full_name}`);
        }
      });
      
    } else {
      console.error(`❌ Failed to list repositories: ${responseCode}`);
      console.error(response.getContentText());
    }
    
  } catch (error) {
    console.error('❌ Error listing repositories:', error.message);
  }
}

/**
 * Step 5: Test the exact API call that's failing
 */
function testExactAPICall() {
  console.log('🔍 Testing the exact API call that\'s failing...');
  
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  const username = 'hjropemaster';
  const repo = 'attendance';
  const filePath = 'data/sheet-data.csv';
  
  const url = `https://api.github.com/repos/${username}/${repo}/contents/${filePath}`;
  
  console.log(`🔗 Testing URL: ${url}`);
  
  const options = {
    method: 'GET',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Accept': 'application/vnd.github.v3+json',
      'User-Agent': 'Google-Apps-Script-Diagnostic'
    },
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    console.log(`📡 Response Code: ${responseCode}`);
    console.log(`📄 Response: ${responseText}`);
    
    if (responseCode === 404) {
      console.log('ℹ️ This could mean:');
      console.log('  1. Repository doesn\'t exist');
      console.log('  2. File doesn\'t exist (normal for first run)');
      console.log('  3. No access to repository');
    }
    
  } catch (error) {
    console.error('❌ Error with exact API call:', error.message);
  }
}

/**
 * Run all diagnostic tests in sequence
 */
function runFullDiagnostic() {
  console.log('🚀 Running full GitHub diagnostic...\n');
  
  console.log('='.repeat(50));
  console.log('STEP 1: Testing basic token validity');
  console.log('='.repeat(50));
  const tokenValid = testGitHubTokenBasic();
  
  if (!tokenValid) {
    console.log('\n❌ Token is invalid. Please generate a new GitHub token.');
    return;
  }
  
  console.log('\n' + '='.repeat(50));
  console.log('STEP 2: Testing repository access');
  console.log('='.repeat(50));
  testRepositoryAccess();
  
  console.log('\n' + '='.repeat(50));
  console.log('STEP 3: Listing your repositories');
  console.log('='.repeat(50));
  listMyRepositories();
  
  console.log('\n' + '='.repeat(50));
  console.log('STEP 4: Testing exact failing API call');
  console.log('='.repeat(50));
  testExactAPICall();
  
  console.log('\n🏁 Diagnostic complete! Check the logs above for issues.');
}

