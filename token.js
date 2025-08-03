/**
 * Improved Google Sheets to GitHub CSV Export Script
 * 
 * This script exports Google Sheets data to a GitHub repository as a CSV file
 * with enhanced error handling, security, and logging capabilities.
 * 
 * Author: Manus AI Assistant
 * Version: 2.0
 * Date: 2025-08-03
 */

// Configuration constants
const CONFIG = {
  GITHUB_USERNAME: 'hjropemaster',
  REPO: 'attendance',
  BRANCH: 'main',
  FILE_PATH: 'data/sheet-data.csv', // Will create data folder if it doesn't exist
  SHEET_NAME: '0MAIN', // Specify which sheet to export
  MAX_RETRIES: 3,
  RETRY_DELAY: 1000, // milliseconds
  TIMEOUT: 30000 // 30 seconds
};

/**
 * Main function to export sheet data to GitHub
 * Call this function to trigger the export
 */
function exportSheetToGitHub() {
  console.log('🚀 Starting GitHub export process...');
  
  try {
    // Validate configuration
    validateConfiguration();
    
    // Get and validate data
    const data = getSheetData();
    if (!data || data.length === 0) {
      throw new Error('No data found to export');
    }
    
    // Convert to CSV
    const csv = convertToCSV(data);
    console.log(`📊 Converted ${data.length} rows to CSV format`);
    
    // Export to GitHub
    const result = exportToGitHub(csv);
    
    if (result.success) {
      console.log('✅ Export completed successfully!');
      console.log(`📁 File updated: ${CONFIG.FILE_PATH}`);
      console.log(`🔗 View at: https://github.com/${CONFIG.GITHUB_USERNAME}/${CONFIG.REPO}/blob/${CONFIG.BRANCH}/${CONFIG.FILE_PATH}`);
      
      // Optional: Send email notification
      sendNotificationEmail(true, result.message);
      
    } else {
      throw new Error(result.message);
    }
    
  } catch (error) {
    console.error('❌ Export failed:', error.message);
    console.error('Stack trace:', error.stack);
    
    // Send error notification
    sendNotificationEmail(false, error.message);
    
    // Re-throw for debugging
    throw error;
  }
}

/**
 * Validates the configuration settings
 */
function validateConfiguration() {
  console.log('🔍 Validating configuration...');
  
  if (!CONFIG.GITHUB_USERNAME || !CONFIG.REPO || !CONFIG.BRANCH || !CONFIG.FILE_PATH) {
    throw new Error('Missing required configuration values');
  }
  
  // Check if GitHub token is properly set
  const token = getGitHubToken();
  if (!token) {
    throw new Error('GitHub token not found. Please set GITHUB_TOKEN in Script Properties.');
  }
  
  if (token.length < 20) {
    throw new Error('GitHub token appears to be invalid (too short)');
  }
  
  console.log('✅ Configuration validated');
}

/**
 * Retrieves and validates sheet data
 */
function getSheetData() {
  console.log(`📋 Getting data from sheet: ${CONFIG.SHEET_NAME}`);
  
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet;
    
    // Try to get specific sheet by name, fallback to active sheet
    if (CONFIG.SHEET_NAME) {
      sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
      if (!sheet) {
        console.warn(`⚠️ Sheet "${CONFIG.SHEET_NAME}" not found, using active sheet`);
        sheet = spreadsheet.getActiveSheet();
      }
    } else {
      sheet = spreadsheet.getActiveSheet();
    }
    
    const range = sheet.getDataRange();
    if (!range) {
      throw new Error('No data range found in sheet');
    }
    
    const data = range.getValues();
    
    // Filter out rows that might contain sensitive information
    const filteredData = data.filter(row => {
      return !row.some(cell => {
        if (typeof cell === 'string') {
          const cellStr = cell.toLowerCase();
          // Filter out rows containing tokens, passwords, or other sensitive data
          return cellStr.includes('ghp_') || 
                 cellStr.includes('github_pat_') ||
                 cellStr.includes('password') ||
                 cellStr.includes('secret') ||
                 cellStr.includes('token');
        }
        return false;
      });
    });
    
    console.log(`📊 Retrieved ${data.length} total rows, ${filteredData.length} after filtering`);
    
    if (filteredData.length !== data.length) {
      console.warn(`⚠️ Filtered out ${data.length - filteredData.length} rows containing sensitive data`);
    }
    
    return filteredData;
    
  } catch (error) {
    throw new Error(`Failed to get sheet data: ${error.message}`);
  }
}

/**
 * Converts array data to properly formatted CSV
 */
function convertToCSV(data) {
  console.log('🔄 Converting data to CSV format...');
  
  try {
    const csvRows = data.map(row => {
      return row.map(cell => {
        // Handle different data types
        let value = '';
        
        if (cell === null || cell === undefined) {
          value = '';
        } else if (cell instanceof Date) {
          value = Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
        } else {
          value = String(cell);
        }
        
        // Escape CSV special characters
        if (value.includes(',') || value.includes('"') || value.includes('\n') || value.includes('\r')) {
          value = '"' + value.replace(/"/g, '""') + '"';
        }
        
        return value;
      }).join(',');
    });
    
    return csvRows.join('\n');
    
  } catch (error) {
    throw new Error(`Failed to convert to CSV: ${error.message}`);
  }
}

/**
 * Exports CSV data to GitHub repository
 */
function exportToGitHub(csvContent) {
  console.log('📤 Exporting to GitHub...');
  
  const token = getGitHubToken();
  const url = `https://api.github.com/repos/${CONFIG.GITHUB_USERNAME}/${CONFIG.REPO}/contents/${CONFIG.FILE_PATH}`;
  
  let attempt = 0;
  
  while (attempt < CONFIG.MAX_RETRIES) {
    attempt++;
    console.log(`🔄 Attempt ${attempt}/${CONFIG.MAX_RETRIES}`);
    
    try {
      // First, try to get the current file to get its SHA
      const currentFile = getCurrentFileInfo(url, token);
      
      // Prepare the payload
      const payload = {
        message: `📊 Update attendance data - ${new Date().toISOString()}`,
        content: Utilities.base64Encode(csvContent),
        branch: CONFIG.BRANCH
      };
      
      // Include SHA if file exists (for updates)
      if (currentFile.sha) {
        payload.sha = currentFile.sha;
        console.log('📝 Updating existing file');
      } else {
        console.log('📄 Creating new file');
      }
      
      // Make the API request
      const options = {
        method: 'PUT',
        contentType: 'application/json',
        headers: {
          'Authorization': `Bearer ${token}`,
          'Accept': 'application/vnd.github.v3+json',
          'User-Agent': 'Google-Apps-Script-Attendance-Export'
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };
      
      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      console.log(`📡 GitHub API Response: ${responseCode}`);
      
      if (responseCode === 200 || responseCode === 201) {
        const result = JSON.parse(responseText);
        return {
          success: true,
          message: `File ${responseCode === 201 ? 'created' : 'updated'} successfully`,
          sha: result.content.sha,
          url: result.content.html_url
        };
      } else if (responseCode === 409) {
        console.warn('⚠️ Conflict detected, retrying...');
        if (attempt < CONFIG.MAX_RETRIES) {
          Utilities.sleep(CONFIG.RETRY_DELAY * attempt);
          continue;
        }
      } else if (responseCode === 401) {
        throw new Error('Authentication failed. Please check your GitHub token.');
      } else if (responseCode === 403) {
        throw new Error('Access forbidden. Please check repository permissions.');
      } else if (responseCode === 404) {
        throw new Error('Repository not found. Please check the repository name and your access.');
      } else {
        throw new Error(`GitHub API error (${responseCode}): ${responseText}`);
      }
      
    } catch (error) {
      console.error(`❌ Attempt ${attempt} failed:`, error.message);
      
      if (attempt >= CONFIG.MAX_RETRIES) {
        throw new Error(`Failed after ${CONFIG.MAX_RETRIES} attempts: ${error.message}`);
      }
      
      // Wait before retrying
      Utilities.sleep(CONFIG.RETRY_DELAY * attempt);
    }
  }
  
  throw new Error('Maximum retry attempts exceeded');
}

/**
 * Gets current file information from GitHub
 */
function getCurrentFileInfo(url, token) {
  try {
    const options = {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Accept': 'application/vnd.github.v3+json',
        'User-Agent': 'Google-Apps-Script-Attendance-Export'
      },
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      const content = JSON.parse(response.getContentText());
      return { sha: content.sha };
    } else if (responseCode === 404) {
      return { sha: null }; // File doesn't exist
    } else {
      console.warn(`⚠️ Unexpected response when getting file info: ${responseCode}`);
      return { sha: null };
    }
    
  } catch (error) {
    console.warn(`⚠️ Could not get current file info: ${error.message}`);
    return { sha: null };
  }
}

/**
 * Safely retrieves GitHub token from Script Properties
 */
function getGitHubToken() {
  try {
    const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
    
    if (!token) {
      console.error('❌ GITHUB_TOKEN not found in Script Properties');
      return null;
    }
    
    // Don't log the actual token for security
    console.log('🔑 GitHub token retrieved successfully');
    return token;
    
  } catch (error) {
    console.error('❌ Failed to retrieve GitHub token:', error.message);
    return null;
  }
}

/**
 * Sends email notification about export status
 */
function sendNotificationEmail(success, message) {
  try {
    const email = Session.getActiveUser().getEmail();
    if (!email) return;
    
    const subject = success ? 
      '✅ GitHub Export Successful' : 
      '❌ GitHub Export Failed';
    
    const body = `
GitHub Export Status Report

Status: ${success ? 'SUCCESS' : 'FAILED'}
Time: ${new Date().toLocaleString()}
Repository: ${CONFIG.GITHUB_USERNAME}/${CONFIG.REPO}
File: ${CONFIG.FILE_PATH}

${success ? 'Details:' : 'Error:'} ${message}

${success ? 
  `View file: https://github.com/${CONFIG.GITHUB_USERNAME}/${CONFIG.REPO}/blob/${CONFIG.BRANCH}/${CONFIG.FILE_PATH}` :
  'Please check the execution log for more details.'
}

---
This is an automated message from your Google Sheets GitHub Export script.
    `;
    
    MailApp.sendEmail(email, subject, body);
    console.log(`📧 Notification email sent to ${email}`);
    
  } catch (error) {
    console.warn(`⚠️ Could not send notification email: ${error.message}`);
  }
}

/**
 * Test function to validate setup without making changes
 */
function testGitHubConnection() {
  console.log('🧪 Testing GitHub connection...');
  
  try {
    validateConfiguration();
    
    const token = getGitHubToken();
    const url = `https://api.github.com/repos/${CONFIG.GITHUB_USERNAME}/${CONFIG.REPO}`;
    
    const options = {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Accept': 'application/vnd.github.v3+json',
        'User-Agent': 'Google-Apps-Script-Attendance-Export'
      },
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      const repo = JSON.parse(response.getContentText());
      console.log('✅ GitHub connection successful!');
      console.log(`📁 Repository: ${repo.full_name}`);
      console.log(`🔒 Private: ${repo.private}`);
      console.log(`⭐ Stars: ${repo.stargazers_count}`);
      return true;
    } else {
      console.error(`❌ GitHub connection failed: ${responseCode}`);
      console.error(response.getContentText());
      return false;
    }
    
  } catch (error) {
    console.error('❌ Test failed:', error.message);
    return false;
  }
}

/**
 * Setup function to help configure the script
 */
function setupScript() {
  console.log('⚙️ Setting up GitHub export script...');
  
  // Check if token is set
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  
  if (!token) {
    console.log('❌ GitHub token not found!');
    console.log('📝 Please follow these steps:');
    console.log('1. Go to GitHub Settings > Developer settings > Personal access tokens');
    console.log('2. Generate a new token with "repo" permissions');
    console.log('3. In Apps Script, go to Project Settings > Script Properties');
    console.log('4. Add property: Key="GITHUB_TOKEN", Value="your_token_here"');
    return false;
  }
  
  console.log('✅ GitHub token found');
  
  // Test connection
  if (testGitHubConnection()) {
    console.log('🎉 Setup complete! You can now run exportSheetToGitHub()');
    return true;
  } else {
    console.log('❌ Setup incomplete. Please check your token and repository settings.');
    return false;
  }
}

