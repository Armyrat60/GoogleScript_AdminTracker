/***********************
 * CONFIGURATION EXAMPLE
 * 
 * This file shows all the settings you need to configure
 * in AdminHoursTracker.js before first use.
 * 
 * This is NOT a working file - it's just for reference!
 * Make these changes directly in AdminHoursTracker.js
 ***********************/

// ========== STEP 1: UPDATE GITHUB URLS ==========

const GITHUB_VERSION_URL = 'https://raw.githubusercontent.com/YOUR_USERNAME/admin-hours-tracker/main/version.json';
const GITHUB_DOWNLOAD_URL = 'https://github.com/YOUR_USERNAME/admin-hours-tracker';

// Replace YOUR_USERNAME with your GitHub username
// Example: 'https://raw.githubusercontent.com/john-doe/admin-hours-tracker/main/version.json'


// ========== STEP 2: SET YOUR BATTLEMETRICS SERVER ID ==========

BM_SERVER_ID: 'YOUR_SERVER_ID_HERE',

// How to find your Server ID:
// 1. Go to your server on BattleMetrics
// 2. Look at the URL: https://www.battlemetrics.com/servers/squad/12345678
// 3. The number at the end (12345678) is your Server ID
// 
// Example: BM_SERVER_ID: '12345678',


// ========== STEP 3: ADD API KEY TO SCRIPT PROPERTIES ==========

// DON'T put your API key in the code!
// Instead, add it to Script Properties:
//
// 1. In Apps Script editor, click Project Settings (gear icon)
// 2. Scroll to "Script Properties"  
// 3. Click "Add script property"
// 4. Property: BATTLEMETRICS_API_KEY
// 5. Value: YOUR_API_TOKEN_HERE
// 6. Click Save
//
// Get your API token from: https://www.battlemetrics.com/developers


// ========== STEP 4: CONFIGURE SHEET NAMES ==========

SOURCE_SHEET_NAME: 'Server Admin Team',     // Name of sheet with your admin data
RANKINGS_SHEET_NAME: 'Admin Rankings',      // Name for tracker display (auto-created)


// ========== STEP 5: VERIFY COLUMN NUMBERS ==========

// Make sure these match where your data is in the source sheet:
SOURCE_COL_NAME: 2,       // Column B - Admin names
SOURCE_COL_TEAM: 3,       // Column C - Team/role
SOURCE_COL_STEAM_ID: 6,   // Column F - Steam IDs (64-bit)
SOURCE_FIRST_DATA_ROW: 2, // First row of data (row 1 should be headers)

// Your source sheet should look like:
// Row 1: [Headers]    Name         Team        ...    Steam ID
// Row 2: [Data]       John Doe     Admin       ...    76561198XXXXXXXXX
// Row 3: [Data]       Jane Smith   Moderator   ...    76561198YYYYYYYYY


// ========== STEP 6: CUSTOMIZE TEAM COLORS (OPTIONAL) ==========

TEAM_COLORS: {
  // Format: 'Team Name': '#HEXCOLOR'
  // Customize team names and colors to match your server
  
  'Owner': '#ad1457',          // Dark pink
  'Executive': '#c62828',      // Dark red
  'Senior Admin': '#1565c0',   // Dark blue
  'Moderator': '#0277bd',      // Dark cyan
  'Admin': '#e65100',          // Dark orange
  'Trial Admin': '#00838f',    // Dark teal
  
  // Add your own teams:
  // 'Your Team Name': '#YOUR_COLOR',
},


// ========== STEP 7: ADJUST HOUR RANGES (OPTIONAL) ==========

COLOR_RANGES: [
  // Hours-based background colors
  // Customize ranges to match your server's activity levels
  
  { min: 0, max: 0, color: '#ffcdd2' },         // Inactive - Light red
  { min: 0.01, max: 10, color: '#fff9c4' },     // Low - Very light yellow
  { min: 10, max: 25, color: '#fff59d' },       // Below average - Light yellow
  { min: 25, max: 50, color: '#ffee58' },       // Average - Medium yellow
  { min: 50, max: 75, color: '#c5e1a5' },       // Good - Light green
  { min: 75, max: 100, color: '#aed581' },      // Very good - Medium light green
  { min: 100, max: 150, color: '#9ccc65' },     // Great - Medium green
  { min: 150, max: 200, color: '#7cb342' },     // Excellent - Good green
  { min: 200, max: 300, color: '#689f38' },     // Outstanding - Strong green
  { min: 300, max: Infinity, color: '#558b2f' } // Exceptional - Dark green
],


// ========== COMPLETE CONFIGURATION CHECKLIST ==========

/*
 * Before running the script, verify you've completed:
 * 
 * ✅ Replaced YOUR_USERNAME in GitHub URLs
 * ✅ Set BM_SERVER_ID to your server ID
 * ✅ Added BATTLEMETRICS_API_KEY to Script Properties
 * ✅ Created "Server Admin Team" sheet with columns B, C, F
 * ✅ Customized team names in TEAM_COLORS
 * ✅ (Optional) Adjusted hour color ranges
 * 
 * Ready? Run: ⚙️ Admin Tracker → 🚀 Complete Initial Setup
 */
