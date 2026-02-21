/***********************
 * CONFIGURATION FILE
 * 
 * Admin Hours Tracker for BattleMetrics
 * Author: ArmyRat60
 * 
 * EDIT THIS FILE TO CONFIGURE YOUR TRACKER
 * 
 * You can either:
 * 1. Use the Setup Wizard (Menu → Setup Wizard) - RECOMMENDED
 * 2. Manually edit the values below
 * 
 * The Setup Wizard will save settings to Script Properties,
 * which take priority over these values.
 ***********************/

var CONFIG = {
  // ========== BATTLEMETRICS SETTINGS ==========
  
  // Your BattleMetrics Server ID
  // Find it in your server URL: https://www.battlemetrics.com/servers/squad/12345678
  // The number at the end (12345678) is your Server ID
  BM_SERVER_ID: 'YOUR_SERVER_ID_HERE',
  
  // BattleMetrics API Token Key (stored in Script Properties)
  // Don't put the actual token here - use Script Properties!
  // Property name: BATTLEMETRICS_API_KEY
  BM_TOKEN_KEY: 'BATTLEMETRICS_API_KEY',
  
  // ========== SHEET CONFIGURATION ==========
  
  // Name of the sheet with your admin data
  SOURCE_SHEET_NAME: 'Server Admin Team',
  
  // Name for the auto-generated rankings sheet
  RANKINGS_SHEET_NAME: 'Admin Rankings',
  
  // ========== COLUMN CONFIGURATION ==========
  
  // Which columns contain your data in the source sheet
  // Column numbers (A=1, B=2, C=3, etc.)
  //
  // NOTE: This configuration reads from your auto-populated columns:
  // - Steam ID (Column F): Manual entry
  // - Team (Column C): Manual entry  
  // - Name (Column U): Auto-populated "STEAM NAME (AUTO)" from BattleMetrics
  //
  // This ensures accurate names pulled directly from BattleMetrics API
  // rather than relying on manual entry in "Team member" column.
  //
  SOURCE_COL_NAME: 21,         // Column U - STEAM NAME (AUTO) - auto-populated from BattleMetrics
  SOURCE_COL_TEAM: 3,          // Column C - Team/role (manual entry)
  SOURCE_COL_STEAM_ID: 6,      // Column F - Steam IDs (64-bit format, manual entry)
  SOURCE_COL_BMID: 22,         // Column V - BMID (AUTO) - auto-populated from BattleMetrics
  SOURCE_COL_LAST_SEEN: 19,    // Column S - Last Seen (Auto) - optional
  SOURCE_COL_FIRST_SEEN: 20,   // Column T - First Seen (Auto) - optional
  SOURCE_COL_EOSID: 23,        // Column W - EOSID (Auto) - optional
  SOURCE_FIRST_DATA_ROW: 2,    // First row of data (row 1 = headers)
  
  // ========== INTERNAL SETTINGS (Advanced) ==========
  
  // Rankings Sheet Data Storage (Visible, Columns AA-AJ)
  // Clean organized structure for easy debugging
  DATA_COL_STEAM_ID: 27,      // Column AA - Steam ID (lookup key)
  DATA_COL_NAME: 28,          // Column AB - Player Name
  DATA_COL_TEAM: 29,          // Column AC - Team/Role
  DATA_COL_BM_ID: 30,         // Column AD - BattleMetrics ID
  DATA_COL_LAST_SEEN: 31,     // Column AE - Last Seen (from BM API)
  DATA_COL_FIRST_SEEN: 32,    // Column AF - First Seen (account creation)
  DATA_COL_EOSID: 33,         // Column AG - Epic Online Services ID
  DATA_COL_HOURS_7D: 34,      // Column AH - 7-Day Hours
  DATA_COL_HOURS_30D: 35,     // Column AI - 30-Day Hours
  DATA_COL_HOURS_90D: 36,     // Column AJ - 90-Day Hours
  DATA_FIRST_ROW: 2,
  
  DISPLAY_START_COL: 1,
  
  // ========== API RATE LIMITS ==========
  
  // BattleMetrics API quotas (don't change)
  MAX_DAILY_API_CALLS: 18000,
  MAX_HOURLY_API_CALLS: 750,
  
  // ========== VISUAL CUSTOMIZATION ==========
  
  // Team colors for text display
  // Add/modify team names and colors to match your server
  // Format: 'Team Name': '#HEXCOLOR'
  // Note: Remove number prefixes (e.g., use "Executive" not "2. Executive")
  TEAM_COLORS: {
    'Owner': '#c62828',          // Dark red
    'Executive': '#ad1457',      // Dark pink
    'Senior Admin': '#1565c0',   // Dark blue
    'Operations Lead': '#00695c',// Dark teal
    'Operations': '#00838f',     // Cyan
    'Tech Team': '#0277bd',      // Blue
    'Admin': '#0d47a1',          // Navy blue (changed for better contrast on red/orange backgrounds)
    'Moderator': '#6a1b9a'       // Purple
    // Add more teams here as needed
  },
  
  // Hours-based color gradient (background colors)
  // Customize ranges to match your server's activity levels
  COLOR_RANGES: [
    { min: 0, max: 0, color: '#e0e0e0' },         // Gray - Inactive (changed from red to avoid conflicts)
    { min: 0.01, max: 10, color: '#fff9c4' },     // Very light yellow
    { min: 10, max: 25, color: '#fff59d' },       // Light yellow
    { min: 25, max: 50, color: '#ffee58' },       // Medium yellow
    { min: 50, max: 75, color: '#c5e1a5' },       // Light green
    { min: 75, max: 100, color: '#aed581' },      // Medium light green
    { min: 100, max: 150, color: '#9ccc65' },     // Medium green
    { min: 150, max: 200, color: '#7cb342' },     // Good green
    { min: 200, max: 300, color: '#689f38' },     // Strong green
    { min: 300, max: Infinity, color: '#558b2f' } // Dark green - Excellent
  ]
};

/***********************
 * CONFIGURATION NOTES
 * 
 * Setup Wizard (Recommended):
 * 1. Menu → ⚙️ Admin Tracker → 🛠️ Setup Wizard
 * 2. Answer a few prompts
 * 3. Everything configured automatically!
 * 
 * Manual Setup:
 * 1. Edit BM_SERVER_ID above
 * 2. Add BATTLEMETRICS_API_KEY to Script Properties:
 *    - Click Project Settings (gear icon)
 *    - Scroll to Script Properties
 *    - Add property: BATTLEMETRICS_API_KEY
 *    - Value: Your BattleMetrics API token
 * 3. Verify column numbers match your sheet
 * 4. Customize team names and colors
 * 5. Run: Menu → Complete Initial Setup
 * 
 * Column Auto-Detection:
 * The Setup Wizard can automatically detect columns
 * by searching for headers like:
 * - "Name", "Admin", "Player"
 * - "Team", "Role", "Rank"
 * - "Steam", "Steam ID", "SteamID"
 ***********************/
