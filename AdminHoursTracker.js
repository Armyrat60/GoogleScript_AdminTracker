/***********************
 * Admin Hours Tracker - Version Controlled
 * 
 * VERSION: 1.0.0
 * 
 * Features:
 * - Automatic update checking from GitHub
 * - Version display in menu
 * - Rotating schedule to avoid timeouts
 * - Dynamic search with conditional formatting
 * - Last update timestamps
 * 
 * GitHub: https://github.com/YOUR_USERNAME/dmh-admin-tracker
 ***********************/

/***********************
 * Admin Hours Tracker for BattleMetrics
 * 
 * VERSION: 1.0.0
 * Author: ArmyRat60
 * GitHub: https://github.com/Armyrat60/GoogleScript_AdminTracker
 * 
 * A Google Apps Script for tracking admin activity hours on BattleMetrics servers.
 * Works with Squad, Rust, ARK, and any game using BattleMetrics.
 * 
 * Features:
 * - Automatic hourly updates with rotating schedule (prevents timeouts)
 * - Real-time sync from source data
 * - Color-coded performance rankings
 * - Last update timestamps
 * - GitHub version checking
 * 
 * Setup Guide: See README.md
 * GitHub: https://github.com/Armyrat60/GoogleScript_AdminTracker
 * 
 * IMPORTANT: Before using, you must:
 * 1. Set BM_SERVER_ID to your BattleMetrics server ID
 * 2. Add your BattleMetrics API token to Script Properties
 * 3. Customize team names/colors in CONFIG section
 ***********************/

const VERSION = '1.0.0';
const GITHUB_VERSION_URL = 'https://raw.githubusercontent.com/Armyrat60/GoogleScript_AdminTracker/main/version.json';
const GITHUB_DOWNLOAD_URL = 'https://github.com/Armyrat60/GoogleScript_AdminTracker';

const CONFIG = {
  // ========== REQUIRED: Update These Settings ==========
  
  // BattleMetrics Server ID (find in your server URL)
  // Example: https://www.battlemetrics.com/servers/squad/12345678
  // Your Server ID would be: 12345678
  BM_SERVER_ID: 'YOUR_SERVER_ID_HERE',
  
  // BattleMetrics API token key name (stored in Script Properties)
  // Get token from: https://www.battlemetrics.com/developers
  BM_TOKEN_KEY: 'BATTLEMETRICS_API_KEY',
  
  // Source sheet name (where your admin data lives)
  SOURCE_SHEET_NAME: 'Server Admin Team',
  
  // Rankings sheet name (auto-created by script)
  RANKINGS_SHEET_NAME: 'Admin Rankings',
  
  // ========== Source Sheet Column Configuration ==========
  
  // Which columns contain your admin data
  SOURCE_COL_NAME: 2,      // Column B - Admin Name
  SOURCE_COL_TEAM: 3,      // Column C - Team/Role
  SOURCE_COL_STEAM_ID: 6,  // Column F - Steam ID (64-bit)
  SOURCE_FIRST_DATA_ROW: 2, // First row of data (row 1 = headers)
  
  // ========== Internal Configuration (Advanced) ==========
  
  // Hidden data storage columns (don't change unless needed)
  DATA_COL_NAME: 25,       // Column Y
  DATA_COL_STEAM_ID: 26,   // Column Z
  DATA_COL_TEAM: 27,       // Column AA
  DATA_COL_BM_ID: 28,      // Column AB
  DATA_COL_HOURS_7D: 29,   // Column AC
  DATA_COL_HOURS_30D: 30,  // Column AD
  DATA_COL_HOURS_90D: 31,  // Column AE
  DATA_FIRST_ROW: 2,
  
  // Search box location
  SEARCH_LABEL_COL: 19,    // Column S
  SEARCH_CELL_COL: 20,     // Column T
  CLEAR_BUTTON_COL: 21,    // Column U
  
  DISPLAY_START_COL: 1,
  
  // BattleMetrics API rate limits (don't change)
  MAX_DAILY_API_CALLS: 18000,
  MAX_HOURLY_API_CALLS: 750,
  
  // ========== Visual Customization ==========
  
  // Team colors for text (customize to match your teams)
  // Add/modify team names and colors as needed
  TEAM_COLORS: {
    'Owner': '#ad1457',          // Dark pink
    'Executive': '#c62828',      // Dark red  
    'Senior Admin': '#1565c0',   // Dark blue
    'Moderator': '#0277bd',      // Dark cyan
    'Admin': '#e65100',          // Dark orange
    'Trial Admin': '#00838f'     // Dark teal
    // Add more teams here as needed
  },
  
  // Color gradient
  COLOR_RANGES: [
    { min: 0, max: 0, color: '#ffcdd2' },
    { min: 0.01, max: 10, color: '#fff9c4' },
    { min: 10, max: 25, color: '#fff59d' },
    { min: 25, max: 50, color: '#ffee58' },
    { min: 50, max: 75, color: '#c5e1a5' },
    { min: 75, max: 100, color: '#aed581' },
    { min: 100, max: 150, color: '#9ccc65' },
    { min: 150, max: 200, color: '#7cb342' },
    { min: 200, max: 300, color: '#689f38' },
    { min: 300, max: Infinity, color: '#558b2f' }
  ],
  
  SEARCH_HIGHLIGHT_COLOR: '#00e5ff'
};

// ========== VERSION CHECKING ==========

function checkForUpdates() {
  try {
    var response = UrlFetchApp.fetch(GITHUB_VERSION_URL, { muteHttpExceptions: true });
    var code = response.getResponseCode();
    
    if (code !== 200) {
      Logger.log('Could not check for updates: ' + code);
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Unable to check for updates. Check internet connection.',
        'Update Check Failed',
        5
      );
      return;
    }
    
    var data = JSON.parse(response.getContentText());
    var latestVersion = data.version;
    var releaseNotes = data.notes || 'See GitHub for details';
    
    Logger.log('Current: ' + VERSION + ', Latest: ' + latestVersion);
    
    if (compareVersions(latestVersion, VERSION) > 0) {
      var ui = SpreadsheetApp.getUi();
      var message = 
        'A new version is available!\n\n' +
        'Current: v' + VERSION + '\n' +
        'Latest: v' + latestVersion + '\n\n' +
        'What\'s new:\n' + releaseNotes + '\n\n' +
        'Click OK to open download page.';
      
      var result = ui.alert('Update Available', message, ui.ButtonSet.OK_CANCEL);
      
      if (result === ui.Button.OK) {
        var htmlOutput = HtmlService.createHtmlOutput(
          '<script>window.open("' + GITHUB_DOWNLOAD_URL + '", "_blank");google.script.host.close();</script>'
        );
        SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Opening GitHub...');
      }
    } else {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'You have the latest version! (v' + VERSION + ')',
        'Up to Date',
        3
      );
    }
  } catch (error) {
    Logger.log('Error checking for updates: ' + error);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Error checking for updates: ' + error.message,
      'Update Check Error',
      5
    );
  }
}

function compareVersions(v1, v2) {
  var parts1 = v1.split('.').map(Number);
  var parts2 = v2.split('.').map(Number);
  
  for (var i = 0; i < 3; i++) {
    if (parts1[i] > parts2[i]) return 1;
    if (parts1[i] < parts2[i]) return -1;
  }
  return 0;
}

function showAbout() {
  var ui = SpreadsheetApp.getUi();
  var message = 
    'DMH Admin Hours Tracker\n' +
    'Version: ' + VERSION + '\n\n' +
    'Features:\n' +
    '• Rotating hourly updates (no timeouts)\n' +
    '• Dynamic search with highlighting\n' +
    '• Real-time sync from source data\n' +
    '• Quota tracking & management\n' +
    '• Last update timestamps\n\n' +
    'GitHub: ' + GITHUB_DOWNLOAD_URL;
  
  ui.alert('About Admin Tracker', message, ui.ButtonSet.OK);
}

// ========== TIMESTAMP HELPERS ==========

function getLastUpdateTimestamp_(metric) {
  try {
    var props = PropertiesService.getScriptProperties();
    var timestamp = props.getProperty('LAST_UPDATE_' + metric);
    if (!timestamp) return 'Never';
    
    var date = new Date(parseInt(timestamp));
    var now = new Date();
    var diffMs = now - date;
    var diffMins = Math.floor(diffMs / 60000);
    
    if (diffMins < 60) {
      return diffMins + ' min ago';
    } else if (diffMins < 1440) {
      return Math.floor(diffMins / 60) + ' hrs ago';
    } else {
      return Math.floor(diffMins / 1440) + ' days ago';
    }
  } catch (error) {
    return 'Unknown';
  }
}

function setLastUpdateTimestamp_(metric) {
  try {
    var props = PropertiesService.getScriptProperties();
    props.setProperty('LAST_UPDATE_' + metric, String(new Date().getTime()));
  } catch (error) {
    Logger.log('Error setting timestamp for ' + metric + ': ' + error);
  }
}

function getNextUpdateInfo_() {
  var currentHour = new Date().getHours();
  var cycle = currentHour % 3;
  var nextUpdate = '';
  
  if (cycle === 0) {
    nextUpdate = '7-day hours';
  } else if (cycle === 1) {
    nextUpdate = '30-day hours';
  } else {
    nextUpdate = '90-day hours';
  }
  
  return 'Next: ' + nextUpdate + ' @ :00';
}

// ========== ENHANCED CUSTOM MENU ==========

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  // Get last update info
  var last7d = getLastUpdateTimestamp_('7D');
  var last30d = getLastUpdateTimestamp_('30D');
  var last90d = getLastUpdateTimestamp_('90D');
  var nextUpdate = getNextUpdateInfo_();
  
  ui.createMenu('⚙️ Admin Tracker (v' + VERSION + ')')
    .addItem('🔄 Check for Updates', 'checkForUpdates')
    .addItem('ℹ️ About', 'showAbout')
    .addSeparator()
    .addSubMenu(ui.createMenu('📊 Update Status')
      .addItem('7-Day Hours: ' + last7d, 'showUpdateInfo')
      .addItem('30-Day Hours: ' + last30d, 'showUpdateInfo')
      .addItem('90-Day Hours: ' + last90d, 'showUpdateInfo')
      .addSeparator()
      .addItem(nextUpdate, 'showUpdateInfo'))
    .addSeparator()
    .addSubMenu(ui.createMenu('ℹ️ Rotating Updates')
      .addItem('How it works...', 'explainRotatingUpdates')
      .addSeparator()
      .addItem('Hour 0,3,6,9,12... → 7-day', 'showUpdateInfo')
      .addItem('Hour 1,4,7,10,13... → 30-day', 'showUpdateInfo')
      .addItem('Hour 2,5,8,11,14... → 90-day', 'showUpdateInfo'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🔧 Manual Updates')
      .addItem('Update 7-Day Hours', 'updateHours7d')
      .addItem('Update 30-Day Hours', 'updateHours30d')
      .addItem('Update 90-Day Hours', 'updateHours90d')
      .addSeparator()
      .addItem('Sync Admin Data', 'STEP1_SyncAdminData')
      .addItem('Populate BM IDs', 'STEP2_PopulateBMIds')
      .addItem('Refresh Display', 'STEP4_CreateRankingsDisplay'))
    .addSeparator()
    .addItem('🚀 Complete Initial Setup', 'SETUP_CompleteInitialSetup')
    .addItem('🔄 Manual Full Update (All)', 'manualFullUpdate')
    .addSeparator()
    // .addItem('🔍 Clear Search', 'clearSearch')  // Hidden - feature in development
    .addItem('📊 Check Quota Status', 'checkQuotaStatus')
    .addItem('🧪 Test API Connection', 'testDirectAPICall')
    .addToUi();
}

function showUpdateInfo() {
  // Dummy function for menu items that are just informational
}

function explainRotatingUpdates() {
  var ui = SpreadsheetApp.getUi();
  
  var message = 
    'ROTATING UPDATE SCHEDULE\n\n' +
    'To avoid timeout errors, the tracker updates ONE metric each hour:\n\n' +
    '🕐 Hours 0,3,6,9,12,15,18,21  →  7-Day Hours\n' +
    '🕑 Hours 1,4,7,10,13,16,19,22  →  30-Day Hours\n' +
    '🕒 Hours 2,5,8,11,14,17,20,23  →  90-Day Hours\n\n' +
    'BENEFITS:\n' +
    '✓ No timeout errors (each run ~10 min)\n' +
    '✓ All metrics update every 3 hours\n' +
    '✓ Fresh data throughout the day\n\n' +
    'MANUAL UPDATES:\n' +
    'Use "Manual Full Update (All)" to update everything at once.';
  
  ui.alert('Rotating Updates Explained', message, ui.ButtonSet.OK);
}

// ========== DYNAMIC SEARCH FUNCTIONALITY ==========

function setupSearchBox_(sheet) {
  var labelCell = sheet.getRange(1, CONFIG.SEARCH_LABEL_COL);
  labelCell.setValue('🔍 SEARCH:');
  labelCell.setBackground('#ff9800');
  labelCell.setFontColor('#ffffff');
  labelCell.setFontWeight('bold');
  labelCell.setFontSize(11);
  labelCell.setHorizontalAlignment('center');
  labelCell.setVerticalAlignment('middle');
  labelCell.setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
  var searchCell = sheet.getRange(1, CONFIG.SEARCH_CELL_COL);
  searchCell.setValue('');
  searchCell.setBackground('#fff3cd');
  searchCell.setFontColor('#000000');
  searchCell.setFontStyle('normal');
  searchCell.setFontSize(11);
  searchCell.setHorizontalAlignment('left');
  searchCell.setVerticalAlignment('middle');
  searchCell.setBorder(true, true, true, true, false, false, '#ff9800', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  searchCell.setNote('Type a name or team to highlight matches instantly.\nNo page refresh needed!');
  
  var clearButton = sheet.getRange(1, CONFIG.CLEAR_BUTTON_COL);
  clearButton.setValue('✖ CLEAR');
  clearButton.setBackground('#f44336');
  clearButton.setFontColor('#ffffff');
  clearButton.setFontWeight('bold');
  clearButton.setFontSize(10);
  clearButton.setHorizontalAlignment('center');
  clearButton.setVerticalAlignment('middle');
  clearButton.setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  clearButton.setNote('Click to clear search box');
  
  sheet.setColumnWidth(CONFIG.SEARCH_LABEL_COL, 120);
  sheet.setColumnWidth(CONFIG.SEARCH_CELL_COL, 200);
  sheet.setColumnWidth(CONFIG.CLEAR_BUTTON_COL, 100);
  sheet.setRowHeight(1, 35);
}

function setupDynamicSearchFormatting_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return;
  
  var searchCellRef = columnToLetter(CONFIG.SEARCH_CELL_COL) + '1';
  var rules = [];
  
  rules.push(createSearchRule_(sheet, 1, 3, lastRow, searchCellRef, 1, 3));
  rules.push(createSearchRule_(sheet, 6, 3, lastRow, searchCellRef, 6, 8));
  rules.push(createSearchRule_(sheet, 11, 3, lastRow, searchCellRef, 11, 13));
  
  sheet.setConditionalFormatRules(rules);
}

function createSearchRule_(sheet, startCol, startRow, endRow, searchCellRef, nameCol, teamCol) {
  var nameColLetter = columnToLetter(nameCol + 1);
  var teamColLetter = columnToLetter(teamCol + 1);
  
  var formula = '=AND(' +
    'LEN(' + searchCellRef + ')>0,' +
    'OR(' +
      'ISNUMBER(SEARCH(' + searchCellRef + ',' + nameColLetter + '3)),' +
      'ISNUMBER(SEARCH(' + searchCellRef + ',' + teamColLetter + '3))' +
    ')' +
  ')';
  
  var range = sheet.getRange(startRow, startCol, endRow - startRow + 1, 4);
  
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .setRanges([range])
    .whenFormulaSatisfied(formula)
    .setBackground(CONFIG.SEARCH_HIGHLIGHT_COLOR)
    .setFontColor('#000000')
    .setBold(true)
    .build();
  
  return rule;
}

function columnToLetter(column) {
  var temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function clearSearch() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.RANKINGS_SHEET_NAME);
    
    if (!sheet) return;
    
    var searchCell = sheet.getRange(1, CONFIG.SEARCH_CELL_COL);
    searchCell.setValue('');
    
    SpreadsheetApp.getActiveSpreadsheet().toast('Search cleared', 'Search', 2);
  } catch (error) {
    Logger.log('Error in clearSearch: ' + error);
  }
}

// ========== QUOTA TRACKING ==========

function getQuotaKey_(type) {
  try {
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var hour = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd-HH');
    if (type === 'daily') return 'quota_daily_' + today;
    if (type === 'hourly') return 'quota_hourly_' + hour;
    return null;
  } catch (error) {
    Logger.log('Error in getQuotaKey_: ' + error);
    return null;
  }
}

function getQuotaUsage_(type) {
  try {
    var key = getQuotaKey_(type);
    if (!key) return 0;
    var cache = CacheService.getScriptCache();
    var value = cache.get(key);
    return value ? parseInt(value) : 0;
  } catch (error) {
    Logger.log('Error in getQuotaUsage_: ' + error);
    return 0;
  }
}

function incrementQuotaUsage_(count) {
  try {
    var cache = CacheService.getScriptCache();
    var dailyKey = getQuotaKey_('daily');
    var dailyCount = getQuotaUsage_('daily') + count;
    cache.put(dailyKey, String(dailyCount), 86400);
    var hourlyKey = getQuotaKey_('hourly');
    var hourlyCount = getQuotaUsage_('hourly') + count;
    cache.put(hourlyKey, String(hourlyCount), 3600);
    Logger.log('Quota - Hourly: ' + hourlyCount + '/' + CONFIG.MAX_HOURLY_API_CALLS + ', Daily: ' + dailyCount + '/' + CONFIG.MAX_DAILY_API_CALLS);
  } catch (error) {
    Logger.log('Error in incrementQuotaUsage_: ' + error);
  }
}

function canMakeApiCall_() {
  try {
    var dailyUsage = getQuotaUsage_('daily');
    var hourlyUsage = getQuotaUsage_('hourly');
    if (dailyUsage >= CONFIG.MAX_DAILY_API_CALLS) {
      Logger.log('⚠ Daily quota limit reached: ' + dailyUsage);
      return false;
    }
    if (hourlyUsage >= CONFIG.MAX_HOURLY_API_CALLS) {
      Logger.log('⚠ Hourly quota limit reached: ' + hourlyUsage);
      return false;
    }
    return true;
  } catch (error) {
    Logger.log('Error in canMakeApiCall_: ' + error);
    return false;
  }
}

// ========== HELPER FUNCTIONS ==========

function getScriptProperty_(key) {
  try {
    var value = PropertiesService.getScriptProperties().getProperty(key);
    if (!value) throw new Error('Missing Script Property: "' + key + '"');
    return value;
  } catch (error) {
    Logger.log('Error getting script property "' + key + '": ' + error);
    throw error;
  }
}

function callBMApi_(endpoint, token) {
  try {
    if (!canMakeApiCall_()) {
      Logger.log('⚠ Quota limit reached, skipping API call');
      return null;
    }
    
    var url = 'https://api.battlemetrics.com' + endpoint;
    if (endpoint.indexOf('?') > -1) {
      url += '&access_token=' + encodeURIComponent(token);
    } else {
      url += '?access_token=' + encodeURIComponent(token);
    }
    
    var response = UrlFetchApp.fetch(url, { 
      muteHttpExceptions: true,
      validateHttpsCertificates: true
    });
    
    incrementQuotaUsage_(1);
    
    var code = response.getResponseCode();
    var body = response.getContentText();
    
    if (code !== 200) {
      Logger.log('BM API Error ' + code + ': ' + body);
      return null;
    }
    
    return JSON.parse(body);
  } catch (error) {
    Logger.log('Error in callBMApi_: ' + error);
    return null;
  }
}

function getBMIdFromSteamId_(steamId, token) {
  try {
    var endpoint = '/players?filter[search]=' + encodeURIComponent(steamId) + 
                   '&filter[servers]=' + CONFIG.BM_SERVER_ID;
    
    var data = callBMApi_(endpoint, token);
    
    if (data && data.data && data.data.length > 0) {
      if (data.data.length > 1) {
        Logger.log('  Multiple matches found for ' + steamId + ', using first result');
      }
      return data.data[0].id;
    }
    return null;
  } catch (error) {
    Logger.log('Error in getBMIdFromSteamId_ for ' + steamId + ': ' + error);
    return null;
  }
}

function getPlayedTimeSeconds_(playerId, serverId, start, stop, token) {
  try {
    var endpoint = '/players/' + playerId + '/time-played-history/' + serverId +
                   '?start=' + encodeURIComponent(start.toISOString()) +
                   '&stop=' + encodeURIComponent(stop.toISOString());
    var data = callBMApi_(endpoint, token);
    if (!data || !data.data) return null;
    var seconds = 0;
    for (var i = 0; i < data.data.length; i++) {
      var value = (data.data[i].attributes && data.data[i].attributes.value) || 0;
      seconds += Number(value);
    }
    return seconds;
  } catch (error) {
    Logger.log('Error in getPlayedTimeSeconds_ for player ' + playerId + ': ' + error);
    return null;
  }
}

// ========== STEP 1: SYNC DATA ==========

function STEP1_SyncAdminData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sourceSheet = ss.getSheetByName(CONFIG.SOURCE_SHEET_NAME);
    var rankingsSheet = ss.getSheetByName(CONFIG.RANKINGS_SHEET_NAME);
    
    if (!sourceSheet) {
      throw new Error('Source sheet "' + CONFIG.SOURCE_SHEET_NAME + '" not found');
    }
    
    if (!rankingsSheet) {
      rankingsSheet = ss.insertSheet(CONFIG.RANKINGS_SHEET_NAME);
      ss.setActiveSheet(rankingsSheet);
      ss.moveActiveSheet(1);
    }
    
    var lastRow = sourceSheet.getLastRow();
    if (lastRow < CONFIG.SOURCE_FIRST_DATA_ROW) {
      Logger.log('No data in source sheet');
      return;
    }
    
    var numRows = lastRow - CONFIG.SOURCE_FIRST_DATA_ROW + 1;
    var names = sourceSheet.getRange(CONFIG.SOURCE_FIRST_DATA_ROW, CONFIG.SOURCE_COL_NAME, numRows, 1).getValues();
    var teams = sourceSheet.getRange(CONFIG.SOURCE_FIRST_DATA_ROW, CONFIG.SOURCE_COL_TEAM, numRows, 1).getValues();
    var steamIds = sourceSheet.getRange(CONFIG.SOURCE_FIRST_DATA_ROW, CONFIG.SOURCE_COL_STEAM_ID, numRows, 1).getValues();
    
    rankingsSheet.getRange(CONFIG.DATA_FIRST_ROW, CONFIG.DATA_COL_NAME, numRows, 1).setValues(names);
    rankingsSheet.getRange(CONFIG.DATA_FIRST_ROW, CONFIG.DATA_COL_STEAM_ID, numRows, 1).setValues(steamIds);
    rankingsSheet.getRange(CONFIG.DATA_FIRST_ROW, CONFIG.DATA_COL_TEAM, numRows, 1).setValues(teams);
    
    rankingsSheet.getRange(1, CONFIG.DATA_COL_NAME).setValue('Name');
    rankingsSheet.getRange(1, CONFIG.DATA_COL_STEAM_ID).setValue('Steam ID');
    rankingsSheet.getRange(1, CONFIG.DATA_COL_TEAM).setValue('Team');
    rankingsSheet.getRange(1, CONFIG.DATA_COL_BM_ID).setValue('BM ID');
    rankingsSheet.getRange(1, CONFIG.DATA_COL_HOURS_7D).setValue('Hours 7d');
    rankingsSheet.getRange(1, CONFIG.DATA_COL_HOURS_30D).setValue('Hours 30d');
    rankingsSheet.getRange(1, CONFIG.DATA_COL_HOURS_90D).setValue('Hours 90d');
    
    rankingsSheet.hideColumns(CONFIG.DATA_COL_NAME, 10);
    
    Logger.log('✅ Synced ' + numRows + ' admins from source sheet');
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Synced ' + numRows + ' admins',
      'Data Sync Complete',
      3
    );
  } catch (error) {
    Logger.log('FATAL ERROR in STEP1_SyncAdminData: ' + error);
    throw error;
  }
}

// ========== STEP 2: POPULATE BM IDs ==========

function STEP2_PopulateBMIds() {
  try {
    var token = getScriptProperty_(CONFIG.BM_TOKEN_KEY);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.RANKINGS_SHEET_NAME);
    
    if (!sheet) throw new Error('Rankings sheet not found');
    
    var lastRow = sheet.getLastRow();
    if (lastRow < CONFIG.DATA_FIRST_ROW) {
      Logger.log('No data rows found');
      return;
    }
    
    var numRows = lastRow - CONFIG.DATA_FIRST_ROW + 1;
    var steamIds = sheet.getRange(CONFIG.DATA_FIRST_ROW, CONFIG.DATA_COL_STEAM_ID, numRows, 1).getValues();
    var output = [];
    var stats = { processed: 0, skipped: 0, errors: 0 };
    
    Logger.log('Starting BM ID population...');
    
    for (var i = 0; i < steamIds.length; i++) {
      try {
        var steamId = String(steamIds[i][0] || '').trim();
        var rowNum = CONFIG.DATA_FIRST_ROW + i;
        
        if (!steamId) {
          output.push(['']);
          stats.skipped++;
          continue;
        }
        
        if (!canMakeApiCall_()) {
          Logger.log('⚠ Quota limit reached, stopping');
          break;
        }
        
        Logger.log('Row ' + rowNum + ' | Steam: ' + steamId);
        var bmId = getBMIdFromSteamId_(steamId, token);
        Utilities.sleep(250);
        
        if (bmId) {
          output.push([bmId]);
          stats.processed++;
          Logger.log('  ✓ BM ID: ' + bmId);
        } else {
          output.push(['Not Found']);
          stats.errors++;
          Logger.log('  ✗ Not found');
        }
      } catch (rowError) {
        Logger.log('Error processing row ' + rowNum + ': ' + rowError);
        output.push(['Error']);
        stats.errors++;
      }
    }
    
    if (output.length > 0) {
      sheet.getRange(CONFIG.DATA_FIRST_ROW, CONFIG.DATA_COL_BM_ID, output.length, 1).setValues(output);
    }
    
    Logger.log('Complete - Processed: ' + stats.processed + ', Skipped: ' + stats.skipped + ', Errors: ' + stats.errors);
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Processed: ' + stats.processed + ', Errors: ' + stats.errors,
      'BM ID Population Complete',
      5
    );
  } catch (error) {
    Logger.log('FATAL ERROR in STEP2_PopulateBMIds: ' + error);
    throw error;
  }
}

// ========== AUTO-POPULATE MISSING BM IDs ==========

function autoPopulateMissingBMIds_() {
  try {
    var token = getScriptProperty_(CONFIG.BM_TOKEN_KEY);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.RANKINGS_SHEET_NAME);
    
    if (!sheet) return;
    
    var lastRow = sheet.getLastRow();
    if (lastRow < CONFIG.DATA_FIRST_ROW) return;
    
    var numRows = lastRow - CONFIG.DATA_FIRST_ROW + 1;
    
    var steamIds = sheet.getRange(CONFIG.DATA_FIRST_ROW, CONFIG.DATA_COL_STEAM_ID, numRows, 1).getValues();
    var bmIds = sheet.getRange(CONFIG.DATA_FIRST_ROW, CONFIG.DATA_COL_BM_ID, numRows, 1).getValues();
    
    var newlyAdded = 0;
    
    for (var i = 0; i < steamIds.length; i++) {
      try {
        var steamId = String(steamIds[i][0] || '').trim();
        var bmId = String(bmIds[i][0] || '').trim();
        var rowNum = CONFIG.DATA_FIRST_ROW + i;
        
        if (bmId || !steamId) {
          continue;
        }
        
        if (!canMakeApiCall_()) {
          Logger.log('⚠ Quota limit reached, stopping BM ID auto-population');
          break;
        }
        
        Logger.log('Auto-populating BM ID | Row ' + rowNum + ' | Steam: ' + steamId);
        
        var newBMId = getBMIdFromSteamId_(steamId, token);
        Utilities.sleep(250);
        
        if (newBMId) {
          sheet.getRange(rowNum, CONFIG.DATA_COL_BM_ID).setValue(newBMId);
          newlyAdded++;
          Logger.log('  ✓ Added BM ID: ' + newBMId);
        } else {
          sheet.getRange(rowNum, CONFIG.DATA_COL_BM_ID).setValue('Not Found');
          Logger.log('  ✗ BM ID not found');
        }
      } catch (rowError) {
        Logger.log('Error auto-populating row ' + rowNum + ': ' + rowError);
      }
    }
    
    if (newlyAdded > 0) {
      Logger.log('✅ Auto-populated ' + newlyAdded + ' new BM IDs');
    }
  } catch (error) {
    Logger.log('Error in autoPopulateMissingBMIds_: ' + error);
  }
}

// ========== STEP 3: UPDATE HOURS (WITH TIMESTAMPS) ==========

function updateHoursForRange_(label, daysBack, outCol, timestampKey) {
  try {
    var token = getScriptProperty_(CONFIG.BM_TOKEN_KEY);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.RANKINGS_SHEET_NAME);
    
    if (!sheet) throw new Error('Rankings sheet not found');
    
    var lastRow = sheet.getLastRow();
    if (lastRow < CONFIG.DATA_FIRST_ROW) {
      Logger.log('No data rows found');
      return;
    }
    
    var numRows = lastRow - CONFIG.DATA_FIRST_ROW + 1;
    var stop = new Date();
    var start = new Date(stop.getTime());
    start.setDate(start.getDate() - daysBack);
    
    Logger.log('[' + label + '] Range: ' + start.toISOString() + ' to ' + stop.toISOString());
    
    var bmIds = sheet.getRange(CONFIG.DATA_FIRST_ROW, CONFIG.DATA_COL_BM_ID, numRows, 1).getValues();
    var output = [];
    var stats = { processed: 0, skipped: 0, errors: 0, quotaStop: false };
    
    for (var i = 0; i < bmIds.length; i++) {
      try {
        var bmId = String(bmIds[i][0] || '').trim();
        var rowNum = CONFIG.DATA_FIRST_ROW + i;
        
        if (!bmId || bmId === 'Not Found' || bmId === 'Error') {
          output.push(['']);
          stats.skipped++;
          continue;
        }
        
        if (!canMakeApiCall_()) {
          Logger.log('[' + label + '] ⚠ Quota limit reached');
          stats.quotaStop = true;
          var existingValue = sheet.getRange(rowNum, outCol).getValue();
          output.push([existingValue || '']);
          continue;
        }
        
        Logger.log('[' + label + '] Row ' + rowNum + ' | BM ID: ' + bmId);
        var seconds = getPlayedTimeSeconds_(bmId, CONFIG.BM_SERVER_ID, start, stop, token);
        Utilities.sleep(250);
        
        if (seconds !== null) {
          var hours = seconds / 3600;
          output.push([hours]);
          stats.processed++;
          Logger.log('  ✓ Hours: ' + hours.toFixed(2));
        } else {
          output.push(['Error']);
          stats.errors++;
          Logger.log('  ✗ Error');
        }
      } catch (rowError) {
        Logger.log('[' + label + '] Error processing row ' + rowNum + ': ' + rowError);
        output.push(['Error']);
        stats.errors++;
      }
    }
    
    if (output.length > 0) {
      sheet.getRange(CONFIG.DATA_FIRST_ROW, outCol, output.length, 1).setValues(output);
    }
    
    // SET TIMESTAMP
    if (stats.processed > 0) {
      setLastUpdateTimestamp_(timestampKey);
    }
    
    Logger.log('[' + label + '] Complete - Processed: ' + stats.processed + ', Skipped: ' + stats.skipped + ', Errors: ' + stats.errors);
    if (stats.quotaStop) {
      Logger.log('  ⚠ Stopped early due to quota');
    }
  } catch (error) {
    Logger.log('FATAL ERROR in updateHoursForRange_ [' + label + ']: ' + error);
    throw error;
  }
}

function updateHours7d() {
  updateHoursForRange_('7d', 7, CONFIG.DATA_COL_HOURS_7D, '7D');
  SpreadsheetApp.getActiveSpreadsheet().toast('7-day hours updated', 'Update Complete', 3);
}

function updateHours30d() {
  updateHoursForRange_('30d', 30, CONFIG.DATA_COL_HOURS_30D, '30D');
  SpreadsheetApp.getActiveSpreadsheet().toast('30-day hours updated', 'Update Complete', 3);
}

function updateHours90d() {
  updateHoursForRange_('90d', 89, CONFIG.DATA_COL_HOURS_90D, '90D');
  SpreadsheetApp.getActiveSpreadsheet().toast('90-day hours updated', 'Update Complete', 3);
}

// ========== STEP 4: CREATE RANKINGS DISPLAY ==========

function STEP4_CreateRankingsDisplay() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.RANKINGS_SHEET_NAME);
    
    if (!sheet) throw new Error('Rankings sheet not found');
    
    var lastRow = sheet.getLastRow();
    if (lastRow < CONFIG.DATA_FIRST_ROW) {
      Logger.log('No data to display');
      return;
    }
    
    var numRows = lastRow - CONFIG.DATA_FIRST_ROW + 1;
    var names = sheet.getRange(CONFIG.DATA_FIRST_ROW, CONFIG.DATA_COL_NAME, numRows, 1).getValues();
    var teams = sheet.getRange(CONFIG.DATA_FIRST_ROW, CONFIG.DATA_COL_TEAM, numRows, 1).getValues();
    var hours7d = sheet.getRange(CONFIG.DATA_FIRST_ROW, CONFIG.DATA_COL_HOURS_7D, numRows, 1).getValues();
    var hours30d = sheet.getRange(CONFIG.DATA_FIRST_ROW, CONFIG.DATA_COL_HOURS_30D, numRows, 1).getValues();
    var hours90d = sheet.getRange(CONFIG.DATA_FIRST_ROW, CONFIG.DATA_COL_HOURS_90D, numRows, 1).getValues();
    
    var allData = [];
    for (var i = 0; i < numRows; i++) {
      var name = String(names[i][0] || '').trim();
      var team = String(teams[i][0] || '').trim();
      var h7 = typeof hours7d[i][0] === 'number' ? hours7d[i][0] : 0;
      var h30 = typeof hours30d[i][0] === 'number' ? hours30d[i][0] : 0;
      var h90 = typeof hours90d[i][0] === 'number' ? hours90d[i][0] : 0;
      
      if (name) {
        allData.push({
          name: name,
          team: team,
          hours7d: h7,
          hours30d: h30,
          hours90d: h90
        });
      }
    }
    
    sheet.getRange(1, 1, sheet.getMaxRows(), 24).clear();
    sheet.clearConditionalFormatRules();
    
    setupSearchBox_(sheet);
    
    var startCol = 1;
    
    createRankingSection_(sheet, allData, 'hours7d', '7-DAY RANKINGS', 1, startCol);
    startCol += 4;
    
    createBlackDivider_(sheet, startCol);
    startCol += 1;
    
    createRankingSection_(sheet, allData, 'hours30d', '30-DAY RANKINGS', 1, startCol);
    startCol += 4;
    
    createBlackDivider_(sheet, startCol);
    startCol += 1;
    
    createRankingSection_(sheet, allData, 'hours90d', '90-DAY RANKINGS', 1, startCol);
    
    createBlackDivider_(sheet, 15);
    createBlackDivider_(sheet, 16);
    createBlackDivider_(sheet, 17);
    createBlackDivider_(sheet, 18);
    
    formatRankingsSheet_(sheet);
    setupDynamicSearchFormatting_(sheet);
    
    Logger.log('✅ Rankings display created');
  } catch (error) {
    Logger.log('Error in STEP4_CreateRankingsDisplay: ' + error);
    throw error;
  }
}

function createBlackDivider_(sheet, col) {
  sheet.setColumnWidth(col, 10);
  var range = sheet.getRange(1, col, sheet.getMaxRows(), 1);
  range.setBackground('#000000');
}

function createRankingSection_(sheet, data, hoursField, title, startRow, startCol) {
  var sorted = data.slice().sort(function(a, b) {
    return b[hoursField] - a[hoursField];
  });
  
  sheet.getRange(startRow, startCol, 1, 4).merge();
  sheet.getRange(startRow, startCol).setValue(title);
  sheet.getRange(startRow, startCol)
    .setFontWeight('bold')
    .setFontSize(14)
    .setBackground('#4285f4')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);
  
  var headers = [['Rank', 'Name', 'Hours', 'Team']];
  sheet.getRange(startRow + 1, startCol, 1, 4).setValues(headers);
  sheet.getRange(startRow + 1, startCol, 1, 4)
    .setFontWeight('bold')
    .setBackground('#d9d9d9')
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  
  if (sorted.length > 0) {
    var output = sorted.map(function(item, index) {
      return [
        index + 1,
        item.name,
        item[hoursField] === 0 ? 0 : item[hoursField].toFixed(2),
        item.team
      ];
    });
    
    sheet.getRange(startRow + 2, startCol, output.length, 4).setValues(output);
    
    applyFullRowColors_(sheet, sorted, hoursField, startRow + 2, startCol);
    applyTeamTextColors_(sheet, sorted, startRow + 2, startCol + 3);
    
    sheet.getRange(startRow + 2, startCol, output.length, 1).setHorizontalAlignment('center');
    sheet.getRange(startRow + 2, startCol + 2, output.length, 1).setHorizontalAlignment('right');
    sheet.getRange(startRow + 2, startCol + 3, output.length, 1).setHorizontalAlignment('center');
    
    sheet.getRange(startRow + 2, startCol, output.length, 4)
      .setBorder(true, true, true, true, false, false, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
  } else {
    sheet.getRange(startRow + 2, startCol).setValue('No admins found');
  }
  
  sheet.setColumnWidth(startCol, 50);
  sheet.setColumnWidth(startCol + 1, 150);
  sheet.setColumnWidth(startCol + 2, 70);
  sheet.setColumnWidth(startCol + 3, 140);
}

function applyFullRowColors_(sheet, sortedData, hoursField, startRow, startCol) {
  for (var i = 0; i < sortedData.length; i++) {
    var hours = sortedData[i][hoursField];
    var color = getColorForHours_(hours);
    
    if (color) {
      var rowRange = sheet.getRange(startRow + i, startCol, 1, 4);
      rowRange.setBackground(color);
    }
  }
}

function applyTeamTextColors_(sheet, sortedData, startRow, teamCol) {
  for (var i = 0; i < sortedData.length; i++) {
    var team = sortedData[i].team;
    var textColor = CONFIG.TEAM_COLORS[team] || '#000000';
    
    var cellRange = sheet.getRange(startRow + i, teamCol);
    cellRange.setFontColor(textColor);
    cellRange.setFontWeight('bold');
    cellRange.setFontSize(10);
  }
}

function getColorForHours_(hours) {
  for (var i = 0; i < CONFIG.COLOR_RANGES.length; i++) {
    var range = CONFIG.COLOR_RANGES[i];
    if (hours >= range.min && hours <= range.max) {
      return range.color;
    }
  }
  return null;
}

function formatRankingsSheet_(sheet) {
  sheet.setFrozenRows(2);
  sheet.autoResizeColumns(1, 24);
}

// ========== ROTATING UPDATE FUNCTION ==========

function updateAllHoursHourly() {
  try {
    var currentHour = new Date().getHours();
    var cycle = currentHour % 3;
    
    Logger.log('=== Starting Rotating Update (Hour: ' + currentHour + ', Cycle: ' + cycle + ') ===');
    Logger.log('Quota - Daily: ' + getQuotaUsage_('daily') + '/' + CONFIG.MAX_DAILY_API_CALLS);
    Logger.log('Quota - Hourly: ' + getQuotaUsage_('hourly') + '/' + CONFIG.MAX_HOURLY_API_CALLS);
    
    STEP1_SyncAdminData();
    autoPopulateMissingBMIds_();
    
    if (cycle === 0) {
      Logger.log('📊 Updating 7-day hours...');
      updateHours7d();
    } else if (cycle === 1) {
      Logger.log('📊 Updating 30-day hours...');
      updateHours30d();
    } else {
      Logger.log('📊 Updating 90-day hours...');
      updateHours90d();
    }
    
    STEP4_CreateRankingsDisplay();
    
    Logger.log('=== Rotating Update Complete ===');
  } catch (error) {
    Logger.log('FATAL ERROR in updateAllHoursHourly: ' + error);
  }
}

function manualFullUpdate() {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Starting manual full update... This may take 10-15 minutes.',
      'Manual Update',
      5
    );
    
    Logger.log('=== Manual Full Update Start ===');
    
    STEP1_SyncAdminData();
    autoPopulateMissingBMIds_();
    
    updateHours7d();
    
    if (canMakeApiCall_()) {
      Utilities.sleep(2000);
      updateHours30d();
    }
    
    if (canMakeApiCall_()) {
      Utilities.sleep(2000);
      updateHours90d();
    }
    
    STEP4_CreateRankingsDisplay();
    
    Logger.log('=== Manual Full Update Complete ===');
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Manual update complete!',
      'Update Complete',
      5
    );
  } catch (error) {
    Logger.log('FATAL ERROR in manualFullUpdate: ' + error);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Error: ' + error.message,
      'Update Failed',
      10
    );
  }
}

// ========== REAL-TIME CHANGE DETECTION ==========

function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();
    var sheetName = sheet.getName();
    
    if (sheetName === CONFIG.RANKINGS_SHEET_NAME) {
      var row = e.range.getRow();
      var col = e.range.getColumn();
      
      if (row === 1 && col === CONFIG.CLEAR_BUTTON_COL) {
        clearSearch();
        return;
      }
    }
    
    if (sheetName !== CONFIG.SOURCE_SHEET_NAME) return;
    
    var row = e.range.getRow();
    var col = e.range.getColumn();
    
    if (row < CONFIG.SOURCE_FIRST_DATA_ROW) return;
    
    if (col === CONFIG.SOURCE_COL_NAME || 
        col === CONFIG.SOURCE_COL_TEAM || 
        col === CONFIG.SOURCE_COL_STEAM_ID) {
      handleSourceDataChange_(sheet, row, col);
    }
  } catch (error) {
    Logger.log('Error in onEdit: ' + error);
  }
}

function handleSourceDataChange_(sheet, row, col) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var rankingsSheet = ss.getSheetByName(CONFIG.RANKINGS_SHEET_NAME);
    
    if (!rankingsSheet) {
      Logger.log('Rankings sheet not found');
      return;
    }
    
    var rankingsRow = row - CONFIG.SOURCE_FIRST_DATA_ROW + CONFIG.DATA_FIRST_ROW;
    
    var name = String(sheet.getRange(row, CONFIG.SOURCE_COL_NAME).getValue() || '').trim();
    var team = String(sheet.getRange(row, CONFIG.SOURCE_COL_TEAM).getValue() || '').trim();
    var steamId = String(sheet.getRange(row, CONFIG.SOURCE_COL_STEAM_ID).getValue() || '').trim();
    
    rankingsSheet.getRange(rankingsRow, CONFIG.DATA_COL_NAME).setValue(name);
    rankingsSheet.getRange(rankingsRow, CONFIG.DATA_COL_TEAM).setValue(team);
    rankingsSheet.getRange(rankingsRow, CONFIG.DATA_COL_STEAM_ID).setValue(steamId);
    
    if (col === CONFIG.SOURCE_COL_STEAM_ID) {
      Logger.log('Row ' + row + ': Steam ID changed to ' + steamId);
      
      rankingsSheet.getRange(rankingsRow, CONFIG.DATA_COL_BM_ID).setValue('');
      rankingsSheet.getRange(rankingsRow, CONFIG.DATA_COL_HOURS_7D).setValue('');
      rankingsSheet.getRange(rankingsRow, CONFIG.DATA_COL_HOURS_30D).setValue('');
      rankingsSheet.getRange(rankingsRow, CONFIG.DATA_COL_HOURS_90D).setValue('');
      
      if (!steamId) {
        Logger.log('  Steam ID cleared');
        return;
      }
      
      try {
        var token = getScriptProperty_(CONFIG.BM_TOKEN_KEY);
        var bmId = getBMIdFromSteamId_(steamId, token);
        
        if (bmId) {
          rankingsSheet.getRange(rankingsRow, CONFIG.DATA_COL_BM_ID).setValue(bmId);
          Logger.log('  ✓ Updated BM ID to ' + bmId);
          
          Utilities.sleep(1000);
          STEP4_CreateRankingsDisplay();
        } else {
          rankingsSheet.getRange(rankingsRow, CONFIG.DATA_COL_BM_ID).setValue('Not Found');
          Logger.log('  ✗ BM ID not found');
        }
      } catch (lookupError) {
        Logger.log('Error looking up BM ID: ' + lookupError);
        rankingsSheet.getRange(rankingsRow, CONFIG.DATA_COL_BM_ID).setValue('Error');
      }
    } else {
      Logger.log('Row ' + row + ': Data updated, refreshing display');
      Utilities.sleep(500);
      STEP4_CreateRankingsDisplay();
    }
  } catch (error) {
    Logger.log('Error in handleSourceDataChange_ for row ' + row + ': ' + error);
  }
}

function refreshRankingsDisplay() {
  STEP4_CreateRankingsDisplay();
  SpreadsheetApp.getActiveSpreadsheet().toast('Rankings refreshed', 'Display Updated', 3);
}

// ========== COMPLETE SETUP FUNCTION ==========

function SETUP_CompleteInitialSetup() {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Starting complete setup... This will take 25-30 minutes.',
      'Setup',
      5
    );
    
    Logger.log('===== COMPLETE SETUP START =====');
    
    Logger.log('Step 1/6: Syncing admin data...');
    STEP1_SyncAdminData();
    Utilities.sleep(2000);
    
    Logger.log('Step 2/6: Populating BM IDs...');
    STEP2_PopulateBMIds();
    Utilities.sleep(2000);
    
    Logger.log('Step 3/6: Updating 7-day hours...');
    updateHours7d();
    Utilities.sleep(2000);
    
    Logger.log('Step 4/6: Updating 30-day hours...');
    updateHours30d();
    Utilities.sleep(2000);
    
    Logger.log('Step 5/6: Updating 90-day hours...');
    updateHours90d();
    Utilities.sleep(2000);
    
    Logger.log('Step 6/6: Creating rankings display...');
    STEP4_CreateRankingsDisplay();
    
    Logger.log('===== COMPLETE SETUP FINISHED =====');
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Setup complete! Rotating updates will run hourly.',
      'Setup Complete',
      10
    );
  } catch (error) {
    Logger.log('FATAL ERROR in SETUP_CompleteInitialSetup: ' + error);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Setup failed: ' + error.message,
      'Setup Error',
      10
    );
  }
}

// ========== UTILITY FUNCTIONS ==========

function checkQuotaStatus() {
  try {
    var daily = getQuotaUsage_('daily');
    var hourly = getQuotaUsage_('hourly');
    Logger.log('=== Quota Status ===');
    Logger.log('Daily: ' + daily + ' / ' + CONFIG.MAX_DAILY_API_CALLS + ' (' + ((daily/CONFIG.MAX_DAILY_API_CALLS)*100).toFixed(1) + '%)');
    Logger.log('Hourly: ' + hourly + ' / ' + CONFIG.MAX_HOURLY_API_CALLS + ' (' + ((hourly/CONFIG.MAX_HOURLY_API_CALLS)*100).toFixed(1) + '%)');
    
    var message = 'Daily: ' + daily + '/' + CONFIG.MAX_DAILY_API_CALLS + '\n' +
                  'Hourly: ' + hourly + '/' + CONFIG.MAX_HOURLY_API_CALLS;
    
    if (daily >= CONFIG.MAX_DAILY_API_CALLS) {
      Logger.log('⚠ WARNING: Daily quota limit reached!');
      message += '\n\n⚠️ Daily quota limit reached!';
    } else if (hourly >= CONFIG.MAX_HOURLY_API_CALLS) {
      Logger.log('⚠ WARNING: Hourly quota limit reached!');
      message += '\n\n⚠️ Hourly quota limit reached!';
    } else {
      Logger.log('✓ Quotas within limits');
      message += '\n\n✅ Quotas within limits';
    }
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      message,
      'Quota Status',
      10
    );
  } catch (error) {
    Logger.log('Error in checkQuotaStatus: ' + error);
  }
}

function testDirectAPICall() {
  try {
    var token = getScriptProperty_(CONFIG.BM_TOKEN_KEY);
    var testBMId = '39394618';
    var serverId = CONFIG.BM_SERVER_ID;
    
    var stop = new Date();
    var start = new Date(stop.getTime());
    start.setDate(start.getDate() - 7);
    
    var url = 'https://api.battlemetrics.com/players/' + testBMId + 
              '/time-played-history/' + serverId +
              '?start=' + start.toISOString() +
              '&stop=' + stop.toISOString() +
              '&access_token=' + encodeURIComponent(token);
    
    Logger.log('Testing BattleMetrics API...');
    Logger.log('URL: ' + url.replace(token, 'HIDDEN'));
    
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var code = response.getResponseCode();
    
    Logger.log('Response Code: ' + code);
    
    if (code === 403) {
      Logger.log('❌ ERROR: Token lacks permissions');
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'API token lacks permissions. Check token settings.',
        'API Test Failed',
        10
      );
    } else if (code === 200) {
      Logger.log('✅ Success! API is working correctly');
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'API connection successful! ✅',
        'API Test Passed',
        5
      );
    } else {
      Logger.log('⚠ Unexpected response: ' + code);
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Unexpected API response: ' + code,
        'API Test Warning',
        10
      );
    }
  } catch (error) {
    Logger.log('Error in testDirectAPICall: ' + error);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Error testing API: ' + error.message,
      'API Test Error',
      10
    );
  }
}
