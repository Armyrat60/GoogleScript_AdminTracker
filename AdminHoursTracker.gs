/***********************
 * Admin Hours Tracker for BattleMetrics
 * 
 * VERSION: 4.0.0
 * Author: ArmyRat60
 * GitHub: https://github.com/Armyrat60/GoogleScript_AdminTracker
 * 
 * A Google Apps Script for tracking admin activity hours on BattleMetrics servers.
 * Works with Squad, Rust, ARK, and any game using BattleMetrics.
 * 
 * Features:
 * - First Time Setup - Single clear path for new users
 * - Formula-based AUTO columns - Always in sync, never fail
 * - Enhanced onEdit - Fetches ALL data when you type Steam ID
 * - Smart Trigger Manager - Safely manages triggers, protects other scripts
 * - Debug Tools - 7 diagnostic tools for troubleshooting
 * - Sequential daily updates - Runs at 2 AM, all data updated
 * - Migration tool - Upgrade from v3.x seamlessly
 * - Organized menu - SETUP / UPDATE / TRIGGERS / SETTINGS
 * 
 * Setup: Menu → ⚙️ Admin Tracker → 🛠️ SETUP → First Time Setup
 * GitHub: https://github.com/Armyrat60/GoogleScript_AdminTracker
 ***********************/

const VERSION = '4.0.0';
const GITHUB_VERSION_URL = 'https://raw.githubusercontent.com/Armyrat60/GoogleScript_AdminTracker/main/version.json';
const GITHUB_DOWNLOAD_URL = 'https://github.com/Armyrat60/GoogleScript_AdminTracker';

// ========== CONFIGURATION HELPER ==========

function getConfig(key) {
  // HYBRID APPROACH: Check Script Properties WITHOUT CONFIG_ prefix first
  // This is what the new Setup Wizard uses
  // Falls back to Config.gs if not found
  
  var scriptPropertyKeys = ['BM_SERVER_ID', 'SOURCE_COL_NAME', 'SOURCE_COL_TEAM', 'SOURCE_COL_STEAM_ID'];
  
  if (scriptPropertyKeys.indexOf(key) !== -1) {
    try {
      var props = PropertiesService.getScriptProperties();
      // Check WITHOUT CONFIG_ prefix (new wizard)
      var value = props.getProperty(key);
      if (value !== null && value !== '') {
        if (!isNaN(value) && value.trim() !== '') {
          Logger.log('Using Script Property (no prefix) for ' + key + ': ' + value);
          return Number(value);
        }
        Logger.log('Using Script Property (no prefix) for ' + key + ': ' + value);
        return value;
      }
    } catch (error) {
      // Ignore and fall through to Config.gs
    }
  }
  
  // Read from CONFIG object from Config.gs (fallback)
  if (typeof CONFIG !== 'undefined' && CONFIG[key] !== undefined) {
    return CONFIG[key];
  }
  
  throw new Error('Configuration value not found in Script Properties or Config.gs: ' + key);
}

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
    'Admin Hours Tracker for BattleMetrics\n' +
    'Version: ' + VERSION + '\n\n' +
    'Features:\n' +
    '• Setup Wizard for easy configuration\n' +
    '• Rotating hourly updates (no timeouts)\n' +
    '• Real-time sync from source data\n' +
    '• Automatic trigger creation\n' +
    '• Last update timestamps\n' +
    '• Automatic version checking\n\n' +
    'GitHub: ' + GITHUB_DOWNLOAD_URL;
  
  ui.alert('About Admin Tracker', message, ui.ButtonSet.OK);
}


// ========== SETUP WIZARD ==========

function setupWizard() {
  var ui = SpreadsheetApp.getUi();
  
  // Welcome message
  ui.alert(
    '🛠️ Admin Hours Tracker Setup',
    'This wizard will guide you through the complete setup.\n\n' +
    'You\'ll need:\n' +
    '✓ Your BattleMetrics Server ID\n' +
    '✓ Your BattleMetrics API Token\n\n' +
    'Setup takes about 2 minutes.\n\n' +
    'Click OK to begin!',
    ui.ButtonSet.OK_CANCEL
  );
  
  // ========== STEP 1: BattleMetrics Server ID ==========
  var serverIdResult = ui.prompt(
    'Step 1/3: BattleMetrics Server ID',
    'Find your Server ID in the BattleMetrics URL:\n' +
    'Example: battlemetrics.com/servers/squad/12345678\n\n' +
    'The number at the end (12345678) is your Server ID.\n\n' +
    'Enter your Server ID:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (serverIdResult.getSelectedButton() !== ui.Button.OK) {
    ui.alert('Setup Cancelled', 'You can run the Setup Wizard again anytime from the menu.', ui.ButtonSet.OK);
    return;
  }
  
  var serverId = serverIdResult.getResponseText().trim();
  if (!serverId || serverId === '' || isNaN(serverId)) {
    ui.alert('Invalid Server ID', 'Please enter a valid numeric Server ID.\n\nRun the Setup Wizard again.', ui.ButtonSet.OK);
    return;
  }
  
  Logger.log('✓ Server ID entered: ' + serverId);
  
  // ========== STEP 2: BattleMetrics API Token ==========
  var apiKeyResult = ui.prompt(
    'Step 2/3: BattleMetrics API Token',
    'Get your API token from:\n' +
    'https://www.battlemetrics.com/developers\n\n' +
    'Your token should look like:\n' +
    'eyJhbGciOiJIUzI1NiIsInR5cCI6...\n\n' +
    'Paste your BattleMetrics API token:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (apiKeyResult.getSelectedButton() !== ui.Button.OK) {
    ui.alert('Setup Cancelled', 'You can run the Setup Wizard again anytime from the menu.', ui.ButtonSet.OK);
    return;
  }
  
  var apiKey = apiKeyResult.getResponseText().trim();
  if (!apiKey || apiKey === '' || apiKey.length < 50) {
    ui.alert('Invalid API Token', 'The API token seems too short or invalid.\n\nPlease check and run the Setup Wizard again.', ui.ButtonSet.OK);
    return;
  }
  
  Logger.log('✓ API token entered (length: ' + apiKey.length + ')');
  
  // ========== STEP 3: Auto-detect Columns ==========
  ui.alert(
    'Step 3/3: Auto-Detecting Columns',
    'The wizard will now scan your "Server Admin Team" sheet\n' +
    'to find all required columns.\n\n' +
    'If columns are missing, you can create them automatically!\n\n' +
    'Click OK to continue.',
    ui.ButtonSet.OK
  );
  
  // Find source sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName('Server Admin Team');
  
  if (!sourceSheet) {
    ui.alert('Error', 'Sheet "Server Admin Team" not found!\n\nPlease create or rename your admin sheet to "Server Admin Team".', ui.ButtonSet.OK);
    return;
  }
  
  // Get all headers
  var headers = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getDisplayValues()[0];
  
  // Define all columns we're looking for with multiple search variations
  var columnDefs = {
    name: {
      label: 'Steam Name (AUTO)',
      searches: ['steam name (auto)', 'steam name auto', 'steamname (auto)', 'steam name', 'name (auto)'],
      col: 0,
      header: '',
      required: true
    },
    team: {
      label: 'Team',
      searches: ['team', 'role', 'rank'],
      col: 0,
      header: '',
      required: true
    },
    steamId: {
      label: 'Steam ID',
      searches: ['steam id', 'steamid', 'discord id', 'steam64'],
      col: 0,
      header: '',
      required: true
    },
    bmId: {
      label: 'BMID (AUTO)',
      searches: ['bmid (auto)', 'bmid auto', 'bm id (auto)', 'battlemetrics id (auto)', 'bmid'],
      col: 0,
      header: '',
      required: false
    },
    lastSeen: {
      label: 'Last Seen (Auto)',
      searches: ['last seen (auto)', 'last seen auto', 'lastseen (auto)', 'last seen'],
      col: 0,
      header: '',
      required: false
    },
    firstSeen: {
      label: 'First Seen (Auto)',
      searches: ['first seen (auto)', 'first seen auto', 'firstseen (auto)', 'first seen'],
      col: 0,
      header: '',
      required: false
    },
    eosId: {
      label: 'EOSID (Auto)',
      searches: ['eosid (auto)', 'eosid auto', 'eos id (auto)', 'epic id (auto)', 'eosid'],
      col: 0,
      header: '',
      required: false
    }
  };
  
  // Auto-detect each column
  for (var key in columnDefs) {
    var def = columnDefs[key];
    for (var i = 0; i < headers.length; i++) {
      var header = String(headers[i]).toLowerCase().trim();
      for (var j = 0; j < def.searches.length; j++) {
        if (header === def.searches[j] || (def.searches[j].indexOf('(auto)') === -1 && header.includes(def.searches[j]))) {
          def.col = i + 1;
          def.header = headers[i];
          break;
        }
      }
      if (def.col > 0) break;
    }
  }
  
  // Check if all REQUIRED columns found
  var missingRequired = [];
  var missingOptional = [];
  
  for (var key in columnDefs) {
    var def = columnDefs[key];
    if (def.col === 0) {
      if (def.required) {
        missingRequired.push(def.label);
      } else {
        missingOptional.push(def.label);
      }
    }
  }
  
  // Build detection report
  var reportMessage = '📊 Column Detection Results:\n\n';
  reportMessage += '✅ FOUND:\n';
  for (var key in columnDefs) {
    if (columnDefs[key].col > 0) {
      reportMessage += '  • ' + columnDefs[key].label + ': Column ' + columnLetter(columnDefs[key].col) + '\n';
    }
  }
  
  if (missingRequired.length > 0 || missingOptional.length > 0) {
    reportMessage += '\n❌ NOT FOUND:\n';
    if (missingRequired.length > 0) {
      reportMessage += '  REQUIRED:\n';
      for (var i = 0; i < missingRequired.length; i++) {
        reportMessage += '    • ' + missingRequired[i] + '\n';
      }
    }
    if (missingOptional.length > 0) {
      reportMessage += '  OPTIONAL (AUTO columns):\n';
      for (var i = 0; i < missingOptional.length; i++) {
        reportMessage += '    • ' + missingOptional[i] + '\n';
      }
    }
  }
  
  // Handle missing columns
  if (missingRequired.length > 0) {
    ui.alert(
      'Missing Required Columns',
      reportMessage + '\n\nYou must have these columns to continue.\n\nPlease add them to your sheet and run Setup Wizard again.',
      ui.ButtonSet.OK
    );
    return;
  }
  
  // If optional columns missing, offer to create them
  if (missingOptional.length > 0) {
    var createResult = ui.alert(
      'Missing AUTO Columns',
      reportMessage + '\n\nWould you like to automatically create the missing AUTO columns?\n\n' +
      '• YES = Create them automatically\n' +
      '• NO = Skip (you can add them later)\n' +
      '• CANCEL = Exit setup',
      ui.ButtonSet.YES_NO_CANCEL
    );
    
    if (createResult === ui.Button.CANCEL) {
      ui.alert('Setup Cancelled', 'You can run the Setup Wizard again anytime.', ui.ButtonSet.OK);
      return;
    }
    
    if (createResult === ui.Button.YES) {
      // Create missing AUTO columns
      var nextCol = sourceSheet.getLastColumn() + 1;
      
      for (var key in columnDefs) {
        var def = columnDefs[key];
        if (def.col === 0 && !def.required) {
          sourceSheet.getRange(1, nextCol).setValue(def.label);
          def.col = nextCol;
          def.header = def.label;
          Logger.log('✓ Created column: ' + def.label + ' at column ' + nextCol);
          nextCol++;
        }
      }
      
      ui.alert(
        '✅ Columns Created!',
        'Successfully created missing AUTO columns!\n\n' +
        'All columns are now ready.',
        ui.ButtonSet.OK
      );
    }
  }
  
  // Final confirmation
  var confirmMessage = '✅ Setup Ready!\n\n';
  confirmMessage += 'Detected Columns:\n';
  confirmMessage += '• Name: Column ' + columnLetter(columnDefs.name.col) + ' (' + columnDefs.name.header + ')\n';
  confirmMessage += '• Team: Column ' + columnLetter(columnDefs.team.col) + ' (' + columnDefs.team.header + ')\n';
  confirmMessage += '• Steam ID: Column ' + columnLetter(columnDefs.steamId.col) + ' (' + columnDefs.steamId.header + ')\n';
  if (columnDefs.bmId.col > 0) {
    confirmMessage += '• BMID: Column ' + columnLetter(columnDefs.bmId.col) + ' (' + columnDefs.bmId.header + ')\n';
  }
  if (columnDefs.lastSeen.col > 0) {
    confirmMessage += '• Last Seen: Column ' + columnLetter(columnDefs.lastSeen.col) + ' (' + columnDefs.lastSeen.header + ')\n';
  }
  if (columnDefs.firstSeen.col > 0) {
    confirmMessage += '• First Seen: Column ' + columnLetter(columnDefs.firstSeen.col) + ' (' + columnDefs.firstSeen.header + ')\n';
  }
  if (columnDefs.eosId.col > 0) {
    confirmMessage += '• EOSID: Column ' + columnLetter(columnDefs.eosId.col) + ' (' + columnDefs.eosId.header + ')\n';
  }
  confirmMessage += '\nProceed with setup?';
  
  var finalConfirm = ui.alert('Confirm Configuration', confirmMessage, ui.ButtonSet.YES_NO);
  
  if (finalConfirm !== ui.Button.YES) {
    ui.alert('Setup Cancelled', 'You can run the Setup Wizard again anytime.', ui.ButtonSet.OK);
    return;
  }
  
  // ========== SAVE EVERYTHING ==========
  try {
    var props = PropertiesService.getScriptProperties();
    
    // Save Server ID and API Key
    props.setProperty('BM_SERVER_ID', serverId);
    props.setProperty('BATTLEMETRICS_API_KEY', apiKey);
    Logger.log('✓ Saved Server ID and API Key');
    
    // Save ALL detected columns
    props.setProperty('SOURCE_COL_NAME', String(columnDefs.name.col));
    props.setProperty('SOURCE_COL_TEAM', String(columnDefs.team.col));
    props.setProperty('SOURCE_COL_STEAM_ID', String(columnDefs.steamId.col));
    Logger.log('✓ Saved required columns');
    
    // Save optional AUTO columns if detected
    if (columnDefs.bmId.col > 0) {
      props.setProperty('SOURCE_COL_BMID', String(columnDefs.bmId.col));
      Logger.log('✓ Saved BMID column: ' + columnDefs.bmId.col);
    }
    if (columnDefs.lastSeen.col > 0) {
      props.setProperty('SOURCE_COL_LAST_SEEN', String(columnDefs.lastSeen.col));
      Logger.log('✓ Saved Last Seen column: ' + columnDefs.lastSeen.col);
    }
    if (columnDefs.firstSeen.col > 0) {
      props.setProperty('SOURCE_COL_FIRST_SEEN', String(columnDefs.firstSeen.col));
      Logger.log('✓ Saved First Seen column: ' + columnDefs.firstSeen.col);
    }
    if (columnDefs.eosId.col > 0) {
      props.setProperty('SOURCE_COL_EOSID', String(columnDefs.eosId.col));
      Logger.log('✓ Saved EOSID column: ' + columnDefs.eosId.col);
    }
    
    // ========== DATA POPULATION CHOICE ==========
    var popupMsg = '🎉 Configuration Saved!\n\n';
    popupMsg += '✅ Server ID saved\n';
    popupMsg += '✅ API Token saved\n';
    popupMsg += '✅ All columns configured\n\n';
    popupMsg += '📊 Next: Populate Your Data\n\n';
    popupMsg += 'Choose how to proceed:\n\n';
    popupMsg += '👉 QUICK SETUP (Recommended)\n';
    popupMsg += '   • Time: ~2 minutes\n';
    popupMsg += '   • Populates: BM IDs, Names, EOSID, Dates\n';
    popupMsg += '   • Hours: Update automatically via triggers\n';
    popupMsg += '   • Rankings: Created by hourly trigger\n\n';
    popupMsg += '👉 FULL SETUP\n';
    popupMsg += '   • Time: ~8 minutes (may timeout!)\n';
    popupMsg += '   • Populates: Everything immediately\n';
    popupMsg += '   • Only for servers with < 30 players\n\n';
    popupMsg += '👉 MANUAL\n';
    popupMsg += '   • Skip for now\n';
    popupMsg += '   • You\'ll populate data manually later\n\n';
    popupMsg += 'What would you like to do?';
    
    var choice = ui.alert(
      'Choose Setup Type',
      popupMsg,
      ui.ButtonSet.YES_NO_CANCEL
    );
    
    if (choice === ui.Button.YES) {
      // Quick Setup
      Logger.log('User chose: Quick Setup');
      try {
        quickSetupData();
      } catch (e) {
        ui.alert(
          'Quick Setup Error',
          'An error occurred during Quick Setup:\n\n' + e + '\n\n' +
          'You can try again from Menu → Complete Initial Setup.',
          ui.ButtonSet.OK
        );
      }
    } else if (choice === ui.Button.NO) {
      // Full Setup
      Logger.log('User chose: Full Setup');
      var warning = ui.alert(
        'Full Setup Warning',
        'Full Setup may timeout for servers with 50+ players.\n\n' +
        'This will:\n' +
        '• Populate ALL data immediately\n' +
        '• Take ~8 minutes\n' +
        '• May hit 6-minute execution limit\n\n' +
        'Continue with Full Setup?',
        ui.ButtonSet.YES_NO
      );
      
      if (warning === ui.Button.YES) {
        try {
          SETUP_CompleteInitialSetup();
        } catch (e) {
          ui.alert(
            'Full Setup Error',
            'An error occurred:\n\n' + e + '\n\n' +
            'Try Quick Setup instead, or populate manually.',
            ui.ButtonSet.OK
          );
        }
      } else {
        ui.alert(
          'Setup Saved',
          'Configuration is saved!\n\n' +
          'Run "Quick Setup" or "Complete Initial Setup"\n' +
          'from the menu when ready.',
          ui.ButtonSet.OK
        );
      }
    } else {
      // Manual
      Logger.log('User chose: Manual setup');
      ui.alert(
        'Configuration Saved!',
        'Your settings are saved!\n\n' +
        'To populate data manually:\n\n' +
        '1. Menu → Setup & Configuration → Quick Setup\n' +
        '   (or Complete Initial Setup)\n\n' +
        '2. Menu → Manual Updates → Sync Admin Data\n\n' +
        '3. Menu → Manual Updates → Populate BM IDs\n\n' +
        '4. Menu → Manual Updates → Update Hours\n\n' +
        'Don\'t forget to create triggers:\n' +
        'Menu → Setup & Configuration → Quick Fix - Auto-Sync Triggers',
        ui.ButtonSet.OK
      );
    }
    
    Logger.log('=== Setup Wizard Completed Successfully ===');
    Logger.log('Server ID: ' + serverId);
    Logger.log('Columns saved:');
    Logger.log('  Name: ' + columnDefs.name.col);
    Logger.log('  Team: ' + columnDefs.team.col);
    Logger.log('  Steam ID: ' + columnDefs.steamId.col);
    Logger.log('  BMID: ' + columnDefs.bmId.col);
    Logger.log('  Last Seen: ' + columnDefs.lastSeen.col);
    Logger.log('  First Seen: ' + columnDefs.firstSeen.col);
    Logger.log('  EOSID: ' + columnDefs.eosId.col);
    
  } catch (error) {
    Logger.log('Error saving configuration: ' + error);
    ui.alert(
      'Save Error',
      'Failed to save configuration:\n' +
      error.toString() + '\n\n' +
      'Please try running the Setup Wizard again.',
      ui.ButtonSet.OK
    );
  }
}

function columnLetter(col) {
  if (!col || col === 0) return '?';
  var letter = '';
  while (col > 0) {
    var temp = (col - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    col = (col - temp - 1) / 26;
  }
  return letter;
}

// ========== AUTO COLUMN FORMULA MANAGEMENT ==========

/**
 * Creates VLOOKUP formulas in AUTO columns that reference Rankings sheet
 * This makes AUTO columns always in sync with Rankings sheet data
 * Called during First Time Setup and can be run manually to repair
 */
function createAutoColumnFormulas() {
  try {
    var ui = SpreadsheetApp.getUi();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sourceSheet = ss.getSheetByName(getConfig("SOURCE_SHEET_NAME"));
    var rankingsSheet = ss.getSheetByName(getConfig("RANKINGS_SHEET_NAME"));
    
    if (!sourceSheet) {
      throw new Error('Source sheet "' + getConfig("SOURCE_SHEET_NAME") + '" not found');
    }
    
    if (!rankingsSheet) {
      throw new Error('Rankings sheet "' + getConfig("RANKINGS_SHEET_NAME") + '" not found. Run First Time Setup first.');
    }
    
    var lastRow = sourceSheet.getLastRow();
    if (lastRow < 2) {
      ui.alert('No Data', 'No data rows found in source sheet.', ui.ButtonSet.OK);
      return;
    }
    
    Logger.log('Creating AUTO column formulas...');
    Logger.log('Source sheet: ' + sourceSheet.getName());
    Logger.log('Rankings sheet: ' + rankingsSheet.getName());
    Logger.log('Last row: ' + lastRow);
    
    // Get column numbers from config
    var steamIdCol = getConfig("SOURCE_COL_STEAM_ID");
    var lastSeenCol = getConfig("SOURCE_COL_LAST_SEEN");
    var firstSeenCol = getConfig("SOURCE_COL_FIRST_SEEN");
    var nameCol = getConfig("SOURCE_COL_NAME");
    var bmidCol = getConfig("SOURCE_COL_BMID");
    var eosidCol = getConfig("SOURCE_COL_EOSID");
    
    // Rankings sheet uses columns AA-AJ (27-36)
    var rankingsRange = '$AA:$AJ';
    
    // Column S - Last Seen (column AE in Rankings = position 5)
    var lastSeenFormula = '=IFERROR(VLOOKUP($' + columnLetter(steamIdCol) + '2,\'Admin Rankings\'!' + rankingsRange + ',5,FALSE),"")';
    sourceSheet.getRange(2, lastSeenCol, lastRow - 1, 1).setFormula(lastSeenFormula);
    Logger.log('✓ Created Last Seen formulas in column ' + columnLetter(lastSeenCol));
    
    // Column T - First Seen (column AF in Rankings = position 6)
    var firstSeenFormula = '=IFERROR(VLOOKUP($' + columnLetter(steamIdCol) + '2,\'Admin Rankings\'!' + rankingsRange + ',6,FALSE),"")';
    sourceSheet.getRange(2, firstSeenCol, lastRow - 1, 1).setFormula(firstSeenFormula);
    Logger.log('✓ Created First Seen formulas in column ' + columnLetter(firstSeenCol));
    
    // Column U - Steam Name (column AB in Rankings = position 2)
    var nameFormula = '=IFERROR(VLOOKUP($' + columnLetter(steamIdCol) + '2,\'Admin Rankings\'!' + rankingsRange + ',2,FALSE),"")';
    sourceSheet.getRange(2, nameCol, lastRow - 1, 1).setFormula(nameFormula);
    Logger.log('✓ Created Name formulas in column ' + columnLetter(nameCol));
    
    // Column V - BMID (column AD in Rankings = position 4)
    var bmidFormula = '=IFERROR(VLOOKUP($' + columnLetter(steamIdCol) + '2,\'Admin Rankings\'!' + rankingsRange + ',4,FALSE),"")';
    sourceSheet.getRange(2, bmidCol, lastRow - 1, 1).setFormula(bmidFormula);
    Logger.log('✓ Created BMID formulas in column ' + columnLetter(bmidCol));
    
    // Column W - EOSID (column AG in Rankings = position 7)
    var eosidFormula = '=IFERROR(VLOOKUP($' + columnLetter(steamIdCol) + '2,\'Admin Rankings\'!' + rankingsRange + ',7,FALSE),"")';
    sourceSheet.getRange(2, eosidCol, lastRow - 1, 1).setFormula(eosidFormula);
    Logger.log('✓ Created EOSID formulas in column ' + columnLetter(eosidCol));
    
    Logger.log('=== AUTO Column Formulas Created Successfully ===');
    
    ui.alert(
      '✅ Formulas Created!',
      'AUTO column formulas have been created.\n\n' +
      'These formulas will automatically pull data from the Rankings sheet:\n' +
      '• Column S: Last Seen\n' +
      '• Column T: First Seen\n' +
      '• Column U: Steam Name\n' +
      '• Column V: BMID\n' +
      '• Column W: EOSID\n\n' +
      'Data will update automatically when Rankings sheet changes!',
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log('ERROR creating AUTO column formulas: ' + error);
    SpreadsheetApp.getUi().alert(
      'Error Creating Formulas',
      'Failed to create formulas:\n\n' + error + '\n\n' +
      'Check that:\n' +
      '1. Rankings sheet exists\n' +
      '2. AUTO columns exist (S, T, U, V, W)\n' +
      '3. Column mappings are correct',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Validates that AUTO column formulas are correct
 * Returns true if all formulas are present and correct
 */
function validateAutoColumnFormulas() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sourceSheet = ss.getSheetByName(getConfig("SOURCE_SHEET_NAME"));
    
    if (!sourceSheet) return false;
    
    var lastSeenCol = getConfig("SOURCE_COL_LAST_SEEN");
    var firstSeenCol = getConfig("SOURCE_COL_FIRST_SEEN");
    var nameCol = getConfig("SOURCE_COL_NAME");
    var bmidCol = getConfig("SOURCE_COL_BMID");
    var eosidCol = getConfig("SOURCE_COL_EOSID");
    
    // Check if formulas exist in row 2
    var formulas = sourceSheet.getRange(2, 1, 1, sourceSheet.getLastColumn()).getFormulas()[0];
    
    var hasLastSeen = formulas[lastSeenCol - 1] && formulas[lastSeenCol - 1].indexOf('VLOOKUP') > -1;
    var hasFirstSeen = formulas[firstSeenCol - 1] && formulas[firstSeenCol - 1].indexOf('VLOOKUP') > -1;
    var hasName = formulas[nameCol - 1] && formulas[nameCol - 1].indexOf('VLOOKUP') > -1;
    var hasBmid = formulas[bmidCol - 1] && formulas[bmidCol - 1].indexOf('VLOOKUP') > -1;
    var hasEosid = formulas[eosidCol - 1] && formulas[eosidCol - 1].indexOf('VLOOKUP') > -1;
    
    return hasLastSeen && hasFirstSeen && hasName && hasBmid && hasEosid;
    
  } catch (error) {
    Logger.log('Error validating formulas: ' + error);
    return false;
  }
}


// ========== TRIGGER MANAGEMENT ==========

/**
 * Smart Trigger Manager - Scans, identifies, and fixes triggers
 * Protects triggers from other scripts
 */
function manageTriggers() {
  var ui = SpreadsheetApp.getUi();
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var allTriggers = ScriptApp.getUserTriggers(ss);
    
    // Admin Tracker function names to look for
    var ourFunctions = [
      'updateAllHoursDaily',
      'updateAllHoursHourly', // Old function
      'onEditHandler',
      'handleSourceEdit',
      'onChangeHandler',
      'handleSheetChange'
    ];
    
    // Classify triggers
    var ourTriggers = [];
    var otherTriggers = [];
    
    for (var i = 0; i < allTriggers.length; i++) {
      var funcName = allTriggers[i].getHandlerFunction();
      var isOurs = false;
      
      for (var j = 0; j < ourFunctions.length; j++) {
        if (funcName === ourFunctions[j]) {
          isOurs = true;
          break;
        }
      }
      
      if (isOurs) {
        ourTriggers.push(allTriggers[i]);
      } else {
        otherTriggers.push(allTriggers[i]);
      }
    }
    
    // Analyze our triggers
    var dailyCount = 0;
    var onEditCount = 0;
    var onChangeCount = 0;
    var oldHourlyCount = 0;
    
    for (var i = 0; i < ourTriggers.length; i++) {
      var funcName = ourTriggers[i].getHandlerFunction();
      if (funcName === 'updateAllHoursDaily') dailyCount++;
      else if (funcName === 'updateAllHoursHourly') oldHourlyCount++;
      else if (funcName === 'onEditHandler' || funcName === 'handleSourceEdit') onEditCount++;
      else if (funcName === 'onChangeHandler' || funcName === 'handleSheetChange') onChangeCount++;
    }
    
    // Determine what needs to be fixed
    var issues = [];
    var fixes = [];
    
    if (dailyCount === 0) {
      issues.push('❌ Daily update trigger - Missing');
      fixes.push('create_daily');
    } else if (dailyCount > 1) {
      issues.push('⚠️ Daily update trigger - ' + dailyCount + ' found (duplicate!)');
      fixes.push('fix_daily');
    } else {
      issues.push('✅ Daily update trigger - OK');
    }
    
    if (oldHourlyCount > 0) {
      issues.push('⚠️ Old hourly trigger - ' + oldHourlyCount + ' found (should be removed)');
      fixes.push('remove_old_hourly');
    }
    
    if (onEditCount === 0) {
      issues.push('❌ onEdit trigger - Missing');
      fixes.push('create_onEdit');
    } else if (onEditCount > 1) {
      issues.push('⚠️ onEdit trigger - ' + onEditCount + ' found (duplicate!)');
      fixes.push('fix_onEdit');
    } else {
      issues.push('✅ onEdit trigger - OK');
    }
    
    if (onChangeCount === 0) {
      issues.push('❌ onChange trigger - Missing');
      fixes.push('create_onChange');
    } else if (onChangeCount > 1) {
      issues.push('⚠️ onChange trigger - ' + onChangeCount + ' found (duplicate!)');
      fixes.push('fix_onChange');
    } else {
      issues.push('✅ onChange trigger - OK');
    }
    
    // Build report
    var report = '🔍 Trigger Status Report\n\n';
    report += '═══ ADMIN TRACKER TRIGGERS ═══\n';
    report += issues.join('\n');
    report += '\n\n';
    
    if (otherTriggers.length > 0) {
      report += '═══ OTHER TRIGGERS (PROTECTED) ═══\n';
      for (var i = 0; i < otherTriggers.length; i++) {
        report += '✓ ' + otherTriggers[i].getHandlerFunction() + ' - Will not modify\n';
      }
      report += '\n';
    }
    
    if (fixes.length === 0) {
      report += '✅ All triggers are configured correctly!';
      ui.alert('Trigger Manager', report, ui.ButtonSet.OK);
      return;
    }
    
    report += '\n❓ Would you like to fix these issues?';
    
    var result = ui.alert('Trigger Manager', report, ui.ButtonSet.YES_NO);
    
    if (result !== ui.Button.YES) {
      return;
    }
    
    // Apply fixes
    var fixedCount = 0;
    
    for (var i = 0; i < fixes.length; i++) {
      var fix = fixes[i];
      
      if (fix === 'create_daily') {
        createDailyTrigger();
        fixedCount++;
      } else if (fix === 'fix_daily') {
        // Delete all daily triggers and recreate one
        for (var j = 0; j < ourTriggers.length; j++) {
          if (ourTriggers[j].getHandlerFunction() === 'updateAllHoursDaily') {
            ScriptApp.deleteTrigger(ourTriggers[j]);
          }
        }
        createDailyTrigger();
        fixedCount++;
      } else if (fix === 'remove_old_hourly') {
        for (var j = 0; j < ourTriggers.length; j++) {
          if (ourTriggers[j].getHandlerFunction() === 'updateAllHoursHourly') {
            ScriptApp.deleteTrigger(ourTriggers[j]);
            fixedCount++;
          }
        }
      } else if (fix === 'create_onEdit') {
        createOnEditTrigger();
        fixedCount++;
      } else if (fix === 'fix_onEdit') {
        for (var j = 0; j < ourTriggers.length; j++) {
          var func = ourTriggers[j].getHandlerFunction();
          if (func === 'onEditHandler' || func === 'handleSourceEdit') {
            ScriptApp.deleteTrigger(ourTriggers[j]);
          }
        }
        createOnEditTrigger();
        fixedCount++;
      } else if (fix === 'create_onChange') {
        createOnChangeTrigger();
        fixedCount++;
      } else if (fix === 'fix_onChange') {
        for (var j = 0; j < ourTriggers.length; j++) {
          var func = ourTriggers[j].getHandlerFunction();
          if (func === 'onChangeHandler' || func === 'handleSheetChange') {
            ScriptApp.deleteTrigger(ourTriggers[j]);
          }
        }
        createOnChangeTrigger();
        fixedCount++;
      }
    }
    
    ui.alert(
      '✅ Triggers Fixed!',
      'Fixed ' + fixedCount + ' issue(s).\n\n' +
      'Your triggers are now configured correctly!\n\n' +
      '✓ Daily update trigger\n' +
      '✓ onEdit trigger\n' +
      '✓ onChange trigger\n\n' +
      'Other scripts\' triggers were NOT modified.',
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log('ERROR in manageTriggers: ' + error);
    ui.alert(
      'Trigger Manager Error',
      'An error occurred:\n\n' + error,
      ui.ButtonSet.OK
    );
  }
}

function deleteAllTriggers() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var triggers = ScriptApp.getUserTriggers(ss);
    var ui = SpreadsheetApp.getUi();
    
    if (triggers.length === 0) {
      ui.alert('No Triggers', 'No triggers found to delete.', ui.ButtonSet.OK);
      return;
    }
    
    var result = ui.alert(
      'Delete All Triggers',
      'This will delete ALL ' + triggers.length + ' trigger(s).\n\n' +
      'This includes:\n' +
      '• Hourly update triggers\n' +
      '• onChange triggers (row inserts)\n' +
      '• onEdit triggers (cell edits)\n' +
      '• Any old/broken triggers\n\n' +
      'You can recreate them with "Quick Fix" after cleanup.\n\n' +
      'Continue?',
      ui.ButtonSet.YES_NO
    );
    
    if (result !== ui.Button.YES) {
      return;
    }
    
    var deleted = 0;
    for (var i = 0; i < triggers.length; i++) {
      ScriptApp.deleteTrigger(triggers[i]);
      deleted++;
      Logger.log('Deleted trigger: ' + triggers[i].getHandlerFunction());
    }
    
    ui.alert(
      'Triggers Deleted',
      '✅ Deleted ' + deleted + ' trigger(s)\n\n' +
      'Run "Quick Fix - Auto-Sync Triggers" to recreate all needed triggers.',
      ui.ButtonSet.OK
    );
    
    Logger.log('✓ Deleted all ' + deleted + ' triggers');
  } catch (error) {
    Logger.log('Error deleting triggers: ' + error);
    SpreadsheetApp.getUi().alert('Error', 'Failed to delete triggers: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function createHourlyTrigger() {
  try {
    // Delete existing triggers first
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'updateAllHoursHourly') {
        ScriptApp.deleteTrigger(triggers[i]);
        Logger.log('Deleted existing trigger');
      }
    }
    
    // Create new trigger
    ScriptApp.newTrigger('updateAllHoursHourly')
      .timeBased()
      .everyHours(1)
      .create();
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Hourly trigger created successfully!',
      'Trigger Created',
      3
    );
    
    Logger.log('✓ Created hourly trigger for updateAllHoursHourly');
  } catch (error) {
    Logger.log('Error creating trigger: ' + error);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Error creating trigger: ' + error.message,
      'Trigger Error',
      5
    );
  }
}

/**
 * Create Daily Trigger - Runs sequential update daily at 2 AM
 */
function createDailyTrigger() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Delete existing daily triggers first
    var triggers = ScriptApp.getUserTriggers(ss);
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'updateAllHoursDaily') {
        ScriptApp.deleteTrigger(triggers[i]);
        Logger.log('Deleted existing daily trigger');
      }
    }
    
    // Create new daily trigger - runs at 2 AM
    ScriptApp.newTrigger('updateAllHoursDaily')
      .timeBased()
      .atHour(2)
      .everyDays(1)
      .create();
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Daily update trigger created! Will run at 2:00 AM daily.',
      'Trigger Created',
      3
    );
    
    Logger.log('✓ Created daily trigger for updateAllHoursDaily');
  } catch (error) {
    Logger.log('Error creating daily trigger: ' + error);
    throw error;
  }
}

function createOnChangeTrigger() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Delete existing onChange triggers first
    var triggers = ScriptApp.getUserTriggers(ss);
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'onChange') {
        ScriptApp.deleteTrigger(triggers[i]);
        Logger.log('Deleted existing onChange trigger');
      }
    }
    
    // Create new onChange trigger
    ScriptApp.newTrigger('onChange')
      .forSpreadsheet(ss)
      .onChange()
      .create();
    
    Logger.log('✓ Created onChange trigger for row insertion detection');
  } catch (error) {
    Logger.log('Error creating onChange trigger: ' + error);
    throw error;
  }
}

function createOnEditTrigger() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Delete existing onEdit triggers first
    var triggers = ScriptApp.getUserTriggers(ss);
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'onEdit') {
        ScriptApp.deleteTrigger(triggers[i]);
        Logger.log('Deleted existing onEdit trigger');
      }
    }
    
    // Create new INSTALLABLE onEdit trigger
    // This ensures it always uses the latest code (not cached!)
    ScriptApp.newTrigger('onEdit')
      .forSpreadsheet(ss)
      .onEdit()
      .create();
    
    Logger.log('✓ Created INSTALLABLE onEdit trigger');
  } catch (error) {
    Logger.log('Error creating onEdit trigger: ' + error);
    throw error;
  }
}

function quickFixTriggers() {
  try {
    var ui = SpreadsheetApp.getUi();
    
    var result = ui.alert(
      'Quick Fix - Auto-Sync Triggers',
      'This will:\n\n' +
      '1. Delete orphaned triggers (onEditRankings_, onEditBMIDTrigger_)\n' +
      '2. Create onChange trigger for row inserts\n' +
      '3. Create onEdit trigger for cell edits\n\n' +
      'Takes 5 seconds. Continue?',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (result !== ui.Button.OK) {
      return;
    }
    
    // Delete orphaned triggers
    deleteOrphanedTriggers();
    Utilities.sleep(1000);
    
    // Create onChange trigger
    createOnChangeTrigger();
    Utilities.sleep(500);
    
    // Create onEdit trigger (INSTALLABLE - always uses latest code!)
    createOnEditTrigger();
    
    ui.alert(
      '✅ Quick Fix Complete!',
      'Auto-sync is now enabled!\n\n' +
      '✅ Orphaned triggers deleted\n' +
      '✅ onChange trigger created (row inserts)\n' +
      '✅ onEdit trigger created (cell edits)\n\n' +
      'Try typing a Steam ID - it should auto-lookup now!',
      ui.ButtonSet.OK
    );
    
    Logger.log('✓ Quick fix complete - all triggers set up');
  } catch (error) {
    Logger.log('Error in quick fix: ' + error);
    SpreadsheetApp.getUi().alert(
      'Error',
      'Quick fix failed: ' + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function deleteOrphanedTriggers() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var triggers = ScriptApp.getUserTriggers(ss);
    var deleted = [];
    
    // List of function names that should NOT have triggers
    var orphanedFunctions = [
      'onEditRankings_',
      'onEditBMIDTrigger_',
      'trigger_initialUpdateServerAdminTeam'  // Old trigger from previous versions
    ];
    
    for (var i = 0; i < triggers.length; i++) {
      var funcName = triggers[i].getHandlerFunction();
      if (orphanedFunctions.indexOf(funcName) !== -1) {
        ScriptApp.deleteTrigger(triggers[i]);
        deleted.push(funcName);
        Logger.log('Deleted orphaned trigger: ' + funcName);
      }
    }
    
    var ui = SpreadsheetApp.getUi();
    if (deleted.length > 0) {
      ui.alert(
        'Cleanup Complete',
        '✅ Deleted ' + deleted.length + ' orphaned trigger(s):\n\n' + deleted.join('\n') + '\n\nYour tracker should work properly now!',
        ui.ButtonSet.OK
      );
      Logger.log('✓ Cleaned up orphaned triggers: ' + deleted.join(', '));
    } else {
      ui.alert(
        'Already Clean',
        'No orphaned triggers found.\n\nYour triggers are clean!',
        ui.ButtonSet.OK
      );
    }
  } catch (error) {
    Logger.log('Error deleting orphaned triggers: ' + error);
    SpreadsheetApp.getUi().alert('Error', 'Failed to cleanup: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function verifyAllTriggers() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var allTriggers = ScriptApp.getUserTriggers(ss);
    var projectTriggers = ScriptApp.getProjectTriggers();
    
    // Expected triggers and their types
    var expectedTriggers = {
      'onChange': 'ON_CHANGE',
      'onEdit': 'ON_EDIT',
      'updateAllHoursHourly': 'CLOCK'
    };
    
    // Triggers that should NOT exist (orphaned/old)
    var badTriggers = [
      'onEditRankings_',
      'onEditBMIDTrigger_',
      'trigger_initialUpdateServerAdminTeam'
    ];
    
    var report = '🔍 TRIGGER VERIFICATION REPORT\n\n';
    report += '══════════════════════════════════\n\n';
    
    // Check expected triggers
    report += '✅ EXPECTED TRIGGERS:\n\n';
    
    var foundExpected = {};
    for (var i = 0; i < allTriggers.length; i++) {
      var funcName = allTriggers[i].getHandlerFunction();
      if (expectedTriggers.hasOwnProperty(funcName)) {
        foundExpected[funcName] = true;
        var triggerType = allTriggers[i].getTriggerSource();
        report += '  ✅ ' + funcName + '\n';
        report += '     Type: ' + triggerType + '\n';
        
        if (triggerType === ScriptApp.TriggerSource.CLOCK) {
          report += '     Frequency: Every hour\n';
        } else if (triggerType === ScriptApp.TriggerSource.SPREADSHEETS) {
          var eventType = allTriggers[i].getEventType();
          report += '     Event: ' + eventType + '\n';
        }
        report += '\n';
      }
    }
    
    // Check for missing expected triggers
    for (var expectedFunc in expectedTriggers) {
      if (!foundExpected[expectedFunc]) {
        report += '  ❌ MISSING: ' + expectedFunc + '\n';
        report += '     Action: Run "Quick Fix" or "Complete Initial Setup"\n\n';
      }
    }
    
    // Check for bad/orphaned triggers
    report += '══════════════════════════════════\n\n';
    report += '❌ ORPHANED/OLD TRIGGERS:\n\n';
    
    var foundBad = [];
    for (var i = 0; i < allTriggers.length; i++) {
      var funcName = allTriggers[i].getHandlerFunction();
      if (badTriggers.indexOf(funcName) !== -1) {
        foundBad.push(funcName);
        report += '  ❌ ' + funcName + '\n';
        report += '     Status: SHOULD BE DELETED\n';
        report += '     Issue: Function no longer exists\n\n';
      }
    }
    
    if (foundBad.length === 0) {
      report += '  ✅ None found - Good!\n\n';
    }
    
    // Check for unexpected triggers
    report += '══════════════════════════════════\n\n';
    report += 'ℹ️ OTHER TRIGGERS:\n\n';
    
    var foundOther = false;
    for (var i = 0; i < allTriggers.length; i++) {
      var funcName = allTriggers[i].getHandlerFunction();
      if (!expectedTriggers.hasOwnProperty(funcName) && badTriggers.indexOf(funcName) === -1) {
        foundOther = true;
        report += '  ⚠️ ' + funcName + '\n';
        var triggerType = allTriggers[i].getTriggerSource();
        report += '     Type: ' + triggerType + '\n';
        report += '     Note: Not in expected list\n\n';
      }
    }
    
    if (!foundOther) {
      report += '  ✅ None found - Clean!\n\n';
    }
    
    // Summary
    report += '══════════════════════════════════\n\n';
    report += '📊 SUMMARY:\n\n';
    report += 'Total triggers: ' + allTriggers.length + '\n';
    report += 'Expected found: ' + Object.keys(foundExpected).length + '/' + Object.keys(expectedTriggers).length + '\n';
    report += 'Orphaned found: ' + foundBad.length + '\n\n';
    
    // Recommendations
    if (foundBad.length > 0) {
      report += '⚠️ ACTION REQUIRED:\n';
      report += 'Run "Delete Orphaned Triggers" to clean up!\n\n';
    } else if (Object.keys(foundExpected).length < Object.keys(expectedTriggers).length) {
      report += '⚠️ ACTION REQUIRED:\n';
      report += 'Run "Quick Fix" to create missing triggers!\n\n';
    } else {
      report += '✅ ALL GOOD!\n';
      report += 'Your triggers are properly configured!\n\n';
    }
    
    var ui = SpreadsheetApp.getUi();
    ui.alert('Trigger Verification', report, ui.ButtonSet.OK);
    
    Logger.log('Trigger verification complete');
    Logger.log(report);
  } catch (error) {
    Logger.log('Error verifying triggers: ' + error);
    SpreadsheetApp.getUi().alert('Error', 'Failed to verify triggers: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function checkTriggerStatus() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var triggers = ScriptApp.getUserTriggers(ss);
    
    var found = {
      hourly: false,
      onChange: false,
      onEdit: false
    };
    
    var status = '⏰ TRIGGER STATUS\n\n';
    status += 'Total triggers: ' + triggers.length + '\n\n';
    
    for (var i = 0; i < triggers.length; i++) {
      var funcName = triggers[i].getHandlerFunction();
      var eventType = triggers[i].getEventType();
      
      status += '✓ ' + funcName + '\n';
      status += '  Event: ' + eventType + '\n\n';
      
      if (funcName === 'updateAllHoursHourly') found.hourly = true;
      if (funcName === 'onChange') found.onChange = true;
      if (funcName === 'onEdit') found.onEdit = true;
    }
    
    status += '═══════════════════════\n\n';
    status += 'EXPECTED TRIGGERS:\n\n';
    status += (found.hourly ? '✅' : '❌') + ' updateAllHoursHourly (hourly updates)\n';
    status += (found.onChange ? '✅' : '❌') + ' onChange (row inserts/deletes)\n';
    status += (found.onEdit ? '✅' : '❌') + ' onEdit (cell edits)\n\n';
    
    if (!found.hourly || !found.onChange || !found.onEdit) {
      status += '⚠️ MISSING TRIGGERS!\n\n';
      status += 'Run "Quick Fix - Auto-Sync Triggers" to fix!';
      
      var ui = SpreadsheetApp.getUi();
      var result = ui.alert(
        'Trigger Status',
        status + '\n\nWould you like to run Quick Fix now?',
        ui.ButtonSet.YES_NO
      );
      
      if (result === ui.Button.YES) {
        quickFixTriggers();
      }
    } else {
      status += '✅ ALL TRIGGERS ACTIVE!';
      SpreadsheetApp.getUi().alert('Trigger Status', status, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    
    Logger.log(status);
  } catch (error) {
    Logger.log('Error checking trigger status: ' + error);
  }
}

// ========== CONFIGURATION MANAGEMENT ==========

function showConfiguration() {
  try {
    var ui = SpreadsheetApp.getUi();
    var props = PropertiesService.getScriptProperties();
    
    // Get current active values
    var serverId = getConfig('BM_SERVER_ID');
    var sheetName = getConfig('SOURCE_SHEET_NAME');
    var nameCol = getConfig('SOURCE_COL_NAME');
    var teamCol = getConfig('SOURCE_COL_TEAM');
    var steamCol = getConfig('SOURCE_COL_STEAM_ID');
    
    // Check where values are coming from
    var serverIdSource = props.getProperty('BM_SERVER_ID') ? 'Wizard' : 'Config.gs';
    var nameColSource = props.getProperty('SOURCE_COL_NAME') ? 'Wizard' : 'Config.gs';
    
    ui.alert(
      'Current Configuration',
      '📊 Active Configuration:\n\n' +
      '• Server ID: ' + serverId + ' (' + serverIdSource + ')\n' +
      '• Source Sheet: ' + sheetName + '\n' +
      '• Name Column: ' + columnLetter(nameCol) + ' (Column ' + nameCol + ') (' + nameColSource + ')\n' +
      '• Team Column: ' + columnLetter(teamCol) + ' (Column ' + teamCol + ')\n' +
      '• Steam ID Column: ' + columnLetter(steamCol) + ' (Column ' + steamCol + ')\n\n' +
      '💡 To change: Run Setup Wizard or edit Config.gs',
      ui.ButtonSet.OK
    );
  } catch (error) {
    Logger.log('Error showing configuration: ' + error);
  }
}

function cleanupOldWizardConfig() {
  var ui = SpreadsheetApp.getUi();
  
  var result = ui.alert(
    'Clean Up Old Wizard Config',
    'This will remove OLD wizard settings (CONFIG_ prefix)\n' +
    'from Script Properties.\n\n' +
    'Properties to remove:\n' +
    '• CONFIG_BM_SERVER_ID (old)\n' +
    '• CONFIG_SOURCE_SHEET_NAME (old)\n' +
    '• CONFIG_SOURCE_COL_NAME (old)\n' +
    '• CONFIG_SOURCE_COL_TEAM (old)\n' +
    '• CONFIG_SOURCE_COL_STEAM_ID (old)\n' +
    '• CONFIG_SOURCE_FIRST_DATA_ROW (old)\n\n' +
    '✅ Your NEW wizard settings will be kept!\n' +
    '(BM_SERVER_ID, SOURCE_COL_NAME, etc.)\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) {
    return;
  }
  
  try {
    var props = PropertiesService.getScriptProperties();
    var removed = [];
    
    // ONLY remove OLD CONFIG_ prefixed properties
    // DO NOT remove new properties (without prefix)
    var oldProps = [
      'CONFIG_BM_SERVER_ID',
      'CONFIG_SOURCE_SHEET_NAME',
      'CONFIG_SOURCE_COL_NAME',
      'CONFIG_SOURCE_COL_TEAM',
      'CONFIG_SOURCE_COL_STEAM_ID',
      'CONFIG_SOURCE_FIRST_DATA_ROW'
    ];
    
    for (var i = 0; i < oldProps.length; i++) {
      if (props.getProperty(oldProps[i])) {
        props.deleteProperty(oldProps[i]);
        removed.push(oldProps[i]);
      }
    }
    
    if (removed.length > 0) {
      ui.alert(
        'Cleanup Complete',
        '✅ Removed ' + removed.length + ' old properties:\n\n' +
        removed.join('\n') + '\n\n' +
        '✅ Your NEW wizard settings are still active!\n' +
        'All settings now use the new naming (no CONFIG_ prefix).',
        ui.ButtonSet.OK
      );
      Logger.log('✓ Cleaned up old CONFIG_ properties: ' + removed.join(', '));
    } else {
      ui.alert(
        'Nothing to Clean',
        'No old CONFIG_ properties found.\n\n' +
        '✅ Your configuration is already clean!',
        ui.ButtonSet.OK
      );
    }
  } catch (error) {
    Logger.log('Error cleaning up properties: ' + error);
    ui.alert('Error', 'Failed to clean up: ' + error.toString(), ui.ButtonSet.OK);
  }
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
    
    if (diffMins < 1) {
      return 'Just now';  // Better than "0 min ago"
    } else if (diffMins < 60) {
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


// ========== CUSTOM MENU ==========

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu('⚙️ Admin Tracker (v' + VERSION + ')')
    
    // === SETUP ===
    .addSubMenu(ui.createMenu('🛠️ SETUP')
      .addItem('First Time Setup', 'firstTimeSetup')
      .addSeparator()
      .addItem('Migrate Old Data (v3.x → v4.0)', 'migrateOldData'))
    
    .addSeparator()
    
    // === UPDATE ===
    .addSubMenu(ui.createMenu('🔄 UPDATE')
      .addItem('Update All Data', 'updateAllData')
      .addSeparator()
      .addItem('Update Server ID', 'updateServerID')
      .addItem('Update API Key', 'updateApiKey')
      .addItem('Update Column Mappings', 'updateColumnMappings')
      .addSeparator()
      .addItem('Sync Admin Data', 'STEP1_SyncAdminData')
      .addItem('Populate BM IDs', 'STEP2_PopulateBMIds')
      .addItem('Update 7-Day Hours', 'updateHours7d')
      .addItem('Update 30-Day Hours', 'updateHours30d')
      .addItem('Update 90-Day Hours', 'updateHours90d')
      .addItem('Refresh Rankings', 'STEP4_CreateRankingsDisplay')
      .addSeparator()
      .addItem('Update Single Player', 'updateSinglePlayer'))
    
    .addSeparator()
    
    // === TRIGGERS ===
    .addSubMenu(ui.createMenu('⏰ TRIGGERS')
      .addItem('Manage Triggers', 'manageTriggers'))
    
    .addSeparator()
    
    // === SETTINGS ===
    .addSubMenu(ui.createMenu('⚙️ SETTINGS')
      .addItem('Configuration', 'showConfiguration')
      .addItem('Debug Tools', 'debugToolsMenu')
      .addItem('About', 'showAbout')
      .addItem('Check for Updates', 'checkForUpdates'))
    
    .addToUi();
}

function showUpdateInfo() {
  // Dummy function for menu items that are just informational
}

// ========== MENU FUNCTIONS ==========

/**
 * Migrate Old Data - For users upgrading from v3.x to v4.0
 * Migrates data from old columns (Y-AE) to new structure (AA-AJ)
 */
function migrateOldData() {
  var ui = SpreadsheetApp.getUi();
  
  try {
    var result = ui.alert(
      'Migrate to New Structure',
      'This will migrate your data from the old column structure (Y-AE)\n' +
      'to the new organized structure (AA-AJ).\n\n' +
      'This is a ONE-TIME migration for users upgrading from v3.x.\n\n' +
      'What will happen:\n' +
      '1. Copy data from old columns to new columns\n' +
      '2. Create AUTO column formulas\n' +
      '3. Verify data integrity\n\n' +
      'Your original data will NOT be deleted (backup).\n\n' +
      'Continue with migration?',
      ui.ButtonSet.YES_NO
    );
    
    if (result !== ui.Button.YES) {
      return;
    }
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var rankingsSheet = ss.getSheetByName(getConfig("RANKINGS_SHEET_NAME"));
    
    if (!rankingsSheet) {
      ui.alert('Error', 'Rankings sheet not found. Run First Time Setup instead.', ui.ButtonSet.OK);
      return;
    }
    
    Logger.log('=== MIGRATION START ===');
    
    // Check if old data exists
    var lastRow = rankingsSheet.getLastRow();
    if (lastRow < 2) {
      ui.alert('No Data', 'No data found to migrate.', ui.ButtonSet.OK);
      return;
    }
    
    // OLD structure: Y=Name(25), Z=SteamID(26), AA=Team(27), AB=BMID(28), AC=7d(29), AD=30d(30), AE=90d(31)
    // NEW structure: AA=SteamID(27), AB=Name(28), AC=Team(29), AD=BMID(30), AE=LastSeen(31), AF=FirstSeen(32), AG=EOSID(33), AH=7d(34), AI=30d(35), AJ=90d(36)
    
    var numRows = lastRow - 1;
    
    // Read OLD data
    Logger.log('Reading old data structure...');
    var oldNames = rankingsSheet.getRange(2, 25, numRows, 1).getValues();        // Y
    var oldSteamIds = rankingsSheet.getRange(2, 26, numRows, 1).getValues();    // Z
    var oldTeams = rankingsSheet.getRange(2, 27, numRows, 1).getValues();       // AA (was team)
    var oldBmIds = rankingsSheet.getRange(2, 28, numRows, 1).getValues();       // AB (was bmid)
    var old7d = rankingsSheet.getRange(2, 29, numRows, 1).getValues();          // AC (was 7d)
    var old30d = rankingsSheet.getRange(2, 30, numRows, 1).getValues();         // AD (was 30d)
    var old90d = rankingsSheet.getRange(2, 31, numRows, 1).getValues();         // AE (was 90d)
    
    // Count non-empty rows
    var dataCount = 0;
    for (var i = 0; i < oldSteamIds.length; i++) {
      if (oldSteamIds[i][0] && String(oldSteamIds[i][0]).trim()) {
        dataCount++;
      }
    }
    
    Logger.log('Found ' + dataCount + ' players in old structure');
    
    if (dataCount === 0) {
      ui.alert('No Data', 'No player data found in old columns.', ui.ButtonSet.OK);
      return;
    }
    
    // Write to NEW structure
    Logger.log('Writing to new data structure (AA-AJ)...');
    rankingsSheet.getRange(2, 27, numRows, 1).setValues(oldSteamIds);  // AA = Steam ID
    rankingsSheet.getRange(2, 28, numRows, 1).setValues(oldNames);     // AB = Name
    rankingsSheet.getRange(2, 29, numRows, 1).setValues(oldTeams);     // AC = Team
    rankingsSheet.getRange(2, 30, numRows, 1).setValues(oldBmIds);     // AD = BM ID
    // AE, AF, AG will be populated by next data update (Last Seen, First Seen, EOSID)
    rankingsSheet.getRange(2, 34, numRows, 1).setValues(old7d);        // AH = 7d hours
    rankingsSheet.getRange(2, 35, numRows, 1).setValues(old30d);       // AI = 30d hours
    rankingsSheet.getRange(2, 36, numRows, 1).setValues(old90d);       // AJ = 90d hours
    
    // Set headers
    Logger.log('Setting new column headers...');
    rankingsSheet.getRange(1, 27).setValue('Steam ID');
    rankingsSheet.getRange(1, 28).setValue('Name');
    rankingsSheet.getRange(1, 29).setValue('Team');
    rankingsSheet.getRange(1, 30).setValue('BM ID');
    rankingsSheet.getRange(1, 31).setValue('Last Seen');
    rankingsSheet.getRange(1, 32).setValue('First Seen');
    rankingsSheet.getRange(1, 33).setValue('EOSID');
    rankingsSheet.getRange(1, 34).setValue('7-Day Hours');
    rankingsSheet.getRange(1, 35).setValue('30-Day Hours');
    rankingsSheet.getRange(1, 36).setValue('90-Day Hours');
    
    Logger.log('✓ Data migrated to new structure');
    
    // Create AUTO column formulas
    Logger.log('Creating AUTO column formulas...');
    createAutoColumnFormulas();
    
    Logger.log('=== MIGRATION COMPLETE ===');
    
    ui.alert(
      '✅ Migration Complete!',
      'Successfully migrated ' + dataCount + ' players!\n\n' +
      '✅ Data copied to new structure (AA-AJ)\n' +
      '✅ Column headers set\n' +
      '✅ AUTO column formulas created\n\n' +
      '📝 Next Steps:\n' +
      '1. Verify data looks correct in Rankings sheet\n' +
      '2. Run UPDATE → Populate BM IDs to add missing fields\n' +
      '3. Old columns (Y-AE) are still there as backup\n\n' +
      'After verifying everything works, you can manually\n' +
      'delete the old columns Y-AE if desired.',
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log('ERROR in migrateOldData: ' + error);
    Logger.log('Stack: ' + error.stack);
    ui.alert(
      'Migration Error',
      'An error occurred during migration:\n\n' + error + '\n\n' +
      'Your data was NOT modified.\n' +
      'Please check the logs for details.',
      ui.ButtonSet.OK
    );
  }
}

/**
 * First Time Setup - Sequenced setup for new users
 * No timeout risk, handles any number of players
 */
function firstTimeSetup() {
  var ui = SpreadsheetApp.getUi();
  
  try {
    var result = ui.alert(
      'First Time Setup',
      'This will guide you through setting up the Admin Hours Tracker.\n\n' +
      'Steps:\n' +
      '1. Configure Server ID & API Key\n' +
      '2. Detect/create columns\n' +
      '3. Set up Rankings sheet\n' +
      '4. Create AUTO column formulas\n' +
      '5. Sync admin data\n' +
      '6. Populate BM IDs\n' +
      '7. Create triggers\n\n' +
      'Hours will update via daily trigger or can be updated manually.\n\n' +
      'Continue?',
      ui.ButtonSet.YES_NO
    );
    
    if (result !== ui.Button.YES) {
      return;
    }
    
    // Run the Setup Wizard (handles config)
    setupWizard();
    
    // After wizard completes, offer to populate data
    var populate = ui.alert(
      'Setup Complete!',
      'Configuration saved!\n\n' +
      'Would you like to populate data now?\n\n' +
      'This will:\n' +
      '• Sync admin data to Rankings sheet\n' +
      '• Populate BM IDs and basic info\n' +
      '• Create AUTO column formulas\n' +
      '• Create automation triggers\n\n' +
      'Hours will be populated by daily trigger or can be updated manually.\n\n' +
      'Populate data now?',
      ui.ButtonSet.YES_NO
    );
    
    if (populate === ui.Button.YES) {
      ui.alert(
        'Starting Data Population',
        'This may take a few minutes depending on the number of players.\n\n' +
        'You can continue working while this runs.',
        ui.ButtonSet.OK
      );
      
      // Sync admin data
      STEP1_SyncAdminData();
      
      // Populate BM IDs
      STEP2_PopulateBMIds();
      
      // Create formulas
      createAutoColumnFormulas();
      
      // Create triggers
      try {
        createDailyTrigger();
      } catch (e) {
        Logger.log('Daily trigger may already exist: ' + e);
      }
      
      try {
        createOnChangeTrigger();
      } catch (e) {
        Logger.log('onChange trigger may already exist: ' + e);
      }
      
      try {
        createOnEditTrigger();
      } catch (e) {
        Logger.log('onEdit trigger may already exist: ' + e);
      }
      
      ui.alert(
        '🎉 Setup Complete!',
        'Your tracker is ready!\n\n' +
        '✅ Data synced\n' +
        '✅ BM IDs populated\n' +
        '✅ AUTO column formulas created\n' +
        '✅ Triggers created\n\n' +
        'Hours will update automatically via daily trigger.\n' +
        'Or run UPDATE → Update All Data to populate hours now.',
        ui.ButtonSet.OK
      );
    }
    
  } catch (error) {
    Logger.log('ERROR in firstTimeSetup: ' + error);
    ui.alert(
      'Setup Error',
      'An error occurred:\n\n' + error + '\n\n' +
      'Check the logs for details.',
      ui.ButtonSet.OK
    );
  }
}

/**
 * Update All Data - Sequential update of all data
 */
function updateAllData() {
  var ui = SpreadsheetApp.getUi();
  
  var result = ui.alert(
    'Update All Data',
    'This will update ALL data sequentially:\n\n' +
    '1. Sync admin data\n' +
    '2. Populate BM IDs\n' +
    '3. Update 7-day hours\n' +
    '4. Update 30-day hours\n' +
    '5. Update 90-day hours\n' +
    '6. Refresh rankings display\n\n' +
    'This may take several minutes.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) {
    return;
  }
  
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('Starting full data update...', 'Update All', -1);
    
    STEP1_SyncAdminData();
    STEP2_PopulateBMIds();
    updateHours7d();
    updateHours30d();
    updateHours90d();
    STEP4_CreateRankingsDisplay();
    
    ui.alert(
      '✅ Update Complete!',
      'All data has been updated successfully!',
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log('ERROR in updateAllData: ' + error);
    ui.alert(
      'Update Error',
      'An error occurred during update:\n\n' + error,
      ui.ButtonSet.OK
    );
  }
}

/**
 * Update Server ID
 */
function updateServerID() {
  var ui = SpreadsheetApp.getUi();
  
  var currentId = getConfig("BM_SERVER_ID");
  var prompt = ui.prompt(
    'Update Server ID',
    'Current Server ID: ' + currentId + '\n\n' +
    'Enter new BattleMetrics Server ID:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (prompt.getSelectedButton() === ui.Button.OK) {
    var newId = prompt.getResponseText().trim();
    if (newId) {
      PropertiesService.getScriptProperties().setProperty('BM_SERVER_ID', newId);
      ui.alert('✅ Server ID Updated', 'Server ID has been updated to: ' + newId, ui.ButtonSet.OK);
    }
  }
}

/**
 * Update API Key
 */
function updateApiKey() {
  var ui = SpreadsheetApp.getUi();
  
  var prompt = ui.prompt(
    'Update API Key',
    'Enter new BattleMetrics API Key:\n\n' +
    '(Your API key is stored securely)',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (prompt.getSelectedButton() === ui.Button.OK) {
    var newKey = prompt.getResponseText().trim();
    if (newKey) {
      PropertiesService.getScriptProperties().setProperty('BATTLEMETRICS_API_KEY', newKey);
      ui.alert('✅ API Key Updated', 'API key has been updated successfully.', ui.ButtonSet.OK);
    }
  }
}

/**
 * Update Column Mappings
 */
function updateColumnMappings() {
  var ui = SpreadsheetApp.getUi();
  
  ui.alert(
    'Update Column Mappings',
    'To update column mappings, run:\n\n' +
    'SETUP → First Time Setup\n\n' +
    'The setup wizard will auto-detect your columns.',
    ui.ButtonSet.OK
  );
}

/**
 * Update Single Player
 */
function updateSinglePlayer() {
  var ui = SpreadsheetApp.getUi();
  
  var prompt = ui.prompt(
    'Update Single Player',
    'Enter Steam ID (64-bit format):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (prompt.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  var steamId = prompt.getResponseText().trim();
  if (!steamId) {
    ui.alert('Error', 'Please enter a valid Steam ID.', ui.ButtonSet.OK);
    return;
  }
  
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('Updating player data...', 'Update', -1);
    
    var token = getScriptProperty_(getConfig("BM_TOKEN_KEY"));
    var serverId = getConfig("BM_SERVER_ID");
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var rankingsSheet = ss.getSheetByName(getConfig("RANKINGS_SHEET_NAME"));
    
    if (!rankingsSheet) {
      throw new Error('Rankings sheet not found');
    }
    
    // Find the row with this Steam ID
    var steamIds = rankingsSheet.getRange(getConfig("DATA_FIRST_ROW"), getConfig("DATA_COL_STEAM_ID"), rankingsSheet.getLastRow(), 1).getValues();
    var rowNum = -1;
    
    for (var i = 0; i < steamIds.length; i++) {
      if (String(steamIds[i][0]).trim() === steamId) {
        rowNum = getConfig("DATA_FIRST_ROW") + i;
        break;
      }
    }
    
    if (rowNum === -1) {
      ui.alert('Not Found', 'Steam ID not found in Rankings sheet.', ui.ButtonSet.OK);
      return;
    }
    
    // Fetch all data
    var result = getBMIdFromSteamId_(steamId, token);
    
    if (result && result.bmId) {
      // Write basic info
      rankingsSheet.getRange(rowNum, getConfig("DATA_COL_BM_ID")).setValue(result.bmId);
      rankingsSheet.getRange(rowNum, getConfig("DATA_COL_NAME")).setValue(result.steamName);
      if (result.lastSeen) {
        rankingsSheet.getRange(rowNum, getConfig("DATA_COL_LAST_SEEN")).setValue(new Date(result.lastSeen));
      }
      if (result.firstSeen) {
        rankingsSheet.getRange(rowNum, getConfig("DATA_COL_FIRST_SEEN")).setValue(new Date(result.firstSeen));
      }
      if (result.eosId) {
        rankingsSheet.getRange(rowNum, getConfig("DATA_COL_EOSID")).setValue(result.eosId);
      }
      
      // Fetch hours
      var now = new Date();
      var start7d = new Date(now.getTime() - (7 * 24 * 60 * 60 * 1000));
      var start30d = new Date(now.getTime() - (30 * 24 * 60 * 60 * 1000));
      var start90d = new Date(now.getTime() - (90 * 24 * 60 * 60 * 1000));
      
      var hours7d = getPlayedTimeSeconds_(result.bmId, serverId, start7d, now, token);
      var hours30d = getPlayedTimeSeconds_(result.bmId, serverId, start30d, now, token);
      var hours90d = getPlayedTimeSeconds_(result.bmId, serverId, start90d, now, token);
      
      if (hours7d !== null) {
        rankingsSheet.getRange(rowNum, getConfig("DATA_COL_HOURS_7D")).setValue(Math.round((hours7d / 3600) * 100) / 100);
      }
      if (hours30d !== null) {
        rankingsSheet.getRange(rowNum, getConfig("DATA_COL_HOURS_30D")).setValue(Math.round((hours30d / 3600) * 100) / 100);
      }
      if (hours90d !== null) {
        rankingsSheet.getRange(rowNum, getConfig("DATA_COL_HOURS_90D")).setValue(Math.round((hours90d / 3600) * 100) / 100);
      }
      
      ui.alert(
        '✅ Player Updated!',
        'Updated data for: ' + result.steamName + '\n\n' +
        'All fields have been refreshed.',
        ui.ButtonSet.OK
      );
      
    } else {
      ui.alert('Error', 'Could not fetch player data from BattleMetrics.', ui.ButtonSet.OK);
    }
    
  } catch (error) {
    Logger.log('ERROR in updateSinglePlayer: ' + error);
    ui.alert('Error', 'An error occurred:\n\n' + error, ui.ButtonSet.OK);
  }
}

/**
 * Debug Tools Menu - Diagnostic and troubleshooting tools
 */
function debugToolsMenu() {
  var ui = SpreadsheetApp.getUi();
  
  var choice = ui.alert(
    '🔧 Debug Tools',
    'Select a diagnostic tool:\n\n' +
    '1. Test API Connection\n' +
    '2. Verify Column Mappings\n' +
    '3. Test Player Lookup\n' +
    '4. Validate Data Integrity\n' +
    '5. Repair AUTO Column Formulas\n' +
    '6. Check API Quota\n' +
    '7. View Recent Logs\n\n' +
    'Enter number (1-7):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (choice !== ui.Button.OK) {
    return;
  }
  
  var toolChoice = ui.prompt(
    'Debug Tools',
    'Enter tool number (1-7):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (toolChoice.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  var tool = toolChoice.getResponseText().trim();
  
  switch(tool) {
    case '1':
      testAPIConnection_();
      break;
    case '2':
      verifyColumnMappings_();
      break;
    case '3':
      testPlayerLookup_();
      break;
    case '4':
      validateDataIntegrity_();
      break;
    case '5':
      createAutoColumnFormulas();
      break;
    case '6':
      checkQuotaStatus();
      break;
    case '7':
      viewRecentLogs_();
      break;
    default:
      ui.alert('Invalid Choice', 'Please enter a number from 1-7.', ui.ButtonSet.OK);
  }
}

function testAPIConnection_() {
  var ui = SpreadsheetApp.getUi();
  
  try {
    var token = getScriptProperty_(getConfig("BM_TOKEN_KEY"));
    var serverId = getConfig("BM_SERVER_ID");
    
    if (!token) {
      ui.alert('Error', 'API Token not configured. Run First Time Setup.', ui.ButtonSet.OK);
      return;
    }
    
    if (!serverId || serverId === 'YOUR_SERVER_ID_HERE') {
      ui.alert('Error', 'Server ID not configured. Run First Time Setup.', ui.ButtonSet.OK);
      return;
    }
    
    ui.alert('Testing...', 'Testing connection to BattleMetrics API...', ui.ButtonSet.OK);
    
    var endpoint = '/servers/' + serverId;
    var data = callBMApi_(endpoint, token);
    
    if (data && data.data) {
      var serverName = data.data.attributes && data.data.attributes.name;
      ui.alert(
        '✅ Connection Successful!',
        'Connected to BattleMetrics API!\n\n' +
        'Server: ' + (serverName || serverId) + '\n\n' +
        'API is working correctly.',
        ui.ButtonSet.OK
      );
    } else {
      ui.alert('Error', 'API returned unexpected data. Check logs.', ui.ButtonSet.OK);
    }
    
  } catch (error) {
    Logger.log('API test error: ' + error);
    ui.alert(
      'Connection Failed',
      'Could not connect to BattleMetrics API:\n\n' + error + '\n\n' +
      'Check:\n' +
      '1. API Token is correct\n' +
      '2. Server ID is correct\n' +
      '3. Internet connection',
      ui.ButtonSet.OK
    );
  }
}

function verifyColumnMappings_() {
  var ui = SpreadsheetApp.getUi();
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sourceSheet = ss.getSheetByName(getConfig("SOURCE_SHEET_NAME"));
    var rankingsSheet = ss.getSheetByName(getConfig("RANKINGS_SHEET_NAME"));
    
    var report = '📊 Column Mapping Report\n\n';
    report += '═══ SOURCE SHEET ═══\n';
    report += 'Name: Column ' + columnLetter(getConfig("SOURCE_COL_NAME")) + ' (' + getConfig("SOURCE_COL_NAME") + ')\n';
    report += 'Team: Column ' + columnLetter(getConfig("SOURCE_COL_TEAM")) + ' (' + getConfig("SOURCE_COL_TEAM") + ')\n';
    report += 'Steam ID: Column ' + columnLetter(getConfig("SOURCE_COL_STEAM_ID")) + ' (' + getConfig("SOURCE_COL_STEAM_ID") + ')\n';
    report += 'BMID: Column ' + columnLetter(getConfig("SOURCE_COL_BMID")) + ' (' + getConfig("SOURCE_COL_BMID") + ')\n';
    report += 'Last Seen: Column ' + columnLetter(getConfig("SOURCE_COL_LAST_SEEN")) + ' (' + getConfig("SOURCE_COL_LAST_SEEN") + ')\n';
    report += 'First Seen: Column ' + columnLetter(getConfig("SOURCE_COL_FIRST_SEEN")) + ' (' + getConfig("SOURCE_COL_FIRST_SEEN") + ')\n';
    report += 'EOSID: Column ' + columnLetter(getConfig("SOURCE_COL_EOSID")) + ' (' + getConfig("SOURCE_COL_EOSID") + ')\n\n';
    
    report += '═══ RANKINGS SHEET (AA-AJ) ═══\n';
    report += 'Steam ID: Column ' + columnLetter(getConfig("DATA_COL_STEAM_ID")) + ' (' + getConfig("DATA_COL_STEAM_ID") + ')\n';
    report += 'Name: Column ' + columnLetter(getConfig("DATA_COL_NAME")) + ' (' + getConfig("DATA_COL_NAME") + ')\n';
    report += 'Team: Column ' + columnLetter(getConfig("DATA_COL_TEAM")) + ' (' + getConfig("DATA_COL_TEAM") + ')\n';
    report += 'BM ID: Column ' + columnLetter(getConfig("DATA_COL_BM_ID")) + ' (' + getConfig("DATA_COL_BM_ID") + ')\n';
    report += 'Last Seen: Column ' + columnLetter(getConfig("DATA_COL_LAST_SEEN")) + ' (' + getConfig("DATA_COL_LAST_SEEN") + ')\n';
    report += 'First Seen: Column ' + columnLetter(getConfig("DATA_COL_FIRST_SEEN")) + ' (' + getConfig("DATA_COL_FIRST_SEEN") + ')\n';
    report += 'EOSID: Column ' + columnLetter(getConfig("DATA_COL_EOSID")) + ' (' + getConfig("DATA_COL_EOSID") + ')\n';
    report += '7-Day Hours: Column ' + columnLetter(getConfig("DATA_COL_HOURS_7D")) + ' (' + getConfig("DATA_COL_HOURS_7D") + ')\n';
    report += '30-Day Hours: Column ' + columnLetter(getConfig("DATA_COL_HOURS_30D")) + ' (' + getConfig("DATA_COL_HOURS_30D") + ')\n';
    report += '90-Day Hours: Column ' + columnLetter(getConfig("DATA_COL_HOURS_90D")) + ' (' + getConfig("DATA_COL_HOURS_90D") + ')\n\n';
    
    report += '═══ SHEET STATUS ═══\n';
    report += 'Source Sheet: ' + (sourceSheet ? '✅ Found' : '❌ Missing') + '\n';
    report += 'Rankings Sheet: ' + (rankingsSheet ? '✅ Found' : '❌ Missing') + '\n';
    
    ui.alert('Column Mappings', report, ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log('Error verifying columns: ' + error);
    ui.alert('Error', 'An error occurred:\n\n' + error, ui.ButtonSet.OK);
  }
}

function testPlayerLookup_() {
  var ui = SpreadsheetApp.getUi();
  
  var prompt = ui.prompt(
    'Test Player Lookup',
    'Enter Steam ID to test:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (prompt.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  var steamId = prompt.getResponseText().trim();
  if (!steamId) {
    return;
  }
  
  try {
    ui.alert('Testing...', 'Looking up player data...', ui.ButtonSet.OK);
    
    var token = getScriptProperty_(getConfig("BM_TOKEN_KEY"));
    var result = getBMIdFromSteamId_(steamId, token);
    
    if (result && result.bmId) {
      var report = '✅ Player Found!\n\n';
      report += 'Steam ID: ' + steamId + '\n';
      report += 'BM ID: ' + result.bmId + '\n';
      report += 'Name: ' + result.steamName + '\n';
      if (result.lastSeen) report += 'Last Seen: ' + result.lastSeen + '\n';
      if (result.firstSeen) report += 'First Seen: ' + result.firstSeen + '\n';
      if (result.eosId) report += 'EOSID: ' + result.eosId + '\n';
      
      ui.alert('Player Lookup', report, ui.ButtonSet.OK);
    } else {
      ui.alert('Not Found', 'Player not found in BattleMetrics.', ui.ButtonSet.OK);
    }
    
  } catch (error) {
    Logger.log('Lookup error: ' + error);
    ui.alert('Error', 'Lookup failed:\n\n' + error, ui.ButtonSet.OK);
  }
}

function validateDataIntegrity_() {
  var ui = SpreadsheetApp.getUi();
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var rankingsSheet = ss.getSheetByName(getConfig("RANKINGS_SHEET_NAME"));
    
    if (!rankingsSheet) {
      ui.alert('Error', 'Rankings sheet not found.', ui.ButtonSet.OK);
      return;
    }
    
    var lastRow = rankingsSheet.getLastRow();
    if (lastRow < 2) {
      ui.alert('No Data', 'No data found in Rankings sheet.', ui.ButtonSet.OK);
      return;
    }
    
    var steamIds = rankingsSheet.getRange(2, getConfig("DATA_COL_STEAM_ID"), lastRow - 1, 1).getValues();
    var bmIds = rankingsSheet.getRange(2, getConfig("DATA_COL_BM_ID"), lastRow - 1, 1).getValues();
    
    var missingBmIds = 0;
    var totalPlayers = 0;
    
    for (var i = 0; i < steamIds.length; i++) {
      if (steamIds[i][0] && String(steamIds[i][0]).trim()) {
        totalPlayers++;
        if (!bmIds[i][0] || String(bmIds[i][0]).trim() === '') {
          missingBmIds++;
        }
      }
    }
    
    var formulasValid = validateAutoColumnFormulas();
    
    var report = '📊 Data Integrity Report\n\n';
    report += 'Total Players: ' + totalPlayers + '\n';
    report += 'Missing BM IDs: ' + missingBmIds + (missingBmIds > 0 ? ' ⚠️' : ' ✅') + '\n';
    report += 'AUTO Formulas: ' + (formulasValid ? '✅ Valid' : '❌ Invalid') + '\n\n';
    
    if (missingBmIds > 0) {
      report += '⚠️ Run UPDATE → Populate BM IDs to fix missing IDs.\n\n';
    }
    
    if (!formulasValid) {
      report += '⚠️ Run Debug Tools → Repair AUTO Column Formulas to fix.\n\n';
    }
    
    if (missingBmIds === 0 && formulasValid) {
      report += '✅ All data integrity checks passed!';
    }
    
    ui.alert('Data Integrity', report, ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log('Validation error: ' + error);
    ui.alert('Error', 'Validation failed:\n\n' + error, ui.ButtonSet.OK);
  }
}

function viewRecentLogs_() {
  var ui = SpreadsheetApp.getUi();
  ui.alert(
    'View Logs',
    'To view recent logs:\n\n' +
    '1. Click Extensions → Apps Script\n' +
    '2. Click "Executions" in the left sidebar\n' +
    '3. Click on any execution to see detailed logs\n\n' +
    'Logs show all script activity including:\n' +
    '• API calls\n' +
    '• Data updates\n' +
    '• Errors and warnings',
    ui.ButtonSet.OK
  );
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
    Logger.log('Quota - Hourly: ' + hourlyCount + '/' + getConfig("MAX_HOURLY_API_CALLS") + ', Daily: ' + dailyCount + '/' + getConfig("MAX_DAILY_API_CALLS"));
  } catch (error) {
    Logger.log('Error in incrementQuotaUsage_: ' + error);
  }
}

function canMakeApiCall_() {
  try {
    var dailyUsage = getQuotaUsage_('daily');
    var hourlyUsage = getQuotaUsage_('hourly');
    if (dailyUsage >= getConfig("MAX_DAILY_API_CALLS")) {
      Logger.log('⚠ Daily quota limit reached: ' + dailyUsage);
      return false;
    }
    if (hourlyUsage >= getConfig("MAX_HOURLY_API_CALLS")) {
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
                   '&filter[servers]=' + getConfig("BM_SERVER_ID") +
                   '&include=identifier';
    
    var data = callBMApi_(endpoint, token);
    
    if (data && data.data && data.data.length > 0) {
      if (data.data.length > 1) {
        Logger.log('  Multiple matches found for ' + steamId + ', using first result');
      }
      
      var player = data.data[0];
      var attrs = player.attributes || {};
      
      // Extract all available data
      var result = {
        bmId: player.id,
        steamName: attrs.name || '',
        lastSeen: null,
        firstSeen: null,
        eosId: ''
      };
      
      // Get last seen (most recent session end or current time if online)
      if (attrs.lastSeen) {
        result.lastSeen = attrs.lastSeen;
      }
      
      // Get first seen (account creation or first server visit)
      if (attrs.createdAt) {
        result.firstSeen = attrs.createdAt;
      }
      
      // Extract EOSID from identifiers if included
      if (data.included && Array.isArray(data.included)) {
        for (var i = 0; i < data.included.length; i++) {
          var included = data.included[i];
          if (included.type === 'identifier' && included.attributes) {
            var identifier = included.attributes.identifier || '';
            // EOSID format check (usually starts with specific pattern)
            if (identifier && (identifier.indexOf('eos:') === 0 || identifier.length === 32)) {
              result.eosId = identifier;
              break;
            }
          }
        }
      }
      
      return result;
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
    var sourceSheet = ss.getSheetByName(getConfig("SOURCE_SHEET_NAME"));
    var rankingsSheet = ss.getSheetByName(getConfig("RANKINGS_SHEET_NAME"));
    
    if (!sourceSheet) {
      throw new Error('Source sheet "' + getConfig("SOURCE_SHEET_NAME") + '" not found');
    }
    
    if (!rankingsSheet) {
      rankingsSheet = ss.insertSheet(getConfig("RANKINGS_SHEET_NAME"));
      ss.setActiveSheet(rankingsSheet);
      ss.moveActiveSheet(1);
      Logger.log('✓ Created new Rankings sheet');
    }
    
    // Ensure sheet has enough columns (need up to AJ = column 36)
    var currentMaxCol = rankingsSheet.getMaxColumns();
    var requiredCols = getConfig("DATA_COL_HOURS_90D") + 5; // AJ + buffer
    if (currentMaxCol < requiredCols) {
      rankingsSheet.insertColumnsAfter(currentMaxCol, requiredCols - currentMaxCol);
      Logger.log('✓ Expanded sheet from ' + currentMaxCol + ' to ' + requiredCols + ' columns');
    }
    
    // Get last row with data in the Steam ID column (most reliable)
    var steamIdCol = getConfig("SOURCE_COL_STEAM_ID");
    var lastRow = sourceSheet.getLastRow();
    
    // Find the actual last row with data by checking Steam ID column
    var allSteamIds = sourceSheet.getRange(getConfig("SOURCE_FIRST_DATA_ROW"), steamIdCol, lastRow - getConfig("SOURCE_FIRST_DATA_ROW") + 1, 1).getValues();
    var actualLastRow = getConfig("SOURCE_FIRST_DATA_ROW") - 1;
    
    for (var i = allSteamIds.length - 1; i >= 0; i--) {
      if (allSteamIds[i][0] && String(allSteamIds[i][0]).trim() !== '') {
        actualLastRow = getConfig("SOURCE_FIRST_DATA_ROW") + i;
        break;
      }
    }
    
    if (actualLastRow < getConfig("SOURCE_FIRST_DATA_ROW")) {
      Logger.log('No data in source sheet');
      return;
    }
    
    var numRows = actualLastRow - getConfig("SOURCE_FIRST_DATA_ROW") + 1;
    Logger.log('Found ' + numRows + ' admins with data (rows ' + getConfig("SOURCE_FIRST_DATA_ROW") + ' to ' + actualLastRow + ')');
    
    // Clear NEW data columns (AA-AJ = columns 27-36) to remove any old data
    Logger.log('🧹 Clearing data columns AA-AJ...');
    var maxRows = rankingsSheet.getMaxRows();
    if (maxRows > 1) {
      var startCol = getConfig("DATA_COL_STEAM_ID"); // Column AA (27)
      var numCols = 10; // AA through AJ
      rankingsSheet.getRange(2, startCol, maxRows - 1, numCols).clearContent();
      Logger.log('✓ Cleared columns AA-AJ (rows 2-' + maxRows + ')');
    }
    
    // Read data from source sheet using getDisplayValues() (handles formulas/checkboxes)
    var steamIds = sourceSheet.getRange(getConfig("SOURCE_FIRST_DATA_ROW"), getConfig("SOURCE_COL_STEAM_ID"), numRows, 1).getDisplayValues();
    var names = sourceSheet.getRange(getConfig("SOURCE_FIRST_DATA_ROW"), getConfig("SOURCE_COL_NAME"), numRows, 1).getDisplayValues();
    var teams = sourceSheet.getRange(getConfig("SOURCE_FIRST_DATA_ROW"), getConfig("SOURCE_COL_TEAM"), numRows, 1).getDisplayValues();
    var bmIds = sourceSheet.getRange(getConfig("SOURCE_FIRST_DATA_ROW"), getConfig("SOURCE_COL_BMID"), numRows, 1).getDisplayValues();
    
    Logger.log('=== Writing to Rankings Sheet (AA-AJ) ===');
    Logger.log('Column AA (Steam ID): ' + getConfig("DATA_COL_STEAM_ID"));
    Logger.log('Column AB (Name): ' + getConfig("DATA_COL_NAME"));
    Logger.log('Column AC (Team): ' + getConfig("DATA_COL_TEAM"));
    Logger.log('Column AD (BM ID): ' + getConfig("DATA_COL_BM_ID"));
    
    // Write to NEW positions (AA-AD)
    rankingsSheet.getRange(getConfig("DATA_FIRST_ROW"), getConfig("DATA_COL_STEAM_ID"), numRows, 1).setValues(steamIds);
    rankingsSheet.getRange(getConfig("DATA_FIRST_ROW"), getConfig("DATA_COL_NAME"), numRows, 1).setValues(names);
    rankingsSheet.getRange(getConfig("DATA_FIRST_ROW"), getConfig("DATA_COL_TEAM"), numRows, 1).setValues(teams);
    rankingsSheet.getRange(getConfig("DATA_FIRST_ROW"), getConfig("DATA_COL_BM_ID"), numRows, 1).setValues(bmIds);
    
    Logger.log('✓ Data written successfully');
    
    // Set headers for ALL columns (AA-AJ)
    rankingsSheet.getRange(1, getConfig("DATA_COL_STEAM_ID")).setValue('Steam ID');
    rankingsSheet.getRange(1, getConfig("DATA_COL_NAME")).setValue('Name');
    rankingsSheet.getRange(1, getConfig("DATA_COL_TEAM")).setValue('Team');
    rankingsSheet.getRange(1, getConfig("DATA_COL_BM_ID")).setValue('BM ID');
    rankingsSheet.getRange(1, getConfig("DATA_COL_LAST_SEEN")).setValue('Last Seen');
    rankingsSheet.getRange(1, getConfig("DATA_COL_FIRST_SEEN")).setValue('First Seen');
    rankingsSheet.getRange(1, getConfig("DATA_COL_EOSID")).setValue('EOSID');
    rankingsSheet.getRange(1, getConfig("DATA_COL_HOURS_7D")).setValue('7-Day Hours');
    rankingsSheet.getRange(1, getConfig("DATA_COL_HOURS_30D")).setValue('30-Day Hours');
    rankingsSheet.getRange(1, getConfig("DATA_COL_HOURS_90D")).setValue('90-Day Hours');
    
    // Make sheet visible (no longer hiding columns)
    // Data is now in visible columns AA-AJ for easy debugging
    Logger.log('✓ Rankings sheet data columns are VISIBLE (AA-AJ)');
    
    Logger.log('✅ Synced ' + numRows + ' admins to Rankings sheet');
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Synced ' + numRows + ' admins to Rankings sheet (AA-AJ)',
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
    var token = getScriptProperty_(getConfig("BM_TOKEN_KEY"));
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var rankingsSheet = ss.getSheetByName(getConfig("RANKINGS_SHEET_NAME"));
    
    if (!rankingsSheet) throw new Error('Rankings sheet not found');
    
    var lastRow = rankingsSheet.getLastRow();
    if (lastRow < getConfig("DATA_FIRST_ROW")) {
      Logger.log('No data rows found');
      return;
    }
    
    var numRows = lastRow - getConfig("DATA_FIRST_ROW") + 1;
    var steamIds = rankingsSheet.getRange(getConfig("DATA_FIRST_ROW"), getConfig("DATA_COL_STEAM_ID"), numRows, 1).getValues();
    
    // Output arrays for Rankings sheet (AA-AJ)
    var bmIdOutput = [];
    var nameOutput = [];
    var lastSeenOutput = [];
    var firstSeenOutput = [];
    var eosIdOutput = [];
    
    var stats = { processed: 0, skipped: 0, errors: 0 };
    
    Logger.log('Starting AUTO field population...');
    Logger.log('Will write to Rankings sheet columns AA-AJ');
    Logger.log('AUTO columns will update via formulas');
    
    for (var i = 0; i < steamIds.length; i++) {
      try {
        var steamId = String(steamIds[i][0] || '').trim();
        var rowNum = getConfig("DATA_FIRST_ROW") + i;
        
        if (!steamId) {
          bmIdOutput.push(['']);
          nameOutput.push(['']);
          lastSeenOutput.push(['']);
          firstSeenOutput.push(['']);
          eosIdOutput.push(['']);
          stats.skipped++;
          continue;
        }
        
        if (!canMakeApiCall_()) {
          Logger.log('⚠ Quota limit reached, stopping');
          break;
        }
        
        Logger.log('Row ' + rowNum + ' | Steam: ' + steamId);
        var result = getBMIdFromSteamId_(steamId, token);
        Utilities.sleep(250);
        
        if (result && result.bmId) {
          // Write ALL fields to Rankings sheet
          bmIdOutput.push([result.bmId]);
          nameOutput.push([result.steamName]);
          lastSeenOutput.push([result.lastSeen ? new Date(result.lastSeen) : '']);
          firstSeenOutput.push([result.firstSeen ? new Date(result.firstSeen) : '']);
          eosIdOutput.push([result.eosId || '']);
          
          stats.processed++;
          Logger.log('  ✓ BM ID: ' + result.bmId);
          Logger.log('  ✓ Steam Name: ' + result.steamName);
          if (result.lastSeen) Logger.log('  ✓ Last Seen: ' + result.lastSeen);
          if (result.firstSeen) Logger.log('  ✓ First Seen: ' + result.firstSeen);
          if (result.eosId) Logger.log('  ✓ EOSID: ' + result.eosId);
        } else {
          bmIdOutput.push(['Not Found']);
          nameOutput.push(['']);
          lastSeenOutput.push(['']);
          firstSeenOutput.push(['']);
          eosIdOutput.push(['']);
          stats.errors++;
          Logger.log('  ✗ Not found');
        }
      } catch (rowError) {
        Logger.log('Error processing row ' + rowNum + ': ' + rowError);
        bmIdOutput.push(['Error']);
        nameOutput.push(['']);
        lastSeenOutput.push(['']);
        firstSeenOutput.push(['']);
        eosIdOutput.push(['']);
        stats.errors++;
      }
    }
    
    // Write ALL fields to Rankings sheet (AA-AJ)
    if (bmIdOutput.length > 0) {
      rankingsSheet.getRange(getConfig("DATA_FIRST_ROW"), getConfig("DATA_COL_BM_ID"), bmIdOutput.length, 1).setValues(bmIdOutput);
      rankingsSheet.getRange(getConfig("DATA_FIRST_ROW"), getConfig("DATA_COL_NAME"), nameOutput.length, 1).setValues(nameOutput);
      rankingsSheet.getRange(getConfig("DATA_FIRST_ROW"), getConfig("DATA_COL_LAST_SEEN"), lastSeenOutput.length, 1).setValues(lastSeenOutput);
      rankingsSheet.getRange(getConfig("DATA_FIRST_ROW"), getConfig("DATA_COL_FIRST_SEEN"), firstSeenOutput.length, 1).setValues(firstSeenOutput);
      rankingsSheet.getRange(getConfig("DATA_FIRST_ROW"), getConfig("DATA_COL_EOSID"), eosIdOutput.length, 1).setValues(eosIdOutput);
      Logger.log('✓ Written ALL fields to Rankings sheet (AA-AJ)');
      Logger.log('✓ AUTO columns will update automatically via formulas');
    }
    
    Logger.log('Complete - Processed: ' + stats.processed + ', Skipped: ' + stats.skipped + ', Errors: ' + stats.errors);
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Populated ALL fields in Rankings sheet!\nProcessed: ' + stats.processed + ', Errors: ' + stats.errors + '\nAUTO columns update via formulas.',
      'Data Population Complete',
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
    var token = getScriptProperty_(getConfig("BM_TOKEN_KEY"));
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var rankingsSheet = ss.getSheetByName(getConfig("RANKINGS_SHEET_NAME"));
    
    if (!rankingsSheet) return;
    
    var lastRow = rankingsSheet.getLastRow();
    if (lastRow < getConfig("DATA_FIRST_ROW")) return;
    
    var numRows = lastRow - getConfig("DATA_FIRST_ROW") + 1;
    
    var steamIds = rankingsSheet.getRange(getConfig("DATA_FIRST_ROW"), getConfig("DATA_COL_STEAM_ID"), numRows, 1).getValues();
    var bmIds = rankingsSheet.getRange(getConfig("DATA_FIRST_ROW"), getConfig("DATA_COL_BM_ID"), numRows, 1).getValues();
    
    var newlyAdded = 0;
    
    Logger.log('Checking for new players needing AUTO field population...');
    
    for (var i = 0; i < steamIds.length; i++) {
      try {
        var steamId = String(steamIds[i][0] || '').trim();
        var bmId = String(bmIds[i][0] || '').trim();
        var rowNum = getConfig("DATA_FIRST_ROW") + i;
        
        if (bmId || !steamId) {
          continue;
        }
        
        if (!canMakeApiCall_()) {
          Logger.log('⚠ Quota limit reached, stopping AUTO field auto-population');
          break;
        }
        
        Logger.log('Auto-populating ALL fields | Row ' + rowNum + ' | Steam: ' + steamId);
        
        var result = getBMIdFromSteamId_(steamId, token);
        Utilities.sleep(250);
        
        if (result && result.bmId) {
          // Write ALL fields to Rankings sheet (AA-AJ)
          rankingsSheet.getRange(rowNum, getConfig("DATA_COL_BM_ID")).setValue(result.bmId);
          rankingsSheet.getRange(rowNum, getConfig("DATA_COL_NAME")).setValue(result.steamName);
          if (result.lastSeen) {
            rankingsSheet.getRange(rowNum, getConfig("DATA_COL_LAST_SEEN")).setValue(new Date(result.lastSeen));
          }
          if (result.firstSeen) {
            rankingsSheet.getRange(rowNum, getConfig("DATA_COL_FIRST_SEEN")).setValue(new Date(result.firstSeen));
          }
          if (result.eosId) {
            rankingsSheet.getRange(rowNum, getConfig("DATA_COL_EOSID")).setValue(result.eosId);
          }
          
          // AUTO columns on Source sheet will update via formulas automatically
          
          newlyAdded++;
          Logger.log('  ✓ Added BM ID: ' + result.bmId);
          Logger.log('  ✓ Added Steam Name: ' + result.steamName);
          if (result.lastSeen) Logger.log('  ✓ Added Last Seen: ' + result.lastSeen);
          if (result.firstSeen) Logger.log('  ✓ Added First Seen: ' + result.firstSeen);
          if (result.eosId) Logger.log('  ✓ Added EOSID: ' + result.eosId);
          Logger.log('  ✓ AUTO columns will update via formulas');
        } else {
          rankingsSheet.getRange(rowNum, getConfig("DATA_COL_BM_ID")).setValue('Not Found');
          Logger.log('  ✗ BM ID not found');
        }
      } catch (rowError) {
        Logger.log('Error auto-populating row ' + rowNum + ': ' + rowError);
      }
    }
    
    if (newlyAdded > 0) {
      Logger.log('✅ Auto-populated ALL fields for ' + newlyAdded + ' new players');
    }
  } catch (error) {
    Logger.log('Error in autoPopulateMissingBMIds_: ' + error);
  }
}

// ========== STEP 3: UPDATE HOURS (WITH TIMESTAMPS) ==========

function updateHoursForRange_(label, daysBack, outCol, timestampKey) {
  try {
    var token = getScriptProperty_(getConfig("BM_TOKEN_KEY"));
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(getConfig("RANKINGS_SHEET_NAME"));
    
    if (!sheet) throw new Error('Rankings sheet not found');
    
    var lastRow = sheet.getLastRow();
    if (lastRow < getConfig("DATA_FIRST_ROW")) {
      Logger.log('No data rows found');
      return;
    }
    
    var numRows = lastRow - getConfig("DATA_FIRST_ROW") + 1;
    var stop = new Date();
    var start = new Date(stop.getTime());
    start.setDate(start.getDate() - daysBack);
    
    Logger.log('[' + label + '] Range: ' + start.toISOString() + ' to ' + stop.toISOString());
    
    var bmIds = sheet.getRange(getConfig("DATA_FIRST_ROW"), getConfig("DATA_COL_BM_ID"), numRows, 1).getValues();
    var output = [];
    var stats = { processed: 0, skipped: 0, errors: 0, quotaStop: false };
    
    for (var i = 0; i < bmIds.length; i++) {
      try {
        var bmId = String(bmIds[i][0] || '').trim();
        var rowNum = getConfig("DATA_FIRST_ROW") + i;
        
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
        var seconds = getPlayedTimeSeconds_(bmId, getConfig("BM_SERVER_ID"), start, stop, token);
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
      sheet.getRange(getConfig("DATA_FIRST_ROW"), outCol, output.length, 1).setValues(output);
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

// Internal update functions (no sync/display - for use by orchestrators)
function updateHours7d_internal_() {
  updateHoursForRange_('7d', 7, getConfig("DATA_COL_HOURS_7D"), '7D');
}

function updateHours30d_internal_() {
  updateHoursForRange_('30d', 30, getConfig("DATA_COL_HOURS_30D"), '30D');
}

function updateHours90d_internal_() {
  updateHoursForRange_('90d', 89, getConfig("DATA_COL_HOURS_90D"), '90D');
}

// Public update functions (with sync/display - for menu calls)
function updateHours7d() {
  STEP1_SyncAdminData();
  updateHours7d_internal_();
  STEP4_CreateRankingsDisplay();
  SpreadsheetApp.getActiveSpreadsheet().toast('7-day hours updated', 'Update Complete', 3);
}

function updateHours30d() {
  STEP1_SyncAdminData();
  updateHours30d_internal_();
  STEP4_CreateRankingsDisplay();
  SpreadsheetApp.getActiveSpreadsheet().toast('30-day hours updated', 'Update Complete', 3);
}

function updateHours90d() {
  STEP1_SyncAdminData();
  updateHours90d_internal_();
  STEP4_CreateRankingsDisplay();
  SpreadsheetApp.getActiveSpreadsheet().toast('90-day hours updated', 'Update Complete', 3);
}

// ========== STEP 4: CREATE RANKINGS DISPLAY ==========

function STEP4_CreateRankingsDisplay() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(getConfig("RANKINGS_SHEET_NAME"));
    
    if (!sheet) throw new Error('Rankings sheet not found');
    
    // DEBUG: Log sheet info
    Logger.log('=== STEP4 DEBUG START ===');
    Logger.log('Sheet name: ' + sheet.getName());
    Logger.log('Sheet max rows: ' + sheet.getMaxRows());
    Logger.log('Sheet max columns: ' + sheet.getMaxColumns());
    Logger.log('Sheet last row: ' + sheet.getLastRow());
    Logger.log('DATA_FIRST_ROW: ' + getConfig("DATA_FIRST_ROW"));
    Logger.log('DATA_COL_NAME: ' + getConfig("DATA_COL_NAME"));
    
    // Use simple getLastRow() - more reliable than complex smart detection
    var lastRow = sheet.getLastRow();
    
    if (lastRow < getConfig("DATA_FIRST_ROW")) {
      Logger.log('⚠️ No data to display - lastRow (' + lastRow + ') < DATA_FIRST_ROW (' + getConfig("DATA_FIRST_ROW") + ')');
      return;
    }
    
    var numRows = lastRow - getConfig("DATA_FIRST_ROW") + 1;
    Logger.log('Reading data from hidden columns...');
    Logger.log('DATA_COL_NAME: ' + getConfig("DATA_COL_NAME"));
    Logger.log('DATA_COL_TEAM: ' + getConfig("DATA_COL_TEAM"));
    Logger.log('Number of rows with data: ' + numRows);
    
    var names = sheet.getRange(getConfig("DATA_FIRST_ROW"), getConfig("DATA_COL_NAME"), numRows, 1).getValues();
    var teams = sheet.getRange(getConfig("DATA_FIRST_ROW"), getConfig("DATA_COL_TEAM"), numRows, 1).getValues();
    var hours7d = sheet.getRange(getConfig("DATA_FIRST_ROW"), getConfig("DATA_COL_HOURS_7D"), numRows, 1).getValues();
    var hours30d = sheet.getRange(getConfig("DATA_FIRST_ROW"), getConfig("DATA_COL_HOURS_30D"), numRows, 1).getValues();
    var hours90d = sheet.getRange(getConfig("DATA_FIRST_ROW"), getConfig("DATA_COL_HOURS_90D"), numRows, 1).getValues();
    
    Logger.log('First name: ' + names[0][0]);
    Logger.log('First team: ' + teams[0][0]);
    
    var allData = [];
    for (var i = 0; i < numRows; i++) {
      var name = String(names[i][0] || '').trim();
      var team = String(teams[i][0] || '').trim();
      var h7 = typeof hours7d[i][0] === 'number' ? hours7d[i][0] : 0;
      var h30 = typeof hours30d[i][0] === 'number' ? hours30d[i][0] : 0;
      var h90 = typeof hours90d[i][0] === 'number' ? hours90d[i][0] : 0;
      
      // Debug first 3 rows
      if (i < 3) {
        Logger.log('Row ' + (i+2) + ' data:');
        Logger.log('  names[' + i + '][0] = ' + names[i][0] + ' (type: ' + typeof names[i][0] + ')');
        Logger.log('  teams[' + i + '][0] = ' + teams[i][0] + ' (type: ' + typeof teams[i][0] + ')');
        Logger.log('  Processed name: "' + name + '"');
        Logger.log('  Processed team: "' + team + '"');
      }
      
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
    
    // Add update timestamp indicators
    addUpdateTimestamps_(sheet);
    
    formatRankingsSheet_(sheet);
    
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

function addUpdateTimestamps_(sheet) {
  // Columns T(20), U(21), V(22) for update timestamps
  var timestampStartCol = 20;
  
  // Headers
  sheet.getRange(1, timestampStartCol).setValue('📅 LAST UPDATES');
  sheet.getRange(1, timestampStartCol)
    .setFontSize(12)
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  
  // Merge header across 3 columns
  sheet.getRange(1, timestampStartCol, 1, 3).merge();
  
  // Sub-headers
  sheet.getRange(2, timestampStartCol).setValue('7-Day');
  sheet.getRange(2, timestampStartCol + 1).setValue('30-Day');
  sheet.getRange(2, timestampStartCol + 2).setValue('90-Day');
  
  // Format sub-headers
  sheet.getRange(2, timestampStartCol, 1, 3)
    .setFontWeight('bold')
    .setBackground('#e8f0fe')
    .setHorizontalAlignment('center')
    .setFontSize(10);
  
  // Get timestamps
  var last7d = getLastUpdateTimestamp_('7D');
  var last30d = getLastUpdateTimestamp_('30D');
  var last90d = getLastUpdateTimestamp_('90D');
  
  // Display timestamps
  sheet.getRange(3, timestampStartCol).setValue(last7d);
  sheet.getRange(3, timestampStartCol + 1).setValue(last30d);
  sheet.getRange(3, timestampStartCol + 2).setValue(last90d);
  
  // Format timestamp values
  sheet.getRange(3, timestampStartCol, 1, 3)
    .setHorizontalAlignment('center')
    .setFontSize(9)
    .setFontColor('#666666')
    .setBackground('#f8f9fa');
  
  // Add border around the whole timestamp section
  sheet.getRange(1, timestampStartCol, 3, 3)
    .setBorder(true, true, true, true, true, true, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
  
  // Set column widths
  sheet.setColumnWidth(timestampStartCol, 100);
  sheet.setColumnWidth(timestampStartCol + 1, 100);
  sheet.setColumnWidth(timestampStartCol + 2, 100);
  
  // Add next update info below
  sheet.getRange(5, timestampStartCol).setValue('Next Update:');
  sheet.getRange(5, timestampStartCol)
    .setFontWeight('bold')
    .setFontSize(9)
    .setFontColor('#666666');
  
  var nextUpdate = getNextUpdateInfo_();
  sheet.getRange(6, timestampStartCol).setValue(nextUpdate);
  sheet.getRange(6, timestampStartCol, 1, 3).merge();
  sheet.getRange(6, timestampStartCol)
    .setHorizontalAlignment('center')
    .setFontSize(9)
    .setFontColor('#1a73e8')
    .setFontWeight('bold');
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
      // Remove number prefix from team name (e.g., "2. Executive" -> "Executive")
      var cleanTeam = String(item.team).replace(/^\d+\.\s*/, '');
      
      return [
        index + 1,
        item.name,
        item[hoursField] === 0 ? 0 : item[hoursField].toFixed(2),
        cleanTeam
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
  // Find the maximum hours value for gradient calculation
  var maxHours = 0;
  for (var i = 0; i < sortedData.length; i++) {
    if (sortedData[i][hoursField] > maxHours) {
      maxHours = sortedData[i][hoursField];
    }
  }
  
  // Apply gradient colors to full rows
  for (var i = 0; i < sortedData.length; i++) {
    var hours = sortedData[i][hoursField];
    var color = getGradientColor_(hours, maxHours);
    
    // Apply color to entire row (all 4 columns)
    var rowRange = sheet.getRange(startRow + i, startCol, 1, 4);
    rowRange.setBackground(color);
  }
}

function getGradientColor_(value, maxValue) {
  // Calculate percentage (0 to 1)
  var percent = maxValue > 0 ? value / maxValue : 0;
  
  // Darker Red (#ea4335) -> Yellow (#ffff00) -> Green (#00ff00)
  var r, g, b;
  
  if (percent < 0.5) {
    // Dark Red to Yellow (0 to 0.5)
    var localPercent = percent * 2; // Scale 0-0.5 to 0-1
    r = 234; // Start with darker red (234 instead of 255)
    g = Math.round(67 + (255 - 67) * localPercent); // From 67 to 255
    b = Math.round(53 * (1 - localPercent)); // From 53 to 0
  } else {
    // Yellow to Green (0.5 to 1.0)
    var localPercent = (percent - 0.5) * 2; // Scale 0.5-1.0 to 0-1
    r = Math.round(255 * (1 - localPercent));
    g = 255;
    b = 0;
  }
  
  // Convert to hex
  var toHex = function(num) {
    var hex = Math.round(num).toString(16);
    return hex.length === 1 ? '0' + hex : hex;
  };
  
  return '#' + toHex(r) + toHex(g) + toHex(b);
}

function applyTeamTextColors_(sheet, sortedData, startRow, teamCol) {
  var teamColors = getConfig("TEAM_COLORS");
  
  for (var i = 0; i < sortedData.length; i++) {
    var team = String(sortedData[i].team);
    // Remove number prefix from team name for color lookup
    var cleanTeam = team.replace(/^\d+\.\s*/, '');
    var textColor = teamColors[cleanTeam] || teamColors[team] || '#000000';
    
    var cellRange = sheet.getRange(startRow + i, teamCol);
    cellRange.setFontColor(textColor);
    cellRange.setFontWeight('bold');
    cellRange.setFontSize(10);
  }
}

function getColorForHours_(hours) {
  for (var i = 0; i < getConfig("COLOR_RANGES").length; i++) {
    var range = getConfig("COLOR_RANGES")[i];
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

// ========== SEQUENTIAL DAILY UPDATE ==========

/**
 * Daily Sequential Update - Runs all updates sequentially
 * Called by daily trigger at 2 AM
 * Replaces rotating hourly updates
 */
function updateAllHoursDaily() {
  try {
    Logger.log('===== DAILY SEQUENTIAL UPDATE START =====');
    Logger.log('Time: ' + new Date());
    Logger.log('Quota - Daily: ' + getQuotaUsage_('daily') + '/' + getConfig("MAX_DAILY_API_CALLS"));
    Logger.log('Quota - Hourly: ' + getQuotaUsage_('hourly') + '/' + getConfig("MAX_HOURLY_API_CALLS"));
    
    // Step 1: Sync admin data
    Logger.log('Step 1: Syncing admin data...');
    STEP1_SyncAdminData();
    
    // Step 2: Auto-populate any missing BM IDs
    Logger.log('Step 2: Auto-populating missing BM IDs...');
    autoPopulateMissingBMIds_();
    
    // Step 3: Update 7-day hours
    Logger.log('Step 3: Updating 7-day hours...');
    updateHours7d_internal_();
    
    // Step 4: Update 30-day hours
    Logger.log('Step 4: Updating 30-day hours...');
    updateHours30d_internal_();
    
    // Step 5: Update 90-day hours
    Logger.log('Step 5: Updating 90-day hours...');
    updateHours90d_internal_();
    
    // Step 6: Refresh rankings display
    Logger.log('Step 6: Refreshing rankings display...');
    STEP4_CreateRankingsDisplay();
    
    Logger.log('===== DAILY SEQUENTIAL UPDATE COMPLETE =====');
    Logger.log('Final Quota - Daily: ' + getQuotaUsage_('daily') + '/' + getConfig("MAX_DAILY_API_CALLS"));
    Logger.log('Final Quota - Hourly: ' + getQuotaUsage_('hourly') + '/' + getConfig("MAX_HOURLY_API_CALLS"));
    
  } catch (error) {
    Logger.log('FATAL ERROR in updateAllHoursDaily: ' + error);
    Logger.log('Stack: ' + error.stack);
  }
}

// ========== ROTATING UPDATE FUNCTION (LEGACY) ==========

function updateAllHoursHourly() {
  try {
    var currentHour = new Date().getHours();
    var cycle = currentHour % 3;
    
    Logger.log('=== Starting Rotating Update (Hour: ' + currentHour + ', Cycle: ' + cycle + ') ===');
    Logger.log('Quota - Daily: ' + getQuotaUsage_('daily') + '/' + getConfig("MAX_DAILY_API_CALLS"));
    Logger.log('Quota - Hourly: ' + getQuotaUsage_('hourly') + '/' + getConfig("MAX_HOURLY_API_CALLS"));
    
    // Sync data ONCE at the start
    STEP1_SyncAdminData();
    autoPopulateMissingBMIds_();
    
    // Update only ONE metric (use internal function to avoid redundant sync/display)
    if (cycle === 0) {
      Logger.log('📊 Updating 7-day hours...');
      updateHours7d_internal_();
    } else if (cycle === 1) {
      Logger.log('📊 Updating 30-day hours...');
      updateHours30d_internal_();
    } else {
      Logger.log('📊 Updating 90-day hours...');
      updateHours90d_internal_();
    }
    
    // Refresh display ONCE at the end
    STEP4_CreateRankingsDisplay();
    
    Logger.log('=== Rotating Update Complete ===');
  } catch (error) {
    Logger.log('FATAL ERROR in updateAllHoursHourly: ' + error);
  }
}

function manualFullUpdate() {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Starting manual full update... This will take 8-10 minutes.',
      'Manual Update',
      5
    );
    
    Logger.log('=== Manual Full Update Start ===');
    
    // Sync data ONCE at the start
    STEP1_SyncAdminData();
    autoPopulateMissingBMIds_();
    
    // Update all three metrics using internal functions (no redundant sync/display)
    Logger.log('📊 Updating 7-day hours...');
    updateHours7d_internal_();
    
    if (canMakeApiCall_()) {
      Utilities.sleep(2000);
      Logger.log('📊 Updating 30-day hours...');
      updateHours30d_internal_();
    }
    
    if (canMakeApiCall_()) {
      Utilities.sleep(2000);
      Logger.log('📊 Updating 90-day hours...');
      updateHours90d_internal_();
    }
    
    // Refresh display ONCE at the end
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
    Logger.log('=== onEdit FIRED ===');
    
    if (!e || !e.range) {
      Logger.log('❌ No event or range - exiting');
      return;
    }
    
    var sheet = e.range.getSheet();
    var sheetName = sheet.getName();
    var row = e.range.getRow();
    var col = e.range.getColumn();
    
    Logger.log('📍 Edit detected:');
    Logger.log('  Sheet: ' + sheetName);
    Logger.log('  Row: ' + row);
    Logger.log('  Column: ' + col);
    
    // Check sheet name
    var sourceSheetName = getConfig("SOURCE_SHEET_NAME");
    Logger.log('  Expected sheet: ' + sourceSheetName);
    if (sheetName !== sourceSheetName) {
      Logger.log('❌ Wrong sheet - exiting');
      return;
    }
    Logger.log('✓ Sheet name matches');
    
    // Check row
    var firstDataRow = getConfig("SOURCE_FIRST_DATA_ROW");
    Logger.log('  First data row: ' + firstDataRow);
    if (row < firstDataRow) {
      Logger.log('❌ Header row - exiting');
      return;
    }
    Logger.log('✓ Row check passed');
    
    // Check column
    var nameCol = getConfig("SOURCE_COL_NAME");
    var teamCol = getConfig("SOURCE_COL_TEAM");
    var steamCol = getConfig("SOURCE_COL_STEAM_ID");
    
    Logger.log('  Checking columns:');
    Logger.log('    Edited column: ' + col);
    Logger.log('    Name column: ' + nameCol);
    Logger.log('    Team column: ' + teamCol);
    Logger.log('    Steam ID column: ' + steamCol);
    
    if (col === nameCol) {
      Logger.log('✓ Name column edited - processing');
      handleSourceDataChange_(sheet, row, col);
    } else if (col === teamCol) {
      Logger.log('✓ Team column edited - processing');
      handleSourceDataChange_(sheet, row, col);
    } else if (col === steamCol) {
      Logger.log('✓ Steam ID column edited - processing');
      handleSourceDataChange_(sheet, row, col);
    } else {
      Logger.log('❌ Column ' + col + ' not in trigger list - exiting');
      Logger.log('   (Only columns ' + nameCol + ', ' + teamCol + ', ' + steamCol + ' trigger updates)');
    }
  } catch (error) {
    Logger.log('❌ ERROR in onEdit: ' + error);
    Logger.log('Stack: ' + error.stack);
  }
}

// Detects when rows are added/deleted (fires on row insertion, deletion, etc.)
function onChange(e) {
  try {
    if (!e) return;
    
    var changeType = e.changeType;
    Logger.log('onChange triggered: ' + changeType);
    
    // ONLY handle INSERT_ROW or REMOVE_ROW (structural changes)
    // DO NOT handle EDIT - that's what onEdit is for!
    // onEdit is much more efficient because it knows exactly which cell was edited
    if (changeType === 'INSERT_ROW' || changeType === 'REMOVE_ROW' || changeType === 'INSERT_GRID') {
      Logger.log('Detected row change, syncing admin data...');
      STEP1_SyncAdminData();
      
      // Auto-populate any missing BM IDs
      Utilities.sleep(1000);
      autoPopulateMissingBMIds_();
      
      // Refresh display
      Utilities.sleep(1000);
      STEP4_CreateRankingsDisplay();
      
      Logger.log('✓ Auto-sync complete');
    } else {
      Logger.log('Change type "' + changeType + '" - letting onEdit handle this');
    }
  } catch (error) {
    Logger.log('Error in onChange: ' + error);
  }
}


function handleSourceDataChange_(sheet, row, col) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var rankingsSheet = ss.getSheetByName(getConfig("RANKINGS_SHEET_NAME"));
    
    if (!rankingsSheet) {
      Logger.log('Rankings sheet not found');
      return;
    }
    
    var rankingsRow = row - getConfig("SOURCE_FIRST_DATA_ROW") + getConfig("DATA_FIRST_ROW");
    
    var name = String(sheet.getRange(row, getConfig("SOURCE_COL_NAME")).getValue() || '').trim();
    var team = String(sheet.getRange(row, getConfig("SOURCE_COL_TEAM")).getValue() || '').trim();
    var steamId = String(sheet.getRange(row, getConfig("SOURCE_COL_STEAM_ID")).getValue() || '').trim();
    
    rankingsSheet.getRange(rankingsRow, getConfig("DATA_COL_NAME")).setValue(name);
    rankingsSheet.getRange(rankingsRow, getConfig("DATA_COL_TEAM")).setValue(team);
    rankingsSheet.getRange(rankingsRow, getConfig("DATA_COL_STEAM_ID")).setValue(steamId);
    
    if (col === getConfig("SOURCE_COL_STEAM_ID")) {
      Logger.log('Row ' + row + ': Steam ID changed to ' + steamId);
      
      // Clear existing data
      rankingsSheet.getRange(rankingsRow, getConfig("DATA_COL_BM_ID")).setValue('');
      rankingsSheet.getRange(rankingsRow, getConfig("DATA_COL_LAST_SEEN")).setValue('');
      rankingsSheet.getRange(rankingsRow, getConfig("DATA_COL_FIRST_SEEN")).setValue('');
      rankingsSheet.getRange(rankingsRow, getConfig("DATA_COL_EOSID")).setValue('');
      rankingsSheet.getRange(rankingsRow, getConfig("DATA_COL_HOURS_7D")).setValue('');
      rankingsSheet.getRange(rankingsRow, getConfig("DATA_COL_HOURS_30D")).setValue('');
      rankingsSheet.getRange(rankingsRow, getConfig("DATA_COL_HOURS_90D")).setValue('');
      
      if (!steamId) {
        Logger.log('  Steam ID cleared');
        return;
      }
      
      try {
        var token = getScriptProperty_(getConfig("BM_TOKEN_KEY"));
        var serverId = getConfig("BM_SERVER_ID");
        
        // STEP 1: Fetch player info (BM ID, Name, EOSID, First/Last Seen)
        Logger.log('  Fetching player info...');
        var result = getBMIdFromSteamId_(steamId, token);
        
        if (result && result.bmId) {
          // Write basic info to Rankings sheet
          rankingsSheet.getRange(rankingsRow, getConfig("DATA_COL_BM_ID")).setValue(result.bmId);
          rankingsSheet.getRange(rankingsRow, getConfig("DATA_COL_NAME")).setValue(result.steamName);
          if (result.lastSeen) {
            rankingsSheet.getRange(rankingsRow, getConfig("DATA_COL_LAST_SEEN")).setValue(new Date(result.lastSeen));
          }
          if (result.firstSeen) {
            rankingsSheet.getRange(rankingsRow, getConfig("DATA_COL_FIRST_SEEN")).setValue(new Date(result.firstSeen));
          }
          if (result.eosId) {
            rankingsSheet.getRange(rankingsRow, getConfig("DATA_COL_EOSID")).setValue(result.eosId);
          }
          
          Logger.log('  ✓ BM ID: ' + result.bmId);
          Logger.log('  ✓ Steam Name: ' + result.steamName);
          if (result.lastSeen) Logger.log('  ✓ Last Seen: ' + result.lastSeen);
          if (result.firstSeen) Logger.log('  ✓ First Seen: ' + result.firstSeen);
          if (result.eosId) Logger.log('  ✓ EOSID: ' + result.eosId);
          
          // STEP 2: Fetch hours (7-day, 30-day, 90-day)
          Logger.log('  Fetching hours data...');
          
          // 7-day hours
          var now = new Date();
          var start7d = new Date(now.getTime() - (7 * 24 * 60 * 60 * 1000));
          var hours7d = getPlayedTimeSeconds_(result.bmId, serverId, start7d, now, token);
          if (hours7d !== null) {
            var hours7dRounded = Math.round((hours7d / 3600) * 100) / 100;
            rankingsSheet.getRange(rankingsRow, getConfig("DATA_COL_HOURS_7D")).setValue(hours7dRounded);
            Logger.log('  ✓ 7-Day Hours: ' + hours7dRounded);
          }
          
          // 30-day hours
          var start30d = new Date(now.getTime() - (30 * 24 * 60 * 60 * 1000));
          var hours30d = getPlayedTimeSeconds_(result.bmId, serverId, start30d, now, token);
          if (hours30d !== null) {
            var hours30dRounded = Math.round((hours30d / 3600) * 100) / 100;
            rankingsSheet.getRange(rankingsRow, getConfig("DATA_COL_HOURS_30D")).setValue(hours30dRounded);
            Logger.log('  ✓ 30-Day Hours: ' + hours30dRounded);
          }
          
          // 90-day hours
          var start90d = new Date(now.getTime() - (90 * 24 * 60 * 60 * 1000));
          var hours90d = getPlayedTimeSeconds_(result.bmId, serverId, start90d, now, token);
          if (hours90d !== null) {
            var hours90dRounded = Math.round((hours90d / 3600) * 100) / 100;
            rankingsSheet.getRange(rankingsRow, getConfig("DATA_COL_HOURS_90D")).setValue(hours90dRounded);
            Logger.log('  ✓ 90-Day Hours: ' + hours90dRounded);
          }
          
          Logger.log('  ✅ ALL data written to Rankings sheet (AA-AJ)');
          Logger.log('  ✅ AUTO columns will update via formulas automatically');
          
          SpreadsheetApp.getActiveSpreadsheet().toast(
            'Fetched ALL data for ' + result.steamName,
            'Player Data Complete',
            3
          );
        } else {
          rankingsSheet.getRange(rankingsRow, getConfig("DATA_COL_BM_ID")).setValue('Not Found');
          Logger.log('  ✗ BM ID not found');
        }
      } catch (lookupError) {
        Logger.log('Error looking up player data: ' + lookupError);
        rankingsSheet.getRange(rankingsRow, getConfig("DATA_COL_BM_ID")).setValue('Error');
      }
    } else {
      Logger.log('Row ' + row + ': Data updated (no API lookup needed)');
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

// ========== QUICK SETUP (NO TIMEOUT RISK) ==========

function quickSetupData() {
  var ui = SpreadsheetApp.getUi();
  
  try {
    var startTime = new Date();
    Logger.log('===== QUICK SETUP START =====');
    
    // Show progress
    var progressMsg = 'Quick Setup Running...\n\n' +
      'This will take ~2 minutes.\n\n' +
      'Steps:\n' +
      '1. Clean up triggers\n' +
      '2. Sync admin data\n' +
      '3. Populate AUTO fields\n' +
      '4. Create triggers\n\n' +
      'Please wait...';
    
    SpreadsheetApp.getActiveSpreadsheet().toast(progressMsg, 'Quick Setup', -1);
    
    // Step 1: Clean up old triggers
    Logger.log('Step 1/4: Cleaning up old triggers...');
    deleteOrphanedTriggers();
    
    // Step 2: Sync admin data (creates Rankings sheet, populates hidden columns)
    Logger.log('Step 2/4: Syncing admin data...');
    STEP1_SyncAdminData();
    
    // Step 3: Populate BM IDs and ALL AUTO fields
    Logger.log('Step 3/4: Populating AUTO fields...');
    STEP2_PopulateBMIds();
    
    // Step 4: Create triggers
    Logger.log('Step 4/4: Creating triggers...');
    
    // Create hourly trigger
    try {
      createHourlyTrigger();
      Logger.log('✓ Hourly trigger created');
    } catch (e) {
      Logger.log('Note: Hourly trigger may already exist: ' + e);
    }
    
    // Create onChange trigger
    try {
      createOnChangeTrigger();
      Logger.log('✓ onChange trigger created');
    } catch (e) {
      Logger.log('Note: onChange trigger may already exist: ' + e);
    }
    
    // Create onEdit trigger
    try {
      createOnEditTrigger();
      Logger.log('✓ onEdit trigger created');
    } catch (e) {
      Logger.log('Note: onEdit trigger may already exist: ' + e);
    }
    
    var endTime = new Date();
    var duration = (endTime - startTime) / 1000;
    
    Logger.log('===== QUICK SETUP COMPLETE =====');
    Logger.log('Duration: ' + duration + ' seconds');
    
    // Success message
    ui.alert(
      '🎉 Quick Setup Complete!',
      'Your tracker is ready!\n\n' +
      '✅ Admin data synced\n' +
      '✅ BM IDs populated\n' +
      '✅ Steam Names populated\n' +
      '✅ EOSID populated\n' +
      '✅ First/Last Seen populated\n' +
      '✅ Triggers created\n\n' +
      '📊 Rankings Display:\n' +
      'Will be created automatically by the hourly trigger.\n' +
      'Or run "Refresh Display" from the menu.\n\n' +
      '⏰ Hours Data:\n' +
      'Will populate automatically over the next few hours.\n\n' +
      '⏱️ Setup completed in ' + Math.round(duration) + ' seconds!',
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log('ERROR in quickSetupData: ' + error);
    Logger.log('Stack: ' + error.stack);
    
    ui.alert(
      'Setup Error',
      'An error occurred during Quick Setup:\n\n' + error + '\n\n' +
      'Please check the logs (Extensions → Apps Script → Executions).',
      ui.ButtonSet.OK
    );
  }
}

// ========== COMPLETE INITIAL SETUP (FULL - MAY TIMEOUT) ==========

function SETUP_CompleteInitialSetup() {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Starting complete setup... This will take 6-7 minutes.',
      'Setup',
      5
    );
    
    Logger.log('===== COMPLETE SETUP START =====');
    
    Logger.log('Step 1/9: Cleaning up old triggers...');
    deleteOrphanedTriggers();
    Utilities.sleep(1000);
    
    Logger.log('Step 2/9: Syncing admin data...');
    STEP1_SyncAdminData();
    Utilities.sleep(2000);
    
    Logger.log('Step 3/9: Populating BM IDs...');
    STEP2_PopulateBMIds();
    Utilities.sleep(2000);
    
    Logger.log('Step 4/9: Updating 7-day hours...');
    updateHours7d_internal_();
    Utilities.sleep(2000);
    
    Logger.log('Step 5/9: Updating 30-day hours...');
    updateHours30d_internal_();
    Utilities.sleep(2000);
    
    Logger.log('Step 6/9: Updating 90-day hours...');
    updateHours90d_internal_();
    Utilities.sleep(2000);
    
    Logger.log('Step 7/9: Creating rankings display...');
    STEP4_CreateRankingsDisplay();
    Utilities.sleep(1000);
    
    Logger.log('Step 8/9: Setting up hourly auto-update trigger...');
    createHourlyTrigger();
    Utilities.sleep(1000);
    
    Logger.log('Step 9/9: Setting up auto-sync triggers...');
    createOnChangeTrigger();  // For INSERT_ROW/REMOVE_ROW
    createOnEditTrigger();     // For EDIT events (INSTALLABLE - always uses latest code!)
    
    Logger.log('===== COMPLETE SETUP FINISHED =====');
    
    var ui = SpreadsheetApp.getUi();
    ui.alert(
      '🎉 Setup Complete!',
      'Your Admin Hours Tracker is fully configured!\n\n' +
      '✅ Old triggers cleaned up\n' +
      '✅ Admin data synced\n' +
      '✅ BattleMetrics IDs populated\n' +
      '✅ 7-day hours updated\n' +
      '✅ 30-day hours updated\n' +
      '✅ 90-day hours updated\n' +
      '✅ Rankings display created\n' +
      '✅ Hourly auto-update trigger activated\n' +
      '✅ Auto-sync triggers activated (onChange + onEdit)\n\n' +
      '⏰ Updates run automatically every hour!\n' +
      '🔄 Edits to Steam IDs auto-lookup BM ID!\n\n' +
      'Your tracker is now live and tracking admin hours.',
      ui.ButtonSet.OK
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
    Logger.log('Daily: ' + daily + ' / ' + getConfig("MAX_DAILY_API_CALLS") + ' (' + ((daily/getConfig("MAX_DAILY_API_CALLS"))*100).toFixed(1) + '%)');
    Logger.log('Hourly: ' + hourly + ' / ' + getConfig("MAX_HOURLY_API_CALLS") + ' (' + ((hourly/getConfig("MAX_HOURLY_API_CALLS"))*100).toFixed(1) + '%)');
    
    var message = 'Daily: ' + daily + '/' + getConfig("MAX_DAILY_API_CALLS") + '\n' +
                  'Hourly: ' + hourly + '/' + getConfig("MAX_HOURLY_API_CALLS");
    
    if (daily >= getConfig("MAX_DAILY_API_CALLS")) {
      Logger.log('⚠ WARNING: Daily quota limit reached!');
      message += '\n\n⚠️ Daily quota limit reached!';
    } else if (hourly >= getConfig("MAX_HOURLY_API_CALLS")) {
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
    var token = getScriptProperty_(getConfig("BM_TOKEN_KEY"));
    var testBMId = '39394618';
    var serverId = getConfig("BM_SERVER_ID");
    
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
