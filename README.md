# Admin Hours Tracker for BattleMetrics

**Track admin activity hours on your BattleMetrics server with automated Google Sheets.**

Perfect for Squad, Rust, ARK, and any game using BattleMetrics for server management.

![Version](https://img.shields.io/badge/version-1.0.0-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)

---

## 🎯 What Does This Do?

This Google Apps Script automatically tracks how many hours your admins/moderators have logged on your BattleMetrics server. It creates beautiful, color-coded rankings that update every hour.

**Perfect for:**
- Server owners tracking admin activity
- Admin teams monitoring performance
- Staff applications (checking candidate activity)
- Monthly/weekly admin reports

---

## ✨ Features

- ✅ **Rotating Hourly Updates** - Updates one metric per hour (avoids Google's 6-minute timeout)
- ✅ **Three Time Periods** - 7-day, 30-day, and 90-day rankings
- ✅ **Dynamic Search** - Instant highlighting without page refresh
- ✅ **Auto Update Checker** - Get notified of new versions
- ✅ **Color-Coded Rankings** - Visual performance indicators (red to green)
- ✅ **Real-time Sync** - Updates automatically when you add/edit admins
- ✅ **Quota Management** - Smart API rate limiting
- ✅ **Last Update Timestamps** - Always know data freshness

---

## 📋 Prerequisites

Before starting, you need:

1. **Google Account** - For Google Sheets
2. **BattleMetrics Account** - With API access
3. **BattleMetrics API Token** - Get from [developers page](https://www.battlemetrics.com/developers)
4. **Server ID** - Find in your server URL

---

## 🚀 Quick Start (15 minutes)

### Step 1: Create Google Sheet

1. Go to [Google Sheets](https://sheets.google.com)
2. Create new spreadsheet
3. Name it: "Admin Hours Tracker" (or whatever you want)

### Step 2: Set Up Source Sheet

1. Rename the first sheet to: **Server Admin Team**
2. Add column headers in row 1:
   - Column A: (Optional - ID or notes)
   - **Column B:** Name
   - **Column C:** Team
   - Column D-E: (Whatever you want)
   - **Column F:** Steam ID

3. Add your admin data starting from row 2

**Example:**
```
| A  | B          | C           | D      | E    | F                 |
|----|------------|-------------|--------|------|-------------------|
| ID | Name       | Team        | Status | ...  | Steam ID          |
| 1  | John Doe   | Admin       | Active | ...  | 76561198012345678 |
| 2  | Jane Smith | Moderator   | Active | ...  | 76561198087654321 |
```

### Step 3: Add Apps Script

1. In your Google Sheet, go to: **Extensions → Apps Script**
2. Delete any default code
3. Download `AdminHoursTracker.js` from this repo
4. Copy ALL the code
5. Paste into Apps Script editor
6. Save (Ctrl+S or Cmd+S)

### Step 4: Configure the Script

Open the code and find the CONFIG section at the top. Change these values:

#### A. Update GitHub URLs (lines 13-14)
```javascript
const GITHUB_VERSION_URL = 'https://raw.githubusercontent.com/YOUR_USERNAME/admin-hours-tracker/main/version.json';
const GITHUB_DOWNLOAD_URL = 'https://github.com/YOUR_USERNAME/admin-hours-tracker';
```
Replace `YOUR_USERNAME` with your GitHub username (or skip if not using GitHub)

#### B. Set Your Server ID (line 24)
```javascript
BM_SERVER_ID: 'YOUR_SERVER_ID_HERE',
```

**How to find your Server ID:**
1. Go to your server on BattleMetrics
2. Look at the URL: `https://www.battlemetrics.com/servers/squad/12345678`
3. The number at the end (`12345678`) is your Server ID

Change to:
```javascript
BM_SERVER_ID: '12345678',
```

#### C. Verify Column Numbers (lines 35-38)
Make sure these match your sheet layout:
```javascript
SOURCE_COL_NAME: 2,      // Column B - Names
SOURCE_COL_TEAM: 3,      // Column C - Teams
SOURCE_COL_STEAM_ID: 6,  // Column F - Steam IDs
```

#### D. Customize Team Colors (lines 64-71)
Add your team names and colors:
```javascript
TEAM_COLORS: {
  'Owner': '#ad1457',
  'Admin': '#e65100',
  'Moderator': '#0277bd',
  'Trial Admin': '#00838f',
  // Add your teams here
},
```

**Save your changes!** (Ctrl+S)

### Step 5: Add BattleMetrics API Key

**IMPORTANT:** Never put your API key directly in the code!

1. In Apps Script editor, click **Project Settings** ⚙️ (gear icon)
2. Scroll down to **Script Properties**
3. Click **Add script property**
4. Enter:
   - **Property:** `BATTLEMETRICS_API_KEY`
   - **Value:** Your BattleMetrics API token
5. Click **Save script properties**

**Get your API token:**
1. Go to [BattleMetrics Developers](https://www.battlemetrics.com/developers)
2. Create new token
3. Give it read permissions
4. Copy the token

### Step 6: Run Initial Setup

1. **Close and reopen your Google Sheet**
2. You'll see a new menu: **⚙️ Admin Tracker (v1.0.0)**
3. Click: **⚙️ Admin Tracker → 🚀 Complete Initial Setup**
4. **Grant permissions** when Google asks:
   - Click "Review Permissions"
   - Choose your Google account
   - Click "Advanced"
   - Click "Go to [Your Project] (unsafe)"
   - Click "Allow"
5. **Wait 25-30 minutes** for initial data load

The script will:
- ✅ Sync your admin data
- ✅ Look up BattleMetrics IDs
- ✅ Fetch 7-day, 30-day, 90-day hours
- ✅ Create beautiful rankings display

### Step 7: Set Up Hourly Trigger

Make it update automatically every hour:

1. In Apps Script editor, click **Triggers** ⏰ (clock icon on left)
2. Click **+ Add Trigger** (bottom right)
3. Configure:
   - **Function:** `updateAllHoursHourly`
   - **Event source:** Time-driven
   - **Type:** Hour timer
   - **Hour interval:** Every hour
4. Click **Save**

**Done!** 🎉 Your tracker will now update automatically every hour!

---

## 📊 How It Works

### Rotating Update Schedule

To avoid Google's 6-minute execution timeout, the tracker updates ONE metric each hour:

| Hour | Updates |
|------|---------|
| 0, 3, 6, 9, 12, 15, 18, 21 | 7-Day Hours |
| 1, 4, 7, 10, 13, 16, 19, 22 | 30-Day Hours |
| 2, 5, 8, 11, 14, 17, 20, 23 | 90-Day Hours |

**Result:** All metrics stay fresh (max 3 hours old) without timing out!

### Sheet Structure

After setup, you'll have two sheets:

1. **Server Admin Team** - Your source data (you edit this)
2. **Admin Rankings** - Auto-generated rankings (don't edit manually)

---

## 🎨 Using the Tracker

### Menu Options

#### 📊 Update Status
- Shows when each metric was last updated
- Shows next scheduled update

#### 🔧 Manual Updates
- Update individual metrics on demand
- Useful after adding new admins

#### 🔄 Check for Updates
- Checks GitHub for newer versions
- One-click download

### Search Feature

1. Type in the **search box** (Column T)
2. Matches highlight instantly in **bright cyan**
3. Searches both names and teams
4. Click **CLEAR** to reset

### Adding New Admins

1. Add new row to **Server Admin Team** sheet
2. Fill in Name, Team, Steam ID
3. Rankings update automatically!

---

## 🔧 Customization

### Change Hour Color Ranges

Edit the `COLOR_RANGES` in the CONFIG section:

```javascript
COLOR_RANGES: [
  { min: 0, max: 10, color: '#fff9c4' },     // Very light yellow
  { min: 10, max: 25, color: '#fff59d' },    // Light yellow
  { min: 25, max: 50, color: '#ffee58' },    // Medium yellow
  // Adjust ranges to match your server's activity
],
```

### Change Column Layout

If your source sheet has data in different columns:

```javascript
SOURCE_COL_NAME: 2,      // Change to your name column
SOURCE_COL_TEAM: 3,      // Change to your team column
SOURCE_COL_STEAM_ID: 6,  // Change to your Steam ID column
```

---

## 🐛 Troubleshooting

### "Missing Script Property: BATTLEMETRICS_API_KEY"
→ Add the API key in **Project Settings → Script Properties**

### Timeout Errors
→ Make sure hourly trigger calls `updateAllHoursHourly` function

### Rankings Not Updating
→ Check **Executions** log in Apps Script for errors  
→ Verify API token has correct permissions

### Search Not Working
→ Close and reopen sheet to reload menu

### No Data Showing
→ Check that Steam IDs are 64-bit format (17 digits)  
→ Run **Test API Connection** from menu

---

## 🔄 Updating

### When New Version Available:

1. Menu shows notification
2. Click: **⚙️ Admin Tracker → 🔄 Check for Updates**
3. Popup shows new version details
4. Click **OK** to open GitHub
5. Download new `AdminHoursTracker.js`
6. In Apps Script: Select all (Ctrl+A) → Paste new code
7. Save (Ctrl+S)
8. Close and reopen sheet

**Your data and settings are preserved!**

---

## 📝 Configuration Reference

See `config.example.js` for detailed configuration guide.

**Required Settings:**
- `BM_SERVER_ID` - Your BattleMetrics server ID
- `BATTLEMETRICS_API_KEY` - In Script Properties
- `SOURCE_COL_*` - Column numbers for your data

**Optional Customization:**
- `TEAM_COLORS` - Team names and colors
- `COLOR_RANGES` - Hours-based color gradient
- `SOURCE_SHEET_NAME` - Source sheet name
- `RANKINGS_SHEET_NAME` - Rankings sheet name

---

## 🤝 Contributing

Found a bug? Have a feature request?

1. Open an issue on GitHub
2. Or submit a pull request

---

## 📜 License

MIT License - Feel free to modify and use for your server!

---

## ⭐ Support

If this helped your server, consider:
- ⭐ Starring this repo
- 🐛 Reporting bugs
- 💡 Suggesting features
- 📖 Improving documentation

---

## 📞 Getting Help

- **Issues:** [GitHub Issues](https://github.com/YOUR_USERNAME/admin-hours-tracker/issues)
- **Discussions:** [GitHub Discussions](https://github.com/YOUR_USERNAME/admin-hours-tracker/discussions)

---

**Made with ❤️ for server admins everywhere**

**Version:** 1.0.0

Google Apps Script for tracking admin activity hours on BattleMetrics Squad servers with rotating updates, dynamic search, and automatic notifications.

---

## Features

✅ **Rotating Hourly Updates** - Avoids timeout errors by updating one metric per hour  
✅ **Dynamic Search** - Instant highlighting without page refresh  
✅ **Last Update Timestamps** - Always know data freshness  
✅ **Auto Update Checker** - Get notified of new versions  
✅ **Quota Management** - Smart API rate limiting  
✅ **Real-time Sync** - Automatically updates when source data changes  
✅ **Color-Coded Rankings** - Visual performance indicators  

---

## Installation

### 1. Prerequisites

- Google account with access to Google Sheets
- BattleMetrics API token with read access
- BattleMetrics Server ID

### 2. Setup Steps

1. **Create Google Sheet**
   - Open [Google Sheets](https://sheets.google.com)
   - Create new spreadsheet
   - Name it something like "DMH Admin Tracker"

2. **Create Source Sheet**
   - Rename first sheet to: `Server Admin Team`
   - Add columns:
     - Column B: Name
     - Column C: Team
     - Column F: Steam ID
   - Add your admin data starting from row 2

3. **Add Apps Script**
   - Go to: Extensions → Apps Script
   - Delete any default code
   - Copy contents of `AdminHoursTracker.js`
   - Paste into Apps Script editor
   - Save (Ctrl+S)

4. **Update GitHub URLs** (Important!)
   - Find these lines at the top of the script:
   ```javascript
   const GITHUB_VERSION_URL = 'https://raw.githubusercontent.com/YOUR_USERNAME/dmh-admin-tracker/main/version.json';
   const GITHUB_DOWNLOAD_URL = 'https://github.com/YOUR_USERNAME/dmh-admin-tracker';
   ```
   - Replace `YOUR_USERNAME` with your GitHub username

5. **Add BattleMetrics API Key**
   - In Apps Script editor, click Project Settings (gear icon)
   - Scroll to "Script Properties"
   - Click "Add script property"
   - Property: `BATTLEMETRICS_API_KEY`
   - Value: Your BattleMetrics API token
   - Click "Save"

6. **Run Initial Setup**
   - Close and reopen your Google Sheet
   - You'll see new menu: ⚙️ Admin Tracker (v1.0.0)
   - Click: ⚙️ Admin Tracker → 🚀 Complete Initial Setup
   - Grant permissions when prompted
   - Wait 25-30 minutes for initial data load

7. **Set Up Hourly Trigger**
   - In Apps Script, click Triggers (clock icon)
   - Click "+ Add Trigger"
   - Settings:
     - Function: `updateAllHoursHourly`
     - Event source: Time-driven
     - Type: Hour timer
     - Hour interval: Every hour
   - Save

---

## Rotating Update Schedule

To avoid Google's 6-minute execution timeout, the tracker updates ONE metric each hour:

- **Hours 0, 3, 6, 9, 12, 15, 18, 21** → 7-Day Hours
- **Hours 1, 4, 7, 10, 13, 16, 19, 22** → 30-Day Hours
- **Hours 2, 5, 8, 11, 14, 17, 20, 23** → 90-Day Hours

**Result:** All metrics stay fresh (max 3 hours old) without timing out!

---

## Usage

### Menu Options

#### 📊 Update Status
- Shows last update time for each metric
- Shows next scheduled update

#### 🔧 Manual Updates
- Update individual metrics on demand
- Useful after adding new admins

#### 🔄 Check for Updates
- Checks GitHub for newer versions
- One-click to download page

### Search Feature

Type in the search box (Column T) to instantly highlight matching admins across all rankings.

**Tips:**
- Search by name or team
- Case-insensitive
- Highlights all matches in bright cyan
- Click "CLEAR" button to reset

---

## Configuration

Edit these at the top of the script if needed:

```javascript
const CONFIG = {
  SOURCE_SHEET_NAME: 'Server Admin Team',  // Your source data sheet name
  BM_SERVER_ID: '27157414',                // Your BattleMetrics server ID
  
  // Team colors (text color only)
  TEAM_COLORS: {
    'Owner': '#ad1457',
    'Executive': '#c62828',
    // ... add your teams here
  }
};
```

---

## Troubleshooting

### "Missing Script Property: BATTLEMETRICS_API_KEY"
→ Add the API key in Project Settings → Script Properties

### Timeout Errors
→ Make sure hourly trigger is calling `updateAllHoursHourly` (not the old function name)

### Search Not Working
→ Close and reopen sheet to reload the menu

### Rankings Not Updating
→ Check Executions log in Apps Script for errors
→ Verify BattleMetrics API token has correct permissions

### Update Check Fails
→ Verify GitHub URLs in script match your repo
→ Make sure `version.json` is in root of repo

---

## Updating to New Versions

### When Update Available:

1. Click: ⚙️ Admin Tracker → 🔄 Check for Updates
2. Popup will show new version details
3. Click OK to open GitHub
4. Download new `AdminHoursTracker.js`
5. In Apps Script, select all code (Ctrl+A)
6. Paste new version
7. Save (Ctrl+S)
8. Close and reopen sheet

**Note:** Your data and settings are preserved!

---

## Version Control on GitHub

### For Repository Owners

When releasing a new version:

1. Update `VERSION` constant in script
2. Update `version.json` with new version number
3. Update `CHANGELOG.md` with release notes
4. Commit and push to main branch
5. Create GitHub release with tag (e.g., `v1.0.1`)
6. Attach `AdminHoursTracker.js` to release

Users will automatically be notified next time they open their sheet!

---

## File Structure

```
dmh-admin-tracker/
├── AdminHoursTracker.js     # Main script
├── version.json             # Version info for update checker
├── README.md                # This file
├── CHANGELOG.md             # Version history
└── .gitignore              # Git ignore rules
```

---

## Support

- **Issues:** Open an issue on GitHub
- **Questions:** Check Discussions tab
- **Updates:** Watch repository for notifications

---

## License

MIT License - Feel free to modify and use for your server!

---

## Credits

Created for Dead Man's Hand Gaming Squad Server  
Developed with ❤️ by Ryan
