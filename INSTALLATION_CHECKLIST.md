# Installation Checklist

Use this checklist to ensure you've completed all setup steps correctly.

---

## ✅ Pre-Installation (5 minutes)

- [ ] **Created Google Sheet** with a descriptive name
- [ ] **Created sheet named** "Server Admin Team" (exactly this name)
- [ ] **Added column headers** in row 1:
  - [ ] Column B: Name
  - [ ] Column C: Team  
  - [ ] Column F: Steam ID
- [ ] **Added at least one admin** in row 2 with:
  - [ ] Admin name
  - [ ] Team name
  - [ ] 64-bit Steam ID (17 digits)
- [ ] **Obtained BattleMetrics API token** from https://www.battlemetrics.com/developers
- [ ] **Found your Server ID** (number from your BattleMetrics server URL)

---

## ✅ Script Installation (5 minutes)

- [ ] **Opened Apps Script** (Extensions → Apps Script)
- [ ] **Deleted default code** (if any)
- [ ] **Copied entire** `AdminHoursTracker.js` file
- [ ] **Pasted into Apps Script editor**
- [ ] **Saved the script** (Ctrl+S or Cmd+S)

---

## ✅ Configuration (3 minutes)

### GitHub URLs (lines 13-14)
- [ ] **Updated GITHUB_VERSION_URL** with your GitHub username
- [ ] **Updated GITHUB_DOWNLOAD_URL** with your GitHub username
- [ ] Or removed these lines if not using GitHub version checking

### BattleMetrics Settings
- [ ] **Set BM_SERVER_ID** (line 24) to your server ID number
  - Example: `BM_SERVER_ID: '12345678',`

### Column Configuration
- [ ] **Verified SOURCE_COL_NAME** matches your name column (should be 2 for column B)
- [ ] **Verified SOURCE_COL_TEAM** matches your team column (should be 3 for column C)
- [ ] **Verified SOURCE_COL_STEAM_ID** matches your Steam ID column (should be 6 for column F)

### Team Colors (Optional)
- [ ] **Customized TEAM_COLORS** with your team names
- [ ] **Verified team names exactly match** what's in your source sheet

### Final Check
- [ ] **Saved all changes** (Ctrl+S or Cmd+S)

---

## ✅ API Key Setup (2 minutes)

- [ ] **Clicked Project Settings** (gear icon) in Apps Script
- [ ] **Scrolled to Script Properties**
- [ ] **Clicked "Add script property"**
- [ ] **Entered property name:** `BATTLEMETRICS_API_KEY`
- [ ] **Pasted your API token** as the value
- [ ] **Clicked "Save script properties"**
- [ ] **Verified token is saved** (should show in list)

---

## ✅ Initial Setup (30 minutes)

- [ ] **Closed and reopened** Google Sheet
- [ ] **New menu appeared:** ⚙️ Admin Tracker (v1.0.0)
- [ ] **Clicked:** ⚙️ Admin Tracker → 🚀 Complete Initial Setup
- [ ] **Granted permissions** when prompted:
  - [ ] Clicked "Review Permissions"
  - [ ] Selected Google account
  - [ ] Clicked "Advanced"
  - [ ] Clicked "Go to [Your Project]"
  - [ ] Clicked "Allow"
- [ ] **Waited 25-30 minutes** for setup to complete
- [ ] **Verified "Admin Rankings" sheet** was created
- [ ] **Verified data appeared** in rankings sheet

---

## ✅ Trigger Setup (2 minutes)

- [ ] **Clicked Triggers** (clock icon) in Apps Script editor
- [ ] **Clicked "+ Add Trigger"**
- [ ] **Selected function:** `updateAllHoursHourly`
- [ ] **Selected event source:** Time-driven
- [ ] **Selected type:** Hour timer
- [ ] **Selected interval:** Every hour
- [ ] **Clicked Save**
- [ ] **Trigger appears in list**

---

## ✅ Verification (5 minutes)

### Test Menu Functions
- [ ] **⚙️ Admin Tracker** menu exists
- [ ] **📊 Update Status** shows timestamps
- [ ] **🔍 Clear Search** works
- [ ] **📊 Check Quota Status** runs without error
- [ ] **🧪 Test API Connection** shows success

### Test Search
- [ ] **Typed a name** in search box (column T)
- [ ] **Match highlighted** in bright cyan
- [ ] **Clicked CLEAR** and highlight removed

### Test Real-time Sync
- [ ] **Edited an admin name** in "Server Admin Team" sheet
- [ ] **Rankings updated automatically**

### Verify Data
- [ ] **7-Day Rankings** show data
- [ ] **30-Day Rankings** show data
- [ ] **90-Day Rankings** show data
- [ ] **Team colors** appear correctly
- [ ] **Hour colors** gradient from red to green

---

## ✅ Optional: GitHub Setup

- [ ] **Created GitHub repository** named `admin-hours-tracker`
- [ ] **Uploaded files** to repo
- [ ] **Updated GitHub URLs** in script
- [ ] **Tested update checker**

---

## 🎉 Final Checklist

If you can check all these, you're fully set up!

- [ ] **Menu works** and shows version number
- [ ] **All three rankings** display data
- [ ] **Search works** and highlights matches
- [ ] **Hourly trigger** is configured
- [ ] **No errors** in execution log
- [ ] **Data updates automatically** every hour

---

## 🐛 If Something's Wrong

Go to the [Troubleshooting section](README.md#-troubleshooting) in README.md

Common issues:
- Missing API key → Add to Script Properties
- No data showing → Check Steam ID format (17 digits)
- Errors in log → Verify API token permissions
- Menu missing → Close and reopen sheet

---

**Still stuck?** Open an issue on GitHub with:
- Screenshots of your setup
- Error messages from execution log
- Description of what's not working

**Happy tracking!** 🎊
