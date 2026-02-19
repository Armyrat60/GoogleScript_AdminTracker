# Changelog

All notable changes to Admin Hours Tracker will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [1.0.0] - 2026-02-18

### Added
- Initial release with full feature set
- Rotating hourly updates to prevent timeout errors
  - Hour 0,3,6,9,12,15,18,21 → 7-day hours
  - Hour 1,4,7,10,13,16,19,22 → 30-day hours
  - Hour 2,5,8,11,14,17,20,23 → 90-day hours
- Automatic GitHub version checking
- Update notification system with one-click download
- Real-time sync from source data
- Color-coded rankings based on hours (gradient from red to green)
- Team-specific text coloring
- Quota tracking and management
- Enhanced menu with:
  - Version display in menu title
  - Update Status submenu
  - Rotating Updates explanation
  - Manual update options
  - About dialog
- Complete initial setup wizard
- Manual full update option
- Quota status checker
- API connection tester

### Technical Details
- BattleMetrics API integration
- Google Sheets conditional formatting for search
- Script Properties for secure API key storage
- Cache-based quota tracking
- Hourly time-based triggers
- OnEdit triggers for real-time sync

---

## [Unreleased]

### Planned Features
- Export rankings to PDF
- Email notifications for quota limits
- Historical data tracking
- Charts and graphs
- Multi-server support
- Discord webhook integration

---

## Version Number Guide

**Format:** MAJOR.MINOR.PATCH

- **MAJOR:** Breaking changes (requires manual intervention)
- **MINOR:** New features (backward compatible)
- **PATCH:** Bug fixes and small improvements

---

## How to Update

1. Check version in menu: ⚙️ Admin Tracker (v1.0.0)
2. Click: 🔄 Check for Updates
3. If update available, click OK to open GitHub
4. Download new AdminHoursTracker.js
5. Replace script code in Apps Script editor
6. Save and refresh sheet

Your data is preserved during updates!
