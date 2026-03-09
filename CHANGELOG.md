# Changelog
All notable changes to this project will be documented in this file.

## [0.16.3] - 2026-03-09
### Added
- **Project Documentation**: Initialized CHANGELOG.md to track version history and system evolution.
- **Version Alignment**: Synced project versioning between Google Apps Script environment and Git repository.

## [0.16.2] - 2026-03-09
### Added
- **DateService Refactor**: Centralized date logic into a dedicated service. Added ddDays and getHebrewTwoWeeks for automated date calculations.
- **Dynamic Mailer**: Implemented automatic Hebrew date range generation for weekly emails.

### Fixed
- **Export Style Sync**: Fixed an issue where exported schedules didn't match the source sheet's theme.
- **Scope Error**: Resolved ReferenceError: weekColor is not defined in the uildWeekBlock function.
- **Drive API Optimization**: Refactored folder lookup logic to reduce redundant Google Drive API calls.
