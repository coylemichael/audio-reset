# Changelog

All notable changes to this project will be documented in this file.

## [v0.9.1] - 2025-12-06

### Added
- Automated release builds with GitHub Actions
- Build attestation for provenance verification
- SHA256 hash in release notes and changelog

## [v0.9] - 2025-12-06

### Added
- Configuration mismatch detection - detects when saved audio config differs from current Windows settings
- Custom mismatch dialog with side-by-side comparison showing "Saved Configuration" vs "Current Windows Settings"
- "Restore Saved" / "Keep Current" buttons to choose which config to apply
- Run in background mode - option to run silently without showing GUI
- Completion notification option - optional popup when audio reset finishes
- Editable install folder with automatic file migration
- Dark theme UI throughout all dialogs
- Blue accent colors for headers and labels
