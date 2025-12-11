# Changelog

All notable changes to this project will be documented in this file.

## [v0.9.6] - 2025-12-11

SHA256: `26F8697B6B6770116D27CA7C6688A1E9B240168A77F17AA6414922E00B0B529F`

### Changed
- Renamed "Fix Audio" button to "Reset Audio" for clarity
- Added clarification in README that tool resets to user-defined baseline
- System tray with right-click context menu (Open Settings / Run Reset / Exit)
- Left-click tray icon opens config GUI after reset
- Safe volume handling - saves current volume, lowers to 20% during reset, restores after

## [v0.9.3] - 2025-12-06

SHA256: `83BC559CB7E28C9AC1778805066902883017BDC3D5235ECC5F4C82E4630A9699`

### Changed
- Improved README documentation
- Fixed release workflow (removed extra code block)
- Restructured Advanced Method instructions

## [v0.9.1] - 2025-12-06

SHA256: `78194058B8C4A04015FE221A659506C4BDA059A6BD8F7628561231464A32045A`

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
