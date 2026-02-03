# Changelog

All notable changes to this project will be documented in this file.

## [1.0.3] - 2026-02-03

### Added
- **Manual Font Name Entry**: Added a text field in the GUI to manually specify the target font name (e.g., "FMAbhaya") if the default name is not found on the system.
- **Robust Font Fallback**: The app now tries multiple variations of the font name (`FMAbhaya`, `FM-Abhaya`, `Abhaya`) automatically.

### Fixed
- Improved error handling for `Server execution failed` (COM Hang), now explicitly advising a PC restart.
- Added specific handling for `RPC_E_SERVERCALL_RETRYLATER` (Application Busy).

## [1.0.2] - 2026-02-03

### Fixed
- Added robust retry mechanism (3 retries, 1s sleep) for identifying Photoshop connection issues when the application is busy (RPC_E_SERVERCALL_RETRYLATER).

## [1.0.1] - 2026-02-03

### Fixed
- Fixed critical `CoInitialize` error when connecting to Photoshop from a background thread.
- Added Robust fallback to `Dispatch` if `GetActiveObject` fails.

## [1.0.0] - 2026-02-03

### Added
- Initial release of SinhalaType Pro.
- Core conversion logic (Unicode to FMAbhaya).
- Support for Kombuwa reordering (Split-Swap-Append).
- Support for Rephaya and Rakaaransaya handling.
- GUI using CustomTkinter with "Dark" theme.
- Photoshop integration via pywin32 COM interface.
- Clipboard auto-copy feature.
