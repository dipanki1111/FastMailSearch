# Changelog

All notable changes to **FastMailSearch** are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [1.0.7] - 2026-11-10

### Release Metadata
- **Commit:** `31bfba0a2ecb3c408fb5b3ddb45d2e52dbe47e2b`
- **Installer:** `FastMailSearch-v1.0.7-Store-Setup.exe`
- **Installer SHA-256:** `F2E6A01D9CA9E49FF00AA993EC9083B25EDECFEFD0B4C068F6E7BC5FF4D101CF`
- **Target Platform:** Windows 10 / 11 (64-bit), Classic Microsoft Outlook Desktop (Win32 COM)

### Changed
- **Indexing Status Presentation Clarification**:
  - Refined status generation in `SearchViewModel` and `IndexHealthService` so that the status text accurately reflects whether a background indexing worker is genuinely active.
  - Eliminated misleading `"Updating · N email(s) remaining"` state when the queue is paused and no indexing worker is executing.
  - Introduced explicit status differentiation:
    - `🔵 Updating • N email(s) remaining`: Active background worker is actively indexing items.
    - `⏸ Idle — Queue Paused (N pending)`: Queue has pending work, but indexing is paused/idle.
    - `✅ Up to date (N emails indexed)`: All store items are fully indexed and queue is clear.
    - `🟡 N indexing issue(s)`: Genuine indexing errors present for specific items.
  - Added cross-process indexing heartbeat validation to ensure remote worker progress is accurately tracked without false positives.

### Added
- **Safe Email Opening Failure Handling**:
  - Hardened email opening workflow in `OutlookEmailOpener` and `DiagnosticsWindow` against invalid, deleted, or unmapped StoreId/EntryId pairs.
  - Added structured error handling around Outlook COM `Session.GetItemFromID` and `Inspector.Display` calls, preventing application or Outlook instability when opening search results for items that were moved or deleted in Outlook.
  - Added user-facing error reporting for inaccessible email records.

---

## [1.0.6] - 2026-11-08

### Release Metadata
- **Commit:** `9c8df33bc511741740353d2706333a6e262e9ed5`
- **Installer:** `FastMailSearch-v1.0.6-Store-Setup.exe`
- **Installer SHA-256:** `1477B51432FE74A2C3A83CC03CF60B74AFEADBE96B19FA5658101650D30D615B`

### Added
- **Automated Pre-Upgrade Database Backup**:
  - Installer automatically creates a timestamped backup of the production database (`FastMailSearch.db`) and SQLite companion files (`.db-wal`, `.db-shm`) in `%USERPROFILE%\Documents\FastMailSearch\` prior to binary replacement.
  - Aborts installation safely if backup creation fails, preventing data corruption during upgrades.
- **Enhanced Diagnostics Window**:
  - Added live index health reporting, failed items count, and store binding diagnostic details.

---

## [1.0.0] - 2026-10-31

### Initial Release
- **Classic Outlook COM Add-In**: Deep Win32 COM integration for Microsoft Outlook 2016, 2019, 2021, and Microsoft 365 Desktop.
- **Local SQLite FTS5 Engine**: Sub-second full-text email search across hundreds of thousands of items without cloud dependencies.
- **Search Capabilities**: Keyword search, exact phrase matching, field filters (`from:`, `to:`, `subject:`, `folder:`), property filters (`hasattachment:`, `isread:`, `isflagged:`), and date range queries.
- **Background Incremental Indexing**: Non-blocking background indexing with real-time Outlook store event listening (`ItemAdd`, `ItemChange`, `ItemRemove`).
- **COM Lifecycle Safety**: Deterministic COM reference release preventing process hangs on Outlook exit.
- **100% Offline Licensing**: RSA cryptographic licensing engine with 30-day evaluation trial.

