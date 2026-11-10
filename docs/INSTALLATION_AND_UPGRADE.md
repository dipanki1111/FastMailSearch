# FastMailSearch Installation & Upgrade Instructions (v1.0.7)

This document provides technical instructions for installing and upgrading **FastMailSearch** on Windows workstations.

---

## 1. Per-User Installation Architecture

FastMailSearch installs strictly under a **per-user scope**, requiring zero administrative privileges:

- **Target Directory:** `%LOCALAPPDATA%\Programs\FastMailSearch\`
- **Registry Integration:**
  - COM Add-in Registration: `HKCU\Software\Microsoft\Office\Outlook\Addins\FastMailSearch.AddIn`
  - InprocServer32 Registration: `HKCU\Software\Classes\CLSID\{4A10F4D0-7E55-442C-B23B-983F78BA48D2}`
- **Security Advantages:**
  - No UAC (User Account Control) prompt or local Administrator rights needed.
  - Complete isolation from other user profiles on shared computers.
  - Native COM registration avoids VSTO runtime security policy barriers.

---

## 2. Interactive Installation Steps

1. **Close Outlook:** Before launching the installer, ensure Microsoft Outlook is completely closed to prevent file locking on existing add-in binaries.
2. **Execute Installer:** Run `FastMailSearch-v1.0.7-Store-Setup.exe` interactively (do not use `/quiet` or `/silent` in production deployments to ensure modal notifications are received).
3. **Automated Safety Preflight:**
   - The installer verifies system prerequisites (.NET Framework 4.8, x64 architecture).
   - The installer checks target path integrity to prevent junction/symlink reparse redirection.
4. **Automated Database Backup:**
   - If an existing database (`FastMailSearch.db`) is discovered in `%USERPROFILE%\Documents\FastMailSearch\`, the installer automatically creates a timestamped pre-upgrade backup:
     `FastMailSearch_Backup_PreUpgrade_YYYYMMDD_HHMMSS.db`
     (including `.db-wal` and `.db-shm` files if present).
   - If backup creation encounters an error, the installer **aborts immediately** to prevent any data loss.
5. **Payload Extraction & Registration:**
   - The 31 release binaries and assets are extracted to `%LOCALAPPDATA%\Programs\FastMailSearch`.
   - The native COM add-in is registered in `HKCU`.
6. **Completion:** A confirmation dialog is displayed. Click **OK** to finish setup.

---

## 3. Upgrade Expectations & Precautions

- **Database Preservation:** Upgrading from v1.0.0 through v1.0.6 to v1.0.7 preserves all existing indexed email records, search history, and settings without requiring re-indexing.
- **Backup Location:** All automated pre-upgrade backups are safely preserved in:
  `%USERPROFILE%\Documents\FastMailSearch\`
  and are never overwritten or deleted by subsequent upgrades.
- **Boundaries & Scope:**
  - FastMailSearch does not offer cross-computer database migration tools or cloud synchronization.
  - Upgrades apply only to the local user profile where the installer is executed.

---

## 4. Rollback & Uninstallation

### Clean Uninstallation
To remove FastMailSearch:
1. Close Classic Outlook.
2. Run the uninstaller or removal script (`scripts/Uninstall-FastMailSearch.ps1`), or remove via Windows Installed Apps.
3. Uninstallation removes binaries from `%LOCALAPPDATA%\Programs\FastMailSearch` and unregisters the add-in from `HKCU`.
4. Your index database in `%USERPROFILE%\Documents\FastMailSearch\` is preserved by default to protect your email index data.

### Rollback Procedure
If it is necessary to roll back to a prior release (such as v1.0.6):
1. Close Classic Outlook.
2. Run the previous verified installer package (e.g. `FastMailSearch-v1.0.6-Store-Setup.exe`).
3. The previous installer will verify the existing database and restore the previous binary version cleanly.

