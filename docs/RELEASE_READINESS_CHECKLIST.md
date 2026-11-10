# FastMailSearch v1.0.7 Release Readiness Checklist

This document certifies release readiness for **FastMailSearch v1.0.7**.

---

## 1. Release Identity & Version Consistency

- **Product Name:** FastMailSearch
- **Release Version:** `v1.0.7`
- **Release Git Commit:** `31bfba0a2ecb3c408fb5b3ddb45d2e52dbe47e2b`
- **Assembly FileVersion:** `1.0.7.0`
- **Assembly ProductVersion:** `1.0.7+31bfba0a2ecb3c408fb5b3ddb45d2e52dbe47e2b`
- **Installer Package:** `FastMailSearch-v1.0.7-Store-Setup.exe`
- **Exact Installer Size:** `4,308,480 bytes`
- **Installer SHA-256:** `F2E6A01D9CA9E49FF00AA993EC9083B25EDECFEFD0B4C068F6E7BC5FF4D101CF`

---

## 2. Supported Platform Statement

- **Operating System:** Windows 10 (64-bit) and Windows 11 (64-bit).
- **Runtime Prerequisites:** Microsoft .NET Framework 4.8.
- **Supported Host Application:** Classic Microsoft Outlook Desktop (Win32 COM) 2016, 2019, 2021, and Microsoft 365 Desktop (32-bit & 64-bit).
- **Explicitly Unsupported Platforms:**
  - New Outlook for Windows (WebView2 / Web client)
  - Outlook on the Web (OWA)
  - Outlook for Mac, iOS, and Android

---

## 3. Documentation Completeness Audit

| Documentation Artifact | File Path | Status | Verification Summary |
| :--- | :--- | :---: | :--- |
| **Changelog** | `CHANGELOG.md` | **Verified** | Accurately documents v1.0.7 status clarification, heartbeat tracking, and safe email opening error handling without invented features. |
| **User Guide** | `docs/USER_GUIDE.md` | **Verified** | Explains installation, ribbon integration, search syntax, status indicators, and non-destructive troubleshooting. |
| **Known Limitations** | `docs/KNOWN_LIMITATIONS.md` | **Verified** | Discloses platform boundaries and explains the 3,429 vs 10,206 queue count context accurately without unverified DB claims. |
| **Install & Upgrade Guide** | `docs/INSTALLATION_AND_UPGRADE.md` | **Verified** | Details per-user setup, automated pre-upgrade backups, and rollback instructions. |

---

## 4. Production Safety & Invariant Verification

- [x] **Pre-Upgrade Backup:** Created and verified bit-for-bit with production database baseline.
- [x] **Production Database Hash:** `32742A9A63E7585A630AE4157D652A83138FD24582DBF2C577A6B5E69BB94CFB` (Untouched).
- [x] **Primary PST Size:** `4,632,609,792 bytes` (Untouched).
- [x] **Process Isolation:** Verified 0 running Outlook or UI processes following testing.
- [x] **Rollback Readiness:** Prior v1.0.6 installer package and automated backups preserved in `%USERPROFILE%\Documents\FastMailSearch\`.

