# FastMail Search

[![Latest Release](https://img.shields.io/badge/Release-v1.0.6-0078D4)](https://github.com/dipanki1111/FastMailSearch/releases/tag/v1.0.6)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%20%7C%2011%20(64--bit)-0078D4)](#system-requirements)
[![Outlook](https://img.shields.io/badge/Outlook-Classic%20Desktop%20(2016--365)-0078D4)](#system-requirements)
[![Tests](https://img.shields.io/badge/Tests-661%20Passed%20%7C%200%20Failed-success)](#run-automated-tests)
[![License](https://img.shields.io/badge/License-Proprietary%20--%20All%20Rights%20Reserved-red)](LICENSE)

**FastMail Search** is a standalone Classic Outlook for Windows add-in and desktop email search engine powered by a local SQLite FTS5 database and modern WPF user interface. It delivers sub-second full-text searches across hundreds of thousands of indexed emails without relying on cloud services, external servers, or web APIs.

---

## Technical Features

- **Classic Outlook Integration**: Deep COM integration with Outlook Classic for Windows (Office 2016 / 2019 / 2021 / 365 Desktop).
- **SQLite FTS5 Full-Text Search**: Local FTS5 database stored by default at `%USERPROFILE%\Documents\FastMailSearch\FastMailSearch.db`.
- **Flexible Search Syntax**:
  - Full-text keywords: `invoice ABC`, `"exact phrase"`
  - Field filters: `from:domain.com`, `to:john`, `subject:status`
  - Property filters: `hasattachment:true`, `isread:false`, `isflagged:true`, `folder:inbox`
  - Date ranges: `after:2026-01-01`, `before:2026-04-01`, `today`, `yesterday`, `thisweek`, `thismonth`
- **Background Incremental Indexer**: Scans mailboxes in non-blocking background batches and listens to real-time Outlook store events (`ItemAdd`, `ItemChange`, `ItemRemove`).
- **COM Lifecycle Safety**: Deterministic COM reference management (`ComReleaseHelper`) preventing Outlook memory leaks, process locks, or hangs.
- **Privacy & Security**: 100% offline operation. All email content and index metadata remain strictly on your local computer.

---

## Compatibility & System Requirements

### Supported Environments
- **Operating System**: Windows 10 or Windows 11 (64-bit recommended)
- **Email Client**: Classic Microsoft Outlook for Windows (Win32 COM / VSTO host)
  - Microsoft Office 2016 (32-bit & 64-bit)
  - Microsoft Office 2019 (32-bit & 64-bit)
  - Microsoft Office 2021 (32-bit & 64-bit)
  - Microsoft 365 Desktop Apps (32-bit & 64-bit)
- **Runtimes**: .NET Framework 4.8 and .NET Desktop Runtime 8.0+

### Unsupported Environments
- **New Outlook for Windows**: The modern web-based / WebView2 Outlook client does not support Win32 COM/VSTO add-in APIs.
- **Cross-Platform**: Outlook for Mac, Outlook on the Web (OWA), and Outlook Mobile (iOS / Android) are not supported.

---

## Public Source & Commercial Licensing

- **License Model**: Proprietary — All Rights Reserved. (See [LICENSE](LICENSE) for full legal terms).
- **Public Source Availability**: Access to this source code repository is provided solely for authorized evaluation, technical review, and security inspection. No open-source license (such as MIT, Apache, or GPL) is granted.
- **Proprietary Licensing Isolation**: To protect commercial license verification integrity, the offline cryptographic license key generator (`FastMailSearch.LicenseGenerator`) and master private signing keys are maintained privately and are excluded from this repository.
- **Free Evaluation**: FastMail Search includes a fully functional 30-day evaluation trial with unrestricted local search and indexing capabilities.
- **Commercial License**: ₹499 for a 1-year commercial license (includes software usage and updates for 1 year from the date of purchase).
- **100% Offline Cryptographic Activation**: Licensing verification operates entirely offline using RSA cryptographic signatures bound to a local Hardware ID.
- **No Telemetry or Phone-Home**: Zero analytics, zero usage tracking, and no background network connections.

---

## Project Architecture

```
FastMailSearch/
│
├── FastMailSearch.slnx                # Visual Studio solution file (all public components)
├── Directory.Build.props              # Harmonized assembly version & build metadata
├── LICENSE                            # Proprietary software license agreement
├── SECURITY.md                        # Security policy & vulnerability reporting guide
├── THIRD-PARTY-NOTICES.md             # Third-party open-source dependency notices
│
├── src/
│   ├── FastMailSearch.Core/           # Core domain models, search parser, interfaces, licensing contracts
│   ├── FastMailSearch.Infrastructure/ # SQLite FTS5 repository, background indexer, Outlook COM adapter, licensing verification
│   ├── FastMailSearch.OutlookAddIn/   # Classic Outlook COM Add-In host & Ribbon XML integration
│   ├── FastMailSearch.UI/             # Modern WPF search application, view models, settings, activation UI
│   └── FastMailSearch.Setup/          # Standalone payload extractor & self-contained installer
│
├── tests/
│   ├── FastMailSearch.Core.Tests/     # Comprehensive unit tests (targeting net8.0 and net48)
│   ├── FastMailSearch.Database.Tests/ # SQLite & FTS5 database integration tests (targeting net8.0 and net48)
│   └── FastMailSearch.PerformanceTests/# Synthetic data benchmark engine (targeting net8.0 and net48)
│
├── tools/
│   └── SyntheticPstIntegrationTest/   # Isolated synthetic PST integration test runner
│
├── scripts/
│   ├── Build-ReleasePackage.ps1       # Release packaging and standalone setup generation script
│   ├── Install-FastMailSearch.ps1     # Standalone registration & installation script
│   └── Uninstall-FastMailSearch.ps1   # Clean uninstallation & registry cleanup script
│
├── website/                           # Static documentation & product website
│   ├── css/                           # Product site stylesheets
│   ├── assets/                        # Web icons and branding assets
│   └── *.html                         # Documentation pages (FAQ, install, support, smartscreen)
│
└── docs/
    └── FINAL_VERIFICATION.md          # Regression test & verification log for pagination hardening
```

---

## Developer Setup & Build Instructions

### Prerequisites
- Windows 10 or Windows 11 (64-bit)
- Visual Studio 2022 (Community, Professional, or Enterprise) with:
  - `.NET Desktop Development` workload
  - `Office/SharePoint development` workload (Visual Studio Tools for Office / COM)
  - `.NET Framework 4.8 Targeting Pack`
- .NET SDK 8.0 or higher
- Classic Microsoft Outlook installed locally (for add-in debugging)

### Build the Solution
Because FastMailSearch targets both modern .NET 8 and .NET Framework 4.8 with COM interop, compile using Visual Studio MSBuild:
```powershell
& "C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe" FastMailSearch.slnx /p:Configuration=Release /v:m
```
*(Or use `msbuild` from a Visual Studio Developer Command Prompt / PowerShell).*

### Run Automated Tests
Execute the unit and database test suites across both supported runtime targets (.NET 8 and .NET Framework 4.8):
```powershell
# Core unit tests (635 tests across net8.0 and net48)
dotnet test tests/FastMailSearch.Core.Tests/FastMailSearch.Core.Tests.csproj -c Release --no-build

# Database & SQLite FTS5 tests (26 tests across net8.0 and net48)
dotnet test tests/FastMailSearch.Database.Tests/FastMailSearch.Database.Tests.csproj -c Release --no-build
```

### Synthetic PST Integration Test
To run integration tests against a synthetic test store without touching personal Outlook mailboxes:
```powershell
& ".\tools\SyntheticPstIntegrationTest\bin\Release\net48\SyntheticPstIntegrationTest.exe"
```

### Build the Standalone Release Package
To package the standalone release installer (`FastMailSearch-v1.0.6-Store-Setup.exe`):
```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\Build-ReleasePackage.ps1 -Version 1.0.6
```

---

## Storage & Configuration

- **SQLite Search Database**: `%USERPROFILE%\Documents\FastMailSearch\FastMailSearch.db`
  *(Path can be customized via the `FASTMAILSEARCH_DB_PATH` environment variable)*
- **Application Diagnostic Logs**: `%LOCALAPPDATA%\FastMailSearch\Logs\`
- **Database Connection Strategy**: Configured with WAL (Write-Ahead Logging) mode, synchronous NORMAL, and connection pooling for concurrent UI search and background indexing.

---

## Contributions & Issue Reporting

- **Source Inspection**: This repository is source-available for technical review, security auditing, and code inspection.
- **Pull Requests**: Because FastMail Search is proprietary commercial software, external pull requests and code contributions cannot be accepted without prior written agreement and copyright assignment.
- **Issues & Feedback**: Bug reports, feature suggestions, and usability feedback are welcomed via [GitHub Issues](https://github.com/dipanki1111/FastMailSearch/issues).

---

## Security & Support

- **Security Disclosure**: Please report any security vulnerabilities responsibly according to [SECURITY.md](SECURITY.md).
- **Support Contact**: [developer.diptesh@gmail.com](mailto:developer.diptesh@gmail.com)
- **Official Website**: [https://fastmailsearch.com](https://fastmailsearch.com) (or browse local `/website/` documentation)
- **GitHub Repository**: [https://github.com/dipanki1111/FastMailSearch](https://github.com/dipanki1111/FastMailSearch)
- **Latest Release**: [v1.0.6 (Release Notes)](https://github.com/dipanki1111/FastMailSearch/releases/tag/v1.0.6)
