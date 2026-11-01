# FastMail Search

**FastMail Search** is a standalone Classic Outlook for Windows add-in and desktop email search engine powered by a local SQLite FTS5 database and modern WPF user interface. It delivers sub-second full-text searches across hundreds of thousands of indexed emails without relying on cloud services, external servers, or web APIs.

---

## Technical Features

- **Classic Outlook Support**: Targets Outlook Classic for Windows (Office 2016 / 2019 / 2021 / 365 Desktop).
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

## Commercial Licensing

- **License Model**: Proprietary — All Rights Reserved. (See [LICENSE](LICENSE) for full terms).
- **Free Evaluation**: Fully functional 30-day free trial with unrestricted local email search and indexing.
- **Commercial License**: ₹499 for a 1-year license (includes software usage and updates for 1 year from the date of purchase).
- **Offline Cryptographic Activation**: Licensing verification operates 100% offline using RSA cryptographic signatures bound to a local Hardware ID.
- **No Telemetry or Tracking**: No automatic online phone-home, no background analytics, and no recurring billing or automatic renewal.
- **Source Code Notice**: Access to this source code repository is provided solely for authorized review and inspection. No open-source license is granted.

---

## Project Architecture

```
FastMailSearch/
│
├── FastMailSearch.slnx                # Visual Studio solution file
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
└── docs/
    └── FINAL_VERIFICATION.md          # Architectural verification log & measured performance metrics
```

---

## Developer Setup & Build Instructions

### Prerequisites
- Windows 10 or Windows 11 (64-bit recommended)
- Classic Microsoft Outlook (Office 365, 2021, 2019, 2016 Desktop)
- Visual Studio 2022 (with .NET Desktop Development, Office/VSTO, and .NET Framework 4.8 targeting pack)
- .NET SDK 8.0 or higher

### Build the Solution
Because FastMailSearch targets both modern .NET 8 and .NET Framework 4.8 with COM interop, compile using Visual Studio MSBuild:
```powershell
msbuild FastMailSearch.slnx /p:Configuration=Release /v:m
```

### Run Automated Tests
Execute the unit and database test suites across both supported frameworks (.NET 8 and .NET Framework 4.8):
```powershell
dotnet test tests/FastMailSearch.Core.Tests/FastMailSearch.Core.Tests.csproj -c Release --no-build
dotnet test tests/FastMailSearch.Database.Tests/FastMailSearch.Database.Tests.csproj -c Release --no-build
```

### Build the Standalone Release Package
To package the standalone release installer (`FastMailSearch-v1.0.6-Store-Setup.exe`):
```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\Build-ReleasePackage.ps1 -Version 1.0.6
```

---

## Storage & Logs Location

- **SQLite Search Database**: `%USERPROFILE%\Documents\FastMailSearch\FastMailSearch.db`
  *(Path can be customized via the `FASTMAILSEARCH_DB_PATH` environment variable)*
- **Application Diagnostic Logs**: `%LOCALAPPDATA%\FastMailSearch\Logs\`

---

## Security & Support

- **Security Disclosure**: Please report any security vulnerabilities responsibly according to [SECURITY.md](SECURITY.md).
- **Support Email**: [developer.diptesh@gmail.com](mailto:developer.diptesh@gmail.com)
- **GitHub Repository**: [https://github.com/dipanki111/FastMailSearch](https://github.com/dipanki111/FastMailSearch)
