# FastMailSearch Known Limitations (v1.0.7)

This document outlines platform constraints, architectural boundaries, and known operational characteristics for **FastMailSearch v1.0.7**.

---

## 1. Platform & Application Compatibility

### Supported
- **Windows OS:** Windows 10 (64-bit) and Windows 11 (64-bit).
- **Outlook Host:** Classic Microsoft Outlook Desktop (Win32 COM / VSTO host).
  - Outlook 2016, 2019, 2021 (32-bit & 64-bit).
  - Microsoft 365 Desktop Apps (32-bit & 64-bit).

### Explicitly Unsupported
- **"New Outlook for Windows"**: The modern web-based / WebView2 client introduced by Microsoft does not support Win32 COM add-ins, native Windows message loops, or local MAPI store hooks.
- **Outlook on the Web (OWA)**: Web browsers cannot host desktop COM add-ins.
- **Outlook for Mac (macOS)**: Apple macOS lacks the Windows COM/MAPI infrastructure utilized by FastMailSearch.
- **Outlook Mobile (iOS & Android)**: Mobile platforms do not support native desktop add-ins.

---

## 2. PendingChanges Queue Metrics & Count Context

During production smoke testing and reconciliation audits of v1.0.7, two distinct counts related to the `PendingChanges` queue were recorded:

### 1. UI Displayed Active Queue Count: 3,429
In the verified production UI smoke test, the search window status bar displayed:
```
Idle — Queue Paused (3,429 pending)
```
- **Architectural Basis:** The UI query (`GetPendingChangesCountAsync()`) specifically filters for actionable work awaiting background indexing:
  ```sql
  SELECT COUNT(*) FROM PendingChanges WHERE Status IN ('Pending', 'Processing');
  ```
- This metric represents the active queue items requiring indexing worker attention.

### 2. Total Database Table Row Count: 10,206
During earlier pre-upgrade reporting in Stage 7AL-7, an unfiltered count of the table reported **10,206 total rows**:
```sql
SELECT COUNT(*) FROM PendingChanges;
```
- **Context:** The `PendingChanges` table retains historical records across the operational lifecycle of the indexer (e.g. items with `Status = 'Completed'` or `Status = 'Failed'`).
- **Safety Boundary Note:** In strict adherence to Stage 7AL-8 read-only safety protocols, **no fresh query was executed against the production database file** during the reconciliation stage.
- **Limitation:** Consequently, no exact numerical breakdown between individual historical statuses (such as `Completed` vs `Failed`) is claimed or assumed beyond the code-confirmed filter logic.

---

## 3. Storage & Operational Boundaries

- **Local Machine Scope:** FastMailSearch maintains an entirely offline, local SQLite database (`%USERPROFILE%\Documents\FastMailSearch\FastMailSearch.db`). It does not sync indexes across multiple machines or mobile devices.
- **Classic Outlook COM Dependency:** Outlook must be configured with accessible profile stores (PST/OST) in Classic Outlook. If an email item is permanently deleted from Outlook outside of FastMailSearch, its search result entry remains in the local index until the next reconciliation cycle cleans the record.

