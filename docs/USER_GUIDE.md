# FastMailSearch User Guide (v1.0.7)

Welcome to **FastMailSearch**, a high-performance offline desktop search engine and add-in for Classic Microsoft Outlook on Windows.

---

## 1. System Requirements & Compatibility

- **Operating System:** Windows 10 or Windows 11 (64-bit).
- **Email Client:** Classic Microsoft Outlook for Windows (Desktop):
  - Microsoft Office 2016 (32-bit & 64-bit)
  - Microsoft Office 2019 (32-bit & 64-bit)
  - Microsoft Office 2021 (32-bit & 64-bit)
  - Microsoft 365 Desktop Apps (32-bit & 64-bit)
- **Runtimes:** .NET Framework 4.8.
- **Unsupported:**
  - New Outlook for Windows (web-based / WebView2 client)
  - Outlook on the Web (OWA)
  - Outlook for Mac, iOS, or Android

---

## 2. Installation & First Launch

### Installation Steps
1. Ensure Microsoft Outlook is completely closed before installing.
2. Run the self-contained installer executable:
   `FastMailSearch-v1.0.7-Store-Setup.exe`
3. The installer operates in per-user scope under:
   `%LOCALAPPDATA%\Programs\FastMailSearch\`
   *(No administrative permissions are required).*
4. **Pre-Upgrade Safety:** The installer automatically checks for existing databases and generates a timestamped backup in `%USERPROFILE%\Documents\FastMailSearch\` before writing any new binaries.
5. Once the success dialog appears, click **OK** to finish setup.

### First Launch
1. Open Classic Microsoft Outlook.
2. In the Outlook Ribbon under the **Home** tab, locate the **FastMail Search** group.
3. Click **Open FastMail Search** to launch the search window.

---

## 3. Search Syntax & Usage

FastMailSearch allows fast, flexible querying across your indexed mailboxes using a local SQLite FTS5 engine:

### Free-Text & Exact Phrase Queries
- `invoice` — Searches for the word "invoice" across subject, sender, recipients, and body.
- `"project kickoff meeting"` — Enclosing terms in double quotes searches for the exact phrase in sequential order.
- `report 2026` — Matches emails containing both "report" and "2026".

### Field Operators
- `from:john@example.com` or `from:john` — Matches sender display name or email address.
- `to:finance@company.com` — Matches direct recipient address.
- `subject:status` — Searches specifically within email subject headers.
- `folder:inbox` or `folder:archive` — Restricts search to a specific Outlook folder.

### Property Filters
- `hasattachment:true` or `hasattachment:false` — Filters emails with or without attachments.
- `isread:false` — Shows only unread messages.
- `isflagged:true` — Shows only flagged/followed-up messages.

### Date Filters
- `after:2026-01-01` — Emails received on or after January 1, 2026.
- `before:2026-06-30` — Emails received on or before June 30, 2026.
- `today`, `yesterday`, `thisweek`, `thismonth` — Dynamic date range shortcuts.

### Opening Emails
- **Double-click** any result row in the search results table to open the email directly in its native Microsoft Outlook Inspector window.
- If an email has been moved or deleted from Outlook since indexing, FastMailSearch safely intercepts the COM lookup and displays a clear message rather than crashing or freezing Outlook.

---

## 4. Indexing Status Messages Explained

FastMailSearch provides clear status indicators in the search window status bar:

| Status Display | Meaning & Behavior |
| :--- | :--- |
| `🔵 Updating • N email(s) remaining` | **Active Indexing:** A background indexing worker is actively reading and processing email items from your Outlook store into the local index. |
| `⏸ Idle — Queue Paused (N pending)` | **Queue Paused:** Pending email changes exist in the queue, but the indexing worker is currently idle or paused. **Important: A paused queue does NOT mean active indexing is taking place.** Background resources are idle. |
| `✅ Up to date (N emails indexed)` | **Fully Synchronized:** All store items have been indexed into the local database, and no pending changes remain in the queue. |
| `🟡 N indexing issue(s)` | **Attention Needed:** Specific items encountered an issue during indexing (e.g. corrupted attachment or inaccessible item). Details can be viewed in the Diagnostics window. |

---

## 5. Non-Destructive Troubleshooting Guidance

If you experience unexpected behavior, follow these non-destructive checks:

### Add-In Ribbon Button Does Not Appear
1. In Classic Outlook, go to **File** > **Options** > **Add-ins**.
2. At the bottom of the window, ensure **Manage: COM Add-ins** is selected, and click **Go...**.
3. Verify that **FastMail Search** is listed and its checkbox is checked.
4. If unchecked, check it and click **OK**.
5. If listed under "Disabled Application Add-ins", select **Manage: Disabled Items** > **Go...**, select **FastMail Search**, and click **Enable**.

### Search Window Reports "Queue Paused"
- If status shows `Idle — Queue Paused`, this is normal when Outlook is closed or when indexing is paused to preserve CPU during intensive user tasks.
- The queue will resume processing when background indexing cycles engage.

### What NOT to Do (Data Safety Precautions)
- **Do not** manually delete, move, or rename `%USERPROFILE%\Documents\FastMailSearch\FastMailSearch.db`.
- **Do not** delete or edit Outlook data files (`.pst` or `.ost`).
- **Do not** run third-party database optimization tools or cleaners against the FastMailSearch database directory.

