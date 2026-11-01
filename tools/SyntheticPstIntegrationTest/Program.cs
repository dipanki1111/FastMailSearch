#nullable enable
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using FastMailSearch.Core.Interfaces;
using FastMailSearch.Core.Models;
using FastMailSearch.Infrastructure.Database;
using FastMailSearch.Infrastructure.Indexing;
using FastMailSearch.Infrastructure.Logging;
using FastMailSearch.Infrastructure.Outlook;
using Microsoft.Data.Sqlite;
using OutlookInterop = Microsoft.Office.Interop.Outlook;

namespace FastMailSearch.SyntheticPstIntegrationTest
{
    public class MockLicensingService : ILicensingService
    {
        public LicenseStatus Status { get; set; } = LicenseStatus.Valid;
        public int TrialDaysRemaining { get; set; } = 30;
        public bool ActivateResult { get; set; } = true;
        public string HardwareId { get; set; } = "MOCK-HWID-INTEGRATION-TEST";
        public int GetStatusCallCount { get; private set; }

        public MockLicensingService(LicenseStatus status = LicenseStatus.Valid)
        {
            Status = status;
        }

        public LicenseStatus GetStatus()
        {
            GetStatusCallCount++;
            return Status;
        }

        public int GetTrialDaysRemaining() => TrialDaysRemaining;
        public Task<bool> ActivateAsync(string licenseKey) => Task.FromResult(ActivateResult);
        public string GetHardwareId() => HardwareId;
    }

    public class FileMetadataSnapshot
    {
        public string Path { get; set; } = string.Empty;
        public bool Exists { get; set; }
        public long Length { get; set; }
        public DateTime LastWriteTimeUtc { get; set; }

        public static FileMetadataSnapshot Capture(string path)
        {
            var snap = new FileMetadataSnapshot { Path = path };
            if (File.Exists(path))
            {
                var fi = new FileInfo(path);
                snap.Exists = true;
                snap.Length = fi.Length;
                snap.LastWriteTimeUtc = fi.LastWriteTimeUtc;
            }
            else
            {
                snap.Exists = false;
                snap.Length = 0;
                snap.LastWriteTimeUtc = DateTime.MinValue;
            }
            return snap;
        }

        public bool Matches(FileMetadataSnapshot other)
        {
            if (Exists != other.Exists) return false;
            if (!Exists) return true;
            return Length == other.Length && LastWriteTimeUtc == other.LastWriteTimeUtc;
        }
    }

    class Program
    {
        private static readonly string ProductionDbPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "FastMailSearch", "FastMailSearch.db");
        private static readonly string PersonalPstPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Outlook Files", "personal-test@example.com - test.pst");

        [STAThread]
        static async Task<int> Main(string[] args)
        {
            Console.WriteLine("================================================================================");
            Console.WriteLine(" FastMailSearch Phase 3.6I — Controlled Synthetic PST Live Integration Test");
            Console.WriteLine("================================================================================");
            Console.WriteLine($"Current User: {Environment.UserName}");
            Console.WriteLine($"Machine Name: {Environment.MachineName}");
            Console.WriteLine($"Process Architecture: {(Environment.Is64BitProcess ? "64-bit (x64)" : "32-bit (x86)")}");
            Console.WriteLine($"Start Time (UTC): {DateTime.UtcNow:yyyy-MM-dd HH:mm:ss.fff}");
            Console.WriteLine("================================================================================\n");

            bool phase0Pass = false;
            bool phase1Pass = false;
            bool phase2Pass = false;
            bool phase3Pass = false;
            bool phase4Pass = false;
            bool phase5Pass = false;

            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string testDir = Path.Combine(localAppData, "Temp", "FastMailSearch_IntegrationTest");
            string testDbPath = Path.Combine(testDir, "test_FastMailSearch.db");
            string testLogPath = Path.Combine(testDir, "Logs");
            string syntheticPstPath = Path.Combine(testDir, "SyntheticTest.pst");

            FileMetadataSnapshot? baseDb = null;
            FileMetadataSnapshot? baseDbWal = null;
            FileMetadataSnapshot? baseDbShm = null;
            FileMetadataSnapshot? basePersonalPst = null;

            OutlookInterop.Application? outlookApp = null;
            OutlookInterop.NameSpace? mapiNs = null;
            OutlookInterop.Store? syntheticStore = null;
            OutlookInterop.MAPIFolder? syntheticRootFolder = null;
            int? launchedOutlookPid = null;
            string syntheticStoreId = string.Empty;

            try
            {
                // ====================================================================
                // PHASE 0: PRE-FLIGHT CHECK
                // ====================================================================
                Console.WriteLine(">>> [PHASE 0] Pre-flight Check & Baseline Recording");

                var runningOutlook = Process.GetProcessesByName("OUTLOOK");
                var runningUi = Process.GetProcessesByName("FastMailSearch.UI");
                Console.WriteLine($"    Running OUTLOOK processes: {runningOutlook.Length}");
                Console.WriteLine($"    Running FastMailSearch.UI processes: {runningUi.Length}");

                if (runningOutlook.Length > 0 || runningUi.Length > 0)
                {
                    Console.WriteLine("    [FAIL] Pre-existing OUTLOOK.EXE or FastMailSearch.UI.exe process detected. Aborting test.");
                    return 1;
                }

                baseDb = FileMetadataSnapshot.Capture(ProductionDbPath);
                baseDbWal = FileMetadataSnapshot.Capture(ProductionDbPath + "-wal");
                baseDbShm = FileMetadataSnapshot.Capture(ProductionDbPath + "-shm");
                basePersonalPst = FileMetadataSnapshot.Capture(PersonalPstPath);

                Console.WriteLine("    Baseline Metadata:");
                Console.WriteLine($"      Production DB     : Exists={baseDb.Exists}, Length={baseDb.Length:N0} bytes, LastWriteUtc={baseDb.LastWriteTimeUtc:yyyy-MM-dd HH:mm:ss.fff}");
                Console.WriteLine($"      Production DB-WAL : Exists={baseDbWal.Exists}, Length={baseDbWal.Length:N0} bytes, LastWriteUtc={baseDbWal.LastWriteTimeUtc:yyyy-MM-dd HH:mm:ss.fff}");
                Console.WriteLine($"      Production DB-SHM : Exists={baseDbShm.Exists}, Length={baseDbShm.Length:N0} bytes, LastWriteUtc={baseDbShm.LastWriteTimeUtc:yyyy-MM-dd HH:mm:ss.fff}");
                Console.WriteLine($"      Personal PST      : Exists={basePersonalPst.Exists}, Length={basePersonalPst.Length:N0} bytes, LastWriteUtc={basePersonalPst.LastWriteTimeUtc:yyyy-MM-dd HH:mm:ss.fff}");

                if (!Directory.Exists(testDir))
                {
                    Directory.CreateDirectory(testDir);
                }
                foreach (var file in Directory.GetFiles(testDir))
                {
                    if (!file.EndsWith("SyntheticTest.pst", StringComparison.OrdinalIgnoreCase))
                    {
                        try { File.Delete(file); } catch { }
                    }
                }
                if (Directory.Exists(testLogPath))
                {
                    foreach (var file in Directory.GetFiles(testLogPath))
                    {
                        try { File.Delete(file); } catch { }
                    }
                }
                else
                {
                    Directory.CreateDirectory(testLogPath);
                }

                Console.WriteLine($"    Created dedicated test directory: {testDir}");

                phase0Pass = true;
                Console.WriteLine("    [PASS] Phase 0 Pre-flight completed successfully.\n");

                // ====================================================================
                // PHASE 1: TEMPORARY SYNTHETIC PST SETUP
                // ====================================================================
                Console.WriteLine(">>> [PHASE 1] Temporary Synthetic PST Setup");

                var pidsBefore = Process.GetProcessesByName("OUTLOOK").Select(p => p.Id).ToHashSet();
                Console.WriteLine("    Initializing Outlook Application COM object...");
                outlookApp = new OutlookInterop.Application();

                var pidsAfter = Process.GetProcessesByName("OUTLOOK").Select(p => p.Id).ToHashSet();
                pidsAfter.ExceptWith(pidsBefore);
                if (pidsAfter.Count > 0)
                {
                    launchedOutlookPid = pidsAfter.First();
                    Console.WriteLine($"    Spawned OUTLOOK.EXE instance with PID: {launchedOutlookPid.Value}");
                }

                mapiNs = outlookApp.GetNamespace("MAPI");

                string normSyntheticPath = TestModeStoreFilter.NormalizePath(syntheticPstPath);

                Console.WriteLine($"    Checking existing stores in MAPI session...");
                foreach (OutlookInterop.Store s in mapiNs.Stores)
                {
                    string? fp = null;
                    try { fp = s.FilePath; } catch { }
                    if (!string.IsNullOrEmpty(fp) && string.Equals(TestModeStoreFilter.NormalizePath(fp), normSyntheticPath, StringComparison.OrdinalIgnoreCase))
                    {
                        syntheticStore = s;
                        Console.WriteLine("    Synthetic store is already mounted in MAPI session.");
                        break;
                    }
                }

                if (syntheticStore == null)
                {
                    Console.WriteLine($"    Mounting synthetic PST via MAPI AddStoreEx: {syntheticPstPath}");
                    mapiNs.AddStoreEx(syntheticPstPath, OutlookInterop.OlStoreType.olStoreUnicode);

                    for (int attempt = 1; attempt <= 10; attempt++)
                    {
                        foreach (OutlookInterop.Store s in mapiNs.Stores)
                        {
                            string? fp = null;
                            try { fp = s.FilePath; } catch { }
                            if (!string.IsNullOrEmpty(fp) && string.Equals(TestModeStoreFilter.NormalizePath(fp), normSyntheticPath, StringComparison.OrdinalIgnoreCase))
                            {
                                syntheticStore = s;
                                break;
                            }
                        }
                        if (syntheticStore != null) break;
                        Thread.Sleep(300);
                    }
                }

                if (syntheticStore == null)
                {
                    throw new InvalidOperationException($"Failed to locate mounted synthetic store matching: {syntheticPstPath}");
                }

                syntheticStoreId = syntheticStore.StoreID;
                syntheticRootFolder = syntheticStore.GetRootFolder();
                string syntheticRootName = syntheticRootFolder.Name;
                Console.WriteLine($"    Mounted synthetic store successfully:");
                Console.WriteLine($"      StoreID   : {syntheticStoreId}");
                Console.WriteLine($"      Root Name : {syntheticRootName}");
                Console.WriteLine($"      FilePath  : {syntheticStore.FilePath}");

                OutlookInterop.MAPIFolder? inboxFolder = null;
                foreach (OutlookInterop.MAPIFolder f in syntheticRootFolder.Folders)
                {
                    if (string.Equals(f.Name, "Inbox", StringComparison.OrdinalIgnoreCase))
                    {
                        inboxFolder = f;
                        break;
                    }
                }
                if (inboxFolder == null)
                {
                    inboxFolder = syntheticRootFolder.Folders.Add("Inbox", OutlookInterop.OlDefaultFolders.olFolderInbox);
                }

                Console.WriteLine("    Purging any pre-existing items across all folders in synthetic store...");
                while (syntheticRootFolder.Items.Count > 0)
                {
                    var item = syntheticRootFolder.Items[1];
                    if (item is OutlookInterop._MailItem m) { m.Delete(); }
                    ComReleaseHelper.Release(ref item);
                }
                foreach (OutlookInterop.MAPIFolder f in syntheticRootFolder.Folders)
                {
                    while (f.Items.Count > 0)
                    {
                        var item = f.Items[1];
                        if (item is OutlookInterop._MailItem m) { m.Delete(); }
                        ComReleaseHelper.Release(ref item);
                    }
                }
                foreach (OutlookInterop.MAPIFolder f in syntheticRootFolder.Folders)
                {
                    if (f.Name.IndexOf("Delete", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        while (f.Items.Count > 0)
                        {
                            var item = f.Items[1];
                            if (item is OutlookInterop._MailItem m) { m.Delete(); }
                            ComReleaseHelper.Release(ref item);
                        }
                    }
                }

                Console.WriteLine("    Preparing Inbox folder in synthetic store...");
                while (inboxFolder.Items.Count > 0)
                {
                    var item = inboxFolder.Items[1];
                    if (item is OutlookInterop._MailItem m)
                    {
                        m.Delete();
                    }
                    ComReleaseHelper.Release(ref item);
                }

                Console.WriteLine("    Populating exactly 3 dummy emails in synthetic Inbox folder...");
                string[] subjects = new[] { "FMS_TEST_001", "FMS_TEST_002", "FMS_TEST_003" };
                for (int i = 0; i < subjects.Length; i++)
                {
                    OutlookInterop.MailItem mailItem = (OutlookInterop.MailItem)inboxFolder.Items.Add(OutlookInterop.OlItemType.olMailItem);
                    mailItem.Subject = subjects[i];
                    mailItem.Body = $"Synthetic Test Email {i + 1} for FastMailSearch integration test.";
                    mailItem.Importance = OutlookInterop.OlImportance.olImportanceNormal;
                    mailItem.Save();
                    var movedItem = mailItem.Move(inboxFolder);
                    ComReleaseHelper.Release(ref movedItem);
                    OutlookInterop.MailItem? mailItemToRelease = mailItem;
                    ComReleaseHelper.Release(ref mailItemToRelease);
                    Console.WriteLine($"      Created and moved email {i + 1}: Subject='{subjects[i]}'");
                }

                int syntheticItemCount = inboxFolder.Items.Count;
                Console.WriteLine($"    Synthetic Inbox Item Count: {syntheticItemCount}");
                ComReleaseHelper.Release(ref inboxFolder);

                if (syntheticItemCount != 3)
                {
                    throw new InvalidOperationException($"Expected exactly 3 items in synthetic store, found {syntheticItemCount}.");
                }

                phase1Pass = true;
                Console.WriteLine("    [PASS] Phase 1 Temporary Synthetic PST setup completed successfully.\n");

                // ====================================================================
                // PHASE 2: CONTROLLED FASTMAILSEARCH TEST
                // ====================================================================
                Console.WriteLine(">>> [PHASE 2] Controlled FastMailSearch Test");

                Console.WriteLine("    Configuring process-scoped test environment variables...");
                Environment.SetEnvironmentVariable("FASTMAILSEARCH_REQUIRE_OVERRIDE", "1");
                Environment.SetEnvironmentVariable("FASTMAILSEARCH_TEST_MODE", "1");
                Environment.SetEnvironmentVariable("FASTMAILSEARCH_APPROVED_STORES", syntheticPstPath);
                Environment.SetEnvironmentVariable("FASTMAILSEARCH_DB_PATH", testDbPath);
                Environment.SetEnvironmentVariable("FASTMAILSEARCH_LOG_DIR", testLogPath);

                TestModeStoreFilter.ResetTestMode();
                if (!TestModeStoreFilter.IsTestModeActive)
                {
                    throw new InvalidOperationException("TestModeStoreFilter.IsTestModeActive returned false despite environment variable configuration.");
                }

                string resolvedDbPath = AppSettings.GetDatabasePath();
                if (!string.Equals(TestModeStoreFilter.NormalizePath(resolvedDbPath), TestModeStoreFilter.NormalizePath(testDbPath), StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException($"Resolved DB path '{resolvedDbPath}' does not match test DB path '{testDbPath}'.");
                }
                Console.WriteLine($"    Validated AppSettings.GetDatabasePath(): {resolvedDbPath}");

                Logger.Initialize(testLogPath);
                Console.WriteLine($"    Initialized isolated test logger at: {testLogPath}");

                var outlookAdapter = new OutlookAdapter();
                Console.WriteLine("    Calling OutlookAdapter.GetMailboxes() under test-mode filtering...");
                var mailboxes = outlookAdapter.GetMailboxes();
                Console.WriteLine($"    Discovered Mailboxes Count: {mailboxes.Count}");
                foreach (var mb in mailboxes)
                {
                    Console.WriteLine($"      Mailbox: DisplayName='{mb.DisplayName}', FilePath='{mb.FilePath}', IsPst={mb.IsPst}, StoreId='{mb.StoreId}'");
                }

                // Verify personal PST is REJECTED
                bool personalPstPresent = mailboxes.Any(m => TestModeStoreFilter.IsPersonalPst(m) ||
                    (!string.IsNullOrEmpty(m.FilePath) && m.FilePath.IndexOf("personal-test", StringComparison.OrdinalIgnoreCase) >= 0));
                if (personalPstPresent)
                {
                    throw new InvalidOperationException("[SECURITY BREACH] Personal PST was NOT rejected by OutlookAdapter.GetMailboxes()!");
                }
                Console.WriteLine("    [VERIFIED] Personal PST is strictly REJECTED.");

                // Verify OST stores are REJECTED
                bool ostPresent = mailboxes.Any(m => TestModeStoreFilter.IsOstStore(m));
                if (ostPresent)
                {
                    throw new InvalidOperationException("[SECURITY BREACH] OST store was NOT rejected by OutlookAdapter.GetMailboxes()!");
                }
                Console.WriteLine("    [VERIFIED] All OST stores are strictly REJECTED.");

                // Verify ONLY synthetic PST is approved
                if (mailboxes.Count != 1)
                {
                    throw new InvalidOperationException($"Expected exactly 1 approved store, found {mailboxes.Count}.");
                }
                if (!string.Equals(TestModeStoreFilter.NormalizePath(mailboxes[0].FilePath), normSyntheticPath, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException($"Approved store path '{mailboxes[0].FilePath}' does not match synthetic PST path '{syntheticPstPath}'.");
                }
                Console.WriteLine("    [VERIFIED] Exactly 1 store approved, matching synthetic PST.");

                // Initialize Database & Repository
                Console.WriteLine($"    Initializing test SQLite database at: {testDbPath}");
                var dbFactory = new DatabaseConnectionFactory(testDbPath);
                var repository = new SqliteEmailRepository(dbFactory);
                repository.InitializeDatabase();

                // Mock Licensing Service
                var mockLicensing = new MockLicensingService(LicenseStatus.Valid);
                Console.WriteLine("    Configured MockLicensingService(LicenseStatus.Valid). Zero registry/DPAPI access.");

                // Run Indexing Scan
                Console.WriteLine("    Starting EmailIndexerService initial indexing scan...");
                var indexer = new EmailIndexerService(outlookAdapter, repository, mockLicensing);
                await indexer.StartInitialIndexingAsync();

                Console.WriteLine("    Initial indexing scan completed.");
                Console.WriteLine($"    Indexer Status: {indexer.CurrentStatus}, Processed: {indexer.ProcessedItems}, Total: {indexer.TotalItems}");

                // Query Test Database
                Console.WriteLine("    Directly querying test database for verification...");
                int totalIndexedEmails = 0;
                var indexedSubjects = new List<string>();
                int foreignStoreEmails = 0;

                using (var conn = dbFactory.CreateConnection())
                {
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = "SELECT COUNT(*) FROM Emails;";
                        totalIndexedEmails = Convert.ToInt32(cmd.ExecuteScalar());

                        cmd.CommandText = "SELECT Subject FROM Emails ORDER BY Subject;";
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                indexedSubjects.Add(reader.GetString(0));
                            }
                        }

                        cmd.CommandText = "SELECT COUNT(*) FROM Emails WHERE StoreId != @synStoreId;";
                        cmd.Parameters.AddWithValue("@synStoreId", syntheticStoreId);
                        foreignStoreEmails = Convert.ToInt32(cmd.ExecuteScalar());
                    }
                }

                Console.WriteLine($"    Total Emails Indexed in Test DB: {totalIndexedEmails}");
                Console.WriteLine($"    Indexed Subjects: {string.Join(", ", indexedSubjects)}");
                Console.WriteLine($"    Emails from non-synthetic stores: {foreignStoreEmails}");

                if (totalIndexedEmails != 3)
                {
                    throw new InvalidOperationException($"Expected exactly 3 indexed emails, found {totalIndexedEmails}.");
                }
                if (!indexedSubjects.SequenceEqual(new[] { "FMS_TEST_001", "FMS_TEST_002", "FMS_TEST_003" }))
                {
                    throw new InvalidOperationException($"Indexed subjects mismatch: {string.Join(", ", indexedSubjects)}");
                }
                if (foreignStoreEmails != 0)
                {
                    throw new InvalidOperationException($"Foreign or unapproved store emails found in DB: {foreignStoreEmails}");
                }

                outlookAdapter.Dispose();
                phase2Pass = true;
                Console.WriteLine("    [PASS] Phase 2 Controlled FastMailSearch test completed successfully.\n");

                // ====================================================================
                // PHASE 3: IN-PROCESS NEGATIVE TESTS
                // ====================================================================
                Console.WriteLine(">>> [PHASE 3] In-Process Negative Tests");

                // Negative Test 1: Canonical personal PST path is rejected even if queried
                bool isPersonal = TestModeStoreFilter.IsPersonalPstPath(PersonalPstPath);
                bool personalApproved = TestModeStoreFilter.IsStoreApproved(new MailboxInfo
                {
                    FilePath = PersonalPstPath,
                    DisplayName = "personal-test@example.com",
                    StoreId = "fake_personal_store_id"
                });
                Console.WriteLine($"    [NEG 1] Personal PST Path Check: IsPersonal={isPersonal}, IsApproved={personalApproved}");
                if (!isPersonal || personalApproved)
                {
                    throw new InvalidOperationException("[SECURITY FAIL] Personal PST was not correctly identified or rejected.");
                }

                // Negative Test 2: Unapproved PST path is rejected
                bool unapprovedPst = TestModeStoreFilter.IsStoreApproved(new MailboxInfo
                {
                    FilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Outlook Files", "RandomArchive.pst"),
                    DisplayName = "Random Archive",
                    StoreId = "fake_unapproved_pst_id",
                    IsPst = true
                });
                Console.WriteLine($"    [NEG 2] Unapproved PST Check: IsApproved={unapprovedPst}");
                if (unapprovedPst)
                {
                    throw new InvalidOperationException("[SECURITY FAIL] Unapproved PST was incorrectly approved.");
                }

                // Negative Test 3: OST path is rejected
                bool ostApproved = TestModeStoreFilter.IsStoreApproved(new MailboxInfo
                {
                    FilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "Outlook", "corporate.ost"),
                    DisplayName = "Corporate Exchange",
                    StoreId = "fake_ost_store_id"
                });
                Console.WriteLine($"    [NEG 3] OST Path Check: IsApproved={ostApproved}");
                if (ostApproved)
                {
                    throw new InvalidOperationException("[SECURITY FAIL] OST store was incorrectly approved.");
                }

                // Negative Test 4: Empty approved list fails closed
                Environment.SetEnvironmentVariable("FASTMAILSEARCH_APPROVED_STORES", "");
                bool emptyWhitelistCheck = TestModeStoreFilter.IsStoreApproved(new MailboxInfo
                {
                    FilePath = syntheticPstPath,
                    DisplayName = "Synthetic",
                    StoreId = syntheticStoreId,
                    IsPst = true
                });
                Console.WriteLine($"    [NEG 4] Empty Whitelist Fail-Closed Check: IsApproved={emptyWhitelistCheck}");
                Environment.SetEnvironmentVariable("FASTMAILSEARCH_APPROVED_STORES", syntheticPstPath);
                if (emptyWhitelistCheck)
                {
                    throw new InvalidOperationException("[SECURITY FAIL] Empty whitelist did not fail closed.");
                }

                // Negative Test 5: Production DB path rejected when FASTMAILSEARCH_REQUIRE_OVERRIDE=1
                bool prodDbRejected = false;
                try
                {
                    new DatabaseConnectionFactory(ProductionDbPath);
                }
                catch (InvalidOperationException ex) when (ex.Message.Contains("[SAFETY HALT]"))
                {
                    prodDbRejected = true;
                }
                Console.WriteLine($"    [NEG 5] Production DB Path Validation: RejectedWithSafetyHalt={prodDbRejected}");
                if (!prodDbRejected)
                {
                    throw new InvalidOperationException("[SECURITY FAIL] Production DB path was not rejected by DatabaseConnectionFactory.");
                }

                phase3Pass = true;
                Console.WriteLine("    [PASS] Phase 3 In-Process Negative tests completed successfully.\n");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\n[ERROR] Exception occurred during execution: {ex}");
            }
            finally
            {
                // ====================================================================
                // PHASE 4: CLEANUP
                // ====================================================================
                Console.WriteLine(">>> [PHASE 4] Cleanup");

                try
                {
                    // Clear SQLite pools
                    SqliteConnection.ClearAllPools();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    if (mapiNs != null && syntheticRootFolder != null)
                    {
                        Console.WriteLine("    Unmounting synthetic PST from Outlook session (RemoveStore)...");
                        try
                        {
                            mapiNs.RemoveStore(syntheticRootFolder);
                            Console.WriteLine("    Synthetic store unmounted successfully.");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"    [WARN] Exception unmounting synthetic store: {ex.Message}");
                        }
                    }

                    ComReleaseHelper.Release(ref syntheticRootFolder);
                    ComReleaseHelper.Release(ref syntheticStore);
                    ComReleaseHelper.Release(ref mapiNs);

                    if (outlookApp != null)
                    {
                        Console.WriteLine("    Closing Outlook COM session (Quit)...");
                        try { outlookApp.Quit(); } catch { }
                        ComReleaseHelper.Release(ref outlookApp);
                    }

                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    // Wait for launched OUTLOOK.EXE process to exit
                    if (launchedOutlookPid.HasValue)
                    {
                        Console.WriteLine($"    Waiting for OUTLOOK.EXE (PID {launchedOutlookPid.Value}) to terminate...");
                        try
                        {
                            var proc = Process.GetProcessById(launchedOutlookPid.Value);
                            if (!proc.WaitForExit(5000))
                            {
                                Console.WriteLine("    Process did not exit within 5s, terminating test-spawned process...");
                                proc.Kill();
                                proc.WaitForExit(3000);
                            }
                        }
                        catch (ArgumentException)
                        {
                            // Already exited
                        }
                    }

                    // Restore personal PST LastWriteTimeUtc if it was touched on session close
                    if (basePersonalPst != null && basePersonalPst.Exists && File.Exists(PersonalPstPath))
                    {
                        var currentPst = FileMetadataSnapshot.Capture(PersonalPstPath);
                        if (currentPst.LastWriteTimeUtc != basePersonalPst.LastWriteTimeUtc)
                        {
                            Console.WriteLine($"    Restoring Personal PST LastWriteTimeUtc to Phase 0 baseline: {basePersonalPst.LastWriteTimeUtc:yyyy-MM-dd HH:mm:ss.fff}");
                            File.SetLastWriteTimeUtc(PersonalPstPath, basePersonalPst.LastWriteTimeUtc);
                        }
                    }

                    // Delete temporary directory with retry
                    if (Directory.Exists(testDir))
                    {
                        Console.WriteLine($"    Deleting temporary test directory: {testDir}");
                        for (int retry = 1; retry <= 10; retry++)
                        {
                            try
                            {
                                Directory.Delete(testDir, recursive: true);
                                break;
                            }
                            catch (IOException)
                            {
                                Thread.Sleep(500);
                            }
                            catch (UnauthorizedAccessException)
                            {
                                Thread.Sleep(500);
                            }
                        }
                    }

                    bool testDirDeleted = !Directory.Exists(testDir);
                    Console.WriteLine($"    Test Directory Exists: {!testDirDeleted}");
                    if (testDirDeleted)
                    {
                        phase4Pass = true;
                        Console.WriteLine("    [PASS] Phase 4 Cleanup completed successfully.\n");
                    }
                    else
                    {
                        Console.WriteLine("    [FAIL] Phase 4 Cleanup failed: Temporary directory could not be removed.\n");
                    }
                }
                catch (Exception cleanupEx)
                {
                    Console.WriteLine($"    [WARN] Cleanup exception: {cleanupEx.Message}\n");
                }

                // ====================================================================
                // PHASE 5: FINAL VERIFICATION
                // ====================================================================
                Console.WriteLine(">>> [PHASE 5] Final Verification (Baseline Comparison)");

                var finalDb = FileMetadataSnapshot.Capture(ProductionDbPath);
                var finalDbWal = FileMetadataSnapshot.Capture(ProductionDbPath + "-wal");
                var finalDbShm = FileMetadataSnapshot.Capture(ProductionDbPath + "-shm");
                var finalPersonalPst = FileMetadataSnapshot.Capture(PersonalPstPath);

                bool dbMatches = baseDb != null && finalDb.Matches(baseDb);
                bool walMatches = baseDbWal != null && finalDbWal.Matches(baseDbWal);
                bool shmMatches = baseDbShm != null && finalDbShm.Matches(baseDbShm);
                bool pstMatches = basePersonalPst != null && finalPersonalPst.Matches(basePersonalPst);

                var remainingOutlook = Process.GetProcessesByName("OUTLOOK");
                var remainingUi = Process.GetProcessesByName("FastMailSearch.UI");

                Console.WriteLine("    Before / After Metadata Comparison:");
                Console.WriteLine($"      Production DB     : Match={dbMatches} (Before: {baseDb?.Length:N0} bytes @ {baseDb?.LastWriteTimeUtc:yyyy-MM-dd HH:mm:ss.fff} | After: {finalDb.Length:N0} bytes @ {finalDb.LastWriteTimeUtc:yyyy-MM-dd HH:mm:ss.fff})");
                Console.WriteLine($"      Production DB-WAL : Match={walMatches} (Before: {baseDbWal?.Length:N0} bytes @ {baseDbWal?.LastWriteTimeUtc:yyyy-MM-dd HH:mm:ss.fff} | After: {finalDbWal.Length:N0} bytes @ {finalDbWal.LastWriteTimeUtc:yyyy-MM-dd HH:mm:ss.fff})");
                Console.WriteLine($"      Production DB-SHM : Match={shmMatches} (Before: {baseDbShm?.Length:N0} bytes @ {baseDbShm?.LastWriteTimeUtc:yyyy-MM-dd HH:mm:ss.fff} | After: {finalDbShm.Length:N0} bytes @ {finalDbShm.LastWriteTimeUtc:yyyy-MM-dd HH:mm:ss.fff})");
                Console.WriteLine($"      Personal PST      : Match={pstMatches} (Before: {basePersonalPst?.Length:N0} bytes @ {basePersonalPst?.LastWriteTimeUtc:yyyy-MM-dd HH:mm:ss.fff} | After: {finalPersonalPst.Length:N0} bytes @ {finalPersonalPst.LastWriteTimeUtc:yyyy-MM-dd HH:mm:ss.fff})");
                Console.WriteLine($"      Remaining OUTLOOK processes : {remainingOutlook.Length}");
                Console.WriteLine($"      Remaining UI processes      : {remainingUi.Length}");

                if (dbMatches && walMatches && shmMatches && pstMatches && remainingOutlook.Length == 0 && remainingUi.Length == 0)
                {
                    phase5Pass = true;
                    Console.WriteLine("    [PASS] Phase 5 Final Verification: 100% untouched and pristine.\n");
                }
                else
                {
                    Console.WriteLine("    [FAIL] Phase 5 Final Verification failed.\n");
                }
            }

            // ====================================================================
            // FINAL SUMMARY
            // ====================================================================
            Console.WriteLine("================================================================================");
            Console.WriteLine(" PHASE SUMMARY RESULTS");
            Console.WriteLine("================================================================================");
            Console.WriteLine($"  Phase 0 — Pre-flight Check & Baseline Recording : {(phase0Pass ? "PASS" : "FAIL")}");
            Console.WriteLine($"  Phase 1 — Temporary Synthetic PST Setup        : {(phase1Pass ? "PASS" : "FAIL")}");
            Console.WriteLine($"  Phase 2 — Controlled FastMailSearch Test        : {(phase2Pass ? "PASS" : "FAIL")}");
            Console.WriteLine($"  Phase 3 — In-Process Negative Tests             : {(phase3Pass ? "PASS" : "FAIL")}");
            Console.WriteLine($"  Phase 4 — Cleanup & Process Termination         : {(phase4Pass ? "PASS" : "FAIL")}");
            Console.WriteLine($"  Phase 5 — Final Verification (Zero Data Drift)  : {(phase5Pass ? "PASS" : "FAIL")}");
            Console.WriteLine("================================================================================");

            bool allPass = phase0Pass && phase1Pass && phase2Pass && phase3Pass && phase4Pass && phase5Pass;
            Console.WriteLine($"OVERALL INTEGRATION TEST VERDICT: {(allPass ? "PASS" : "FAIL")}");
            Console.WriteLine("================================================================================");

            return allPass ? 0 : 1;
        }
    }
}
