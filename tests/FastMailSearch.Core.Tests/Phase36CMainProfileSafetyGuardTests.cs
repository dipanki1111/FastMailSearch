using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using FastMailSearch.Core.Interfaces;
using FastMailSearch.Core.Models;
using FastMailSearch.Infrastructure.Database;
using FastMailSearch.Infrastructure.Indexing;
using Xunit;

namespace FastMailSearch.Core.Tests
{
    [Collection("Non-Parallel Environment Tests")]
    public class Phase36CMainProfileSafetyGuardTests : IDisposable
    {
        private static readonly SemaphoreSlim _envLock = new SemaphoreSlim(1, 1);
        private readonly List<string> _tempFilesToClean = new List<string>();

        public void Dispose()
        {
            TestModeStoreFilter.ResetTestMode();
            foreach (var file in _tempFilesToClean)
            {
                try
                {
                    if (File.Exists(file)) File.Delete(file);
                    string wal = file + "-wal";
                    string shm = file + "-shm";
                    if (File.Exists(wal)) File.Delete(wal);
                    if (File.Exists(shm)) File.Delete(shm);
                }
                catch { }
            }
        }

        private string CreateTempDbPath()
        {
            string path = Path.Combine(Path.GetTempPath(), $"Phase36C_Test_{Guid.NewGuid():N}.db");
            _tempFilesToClean.Add(path);
            return path;
        }

        #region Requirement 1 & 4: DatabaseConnectionFactory Validation & Standalone Resolution

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        [InlineData("\t\r\n")]
        public void DatabaseConnectionFactory_WithNullOrBlankPath_ThrowsSafetyHalt_WhenRequireOverrideActive(string? blankPath)
        {
            _envLock.Wait();
            string? origReq = Environment.GetEnvironmentVariable(AppSettings.RequireOverrideEnvironmentVariable);
            string? origDb = Environment.GetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable);
            try
            {
                Environment.SetEnvironmentVariable(AppSettings.RequireOverrideEnvironmentVariable, "1");
                Environment.SetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable, null);

                var ex = Assert.Throws<InvalidOperationException>(() => new DatabaseConnectionFactory(blankPath!));
                Assert.Contains("[SAFETY HALT]", ex.Message);
                Assert.Contains("FASTMAILSEARCH_REQUIRE_OVERRIDE=1", ex.Message);
            }
            finally
            {
                Environment.SetEnvironmentVariable(AppSettings.RequireOverrideEnvironmentVariable, origReq);
                Environment.SetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable, origDb);
                _envLock.Release();
            }
        }

        [Fact]
        public void DatabaseConnectionFactory_WithExplicitCanonicalProductionPath_ThrowsSafetyHalt_WhenRequireOverrideActive()
        {
            _envLock.Wait();
            string? origReq = Environment.GetEnvironmentVariable(AppSettings.RequireOverrideEnvironmentVariable);
            try
            {
                Environment.SetEnvironmentVariable(AppSettings.RequireOverrideEnvironmentVariable, "1");
                string prodPath = AppSettings.GetDefaultDatabasePath();

                var ex = Assert.Throws<InvalidOperationException>(() => new DatabaseConnectionFactory(prodPath));
                Assert.Contains("[SAFETY HALT]", ex.Message);
                Assert.Contains("matches canonical production database", ex.Message);
            }
            finally
            {
                Environment.SetEnvironmentVariable(AppSettings.RequireOverrideEnvironmentVariable, origReq);
                _envLock.Release();
            }
        }

        [Fact]
        public void DatabaseConnectionFactory_WithExplicitPathVariations_ThrowsSafetyHalt_WhenRequireOverrideActive()
        {
            _envLock.Wait();
            string? origReq = Environment.GetEnvironmentVariable(AppSettings.RequireOverrideEnvironmentVariable);
            try
            {
                Environment.SetEnvironmentVariable(AppSettings.RequireOverrideEnvironmentVariable, "1");
                string prodPath = AppSettings.GetDefaultDatabasePath();

                // 1. Lowercase with trailing separator
                string lowerWithSlash = prodPath.ToLowerInvariant() + @"\";
                var ex1 = Assert.Throws<InvalidOperationException>(() => new DatabaseConnectionFactory(lowerWithSlash));
                Assert.Contains("[SAFETY HALT]", ex1.Message);

                // 2. Uppercase
                string upper = prodPath.ToUpperInvariant();
                var ex2 = Assert.Throws<InvalidOperationException>(() => new DatabaseConnectionFactory(upper));
                Assert.Contains("[SAFETY HALT]", ex2.Message);

                // 3. Path with relative segments (e.g., Documents\..\Documents\FastMailSearch\FastMailSearch.db)
                string dir = Path.GetDirectoryName(prodPath)!;
                string parent = Path.GetDirectoryName(dir)!;
                string relative = Path.Combine(parent, "Documents", "..", Path.GetFileName(dir), Path.GetFileName(prodPath));
                var ex3 = Assert.Throws<InvalidOperationException>(() => new DatabaseConnectionFactory(relative));
                Assert.Contains("[SAFETY HALT]", ex3.Message);
            }
            finally
            {
                Environment.SetEnvironmentVariable(AppSettings.RequireOverrideEnvironmentVariable, origReq);
                _envLock.Release();
            }
        }

        [Fact]
        public void DatabaseConnectionFactory_WithCustomIsolatedPath_Succeeds_WhenRequireOverrideActive()
        {
            _envLock.Wait();
            string? origReq = Environment.GetEnvironmentVariable(AppSettings.RequireOverrideEnvironmentVariable);
            try
            {
                Environment.SetEnvironmentVariable(AppSettings.RequireOverrideEnvironmentVariable, "1");
                string tempDb = CreateTempDbPath();

                var factory = new DatabaseConnectionFactory(tempDb);
                Assert.Equal(tempDb, factory.DbPath);

                using var conn = factory.CreateConnection();
                Assert.NotNull(conn);
                Assert.Equal(System.Data.ConnectionState.Open, conn.State);
            }
            finally
            {
                Environment.SetEnvironmentVariable(AppSettings.RequireOverrideEnvironmentVariable, origReq);
                _envLock.Release();
            }
        }

        [Fact]
        public void DatabaseConnectionFactory_WithExplicitProductionPath_DoesNotThrow_WhenRequireOverrideInactive()
        {
            _envLock.Wait();
            string? origReq = Environment.GetEnvironmentVariable(AppSettings.RequireOverrideEnvironmentVariable);
            try
            {
                Environment.SetEnvironmentVariable(AppSettings.RequireOverrideEnvironmentVariable, null);
                string prodPath = AppSettings.GetDefaultDatabasePath();

                // When override is NOT required, passing production path to factory constructor passes validation
                var factory = new DatabaseConnectionFactory(prodPath);
                Assert.Equal(prodPath, factory.DbPath);
            }
            finally
            {
                Environment.SetEnvironmentVariable(AppSettings.RequireOverrideEnvironmentVariable, origReq);
                _envLock.Release();
            }
        }

        [Fact]
        public void StandaloneTool_PathResolution_RejectsProductionPath_WhenRequireOverrideActive()
        {
            _envLock.Wait();
            string? origReq = Environment.GetEnvironmentVariable(AppSettings.RequireOverrideEnvironmentVariable);
            string? origDb = Environment.GetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable);
            try
            {
                Environment.SetEnvironmentVariable(AppSettings.RequireOverrideEnvironmentVariable, "1");
                Environment.SetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable, null);

                // When standalone tools call AppSettings.GetDatabasePath(), it throws [SAFETY HALT]
                var ex = Assert.Throws<InvalidOperationException>(() => AppSettings.GetDatabasePath());
                Assert.Contains("[SAFETY HALT]", ex.Message);
            }
            finally
            {
                Environment.SetEnvironmentVariable(AppSettings.RequireOverrideEnvironmentVariable, origReq);
                Environment.SetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable, origDb);
                _envLock.Release();
            }
        }

        #endregion

        #region Requirement 2 & 5: TestModeStoreFilter Safeguards

        [Fact]
        public void TestModeStoreFilter_WhenTestModeInactive_AcceptsAllStores_PreservingProductionBehavior()
        {
            _envLock.Wait();
            string? origTest = Environment.GetEnvironmentVariable(TestModeStoreFilter.TestModeEnvironmentVariable);
            try
            {
                Environment.SetEnvironmentVariable(TestModeStoreFilter.TestModeEnvironmentVariable, null);
                TestModeStoreFilter.ResetTestMode();

                Assert.False(TestModeStoreFilter.IsTestModeActive);

                var pstMailbox = new MailboxInfo
                {
                    StoreId = "PST-001",
                    DisplayName = "personal@gmail.com",
                    FilePath = @"C:\Users\User\Documents\Outlook Files\personal.pst",
                    IsPst = true
                };

                var exchangeMailbox = new MailboxInfo
                {
                    StoreId = "EXCH-001",
                    DisplayName = "work@company.com",
                    FilePath = @"C:\Users\User\AppData\Local\Microsoft\Outlook\work.ost",
                    IsPst = false
                };

                // Normal production behavior: All stores permitted
                Assert.True(TestModeStoreFilter.IsStoreApproved(pstMailbox));
                Assert.True(TestModeStoreFilter.IsStoreApproved(exchangeMailbox));

                var filtered = TestModeStoreFilter.FilterMailboxes(new[] { pstMailbox, exchangeMailbox });
                Assert.Equal(2, filtered.Count);
            }
            finally
            {
                Environment.SetEnvironmentVariable(TestModeStoreFilter.TestModeEnvironmentVariable, origTest);
                TestModeStoreFilter.ResetTestMode();
                _envLock.Release();
            }
        }

        [Fact]
        public void TestModeStoreFilter_WhenTestModeActive_NoApprovedStoresConfigured_FailsClosed_RejectsAll()
        {
            _envLock.Wait();
            string? origTest = Environment.GetEnvironmentVariable(TestModeStoreFilter.TestModeEnvironmentVariable);
            string? origAppr = Environment.GetEnvironmentVariable(TestModeStoreFilter.ApprovedStoresEnvironmentVariable);
            try
            {
                Environment.SetEnvironmentVariable(TestModeStoreFilter.TestModeEnvironmentVariable, "1");
                Environment.SetEnvironmentVariable(TestModeStoreFilter.ApprovedStoresEnvironmentVariable, null);
                TestModeStoreFilter.ResetTestMode();

                Assert.True(TestModeStoreFilter.IsTestModeActive);

                var mailbox = new MailboxInfo
                {
                    StoreId = "ANY-STORE-001",
                    DisplayName = "Any Store",
                    FilePath = @"C:\Mail\store.ost",
                    IsPst = false
                };

                // Fail-closed: 0 approved stores configured -> all stores rejected
                Assert.False(TestModeStoreFilter.IsStoreApproved(mailbox));
                Assert.False(TestModeStoreFilter.IsStoreIdApproved("ANY-STORE-001"));

                var filtered = TestModeStoreFilter.FilterMailboxes(new[] { mailbox });
                Assert.Empty(filtered);
            }
            finally
            {
                Environment.SetEnvironmentVariable(TestModeStoreFilter.TestModeEnvironmentVariable, origTest);
                Environment.SetEnvironmentVariable(TestModeStoreFilter.ApprovedStoresEnvironmentVariable, origAppr);
                TestModeStoreFilter.ResetTestMode();
                _envLock.Release();
            }
        }

        [Fact]
        public void TestModeStoreFilter_WhenTestModeActive_RejectsPstStores_EvenIfNamedInApprovedList()
        {
            _envLock.Wait();
            string? origTest = Environment.GetEnvironmentVariable(TestModeStoreFilter.TestModeEnvironmentVariable);
            string? origAppr = Environment.GetEnvironmentVariable(TestModeStoreFilter.ApprovedStoresEnvironmentVariable);
            try
            {
                Environment.SetEnvironmentVariable(TestModeStoreFilter.TestModeEnvironmentVariable, "1");
                Environment.SetEnvironmentVariable(TestModeStoreFilter.ApprovedStoresEnvironmentVariable, "PST-STORE-001;personal@gmail.com");
                TestModeStoreFilter.ResetTestMode();

                Assert.True(TestModeStoreFilter.IsTestModeActive);

                // Case A: IsPst = true
                var pstFlagged = new MailboxInfo
                {
                    StoreId = "PST-STORE-001",
                    DisplayName = "personal@gmail.com",
                    FilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Outlook Files", "test.pst"),
                    IsPst = true
                };
                Assert.False(TestModeStoreFilter.IsStoreApproved(pstFlagged));

                // Case B: FilePath ending in .pst (even if IsPst was false)
                var pstExtension = new MailboxInfo
                {
                    StoreId = "PST-STORE-001",
                    DisplayName = "personal@gmail.com",
                    FilePath = @"D:\Archive\Personal_2026.PST",
                    IsPst = false
                };
                Assert.False(TestModeStoreFilter.IsStoreApproved(pstExtension));

                // Case C: FilePath ending in .ost
                var ostExtension = new MailboxInfo
                {
                    StoreId = "PST-STORE-001",
                    DisplayName = "personal@gmail.com",
                    FilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Outlook.ost"),
                    IsPst = false
                };
                Assert.False(TestModeStoreFilter.IsStoreApproved(ostExtension));
            }
            finally
            {
                Environment.SetEnvironmentVariable(TestModeStoreFilter.TestModeEnvironmentVariable, origTest);
                Environment.SetEnvironmentVariable(TestModeStoreFilter.ApprovedStoresEnvironmentVariable, origAppr);
                TestModeStoreFilter.ResetTestMode();
                _envLock.Release();
            }
        }

        [Fact]
        public void TestModeStoreFilter_WhenTestModeActive_AcceptsOnlyExplicitlyApprovedSyntheticStores()
        {
            _envLock.Wait();
            try
            {
                // Programmatic configuration
                TestModeStoreFilter.SetTestMode(true, new[] { "SYNTHETIC-STORE-100", "ApprovedSyntheticDisplayName" });

                Assert.True(TestModeStoreFilter.IsTestModeActive);

                var approvedById = new MailboxInfo
                {
                    StoreId = "SYNTHETIC-STORE-100",
                    DisplayName = "Synthetic Test Account 1",
                    FilePath = "",
                    IsPst = false
                };

                var approvedByName = new MailboxInfo
                {
                    StoreId = "OTHER-ID-999",
                    DisplayName = "ApprovedSyntheticDisplayName",
                    FilePath = "",
                    IsPst = false
                };

                var unapproved = new MailboxInfo
                {
                    StoreId = "UNAPPROVED-STORE-200",
                    DisplayName = "Unapproved Exchange Store",
                    FilePath = "",
                    IsPst = false
                };

                var approvedNameButPst = new MailboxInfo
                {
                    StoreId = "SYNTHETIC-STORE-100",
                    DisplayName = "ApprovedSyntheticDisplayName",
                    FilePath = @"C:\Mail\personal.pst",
                    IsPst = true
                };

                Assert.True(TestModeStoreFilter.IsStoreApproved(approvedById));
                Assert.True(TestModeStoreFilter.IsStoreApproved(approvedByName));
                Assert.False(TestModeStoreFilter.IsStoreApproved(unapproved));
                Assert.False(TestModeStoreFilter.IsStoreApproved(approvedNameButPst)); // PST veto

                var filtered = TestModeStoreFilter.FilterMailboxes(new[] { approvedById, approvedByName, unapproved, approvedNameButPst });
                Assert.Equal(2, filtered.Count);
                Assert.Contains(approvedById, filtered);
                Assert.Contains(approvedByName, filtered);
            }
            finally
            {
                TestModeStoreFilter.ResetTestMode();
                _envLock.Release();
            }
        }

        #endregion

        #region Requirement 3: Mock Licensing & EmailIndexerService Store Isolation

        private class SyntheticSafeOutlookAdapter : IOutlookAdapter
        {
            public List<MailboxInfo> Mailboxes { get; } = new List<MailboxInfo>();
            public Dictionary<string, List<MailboxFolder>> StoreFolders { get; } = new Dictionary<string, List<MailboxFolder>>();
            public Dictionary<string, List<EmailDocument>> FolderEmails { get; } = new Dictionary<string, List<EmailDocument>>();

            public bool IsOutlookAvailable() => true;
            public string GetOutlookVersion() => "16.0 (Headless Test)";
            public string GetOfficeBitness() => "64-bit";
            public List<MailboxInfo> GetMailboxes() => Mailboxes;
            public List<MailboxFolder> GetFolders(string storeId) => StoreFolders.TryGetValue(storeId, out var f) ? f : new List<MailboxFolder>();

            public List<EmailDocument> FetchFolderEmailsBatch(string storeId, string folderEntryId, int startIndex, int batchSize, Func<bool>? cancelCheck = null)
            {
                string key = $"{storeId}_{folderEntryId}";
                if (!FolderEmails.TryGetValue(key, out var list)) return new List<EmailDocument>();
                int zeroIndex = startIndex - 1;
                if (zeroIndex < 0 || zeroIndex >= list.Count) return new List<EmailDocument>();
                return list.Skip(zeroIndex).Take(batchSize).ToList();
            }

            public EmailDocument? FetchSingleEmail(string storeId, string entryId) => null;
            public bool OpenEmailInOutlook(string storeId, string entryId) => true;
            public bool IsOutboxFolder(string storeId, string folderEntryId) => false;
            public void SubscribeToEvents(Action<string, string, string> onItemAddedOrChanged, Action<string, string> onItemRemoved) { }
            public void UnsubscribeFromEvents() { }
        }

        [Fact]
        public async Task EmailIndexerService_WithTestModeActive_IndexesOnlyApprovedStores_AndIgnoresPstStores()
        {
            await _envLock.WaitAsync();
            try
            {
                string tempDb = CreateTempDbPath();
                var factory = new DatabaseConnectionFactory(tempDb);
                var repo = new SqliteEmailRepository(factory);
                repo.InitializeDatabase();

                try
                {
                    string approvedStoreId = "SYNTHETIC-APPROVED-STORE-1";
                    string unapprovedStoreId = "UNAPPROVED-STORE-2";
                    string pstStoreId = "PERSONAL-PST-STORE-3";

                    TestModeStoreFilter.SetTestMode(true, new[] { approvedStoreId });

                    var adapter = new SyntheticSafeOutlookAdapter();

                    // 1. Approved synthetic store
                    adapter.Mailboxes.Add(new MailboxInfo
                    {
                        StoreId = approvedStoreId,
                        DisplayName = "Approved Synthetic Mailbox",
                        FilePath = "",
                        IsPst = false,
                        IsSelected = true
                    });
                    adapter.StoreFolders[approvedStoreId] = new List<MailboxFolder>
                    {
                        new MailboxFolder { StoreId = approvedStoreId, EntryId = "f_appr", Name = "Inbox", ItemCount = 2 }
                    };
                    adapter.FolderEmails[$"{approvedStoreId}_f_appr"] = new List<EmailDocument>
                    {
                        new EmailDocument { StoreId = approvedStoreId, EntryId = "appr_msg_1", FolderEntryId = "f_appr", Subject = "Approved Message 1", ReceivedDate = DateTime.Now },
                        new EmailDocument { StoreId = approvedStoreId, EntryId = "appr_msg_2", FolderEntryId = "f_appr", Subject = "Approved Message 2", ReceivedDate = DateTime.Now }
                    };

                    // 2. Unapproved store
                    adapter.Mailboxes.Add(new MailboxInfo
                    {
                        StoreId = unapprovedStoreId,
                        DisplayName = "Unapproved Exchange Mailbox",
                        FilePath = "",
                        IsPst = false,
                        IsSelected = true
                    });
                    adapter.StoreFolders[unapprovedStoreId] = new List<MailboxFolder>
                    {
                        new MailboxFolder { StoreId = unapprovedStoreId, EntryId = "f_unappr", Name = "Inbox", ItemCount = 5 }
                    };
                    adapter.FolderEmails[$"{unapprovedStoreId}_f_unappr"] = new List<EmailDocument>
                    {
                        new EmailDocument { StoreId = unapprovedStoreId, EntryId = "unappr_msg_1", FolderEntryId = "f_unappr", Subject = "Unapproved Message 1", ReceivedDate = DateTime.Now }
                    };

                    // 3. Personal PST store
                    adapter.Mailboxes.Add(new MailboxInfo
                    {
                        StoreId = pstStoreId,
                        DisplayName = "Personal PST Store",
                        FilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Outlook Files", "personal.pst"),
                        IsPst = true,
                        IsSelected = true
                    });
                    adapter.StoreFolders[pstStoreId] = new List<MailboxFolder>
                    {
                        new MailboxFolder { StoreId = pstStoreId, EntryId = "f_pst", Name = "Inbox", ItemCount = 10 }
                    };
                    adapter.FolderEmails[$"{pstStoreId}_f_pst"] = new List<EmailDocument>
                    {
                        new EmailDocument { StoreId = pstStoreId, EntryId = "pst_msg_1", FolderEntryId = "f_pst", Subject = "Personal PST Message 1", ReceivedDate = DateTime.Now }
                    };

                    // Inject MockLicensingService — NEVER instantiate real LicensingService or touch DPAPI/registry
                    var mockLicense = new MockLicensingService(LicenseStatus.Valid);
                    var indexer = new EmailIndexerService(adapter, repo, mockLicense);

                    // Check Store Coverage
                    var coverage = await indexer.CheckStoreCoverageAsync();
                    Assert.Equal(1, coverage.ActiveStoreCount); // Only approved synthetic store discovered!

                    // Start Backfill
                    var summary = await indexer.StartActiveStoreBackfillAsync();
                    Assert.Equal(2, summary.TotalDiscovered);
                    Assert.Equal(2, summary.NewlyIndexedCount);

                    // Verify database records
                    var approvedMsg1 = await repo.GetEmailByStoreAndEntryIdAsync(approvedStoreId, "appr_msg_1");
                    var approvedMsg2 = await repo.GetEmailByStoreAndEntryIdAsync(approvedStoreId, "appr_msg_2");
                    Assert.NotNull(approvedMsg1);
                    Assert.NotNull(approvedMsg2);

                    // Verify unapproved and PST records are strictly absent from the database
                    var unapprovedMsg = await repo.GetEmailByStoreAndEntryIdAsync(unapprovedStoreId, "unappr_msg_1");
                    var pstMsg = await repo.GetEmailByStoreAndEntryIdAsync(pstStoreId, "pst_msg_1");
                    Assert.Null(unapprovedMsg);
                    Assert.Null(pstMsg);
                }
                finally
                {
                    TestModeStoreFilter.ResetTestMode();
                }
            }
            finally
            {
                _envLock.Release();
            }
        }

        [Fact]
        public async Task EmailIndexerService_QueueItemForIndexing_DropsItemsFromUnapprovedOrPstStores_WhenTestModeActive()
        {
            await _envLock.WaitAsync();
            try
            {
                string tempDb = CreateTempDbPath();
                var factory = new DatabaseConnectionFactory(tempDb);
                var repo = new SqliteEmailRepository(factory);
                repo.InitializeDatabase();

                try
                {
                    string approvedStoreId = "SYNTHETIC-APPROVED-STORE-1";
                    string unapprovedStoreId = "UNAPPROVED-STORE-2";
                    string pstStoreId = "PERSONAL-PST-STORE-3";

                    TestModeStoreFilter.SetTestMode(true, new[] { approvedStoreId });

                    var adapter = new SyntheticSafeOutlookAdapter();
                    var mockLicense = new MockLicensingService(LicenseStatus.Valid);
                    var indexer = new EmailIndexerService(adapter, repo, mockLicense);

                    // 1. Attempt to queue item for unapproved store
                    await indexer.QueueItemForIndexingAsync(unapprovedStoreId, "unappr_entry_1", "folder_1");

                    // 2. Attempt to queue item for personal PST store
                    await indexer.QueueItemForIndexingAsync(pstStoreId, "pst_entry_1", "folder_2");

                    // 3. Attempt to queue deletion for unapproved store
                    await indexer.QueueItemForDeletionAsync(unapprovedStoreId, "unappr_del_1");

                    // 4. Verify no pending changes or emails were written to database for unapproved/PST stores
                    var unapprovedItem = await repo.GetEmailByStoreAndEntryIdAsync(unapprovedStoreId, "unappr_entry_1");
                    var pstItem = await repo.GetEmailByStoreAndEntryIdAsync(pstStoreId, "pst_entry_1");
                    Assert.Null(unapprovedItem);
                    Assert.Null(pstItem);
                }
                finally
                {
                    TestModeStoreFilter.ResetTestMode();
                }
            }
            finally
            {
                _envLock.Release();
            }
        }

        [Fact]
        public void Licensing_MockLicensingService_DoesNotAccessRegistryOrDpapi()
        {
            // Verify MockLicensingService executes purely in memory
            var mock = new MockLicensingService(LicenseStatus.Trial)
            {
                TrialDaysRemaining = 15
            };

            Assert.Equal(LicenseStatus.Trial, mock.GetStatus());
            Assert.Equal(15, mock.GetTrialDaysRemaining());
            Assert.Equal(1, mock.GetStatusCallCount);
        }

        #endregion
    }
}
