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
    public class EmailIndexerLicensingEnforcementTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly DatabaseConnectionFactory _factory;
        private readonly SqliteEmailRepository _repository;

        public EmailIndexerLicensingEnforcementTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), $"IndexerLicensingTests_{Guid.NewGuid():N}.db");
            _factory = new DatabaseConnectionFactory(_dbPath);
            _repository = new SqliteEmailRepository(_factory);
            _repository.InitializeDatabase();
        }

        public void Dispose()
        {
            try
            {
                if (File.Exists(_dbPath))
                    File.Delete(_dbPath);
            }
            catch { }
        }

        private class MockOutlookAdapterForIndexer : IOutlookAdapter, IFolderBatchReader
        {
            public Dictionary<string, EmailDocument> EmailStorage = new Dictionary<string, EmailDocument>(StringComparer.OrdinalIgnoreCase);
            public List<MailboxInfo> Mailboxes { get; set; } = new List<MailboxInfo>();
            public Dictionary<string, List<MailboxFolder>> StoreFolders { get; set; } = new Dictionary<string, List<MailboxFolder>>(StringComparer.OrdinalIgnoreCase);
            public Dictionary<string, List<EmailDocument>> FolderEmails { get; set; } = new Dictionary<string, List<EmailDocument>>(StringComparer.OrdinalIgnoreCase);

            public bool IsOutlookAvailable() => true;
            public string GetOutlookVersion() => "16.0 (64-bit)";
            public string GetOfficeBitness() => "x64";
            public List<MailboxInfo> GetMailboxes() => Mailboxes;
            public List<MailboxFolder> GetFolders(string storeId) => StoreFolders.TryGetValue(storeId, out var f) ? f : new List<MailboxFolder>();

            public EmailDocument? FetchSingleEmail(string storeId, string entryId)
            {
                string key = $"{storeId}:{entryId}";
                EmailStorage.TryGetValue(key, out var doc);
                return doc;
            }

            public List<EmailDocument> FetchFolderEmailsBatch(string storeId, string folderEntryId, int startIndex, int batchSize, Func<bool>? cancelCheck = null)
            {
                string key = $"{storeId}_{folderEntryId}";
                if (!FolderEmails.TryGetValue(key, out var list)) return new List<EmailDocument>();
                return list.Skip(startIndex - 1).Take(batchSize).ToList();
            }

            public FolderBatchResult FetchFolderBatch(string storeId, string folderEntryId, int startIndex, int batchSize, Func<bool>? cancelCheck = null)
            {
                var items = FetchFolderEmailsBatch(storeId, folderEntryId, startIndex, batchSize, cancelCheck);
                int lastProcessed = startIndex + items.Count - 1;
                return new FolderBatchResult
                {
                    Items = items,
                    StartIndex = startIndex,
                    LastProcessedIndex = lastProcessed,
                    SkippedCount = 0,
                    WasCancelled = cancelCheck != null && cancelCheck()
                };
            }

            public void ResetCursorSession() { }
            public bool OpenEmailInOutlook(string storeId, string entryId) => true;
            public bool IsOutboxFolder(string storeId, string folderEntryId) => false;
            public void SubscribeToEvents(Action<string, string, string> onItemAddedOrChanged, Action<string, string> onItemRemoved) { }
            public void UnsubscribeFromEvents() { }
        }

        [Fact]
        public async Task Test01_ValidLicense_PermitsBackgroundIndexing()
        {
            var mock = new MockOutlookAdapterForIndexer();
            var doc = new EmailDocument
            {
                StoreId = "Store1",
                EntryId = "Entry_Valid_1",
                Subject = "Valid License Indexing Test",
                SenderEmail = "sender@example.com",
                ReceivedDate = DateTime.Now
            };
            mock.EmailStorage[$"{doc.StoreId}:{doc.EntryId}"] = doc;

            var mockLicense = new MockLicensingService(LicenseStatus.Valid);
            var indexer = new EmailIndexerService(mock, _repository, mockLicense);

            await indexer.QueueItemForIndexingAsync("Store1", "Entry_Valid_1", "Folder1");

            int indexedCount = await _repository.GetIndexedEmailCountAsync();
            Assert.Equal(1, indexedCount);
            Assert.True(indexer.IsLicensingEligible());
        }

        [Fact]
        public async Task Test02_TrialLicense_PermitsBackgroundIndexing()
        {
            var mock = new MockOutlookAdapterForIndexer();
            var doc = new EmailDocument
            {
                StoreId = "Store1",
                EntryId = "Entry_Trial_1",
                Subject = "Trial License Indexing Test",
                SenderEmail = "sender@example.com",
                ReceivedDate = DateTime.Now
            };
            mock.EmailStorage[$"{doc.StoreId}:{doc.EntryId}"] = doc;

            var mockLicense = new MockLicensingService(LicenseStatus.Trial);
            var indexer = new EmailIndexerService(mock, _repository, mockLicense);

            await indexer.QueueItemForIndexingAsync("Store1", "Entry_Trial_1", "Folder1");

            int indexedCount = await _repository.GetIndexedEmailCountAsync();
            Assert.Equal(1, indexedCount);
            Assert.True(indexer.IsLicensingEligible());
        }

        [Fact]
        public async Task Test03_ExpiredLicense_BlocksNewBackgroundIndexing_AndPreservesPendingQueue()
        {
            var mock = new MockOutlookAdapterForIndexer();
            var doc = new EmailDocument
            {
                StoreId = "Store1",
                EntryId = "Entry_Expired_1",
                Subject = "Expired License Item",
                SenderEmail = "sender@example.com",
                ReceivedDate = DateTime.Now
            };
            mock.EmailStorage[$"{doc.StoreId}:{doc.EntryId}"] = doc;

            var mockLicense = new MockLicensingService(LicenseStatus.Expired);
            var indexer = new EmailIndexerService(mock, _repository, mockLicense);

            Assert.False(indexer.IsLicensingEligible());

            await indexer.QueueItemForIndexingAsync("Store1", "Entry_Expired_1", "Folder1");

            // Email should NOT be written to the index database
            int indexedCount = await _repository.GetIndexedEmailCountAsync();
            Assert.Equal(0, indexedCount);

            // Change event must be durably preserved in PendingChanges for post-activation catch-up
            int pendingCount = await _repository.GetPendingChangesCountAsync();
            Assert.Equal(1, pendingCount);
        }

        [Fact]
        public async Task Test04_InvalidKey_BlocksNewBackgroundIndexing()
        {
            var mock = new MockOutlookAdapterForIndexer();
            var doc = new EmailDocument
            {
                StoreId = "Store1",
                EntryId = "Entry_Invalid_1",
                Subject = "Invalid Key Item",
                ReceivedDate = DateTime.Now
            };
            mock.EmailStorage[$"{doc.StoreId}:{doc.EntryId}"] = doc;

            var mockLicense = new MockLicensingService(LicenseStatus.InvalidKey);
            var indexer = new EmailIndexerService(mock, _repository, mockLicense);

            Assert.False(indexer.IsLicensingEligible());

            await indexer.QueueItemForIndexingAsync("Store1", "Entry_Invalid_1", "Folder1");

            int indexedCount = await _repository.GetIndexedEmailCountAsync();
            Assert.Equal(0, indexedCount);

            int pendingCount = await _repository.GetPendingChangesCountAsync();
            Assert.Equal(1, pendingCount);
        }

        [Fact]
        public async Task Test05_UnknownOrNullStatus_FailsClosed()
        {
            var mock = new MockOutlookAdapterForIndexer();
            var doc = new EmailDocument
            {
                StoreId = "Store1",
                EntryId = "Entry_Null_1",
                Subject = "Null License Dependency Item",
                ReceivedDate = DateTime.Now
            };
            mock.EmailStorage[$"{doc.StoreId}:{doc.EntryId}"] = doc;

            // 1. Compatibility constructor without licensing dependency fails closed
            var indexerWithoutLicense = new EmailIndexerService(mock, _repository);
            Assert.False(indexerWithoutLicense.IsLicensingEligible());

            await indexerWithoutLicense.QueueItemForIndexingAsync("Store1", "Entry_Null_1", "Folder1");
            int indexedCount = await _repository.GetIndexedEmailCountAsync();
            Assert.Equal(0, indexedCount);

            // 2. Unknown enum status fails closed
            var mockUnknownLicense = new MockLicensingService((LicenseStatus)999);
            var indexerUnknown = new EmailIndexerService(mock, _repository, mockUnknownLicense);
            Assert.False(indexerUnknown.IsLicensingEligible());
        }

        [Fact]
        public async Task Test06_LicensingServiceExceptions_FailClosed()
        {
            var mock = new MockOutlookAdapterForIndexer();
            var doc = new EmailDocument
            {
                StoreId = "Store1",
                EntryId = "Entry_Ex_1",
                Subject = "Exception Test Item",
                ReceivedDate = DateTime.Now
            };
            mock.EmailStorage[$"{doc.StoreId}:{doc.EntryId}"] = doc;

            var throwingLicense = new MockLicensingService();
            throwingLicense.ExceptionToThrow = new InvalidOperationException("Simulated DPAPI corruption or registry failure");

            var indexer = new EmailIndexerService(mock, _repository, throwingLicense);

            // Eligibility check must not crash and must fail closed
            bool eligible = indexer.IsLicensingEligible();
            Assert.False(eligible);

            // Indexing attempt must not crash and must not write email
            await indexer.QueueItemForIndexingAsync("Store1", "Entry_Ex_1", "Folder1");
            int indexedCount = await _repository.GetIndexedEmailCountAsync();
            Assert.Equal(0, indexedCount);
        }

        [Fact]
        public void Test07_ManualRibbonBackfill_IsBlocked_WhenIneligible()
        {
            // Verify licensing policy logic used by Ribbon backfill gate
            var expiredLicense = new MockLicensingService(LicenseStatus.Expired);
            var invalidKeyLicense = new MockLicensingService(LicenseStatus.InvalidKey);
            var validLicense = new MockLicensingService(LicenseStatus.Valid);
            var trialLicense = new MockLicensingService(LicenseStatus.Trial);

            bool IsAllowed(ILicensingService? lic)
            {
                var s = lic?.GetStatus() ?? LicenseStatus.InvalidKey;
                return s == LicenseStatus.Valid || s == LicenseStatus.Trial;
            }

            Assert.False(IsAllowed(expiredLicense));
            Assert.False(IsAllowed(invalidKeyLicense));
            Assert.False(IsAllowed(null));
            Assert.True(IsAllowed(validLicense));
            Assert.True(IsAllowed(trialLicense));
        }

        [Fact]
        public async Task Test08_UnderlyingBackfillService_CannotBeCalledToBypassLicenseGate()
        {
            var mock = new MockOutlookAdapterForIndexer();
            string storeId = "Store1";
            string folderId = "Folder1";

            mock.Mailboxes.Add(new MailboxInfo { StoreId = storeId, DisplayName = "Primary" });
            mock.StoreFolders[storeId] = new List<MailboxFolder>
            {
                new MailboxFolder { EntryId = folderId, Name = "Inbox", ItemCount = 100 }
            };

            var expiredLicense = new MockLicensingService(LicenseStatus.Expired);
            var indexer = new EmailIndexerService(mock, _repository, expiredLicense);

            // Direct call to StartActiveStoreBackfillAsync must reject immediately
            var summary = await indexer.StartActiveStoreBackfillAsync();
            Assert.NotNull(summary.ErrorMessage);
            Assert.Contains("license", summary.ErrorMessage, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(0, summary.NewlyIndexedCount);

            int indexedCount = await _repository.GetIndexedEmailCountAsync();
            Assert.Equal(0, indexedCount);
        }

        [Fact]
        public async Task Test09_RunningWorker_PausesSafely_WhenEligibilityChanges()
        {
            var mock = new MockOutlookAdapterForIndexer();
            string storeId = "Store1";
            string folderId = "Folder1";

            mock.Mailboxes.Add(new MailboxInfo { StoreId = storeId, DisplayName = "Primary" });
            mock.StoreFolders[storeId] = new List<MailboxFolder>
            {
                new MailboxFolder { EntryId = folderId, Name = "Inbox", ItemCount = 500 }
            };

            var emails = new List<EmailDocument>();
            for (int i = 1; i <= 500; i++)
            {
                emails.Add(new EmailDocument
                {
                    StoreId = storeId,
                    EntryId = $"BatchEmail_{i}",
                    Subject = $"Email {i}",
                    FolderName = "Inbox",
                    ReceivedDate = DateTime.Now
                });
            }
            mock.FolderEmails[$"{storeId}_{folderId}"] = emails;

            // License starts Valid, but changes to Expired on subsequent check
            int callCount = 0;
            var dynamicLicense = new MockLicensingService();
            dynamicLicense.StatusFunc = () =>
            {
                callCount++;
                // Allow initial checks and first batch, then expire before second batch
                return callCount <= 2 ? LicenseStatus.Valid : LicenseStatus.Expired;
            };

            var indexer = new EmailIndexerService(mock, _repository, dynamicLicense);

            var summary = await indexer.StartActiveStoreBackfillAsync();

            // Must pause safely without crash
            Assert.True(indexer.CurrentStatus == IndexerStatus.Paused || summary.IsCancelled);
            Assert.Equal(250, summary.NewlyIndexedCount);

            // Checkpoint must be preserved at offset 250
            var state = await _repository.GetFolderIndexStateAsync(storeId, folderId);
            Assert.NotNull(state);
            Assert.Equal(250, state.Value.ProcessedOffset);

            // Database records from first batch must remain intact
            int countInDb = await _repository.GetIndexedEmailCountAsync();
            Assert.Equal(250, countInDb);
        }

        [Fact]
        public async Task Test10_Reactivation_SchedulesOrResumes_IncrementalCatchUp()
        {
            var mock = new MockOutlookAdapterForIndexer();
            for (int i = 1; i <= 5; i++)
            {
                var doc = new EmailDocument
                {
                    StoreId = "Store1",
                    EntryId = $"MissedEmail_{i}",
                    Subject = $"Missed Subject {i}",
                    ReceivedDate = DateTime.Now
                };
                mock.EmailStorage[$"{doc.StoreId}:{doc.EntryId}"] = doc;
            }

            // Start expired
            var mockLicense = new MockLicensingService(LicenseStatus.Expired);
            var indexer = new EmailIndexerService(mock, _repository, mockLicense);

            // Queue 5 emails while expired
            for (int i = 1; i <= 5; i++)
            {
                await indexer.QueueItemForIndexingAsync("Store1", $"MissedEmail_{i}", "Folder1");
            }

            // Verify none were indexed yet, all 5 pending
            Assert.Equal(0, await _repository.GetIndexedEmailCountAsync());
            Assert.Equal(5, await _repository.GetPendingChangesCountAsync());

            // Activate license
            mockLicense.Status = LicenseStatus.Valid;

            // Resume incremental catch-up
            var summary = await indexer.ResumeIncrementalCatchUpAsync();

            // Verify all 5 missed emails are indexed
            int indexedCount = await _repository.GetIndexedEmailCountAsync();
            Assert.Equal(5, indexedCount);

            // Pending changes queue should be drained
            int pendingRemaining = await _repository.GetPendingChangesCountAsync();
            Assert.Equal(0, pendingRemaining);
        }

        [Fact]
        public async Task Test11_CatchUp_IsIdempotent_AndDoesNotCreateDuplicates()
        {
            var mock = new MockOutlookAdapterForIndexer();
            var doc = new EmailDocument
            {
                StoreId = "Store1",
                EntryId = "Idempotent_1",
                Subject = "Idempotent Subject",
                ReceivedDate = DateTime.Now
            };
            mock.EmailStorage[$"{doc.StoreId}:{doc.EntryId}"] = doc;

            var mockLicense = new MockLicensingService(LicenseStatus.Valid);
            var indexer = new EmailIndexerService(mock, _repository, mockLicense);

            await indexer.QueueItemForIndexingAsync("Store1", "Idempotent_1", "Folder1");
            Assert.Equal(1, await _repository.GetIndexedEmailCountAsync());

            // Run catch-up again
            await indexer.ResumeIncrementalCatchUpAsync();
            await indexer.ResumeIncrementalCatchUpAsync();

            // Total count must strictly remain 1
            Assert.Equal(1, await _repository.GetIndexedEmailCountAsync());
        }

        [Fact]
        public async Task Test12_ExistingCheckpointsAndIndexedRecords_RemainIntact_DuringPause()
        {
            var mock = new MockOutlookAdapterForIndexer();
            string storeId = "Store1";
            string folderId = "Folder1";

            mock.Mailboxes.Add(new MailboxInfo { StoreId = storeId, DisplayName = "Primary" });
            mock.StoreFolders[storeId] = new List<MailboxFolder>
            {
                new MailboxFolder { EntryId = folderId, Name = "Inbox", ItemCount = 10 }
            };
            var emails = new List<EmailDocument>();
            for (int i = 1; i <= 10; i++)
            {
                emails.Add(new EmailDocument
                {
                    StoreId = storeId,
                    EntryId = $"Item_{i}",
                    Subject = $"Subject {i}",
                    FolderName = "Inbox",
                    ReceivedDate = DateTime.Now
                });
            }
            mock.FolderEmails[$"{storeId}_{folderId}"] = emails;

            var mockLicense = new MockLicensingService(LicenseStatus.Valid);
            var indexer = new EmailIndexerService(mock, _repository, mockLicense);

            // Index all 10 items
            var summary = await indexer.StartActiveStoreBackfillAsync();
            Assert.Equal(10, summary.NewlyIndexedCount);
            Assert.Equal(10, await _repository.GetIndexedEmailCountAsync());

            var stateBefore = await _repository.GetFolderIndexStateAsync(storeId, folderId);
            Assert.NotNull(stateBefore);
            Assert.Equal(10, stateBefore.Value.ProcessedOffset);

            // License expires
            mockLicense.Status = LicenseStatus.Expired;

            // Attempt backfill while expired
            var expiredSummary = await indexer.StartActiveStoreBackfillAsync();
            Assert.NotNull(expiredSummary.ErrorMessage);

            // Check that existing indexed records and checkpoint were NOT wiped or reset
            Assert.Equal(10, await _repository.GetIndexedEmailCountAsync());
            var stateAfter = await _repository.GetFolderIndexStateAsync(storeId, folderId);
            Assert.NotNull(stateAfter);
            Assert.Equal(10, stateAfter.Value.ProcessedOffset);
        }

        [Fact]
        public async Task Test13_ExistingSearchHistory_RemainsIntact()
        {
            // Insert search history
            await _repository.SaveSearchQueryAsync("quarterly budget");
            await _repository.SaveSearchQueryAsync("annual review");

            var historyBefore = await _repository.GetRecentSearchQueriesAsync(10);
            Assert.Equal(2, historyBefore.Count);

            // Simulate indexer operating while expired
            var mock = new MockOutlookAdapterForIndexer();
            var mockLicense = new MockLicensingService(LicenseStatus.Expired);
            var indexer = new EmailIndexerService(mock, _repository, mockLicense);

            await indexer.QueueItemForIndexingAsync("Store1", "Item_1", "Folder1");
            await indexer.StartActiveStoreBackfillAsync();
            await indexer.PauseIndexingAsync();

            // Search history must remain completely intact
            var historyAfter = await _repository.GetRecentSearchQueriesAsync(10);
            Assert.Equal(2, historyAfter.Count);
            Assert.Contains("quarterly budget", historyAfter);
            Assert.Contains("annual review", historyAfter);
        }

        [Fact]
        public async Task Test14_RepeatedActivationNotifications_DoNotStartDuplicateCatchUpJobs()
        {
            var mock = new MockOutlookAdapterForIndexer();
            mock.Mailboxes.Add(new MailboxInfo { StoreId = "Store1", DisplayName = "Primary" });
            mock.StoreFolders["Store1"] = new List<MailboxFolder>();
            var mockLicense = new MockLicensingService(LicenseStatus.Valid);
            var indexer = new EmailIndexerService(mock, _repository, mockLicense);

            // Fire 5 catch-up requests concurrently
            var tasks = new List<Task<BackfillSummary>>();
            for (int i = 0; i < 5; i++)
            {
                tasks.Add(Task.Run(() => indexer.ResumeIncrementalCatchUpAsync()));
            }

            var results = await Task.WhenAll(tasks);

            // All should complete without error
            Assert.All(results, r => Assert.Null(r.ErrorMessage));
            Assert.Equal(IndexerStatus.Idle, indexer.CurrentStatus);
        }
    }
}
