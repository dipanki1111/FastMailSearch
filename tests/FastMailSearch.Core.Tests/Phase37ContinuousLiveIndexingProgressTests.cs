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
using FastMailSearch.Infrastructure.Services;
using Xunit;

namespace FastMailSearch.Core.Tests
{
    public class Phase37ContinuousLiveIndexingProgressTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly DatabaseConnectionFactory _factory;
        private readonly SqliteEmailRepository _repository;

        public Phase37ContinuousLiveIndexingProgressTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), $"Phase37Tests_{Guid.NewGuid():N}.db");
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

        private class MockOutlookAdapter37 : IOutlookAdapter
        {
            public List<MailboxInfo> Mailboxes { get; set; } = new List<MailboxInfo>();
            public Dictionary<string, List<MailboxFolder>> StoreFolders { get; set; } = new Dictionary<string, List<MailboxFolder>>();
            public Dictionary<string, List<EmailDocument>> FolderEmails { get; set; } = new Dictionary<string, List<EmailDocument>>();

            public bool IsOutlookAvailable() => true;
            public string GetOutlookVersion() => "16.0";
            public string GetOfficeBitness() => "64-bit";
            public List<MailboxInfo> GetMailboxes() => Mailboxes;

            public List<MailboxFolder> GetFolders(string storeId)
            {
                return StoreFolders.TryGetValue(storeId, out var folders) ? folders : new List<MailboxFolder>();
            }

            public List<EmailDocument> FetchFolderEmailsBatch(string storeId, string folderEntryId, int startIndex, int batchSize, Func<bool>? cancelCheck = null)
            {
                if (cancelCheck != null && cancelCheck()) return new List<EmailDocument>();
                string key = $"{storeId}:{folderEntryId}";
                if (!FolderEmails.TryGetValue(key, out var emails)) return new List<EmailDocument>();

                int zeroIndex = startIndex - 1;
                if (zeroIndex < 0 || zeroIndex >= emails.Count) return new List<EmailDocument>();
                int count = Math.Min(batchSize, emails.Count - zeroIndex);
                return emails.GetRange(zeroIndex, count);
            }

            public EmailDocument? FetchSingleEmail(string storeId, string entryId) => null;
            public bool OpenEmailInOutlook(string storeId, string entryId) => true;
            public bool IsOutboxFolder(string storeId, string folderEntryId) => false;
            public void SubscribeToEvents(Action<string, string, string> onItemAddedOrChanged, Action<string, string> onItemRemoved) { }
            public void UnsubscribeFromEvents() { }
        }

        [Fact]
        public async Task Test01_Batch_Progress_Publishes_All_Enriched_Counts()
        {
            var outlook = new MockOutlookAdapter37();
            outlook.Mailboxes.Add(new MailboxInfo { StoreId = "store-1", DisplayName = "Primary" });
            outlook.StoreFolders["store-1"] = new List<MailboxFolder>
            {
                new MailboxFolder { StoreId = "store-1", EntryId = "f-inbox", Name = "Inbox", ItemCount = 500 }
            };

            var emails = new List<EmailDocument>();
            for (int i = 1; i <= 500; i++)
            {
                emails.Add(new EmailDocument
                {
                    StoreId = "store-1",
                    EntryId = $"msg-{i}",
                    FolderEntryId = "f-inbox",
                    FolderName = "Inbox",
                    Subject = $"Subject {i}",
                    ReceivedDate = DateTime.Now
                });
            }
            outlook.FolderEmails["store-1:f-inbox"] = emails;

            var indexer = new EmailIndexerService(outlook, _repository, new MockLicensingService(LicenseStatus.Valid));
            var capturedEvents = new List<IndexerProgressEventArgs>();
            indexer.ProgressChanged += (s, e) => capturedEvents.Add(e);

            var summary = await indexer.StartActiveStoreBackfillAsync();
            Assert.False(summary.IsCancelled);
            Assert.Equal(500, summary.NewlyIndexedCount);

            // Filter batch progress events during active indexing
            var batchEvents = capturedEvents.Where(e => e.Status == IndexerStatus.Indexing && e.ProcessedItems > 0).ToList();
            Assert.NotEmpty(batchEvents);

            foreach (var ev in batchEvents)
            {
                Assert.Equal(500, ev.TotalDiscovered);
                Assert.Equal(ev.ProcessedItems, ev.ScannedCount);
                Assert.True(ev.RemainingCount >= 0);
                Assert.Equal(500, ev.ScannedCount + ev.RemainingCount);
                Assert.StartsWith("Backfilling: ", ev.CurrentOperation);
                Assert.True(ev.Percentage >= 0.0 && ev.Percentage <= 100.0);
            }

            var finalBatch = batchEvents.Last();
            Assert.Equal(500, finalBatch.ScannedCount);
            Assert.Equal(0, finalBatch.RemainingCount);
        }

        [Fact]
        public async Task Test02_Idle_State_Accuracy_When_No_Active_Work()
        {
            var outlook = new MockOutlookAdapter37();
            var indexer = new EmailIndexerService(outlook, _repository);
            var health = new IndexHealthService(_repository, indexer, outlook, _factory);

            var summary = await health.GetHealthSummaryAsync();
            Assert.Equal("Idle", summary.CurrentOperation);
            Assert.Equal(IndexerHealth.Healthy, summary.HealthState);
            Assert.Contains("Ready", summary.StatusState);

            // Stage 7AL-3: When queue has items but indexer is idle, accurately reports Queue Paused
            await _repository.EnqueuePendingChangeAsync("S1", "E1", "Add", "F1");
            summary = await health.GetHealthSummaryAsync();
            Assert.Equal("Idle — Queue Paused", summary.CurrentOperation);
            Assert.Equal(IndexerHealth.Paused, summary.HealthState);
            Assert.Contains("Idle — Queue Paused", summary.StatusState);
        }

        [Fact]
        public async Task Test03_CrossProcess_Active_Indexing_Detected_While_InProcess_Indexer_Is_Idle()
        {
            // Simulate OUTLOOK.EXE writing live progress into shared SQLite DB
            await _repository.SaveLiveProgressAsync(
                "Indexing",
                "Backfilling: Inbox (20,750/58,052)",
                20750,
                58052,
                "Inbox",
                20750,
                58052,
                20750,
                0,
                0,
                3429
            );
            await _repository.SaveFolderIndexStateAsync("pop-store", "f-inbox", "Indexing", 20750, 58052);

            // UI process has an indexer whose local IsIndexing is false
            var outlook = new MockOutlookAdapter37();
            var localUiIndexer = new EmailIndexerService(outlook, _repository);
            Assert.False(localUiIndexer.IsIndexing);

            var health = new IndexHealthService(_repository, localUiIndexer, outlook, _factory);
            var summary = await health.GetHealthSummaryAsync();

            // Must accurately reflect the active backfill in Outlook
            Assert.Equal(IndexerHealth.Updating, summary.HealthState);
            Assert.Equal("Backfilling: Inbox (20,750/58,052)", summary.CurrentOperation);
            Assert.Equal(58052, summary.TotalDiscoveredOutlookItems);
            Assert.Equal(20750, summary.ProcessedOffset);
            Assert.Equal(37302, summary.RemainingToScan);
            Assert.True(summary.ProgressPercentage >= 35.7 && summary.ProgressPercentage <= 35.8);
            Assert.Contains("Backfill in Progress", summary.StatusState);
            Assert.Contains("37,302 remaining", summary.StatusState);
        }

        [Fact]
        public async Task Test04_Stale_CrossProcess_Indexing_Reverts_To_Idle()
        {
            // Simulate an interrupted backfill from 10 minutes ago
            using (var conn = _factory.CreateConnection())
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = @"
                    INSERT INTO GlobalState (Key, Value, UpdatedAt) VALUES ('IndexState', 'Indexing', @OldTime);
                    INSERT INTO GlobalState (Key, Value, UpdatedAt) VALUES ('CurrentOperation', 'Backfilling: Inbox', @OldTime);
                    INSERT INTO GlobalState (Key, Value, UpdatedAt) VALUES ('LastProcessedCount', '20750', @OldTime);
                    INSERT INTO GlobalState (Key, Value, UpdatedAt) VALUES ('LastTotalCount', '58052', @OldTime);
                    INSERT INTO GlobalState (Key, Value, UpdatedAt) VALUES ('LastSync', @OldTime, @OldTime);
                    INSERT INTO IndexState (StoreId, FolderId, LastSync, Status, ItemCount, ProcessedOffset)
                    VALUES ('pop-store', 'f-inbox', @OldTime, 'Indexing', 58052, 20750);";
                string oldTimeStr = DateTime.Now.AddMinutes(-10).ToString("o");
                cmd.Parameters.AddWithValue("@OldTime", oldTimeStr);
                cmd.ExecuteNonQuery();
            }

            var outlook = new MockOutlookAdapter37();
            var localUiIndexer = new EmailIndexerService(outlook, _repository);
            var health = new IndexHealthService(_repository, localUiIndexer, outlook, _factory);

            var summary = await health.GetHealthSummaryAsync();

            // Must report Idle because the background worker is not actively syncing
            Assert.Equal("Idle", summary.CurrentOperation);
            Assert.Equal(IndexerHealth.Healthy, summary.HealthState);
            // Real checkpoint is preserved safely
            Assert.Equal(58052, summary.TotalDiscoveredOutlookItems);
            Assert.Equal(20750, summary.ProcessedOffset);
            Assert.Equal(37302, summary.RemainingToScan);
            Assert.Contains("Backfill recommended", summary.MailboxCoverageStatus);
        }

        [Fact]
        public async Task Test05_Cancellation_And_Resume_Checkpoints_Safely_Across_Restarts()
        {
            var outlook = new MockOutlookAdapter37();
            outlook.Mailboxes.Add(new MailboxInfo { StoreId = "pop-store", DisplayName = "POP Mailbox" });
            outlook.StoreFolders["pop-store"] = new List<MailboxFolder>
            {
                new MailboxFolder { StoreId = "pop-store", EntryId = "f-inbox", Name = "Inbox", ItemCount = 750 }
            };

            var emails = new List<EmailDocument>();
            for (int i = 1; i <= 750; i++)
            {
                emails.Add(new EmailDocument
                {
                    StoreId = "pop-store",
                    EntryId = $"msg-{i}",
                    FolderEntryId = "f-inbox",
                    FolderName = "Inbox",
                    Subject = $"Subject {i}",
                    ReceivedDate = DateTime.Now
                });
            }
            outlook.FolderEmails["pop-store:f-inbox"] = emails;

            // Phase 1: Start backfill, cancel after first batch (250 items)
            var indexer1 = new EmailIndexerService(outlook, _repository, new MockLicensingService(LicenseStatus.Valid));
            var cts1 = new CancellationTokenSource();
            int batchCount = 0;
            indexer1.ProgressChanged += (s, e) =>
            {
                if (e.Status == IndexerStatus.Indexing && e.ProcessedItems >= 250)
                {
                    batchCount++;
                    cts1.Cancel();
                }
            };

            var summary1 = await indexer1.StartActiveStoreBackfillAsync(cts1.Token);
            Assert.True(summary1.IsCancelled);

            // Verify checkpoint saved in SQLite
            var checkpoint = await _repository.GetFolderIndexStateAsync("pop-store", "f-inbox");
            Assert.NotNull(checkpoint);
            Assert.Equal(250, checkpoint.Value.ProcessedOffset);
            Assert.Equal(750, checkpoint.Value.TotalItems);

            // Phase 2: Create brand new indexer instance (simulating application restart)
            var indexer2 = new EmailIndexerService(outlook, _repository, new MockLicensingService(LicenseStatus.Valid));
            var resumeEvents = new List<IndexerProgressEventArgs>();
            indexer2.ProgressChanged += (s, e) => resumeEvents.Add(e);

            var summary2 = await indexer2.StartActiveStoreBackfillAsync();
            Assert.False(summary2.IsCancelled);

            // Verify resumed from 250
            var firstResumeEvent = resumeEvents.First(e => e.Status == IndexerStatus.Indexing && e.ProcessedItems > 0);
            Assert.Equal(250, firstResumeEvent.ProcessedItems);

            // Final state reaches 750/750
            var finalEvent = resumeEvents.Last(e => e.Status == IndexerStatus.Indexing);
            Assert.Equal(750, finalEvent.ProcessedItems);
            Assert.Equal(0, finalEvent.RemainingCount);

            // Ensure no duplicate emails in database
            int totalInDb = await _repository.GetIndexedEmailCountAsync();
            Assert.Equal(750, totalInDb);
        }

        [Fact]
        public async Task Test06_Count_Consistency_Disambiguates_Mailbox_Total_DbCount_And_Queue()
        {
            // Seed database:
            // 1. Folder checkpoint has ItemCount = 58,052, ProcessedOffset = 12,250
            await _repository.SaveFolderIndexStateAsync("pop-store", "f-inbox", "Indexing", 12250, 58052);

            // 2. Queue has 3,429 items
            for (int i = 0; i < 50; i++)
            {
                await _repository.EnqueuePendingChangeAsync("pop-store", $"queue-{i}", "Add", "f-inbox");
            }

            var outlook = new MockOutlookAdapter37();
            var indexer = new EmailIndexerService(outlook, _repository);
            var health = new IndexHealthService(_repository, indexer, outlook, _factory);

            var summary = await health.GetHealthSummaryAsync();

            // Total discovered mailbox emails must be verified from checkpoint (58,052)
            Assert.Equal(58052, summary.TotalDiscoveredOutlookItems);
            Assert.NotEqual(summary.TotalDiscoveredOutlookItems, summary.IndexedEmailCount);
            Assert.NotEqual(summary.TotalDiscoveredOutlookItems, summary.PendingQueueCount);

            // Remaining to scan must be calculated from real checkpoint (58,052 - 12,250 = 45,802)
            Assert.Equal(45802, summary.RemainingToScan);
            Assert.Equal(12250, summary.ProcessedOffset);

            // Pending queue is separate
            Assert.Equal(50, summary.PendingQueueCount);
            Assert.Equal(0, summary.FailedItemCount);
        }

        [Fact]
        public async Task Test07_MultiStore_Checkpoints_Isolated_To_Active_Store()
        {
            // Seed database:
            // 1. Multiple completed folders from an inactive/previous store (e.g. OST account with 22,306 items across 24 folders)
            for (int i = 1; i <= 24; i++)
            {
                int count = (i == 24) ? 306 : 1000;
                await _repository.SaveFolderIndexStateAsync("ost-store", $"f-ost-{i}", "Completed", count, count);
            }

            // 2. Active POP PST store being indexed (58,052 total items, checkpoint offset 23,250)
            await _repository.SaveFolderIndexStateAsync("pop-pst-store", "f-inbox", "Indexing", 23250, 58052);

            // 3. Pending queue has 3,429 items
            for (int i = 0; i < 10; i++)
            {
                await _repository.EnqueuePendingChangeAsync("pop-pst-store", $"queue-{i}", "Add", "f-inbox");
            }

            var outlook = new MockOutlookAdapter37();
            var indexer = new EmailIndexerService(outlook, _repository);
            var health = new IndexHealthService(_repository, indexer, outlook, _factory);

            var summary = await health.GetHealthSummaryAsync();

            // Checkpoints must isolate the active indexing store: 23,250 of 58,052 (NOT 22,306 + 23,250 = 45,556 / 79%)
            Assert.Equal(58052, summary.TotalDiscoveredOutlookItems);
            Assert.Equal(23250, summary.ProcessedOffset);
            Assert.Equal(23250, summary.ScannedCount);
            Assert.Equal(34802, summary.RemainingToScan); // 58,052 - 23,250
            Assert.True(summary.ProgressPercentage >= 40.0 && summary.ProgressPercentage <= 40.2, $"Expected ~40.1% but got {summary.ProgressPercentage}%");
            Assert.Equal(10, summary.PendingQueueCount);
            Assert.Equal(0, summary.FailedItemCount);
        }
    }
}
