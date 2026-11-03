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
    public class Phase4K5ABoundedIngestionSafetyTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly DatabaseConnectionFactory _factory;
        private readonly SqliteEmailRepository _repository;

        public Phase4K5ABoundedIngestionSafetyTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), $"Phase4K5ABounded_{Guid.NewGuid():N}.db");
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
                string wal = _dbPath + "-wal";
                if (File.Exists(wal)) File.Delete(wal);
                string shm = _dbPath + "-shm";
                if (File.Exists(shm)) File.Delete(shm);
            }
            catch { }
        }

        private int GetDatabaseEmailCount()
        {
            using var conn = _factory.CreateConnection();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT COUNT(*) FROM Emails;";
            return Convert.ToInt32(cmd.ExecuteScalar());
        }

        private int GetDistinctEntryIdCount()
        {
            using var conn = _factory.CreateConnection();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT COUNT(DISTINCT EntryId) FROM Emails;";
            return Convert.ToInt32(cmd.ExecuteScalar());
        }

        /// <summary>
        /// Mock adapter supporting configurable folders and tracking batch request limits.
        /// </summary>
        private class MockBoundedOutlookAdapter : IOutlookAdapter, IFolderBatchReader
        {
            public List<MailboxFolder> Folders { get; set; } = new List<MailboxFolder>();
            public Dictionary<string, List<EmailDocument>> FolderEmails { get; set; } = new Dictionary<string, List<EmailDocument>>();

            public int TotalItemsFetched { get; private set; }
            public int MaxBatchSizeRequested { get; private set; }
            public List<int> RequestedBatchSizes { get; } = new List<int>();

            public Action<int>? OnItemReadCallback { get; set; }
            public Action<int>? OnBatchFetchedCallback { get; set; }

            public bool IsOutlookAvailable() => true;
            public string GetOutlookVersion() => "16.0 (64-bit)";
            public string GetOfficeBitness() => "x64";
            public List<MailboxInfo> GetMailboxes() => new List<MailboxInfo>
            {
                new MailboxInfo { StoreId = "store1", DisplayName = "POP Archive", IsPst = true }
            };

            public List<MailboxFolder> GetFolders(string storeId) => Folders;

            public void ResetCursorSession() { }

            public FolderBatchResult FetchFolderBatch(string storeId, string folderEntryId, int startIndex, int batchSize, Func<bool>? cancelCheck = null)
            {
                RequestedBatchSizes.Add(batchSize);
                if (batchSize > MaxBatchSizeRequested)
                    MaxBatchSizeRequested = batchSize;

                var result = new FolderBatchResult
                {
                    StartIndex = startIndex,
                    LastProcessedIndex = startIndex - 1,
                    ProcessedCount = 0,
                    IndexedCount = 0,
                    SkippedCount = 0
                };

                if (!FolderEmails.TryGetValue(folderEntryId, out var allEmails) || allEmails.Count == 0)
                {
                    return result;
                }

                int total = allEmails.Count;
                if (startIndex > total) return result;

                int end = Math.Min(startIndex + batchSize - 1, total);
                var batchList = new List<EmailDocument>();

                for (int i = startIndex; i <= end; i++)
                {
                    if (cancelCheck != null && cancelCheck())
                    {
                        result.WasCancelled = true;
                        break;
                    }

                    OnItemReadCallback?.Invoke(i);

                    var email = allEmails[i - 1];
                    batchList.Add(email);
                    TotalItemsFetched++;
                    result.LastProcessedIndex = i;
                    result.ProcessedCount++;
                    result.IndexedCount++;
                }

                result.Items = batchList;
                OnBatchFetchedCallback?.Invoke(result.IndexedCount);
                return result;
            }

            public List<EmailDocument> FetchFolderEmailsBatch(string storeId, string folderEntryId, int startIndex, int batchSize, Func<bool>? cancelCheck = null)
            {
                var r = FetchFolderBatch(storeId, folderEntryId, startIndex, batchSize, cancelCheck);
                return r.Items;
            }

            public EmailDocument? FetchSingleEmail(string storeId, string entryId) => null;
            public bool OpenEmailInOutlook(string storeId, string entryId) => false;
            public bool IsOutboxFolder(string storeId, string folderEntryId) => false;
            public void SubscribeToEvents(Action<string, string, string> onItemAddedOrChanged, Action<string, string> onItemRemoved) { }
            public void UnsubscribeFromEvents() { }
        }

        private static List<EmailDocument> GenerateEmails(string folderEntryId, string folderName, int count)
        {
            var list = new List<EmailDocument>();
            for (int i = 1; i <= count; i++)
            {
                list.Add(new EmailDocument
                {
                    StoreId = "store1",
                    EntryId = $"ENTRY_{folderEntryId}_{i:D6}",
                    FolderEntryId = folderEntryId,
                    FolderName = folderName,
                    Subject = $"Synthetic Subject {folderName} #{i}",
                    SenderEmail = "pilot@test.local",
                    SenderName = "Pilot Tester",
                    ReceivedDate = DateTime.Now.AddMinutes(-i),
                    BodyText = $"Synthetic Body for {folderName} item {i}"
                });
            }
            return list;
        }

        [Fact]
        public async Task Test01_MaxItems_Null_Preserves_Existing_Unlimited_Behavior()
        {
            var adapter = new MockBoundedOutlookAdapter();
            adapter.Folders.Add(new MailboxFolder { StoreId = "store1", EntryId = "f1", Name = "Inbox", ItemCount = 300 });
            adapter.FolderEmails["f1"] = GenerateEmails("f1", "Inbox", 300);

            var indexer = new EmailIndexerService(adapter, _repository, new MockLicensingService(LicenseStatus.Valid));
            var summary = await indexer.StartActiveStoreBackfillAsync(CancellationToken.None, maxItems: null);

            Assert.False(summary.IsCancelled);
            Assert.Equal(300, summary.NewlyIndexedCount);
            Assert.Equal(300, GetDatabaseEmailCount());

            var state = await _repository.GetFolderIndexStateAsync("store1", "f1");
            Assert.True(state.HasValue);
            Assert.Equal("Completed", state.Value.Status);
            Assert.Equal(300, state.Value.ProcessedOffset);
        }

        [Fact]
        public async Task Test02_MaxItems_One_Fetches_And_Persists_Exactly_One_Item()
        {
            var adapter = new MockBoundedOutlookAdapter();
            adapter.Folders.Add(new MailboxFolder { StoreId = "store1", EntryId = "f1", Name = "Inbox", ItemCount = 100 });
            adapter.FolderEmails["f1"] = GenerateEmails("f1", "Inbox", 100);

            var indexer = new EmailIndexerService(adapter, _repository, new MockLicensingService(LicenseStatus.Valid));
            var summary = await indexer.StartActiveStoreBackfillAsync(CancellationToken.None, maxItems: 1);

            Assert.False(summary.IsCancelled);
            Assert.Equal(1, summary.NewlyIndexedCount);
            Assert.Equal(1, GetDatabaseEmailCount());
            Assert.Equal(1, adapter.TotalItemsFetched);
            Assert.Equal(1, adapter.MaxBatchSizeRequested);

            var state = await _repository.GetFolderIndexStateAsync("store1", "f1");
            Assert.True(state.HasValue);
            Assert.Equal("Indexing", state.Value.Status);
            Assert.Equal(1, state.Value.ProcessedOffset);
        }

        [Fact]
        public async Task Test03_MaxItems_Fifty_Fetches_And_Persists_Exactly_Fifty_Items()
        {
            var adapter = new MockBoundedOutlookAdapter();
            adapter.Folders.Add(new MailboxFolder { StoreId = "store1", EntryId = "f1", Name = "Inbox", ItemCount = 200 });
            adapter.FolderEmails["f1"] = GenerateEmails("f1", "Inbox", 200);

            var indexer = new EmailIndexerService(adapter, _repository, new MockLicensingService(LicenseStatus.Valid));
            var summary = await indexer.StartActiveStoreBackfillAsync(CancellationToken.None, maxItems: 50);

            Assert.False(summary.IsCancelled);
            Assert.Equal(50, summary.NewlyIndexedCount);
            Assert.Equal(50, GetDatabaseEmailCount());
            Assert.Equal(50, adapter.TotalItemsFetched);
            Assert.Equal(50, adapter.MaxBatchSizeRequested);

            var state = await _repository.GetFolderIndexStateAsync("store1", "f1");
            Assert.True(state.HasValue);
            Assert.Equal("Indexing", state.Value.Status);
            Assert.Equal(50, state.Value.ProcessedOffset);
        }

        [Fact]
        public async Task Test04_Limit_Reached_Across_Multiple_Folders()
        {
            var adapter = new MockBoundedOutlookAdapter();
            adapter.Folders.Add(new MailboxFolder { StoreId = "store1", EntryId = "fA", Name = "FolderA", ItemCount = 20 });
            adapter.Folders.Add(new MailboxFolder { StoreId = "store1", EntryId = "fB", Name = "FolderB", ItemCount = 40 });
            adapter.Folders.Add(new MailboxFolder { StoreId = "store1", EntryId = "fC", Name = "FolderC", ItemCount = 50 });

            adapter.FolderEmails["fA"] = GenerateEmails("fA", "FolderA", 20);
            adapter.FolderEmails["fB"] = GenerateEmails("fB", "FolderB", 40);
            adapter.FolderEmails["fC"] = GenerateEmails("fC", "FolderC", 50);

            var indexer = new EmailIndexerService(adapter, _repository, new MockLicensingService(LicenseStatus.Valid));
            var summary = await indexer.StartActiveStoreBackfillAsync(CancellationToken.None, maxItems: 50);

            Assert.False(summary.IsCancelled);
            Assert.Equal(50, summary.NewlyIndexedCount);
            Assert.Equal(50, GetDatabaseEmailCount());
            Assert.Equal(50, adapter.TotalItemsFetched);

            // FolderA was fully completed (20 items)
            var stateA = await _repository.GetFolderIndexStateAsync("store1", "fA");
            Assert.True(stateA.HasValue);
            Assert.Equal("Completed", stateA.Value.Status);
            Assert.Equal(20, stateA.Value.ProcessedOffset);

            // FolderB had 30 items indexed (reaches limit of 50 total)
            var stateB = await _repository.GetFolderIndexStateAsync("store1", "fB");
            Assert.True(stateB.HasValue);
            Assert.Equal("Indexing", stateB.Value.Status);
            Assert.Equal(30, stateB.Value.ProcessedOffset);

            // FolderC was never even accessed
            var stateC = await _repository.GetFolderIndexStateAsync("store1", "fC");
            Assert.False(stateC.HasValue);
        }

        [Fact]
        public async Task Test05_Empty_Folder_Handling_With_Limit()
        {
            var adapter = new MockBoundedOutlookAdapter();
            adapter.Folders.Add(new MailboxFolder { StoreId = "store1", EntryId = "fEmpty", Name = "EmptyFolder", ItemCount = 0 });
            adapter.Folders.Add(new MailboxFolder { StoreId = "store1", EntryId = "fReal", Name = "RealFolder", ItemCount = 30 });

            adapter.FolderEmails["fEmpty"] = new List<EmailDocument>();
            adapter.FolderEmails["fReal"] = GenerateEmails("fReal", "RealFolder", 30);

            var indexer = new EmailIndexerService(adapter, _repository, new MockLicensingService(LicenseStatus.Valid));
            var summary = await indexer.StartActiveStoreBackfillAsync(CancellationToken.None, maxItems: 15);

            Assert.False(summary.IsCancelled);
            Assert.Equal(15, summary.NewlyIndexedCount);
            Assert.Equal(15, GetDatabaseEmailCount());

            var stateEmpty = await _repository.GetFolderIndexStateAsync("store1", "fEmpty");
            Assert.True(stateEmpty.HasValue);
            Assert.Equal("Completed", stateEmpty.Value.Status);
            Assert.Equal(0, stateEmpty.Value.ProcessedOffset);

            var stateReal = await _repository.GetFolderIndexStateAsync("store1", "fReal");
            Assert.True(stateReal.HasValue);
            Assert.Equal("Indexing", stateReal.Value.Status);
            Assert.Equal(15, stateReal.Value.ProcessedOffset);
        }

        [Fact]
        public async Task Test06_Cancellation_Before_Limit()
        {
            var adapter = new MockBoundedOutlookAdapter();
            adapter.Folders.Add(new MailboxFolder { StoreId = "store1", EntryId = "f1", Name = "Inbox", ItemCount = 100 });
            adapter.FolderEmails["f1"] = GenerateEmails("f1", "Inbox", 100);

            using var cts = new CancellationTokenSource();
            int readCount = 0;
            adapter.OnItemReadCallback = (itemIdx) =>
            {
                readCount++;
                if (readCount == 10)
                {
                    cts.Cancel();
                }
            };

            var indexer = new EmailIndexerService(adapter, _repository, new MockLicensingService(LicenseStatus.Valid));
            var summary = await indexer.StartActiveStoreBackfillAsync(cts.Token, maxItems: 50);

            Assert.True(summary.IsCancelled);
            Assert.True(GetDatabaseEmailCount() <= 50);
        }

        [Fact]
        public async Task Test07_Cancellation_At_Batch_Boundary()
        {
            var adapter = new MockBoundedOutlookAdapter();
            adapter.Folders.Add(new MailboxFolder { StoreId = "store1", EntryId = "f1", Name = "Folder1", ItemCount = 25 });
            adapter.Folders.Add(new MailboxFolder { StoreId = "store1", EntryId = "f2", Name = "Folder2", ItemCount = 75 });
            adapter.FolderEmails["f1"] = GenerateEmails("f1", "Folder1", 25);
            adapter.FolderEmails["f2"] = GenerateEmails("f2", "Folder2", 75);

            using var cts = new CancellationTokenSource();
            var indexer = new EmailIndexerService(adapter, _repository, new MockLicensingService(LicenseStatus.Valid));

            // Cancel right when Folder 1 completes persistence at the boundary
            indexer.ProgressChanged += (s, e) =>
            {
                if (e.FolderProcessedItems == 25 && e.CurrentFolder == "Folder1")
                {
                    cts.Cancel();
                }
            };

            var summary = await indexer.StartActiveStoreBackfillAsync(cts.Token, maxItems: 50);

            Assert.True(summary.IsCancelled);
            Assert.Equal(25, summary.NewlyIndexedCount);
            Assert.Equal(25, GetDatabaseEmailCount());

            var state1 = await _repository.GetFolderIndexStateAsync("store1", "f1");
            Assert.True(state1.HasValue);
            Assert.Equal(25, state1.Value.ProcessedOffset);

            var state2 = await _repository.GetFolderIndexStateAsync("store1", "f2");
            Assert.False(state2.HasValue);
        }

        [Fact]
        public async Task Test08_Resume_After_Bounded_Indexing()
        {
            var adapter = new MockBoundedOutlookAdapter();
            adapter.Folders.Add(new MailboxFolder { StoreId = "store1", EntryId = "f1", Name = "Inbox", ItemCount = 100 });
            adapter.FolderEmails["f1"] = GenerateEmails("f1", "Inbox", 100);

            // Run 1: Bound to 25 items
            var indexer1 = new EmailIndexerService(adapter, _repository, new MockLicensingService(LicenseStatus.Valid));
            var sum1 = await indexer1.StartActiveStoreBackfillAsync(CancellationToken.None, maxItems: 25);
            Assert.Equal(25, sum1.NewlyIndexedCount);
            Assert.Equal(25, GetDatabaseEmailCount());

            var state1 = await _repository.GetFolderIndexStateAsync("store1", "f1");
            Assert.Equal(25, state1?.ProcessedOffset);

            // Run 2: Resume with next bound of 25 items (items 26..50)
            var indexer2 = new EmailIndexerService(adapter, _repository, new MockLicensingService(LicenseStatus.Valid));
            var sum2 = await indexer2.StartActiveStoreBackfillAsync(CancellationToken.None, maxItems: 25);
            Assert.Equal(25, sum2.NewlyIndexedCount);
            Assert.Equal(50, GetDatabaseEmailCount());

            var state2 = await _repository.GetFolderIndexStateAsync("store1", "f1");
            Assert.Equal(50, state2?.ProcessedOffset);

            // Run 3: Full resume with maxItems = null (items 51..100)
            var indexer3 = new EmailIndexerService(adapter, _repository, new MockLicensingService(LicenseStatus.Valid));
            var sum3 = await indexer3.StartActiveStoreBackfillAsync(CancellationToken.None, maxItems: null);
            Assert.Equal(50, sum3.NewlyIndexedCount);
            Assert.Equal(100, GetDatabaseEmailCount());

            var state3 = await _repository.GetFolderIndexStateAsync("store1", "f1");
            Assert.Equal("Completed", state3?.Status);
            Assert.Equal(100, state3?.ProcessedOffset);

            // Confirm all 100 items are distinct
            Assert.Equal(100, GetDistinctEntryIdCount());
        }

        [Fact]
        public async Task Test09_Duplicate_Prevention_On_Retry()
        {
            var adapter = new MockBoundedOutlookAdapter();
            adapter.Folders.Add(new MailboxFolder { StoreId = "store1", EntryId = "f1", Name = "Inbox", ItemCount = 50 });
            adapter.FolderEmails["f1"] = GenerateEmails("f1", "Inbox", 50);

            // Run 1: Index 25 items
            var indexer1 = new EmailIndexerService(adapter, _repository, new MockLicensingService(LicenseStatus.Valid));
            await indexer1.StartActiveStoreBackfillAsync(CancellationToken.None, maxItems: 25);
            Assert.Equal(25, GetDatabaseEmailCount());

            // Force folder offset back to 0 (simulating re-try or forced re-scan)
            await _repository.SaveFolderIndexStateAsync("store1", "f1", "Indexing", 0, 50);

            // Run 2: Re-index 25 items over same store
            var indexer2 = new EmailIndexerService(adapter, _repository, new MockLicensingService(LicenseStatus.Valid));
            await indexer2.StartActiveStoreBackfillAsync(CancellationToken.None, maxItems: 25);

            // Total count must still be exactly 25 due to upsert idempotency and UNIQUE constraint
            Assert.Equal(25, GetDatabaseEmailCount());
            Assert.Equal(25, GetDistinctEntryIdCount());
        }

        [Fact]
        public async Task Test10_No_More_Than_MaxItems_Fetched_And_Persisted()
        {
            var adapter = new MockBoundedOutlookAdapter();
            adapter.Folders.Add(new MailboxFolder { StoreId = "store1", EntryId = "f1", Name = "Inbox", ItemCount = 500 });
            adapter.FolderEmails["f1"] = GenerateEmails("f1", "Inbox", 500);

            int targetLimit = 37;
            var indexer = new EmailIndexerService(adapter, _repository, new MockLicensingService(LicenseStatus.Valid));
            var summary = await indexer.StartActiveStoreBackfillAsync(CancellationToken.None, maxItems: targetLimit);

            Assert.Equal(targetLimit, summary.NewlyIndexedCount);
            Assert.Equal(targetLimit, GetDatabaseEmailCount());
            Assert.Equal(targetLimit, adapter.TotalItemsFetched);
            Assert.True(adapter.MaxBatchSizeRequested <= targetLimit);
        }

        [Fact]
        public async Task Test11_StartInitialIndexing_MaxItems_Respected()
        {
            var adapter = new MockBoundedOutlookAdapter();
            adapter.Folders.Add(new MailboxFolder { StoreId = "store1", EntryId = "f1", Name = "Inbox", ItemCount = 80 });
            adapter.FolderEmails["f1"] = GenerateEmails("f1", "Inbox", 80);

            var indexer = new EmailIndexerService(adapter, _repository, new MockLicensingService(LicenseStatus.Valid));
            await indexer.StartInitialIndexingAsync(CancellationToken.None, maxItems: 12);

            Assert.Equal(12, GetDatabaseEmailCount());
            Assert.Equal(12, GetDistinctEntryIdCount());
        }

        [Theory]
        [InlineData(0)]
        [InlineData(-1)]
        [InlineData(-100)]
        public async Task Test12_Zero_And_Negative_MaxItems_Returns_Immediately_With_Zero_Indexed(int nonPositiveLimit)
        {
            var adapter = new MockBoundedOutlookAdapter();
            adapter.Folders.Add(new MailboxFolder { StoreId = "store1", EntryId = "f1", Name = "Inbox", ItemCount = 100 });
            adapter.FolderEmails["f1"] = GenerateEmails("f1", "Inbox", 100);

            var indexer = new EmailIndexerService(adapter, _repository, new MockLicensingService(LicenseStatus.Valid));
            var summary = await indexer.StartActiveStoreBackfillAsync(CancellationToken.None, maxItems: nonPositiveLimit);

            Assert.False(summary.IsCancelled);
            Assert.Equal(0, summary.NewlyIndexedCount);
            Assert.Equal(0, GetDatabaseEmailCount());
            Assert.Equal(0, adapter.TotalItemsFetched);

            // Also verify StartInitialIndexingAsync with non-positive limit
            await indexer.StartInitialIndexingAsync(CancellationToken.None, maxItems: nonPositiveLimit);
            Assert.Equal(0, GetDatabaseEmailCount());
        }
    }
}
