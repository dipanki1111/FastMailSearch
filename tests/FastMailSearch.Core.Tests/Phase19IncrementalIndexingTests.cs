using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using FastMailSearch.Core.Interfaces;
using FastMailSearch.Core.Models;
using FastMailSearch.Core.Services;
using FastMailSearch.Infrastructure.Database;
using FastMailSearch.Infrastructure.Indexing;
using Xunit;

namespace FastMailSearch.Core.Tests
{
    public class Phase19IncrementalIndexingTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly DatabaseConnectionFactory _factory;
        private readonly SqliteEmailRepository _repository;
        private readonly SearchService _searchService;

        public Phase19IncrementalIndexingTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), $"Phase19Tests_{Guid.NewGuid():N}.db");
            _factory = new DatabaseConnectionFactory(_dbPath);
            _repository = new SqliteEmailRepository(_factory);
            _repository.InitializeDatabase();
            _searchService = new SearchService(_repository, _repository, new MockLicensingService());
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

        private class MockOutlookAdapter : IOutlookAdapter
        {
            public Dictionary<string, EmailDocument> EmailStorage = new Dictionary<string, EmailDocument>();
            public int FetchAttemptsCount = 0;
            public int TransientFailuresRemaining = 0;

            public bool IsOutlookAvailable() => true;
            public string GetOutlookVersion() => "16.0.15000.20000";
            public string GetOfficeBitness() => "64-bit";
            public List<MailboxInfo> GetMailboxes() => new List<MailboxInfo>
            {
                new MailboxInfo { StoreId = "Store1", DisplayName = "Test Mailbox" }
            };
            public List<MailboxFolder> GetFolders(string storeId) => new List<MailboxFolder>
            {
                new MailboxFolder { StoreId = storeId, EntryId = "Folder1", Name = "Inbox", ItemCount = 0 }
            };
            public List<EmailDocument> FetchFolderEmailsBatch(string storeId, string folderEntryId, int startIndex, int batchSize, Func<bool>? cancelCheck = null)
            {
                return EmailStorage.Values.Skip(startIndex - 1).Take(batchSize).ToList();
            }
            public EmailDocument? FetchSingleEmail(string storeId, string entryId)
            {
                FetchAttemptsCount++;
                if (TransientFailuresRemaining > 0)
                {
                    TransientFailuresRemaining--;
                    return null;
                }
                string key = $"{storeId}:{entryId}";
                return EmailStorage.TryGetValue(key, out var doc) ? doc : null;
            }
            public bool OpenEmailInOutlook(string storeId, string entryId) => true;
            public bool IsOutboxFolder(string storeId, string folderEntryId) => string.Equals(folderEntryId, "Outbox", StringComparison.OrdinalIgnoreCase);
            public void SubscribeToEvents(Action<string, string, string> onItemAddedOrChanged, Action<string, string> onItemRemoved) { }
            public void UnsubscribeFromEvents() { }
        }

        [Fact]
        public async Task IncrementalNewItem_IndexesSuccessfully()
        {
            var mockOutlook = new MockOutlookAdapter();
            var doc = new EmailDocument
            {
                StoreId = "Store1",
                EntryId = "Entry_New_1",
                Subject = "Quarterly Financial Analysis",
                SenderName = "CFO Office",
                SenderEmail = "cfo@company.com",
                Recipients = "all@company.com",
                BodyText = "Attached is the Q3 analysis report for review.",
                FolderName = "Inbox",
                FolderEntryId = "Folder1",
                ReceivedDate = DateTime.Now
            };
            mockOutlook.EmailStorage[$"{doc.StoreId}:{doc.EntryId}"] = doc;

            var indexer = new EmailIndexerService(mockOutlook, _repository, new MockLicensingService());
            await indexer.QueueItemForIndexingAsync("Store1", "Entry_New_1", "Folder1");

            var search = await _searchService.ExecuteSearchAsync("Quarterly Financial Analysis");
            Assert.Single(search.Items);
            Assert.Equal("Quarterly Financial Analysis", search.Items[0].Subject);
        }

        [Fact]
        public async Task IncrementalUpdate_UpdatesExistingRowWithoutDuplicate()
        {
            var mockOutlook = new MockOutlookAdapter();
            var doc = new EmailDocument
            {
                StoreId = "Store1",
                EntryId = "Entry_Update_1",
                Subject = "Original Subject Line",
                SenderName = "Alice",
                SenderEmail = "alice@example.com",
                BodyText = "Initial draft body text.",
                FolderName = "Inbox",
                FolderEntryId = "Folder1",
                ReceivedDate = DateTime.Now
            };
            mockOutlook.EmailStorage[$"{doc.StoreId}:{doc.EntryId}"] = doc;

            var indexer = new EmailIndexerService(mockOutlook, _repository, new MockLicensingService());
            await indexer.QueueItemForIndexingAsync("Store1", "Entry_Update_1", "Folder1");

            // Update subject and body text
            doc.Subject = "REVISED Subject Line";
            doc.BodyText = "Updated final body text content.";
            await indexer.QueueItemForIndexingAsync("Store1", "Entry_Update_1", "Folder1");

            var meta = await _repository.GetIndexMetadataAsync();
            Assert.Equal(1, meta.TotalEmailsIndexed);

            var oldSearch = await _searchService.ExecuteSearchAsync("\"Original Subject Line\"");
            Assert.Empty(oldSearch.Items);

            var newSearch = await _searchService.ExecuteSearchAsync("\"REVISED Subject Line\"");
            Assert.Single(newSearch.Items);
            Assert.Equal("REVISED Subject Line", newSearch.Items[0].Subject);
        }

        [Fact]
        public async Task IncrementalDelete_RemovesFromDatabaseAndFts()
        {
            var mockOutlook = new MockOutlookAdapter();
            var doc = new EmailDocument
            {
                StoreId = "Store1",
                EntryId = "Entry_Del_1",
                Subject = "Confidential Strategy Document",
                BodyText = "Do not share externally.",
                FolderName = "Inbox",
                FolderEntryId = "Folder1",
                ReceivedDate = DateTime.Now
            };
            mockOutlook.EmailStorage[$"{doc.StoreId}:{doc.EntryId}"] = doc;

            var indexer = new EmailIndexerService(mockOutlook, _repository, new MockLicensingService());
            await indexer.QueueItemForIndexingAsync("Store1", "Entry_Del_1", "Folder1");

            var searchBefore = await _searchService.ExecuteSearchAsync("Confidential Strategy");
            Assert.Single(searchBefore.Items);

            // Trigger delete
            await indexer.QueueItemForDeletionAsync("Store1", "Entry_Del_1");

            var searchAfter = await _searchService.ExecuteSearchAsync("Confidential Strategy");
            Assert.Empty(searchAfter.Items);
        }

        [Fact]
        public async Task IncrementalMove_UpdatesFolderMetadataWithoutDuplicate()
        {
            var mockOutlook = new MockOutlookAdapter();
            var doc = new EmailDocument
            {
                StoreId = "Store1",
                EntryId = "Entry_Move_1",
                Subject = "Moved Email Topic",
                FolderName = "Inbox",
                FolderEntryId = "Folder1",
                ReceivedDate = DateTime.Now
            };
            mockOutlook.EmailStorage[$"{doc.StoreId}:{doc.EntryId}"] = doc;

            var indexer = new EmailIndexerService(mockOutlook, _repository, new MockLicensingService());
            await indexer.QueueItemForIndexingAsync("Store1", "Entry_Move_1", "Folder1");

            // Move to Archive folder
            doc.FolderName = "Archive";
            doc.FolderEntryId = "Folder2";
            await indexer.QueueItemForIndexingAsync("Store1", "Entry_Move_1", "Folder2");

            var meta = await _repository.GetIndexMetadataAsync();
            Assert.Equal(1, meta.TotalEmailsIndexed);

            var search = await _searchService.ExecuteSearchAsync("Moved Email Topic");
            Assert.Single(search.Items);
            Assert.Equal("Archive", search.Items[0].FolderName);
        }

        [Fact]
        public async Task IncrementalDuplicate_10_50_100_EventsResultInSingleDatabaseRecord()
        {
            var mockOutlook = new MockOutlookAdapter();
            var doc = new EmailDocument
            {
                StoreId = "Store1",
                EntryId = "Entry_Dup_1",
                Subject = "High Frequency Event Burst",
                FolderName = "Inbox",
                FolderEntryId = "Folder1",
                ReceivedDate = DateTime.Now
            };
            mockOutlook.EmailStorage[$"{doc.StoreId}:{doc.EntryId}"] = doc;

            var indexer = new EmailIndexerService(mockOutlook, _repository, new MockLicensingService());

            // Queue 100 duplicate events
            for (int i = 0; i < 100; i++)
            {
                await indexer.QueueItemForIndexingAsync("Store1", "Entry_Dup_1", "Folder1");
            }

            var meta = await _repository.GetIndexMetadataAsync();
            Assert.Equal(1, meta.TotalEmailsIndexed);

            var search = await _searchService.ExecuteSearchAsync("High Frequency Event");
            Assert.Single(search.Items);
        }

        [Fact]
        public async Task IncrementalBurst_Handles1000IncomingItemsConsistently()
        {
            var mockOutlook = new MockOutlookAdapter();
            var indexer = new EmailIndexerService(mockOutlook, _repository, new MockLicensingService());

            int burstCount = 1000;
            for (int i = 1; i <= burstCount; i++)
            {
                var doc = new EmailDocument
                {
                    StoreId = "Store1",
                    EntryId = $"Burst_Entry_{i}",
                    Subject = $"Burst Notification Item {i}",
                    BodyText = $"Content for burst performance test item number {i}.",
                    FolderName = "Inbox",
                    FolderEntryId = "Folder1",
                    ReceivedDate = DateTime.Now
                };
                mockOutlook.EmailStorage[$"{doc.StoreId}:{doc.EntryId}"] = doc;
            }

            for (int i = 1; i <= burstCount; i++)
            {
                await indexer.QueueItemForIndexingAsync("Store1", $"Burst_Entry_{i}", "Folder1");
            }

            var meta = await _repository.GetIndexMetadataAsync();
            Assert.Equal(burstCount, meta.TotalEmailsIndexed);

            var search = await _repository.SearchAsync(new SearchQuery { RawQuery = "Burst Notification", Limit = 2000 });
            Assert.Equal(burstCount, search.TotalCount);
        }

        [Fact]
        public async Task IncrementalRetry_RetriesTransientFailuresUpTo3Times()
        {
            var mockOutlook = new MockOutlookAdapter
            {
                TransientFailuresRemaining = 2 // Fails twice, succeeds on 3rd attempt
            };

            var doc = new EmailDocument
            {
                StoreId = "Store1",
                EntryId = "Entry_Retry_1",
                Subject = "Transient COM Error Recovery",
                FolderName = "Inbox",
                FolderEntryId = "Folder1",
                ReceivedDate = DateTime.Now
            };
            mockOutlook.EmailStorage[$"{doc.StoreId}:{doc.EntryId}"] = doc;

            var indexer = new EmailIndexerService(mockOutlook, _repository, new MockLicensingService());
            await indexer.QueueItemForIndexingAsync("Store1", "Entry_Retry_1", "Folder1");

            int attempts = 0;
            var search = await _searchService.ExecuteSearchAsync("Transient COM Error");
            while (search.Items.Count == 0 && attempts++ < 30)
            {
                await Task.Delay(200);
                search = await _searchService.ExecuteSearchAsync("Transient COM Error");
            }

            Assert.Single(search.Items);
            Assert.True(mockOutlook.FetchAttemptsCount >= 3);
        }

        [Fact]
        public async Task FtsConsistency_VerifySearchQueryModifications()
        {
            var mockOutlook = new MockOutlookAdapter();
            var doc = new EmailDocument
            {
                StoreId = "Store1",
                EntryId = "FTS_Item_1",
                Subject = "Project Alpha Status",
                BodyText = "Initial keyword: ZEUS",
                FolderName = "Inbox",
                FolderEntryId = "Folder1",
                ReceivedDate = DateTime.Now
            };
            mockOutlook.EmailStorage[$"{doc.StoreId}:{doc.EntryId}"] = doc;

            var indexer = new EmailIndexerService(mockOutlook, _repository, new MockLicensingService());
            await indexer.QueueItemForIndexingAsync("Store1", "FTS_Item_1", "Folder1");

            // 1. Initial term search
            var r1 = await _searchService.ExecuteSearchAsync("ZEUS");
            Assert.Single(r1.Items);

            // 2. Modify body keyword from ZEUS to POSEIDON
            doc.BodyText = "Updated keyword: POSEIDON";
            await indexer.QueueItemForIndexingAsync("Store1", "FTS_Item_1", "Folder1");

            // Old search term must NOT match
            var r2Old = await _searchService.ExecuteSearchAsync("ZEUS");
            Assert.Empty(r2Old.Items);

            // New search term MUST match
            var r2New = await _searchService.ExecuteSearchAsync("POSEIDON");
            Assert.Single(r2New.Items);
        }

        [Fact]
        public async Task IncrementalSearchConcurrency_SearchRemainsResponsiveDuringQueueProcessing()
        {
            var mockOutlook = new MockOutlookAdapter();
            var indexer = new EmailIndexerService(mockOutlook, _repository, new MockLicensingService());

            // Pre-seed database with 100 items
            for (int i = 1; i <= 100; i++)
            {
                var doc = new EmailDocument
                {
                    StoreId = "Store1",
                    EntryId = $"Pre_Entry_{i}",
                    Subject = $"Invoice Confirmation {i}",
                    FolderName = "Inbox",
                    FolderEntryId = "Folder1",
                    ReceivedDate = DateTime.Now
                };
                mockOutlook.EmailStorage[$"{doc.StoreId}:{doc.EntryId}"] = doc;
                await indexer.QueueItemForIndexingAsync("Store1", $"Pre_Entry_{i}", "Folder1");
            }

            // Perform searches directly against repository while queueing 100 new items
            var searchTasks = new List<Task<SearchResultContainer>>();
            for (int i = 101; i <= 200; i++)
            {
                var doc = new EmailDocument
                {
                    StoreId = "Store1",
                    EntryId = $"Pre_Entry_{i}",
                    Subject = $"Invoice Confirmation {i}",
                    FolderName = "Inbox",
                    FolderEntryId = "Folder1",
                    ReceivedDate = DateTime.Now
                };
                mockOutlook.EmailStorage[$"{doc.StoreId}:{doc.EntryId}"] = doc;
                _ = indexer.QueueItemForIndexingAsync("Store1", $"Pre_Entry_{i}", "Folder1");

                searchTasks.Add(_repository.SearchAsync(new SearchQuery { RawQuery = "Invoice Confirmation", Limit = 200 }));
            }

            var results = await Task.WhenAll(searchTasks);
            Assert.NotEmpty(results);
            Assert.All(results, res => Assert.True(res.TotalCount >= 100));
        }
    }
}
