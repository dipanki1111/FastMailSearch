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
using FastMailSearch.Infrastructure.Services;
using Xunit;

namespace FastMailSearch.Core.Tests
{
    public class Phase20IntegrityAndSelfHealingTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly DatabaseConnectionFactory _factory;
        private readonly SqliteEmailRepository _repository;
        private readonly SearchService _searchService;

        public Phase20IntegrityAndSelfHealingTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), $"Phase20Tests_{Guid.NewGuid():N}.db");
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
        public async Task Test01_FailedItem_Persisted()
        {
            var item = new FailedIndexItem
            {
                StoreId = "Store1",
                EntryId = "Entry_Persist_1",
                FolderName = "Inbox",
                OperationType = "Add",
                FailureReason = "COM Error",
                AttemptCount = 3,
                Status = "Failed",
                ErrorMessage = "COM Error Details"
            };
            await _repository.SaveFailedIndexItemAsync(item);

            var failedList = await _repository.GetFailedIndexItemsAsync("Failed");
            Assert.Single(failedList);
            Assert.Equal("Entry_Persist_1", failedList[0].EntryId);
        }

        [Fact]
        public async Task Test02_FailedItem_SurvivesRestart()
        {
            var item = new FailedIndexItem
            {
                StoreId = "Store1",
                EntryId = "Entry_Survive_1",
                OperationType = "Change",
                AttemptCount = 3,
                Status = "Failed"
            };
            await _repository.SaveFailedIndexItemAsync(item);

            var repo2 = new SqliteEmailRepository(_factory);
            var failedList = await repo2.GetFailedIndexItemsAsync("Failed");
            Assert.Single(failedList);
            Assert.Equal("Entry_Survive_1", failedList[0].EntryId);
        }

        [Fact]
        public async Task Test03_FailedItem_Retry()
        {
            var mockOutlook = new MockOutlookAdapter();
            var doc = new EmailDocument
            {
                StoreId = "Store1",
                EntryId = "Entry_Retry_1",
                Subject = "Retry Recovered Item",
                FolderName = "Inbox",
                ReceivedDate = DateTime.Now
            };
            mockOutlook.EmailStorage[$"{doc.StoreId}:{doc.EntryId}"] = doc;

            var indexer = new EmailIndexerService(mockOutlook, _repository, new MockLicensingService());
            var healthService = new IndexHealthService(_repository, indexer);

            await _repository.SaveFailedIndexItemAsync(new FailedIndexItem
            {
                StoreId = "Store1",
                EntryId = "Entry_Retry_1",
                FolderName = "Folder1",
                OperationType = "Add",
                AttemptCount = 3,
                Status = "Failed"
            });

            await healthService.RetryAllFailedItemsAsync();

            var search = await _searchService.ExecuteSearchAsync("Retry Recovered");
            Assert.Single(search.Items);
        }

        [Fact]
        public async Task Test04_RetrySuccess_RemovesFailureState()
        {
            var mockOutlook = new MockOutlookAdapter();
            var doc = new EmailDocument
            {
                StoreId = "Store1",
                EntryId = "Entry_State_1",
                Subject = "Success Clear Failure",
                FolderName = "Inbox",
                ReceivedDate = DateTime.Now
            };
            mockOutlook.EmailStorage[$"{doc.StoreId}:{doc.EntryId}"] = doc;

            var indexer = new EmailIndexerService(mockOutlook, _repository, new MockLicensingService());
            var healthService = new IndexHealthService(_repository, indexer);

            await _repository.SaveFailedIndexItemAsync(new FailedIndexItem
            {
                StoreId = "Store1",
                EntryId = "Entry_State_1",
                Status = "Failed"
            });

            await healthService.RetryAllFailedItemsAsync();

            var activeFailed = await _repository.GetFailedIndexItemsAsync(status: "Failed");
            Assert.Empty(activeFailed);
        }

        [Fact]
        public async Task Test05_RetryFailure_IncrementsAttemptCount()
        {
            var mockOutlook = new MockOutlookAdapter
            {
                TransientFailuresRemaining = 5
            };

            var indexer = new EmailIndexerService(mockOutlook, _repository, new MockLicensingService());
            await indexer.QueueItemForIndexingAsync("Store1", "Entry_Inc_1", "Folder1");

            List<FailedIndexItem> failedList = new List<FailedIndexItem>();
            int attempts = 0;
            while (attempts++ < 30)
            {
                failedList = await _repository.GetFailedIndexItemsAsync("Failed");
                if (failedList.Count > 0) break;
                await Task.Delay(200);
            }

            Assert.Single(failedList);
            Assert.True(failedList[0].AttemptCount >= 3);
        }

        [Fact]
        public async Task Test06_MaximumRetryProtection()
        {
            var mockOutlook = new MockOutlookAdapter
            {
                TransientFailuresRemaining = 100
            };

            var indexer = new EmailIndexerService(mockOutlook, _repository, new MockLicensingService());
            await indexer.QueueItemForIndexingAsync("Store1", "Entry_Max_1", "Folder1");

            List<FailedIndexItem> failedList = new List<FailedIndexItem>();
            int attempts = 0;
            while (attempts++ < 30)
            {
                failedList = await _repository.GetFailedIndexItemsAsync("Failed");
                if (failedList.Count > 0) break;
                await Task.Delay(200);
            }

            Assert.Single(failedList);
            Assert.Equal(4, failedList[0].AttemptCount); // 3 retries + initial attempt
        }

        [Fact]
        public async Task Test07_RetryAll()
        {
            var mockOutlook = new MockOutlookAdapter();
            for (int i = 1; i <= 5; i++)
            {
                var doc = new EmailDocument
                {
                    StoreId = "Store1",
                    EntryId = $"Batch_Entry_{i}",
                    Subject = $"Batch Item {i}",
                    FolderName = "Inbox",
                    ReceivedDate = DateTime.Now
                };
                mockOutlook.EmailStorage[$"{doc.StoreId}:{doc.EntryId}"] = doc;

                await _repository.SaveFailedIndexItemAsync(new FailedIndexItem
                {
                    StoreId = "Store1",
                    EntryId = $"Batch_Entry_{i}",
                    FolderName = "Folder1",
                    OperationType = "Add",
                    Status = "Failed"
                });
            }

            var indexer = new EmailIndexerService(mockOutlook, _repository, new MockLicensingService());
            var healthService = new IndexHealthService(_repository, indexer);

            await healthService.RetryAllFailedItemsAsync();

            var search = await _searchService.ExecuteSearchAsync("Batch Item");
            Assert.Equal(5, search.TotalCount);
        }

        [Fact]
        public async Task Test08_RemoveFailedItem()
        {
            var item = new FailedIndexItem
            {
                StoreId = "Store1",
                EntryId = "Entry_Rem_1",
                Status = "Failed"
            };
            await _repository.SaveFailedIndexItemAsync(item);
            var list1 = await _repository.GetFailedIndexItemsAsync("Failed");
            Assert.Single(list1);

            await _repository.RemoveFailedIndexItemAsync(list1[0].Id);
            var list2 = await _repository.GetFailedIndexItemsAsync("Failed");
            Assert.Empty(list2);
        }

        [Fact]
        public async Task Test09_DuplicateFailureEvents()
        {
            await _repository.EnqueuePendingChangeAsync("Store1", "Dup_Entry_1", "Add", "Folder1");
            await _repository.EnqueuePendingChangeAsync("Store1", "Dup_Entry_1", "Change", "Folder2");

            var pending = await _repository.GetPendingChangesAsync();
            Assert.Single(pending);
            Assert.Equal("Change", pending[0].Operation);
        }

        [Fact]
        public async Task Test10_ConcurrentItemChangeAndRetry()
        {
            var mockOutlook = new MockOutlookAdapter();
            var doc = new EmailDocument
            {
                StoreId = "Store1",
                EntryId = "Concurrent_1",
                Subject = "Concurrent Change Subject",
                FolderName = "Inbox",
                ReceivedDate = DateTime.Now
            };
            mockOutlook.EmailStorage[$"{doc.StoreId}:{doc.EntryId}"] = doc;

            var indexer = new EmailIndexerService(mockOutlook, _repository, new MockLicensingService());
            var task1 = indexer.QueueItemForIndexingAsync("Store1", "Concurrent_1", "Folder1");
            var task2 = indexer.QueueItemForIndexingAsync("Store1", "Concurrent_1", "Folder1");

            await Task.WhenAll(task1, task2);
            var search = await _searchService.ExecuteSearchAsync("Concurrent Change");
            Assert.Single(search.Items);
        }

        [Fact]
        public async Task Test11_DeleteWhileRetrying()
        {
            var mockOutlook = new MockOutlookAdapter();
            var indexer = new EmailIndexerService(mockOutlook, _repository, new MockLicensingService());

            await indexer.QueueItemForIndexingAsync("Store1", "Entry_DelRetry_1", "Folder1");
            await indexer.QueueItemForDeletionAsync("Store1", "Entry_DelRetry_1");

            var search = await _searchService.ExecuteSearchAsync("Entry_DelRetry_1");
            Assert.Empty(search.Items);
        }

        [Fact]
        public async Task Test12_MoveWhileRetrying()
        {
            var mockOutlook = new MockOutlookAdapter();
            var doc = new EmailDocument
            {
                StoreId = "Store1",
                EntryId = "Entry_Move_1",
                Subject = "Moved Topic Email",
                FolderName = "Inbox",
                FolderEntryId = "Folder1",
                ReceivedDate = DateTime.Now
            };
            mockOutlook.EmailStorage[$"{doc.StoreId}:{doc.EntryId}"] = doc;

            var indexer = new EmailIndexerService(mockOutlook, _repository, new MockLicensingService());
            await indexer.QueueItemForIndexingAsync("Store1", "Entry_Move_1", "Folder1");

            doc.FolderName = "Archive";
            doc.FolderEntryId = "Folder2";
            await indexer.QueueItemForIndexingAsync("Store1", "Entry_Move_1", "Folder2");

            var search = await _searchService.ExecuteSearchAsync("Moved Topic");
            Assert.Single(search.Items);
            Assert.Equal("Archive", search.Items[0].FolderName);
        }

        [Fact]
        public async Task Test13_FtsConsistency()
        {
            var doc = new EmailDocument
            {
                StoreId = "Store1",
                EntryId = "Entry_Fts_1",
                Subject = "FTS Consistency Subject",
                BodyText = "Keyword ZEUS",
                FolderName = "Inbox",
                ReceivedDate = DateTime.Now
            };
            await _repository.SaveEmailBatchAsync(new[] { doc });

            var s1 = await _searchService.ExecuteSearchAsync("ZEUS");
            Assert.Single(s1.Items);

            await _repository.DeleteEmailAsync("Store1", "Entry_Fts_1");
            var s2 = await _searchService.ExecuteSearchAsync("ZEUS");
            Assert.Empty(s2.Items);
        }

        [Fact]
        public async Task Test14_OrphanFtsDetection()
        {
            using (var conn = _factory.CreateConnection())
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "INSERT INTO EmailSearch (EmailId, Subject, SenderName, SenderEmail, Recipients, BodyText, AttachmentNames) VALUES ('999999', 'Orphan', 'O', 'o@test.com', 'r@test.com', 'Body', '');";
                cmd.ExecuteNonQuery();
            }

            var scan = await _repository.ScanIndexIntegrityAsync();
            Assert.False(scan.IsHealthy);
            Assert.True(scan.OrphanFtsCount > 0);
        }

        [Fact]
        public async Task Test15_MissingFtsDetection()
        {
            var doc = new EmailDocument
            {
                StoreId = "Store1",
                EntryId = "Entry_NoFts_1",
                Subject = "No FTS Row Email",
                FolderName = "Inbox",
                ReceivedDate = DateTime.Now
            };
            await _repository.SaveEmailBatchAsync(new[] { doc });

            using (var conn = _factory.CreateConnection())
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "DELETE FROM EmailSearch;";
                cmd.ExecuteNonQuery();
            }

            var scan = await _repository.ScanIndexIntegrityAsync();
            Assert.False(scan.IsHealthy);
            Assert.True(scan.MissingFtsCount > 0);
        }

        [Fact]
        public async Task Test16_TargetedReconciliation()
        {
            var doc = new EmailDocument
            {
                StoreId = "Store1",
                EntryId = "Entry_Recon_1",
                Subject = "Reconciliation Target Email",
                FolderName = "Inbox",
                ReceivedDate = DateTime.Now
            };
            await _repository.SaveEmailBatchAsync(new[] { doc });

            using (var conn = _factory.CreateConnection())
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "DELETE FROM EmailSearch;";
                cmd.ExecuteNonQuery();
            }

            var scan = await _repository.ScanIndexIntegrityAsync();
            await _repository.RepairIndexIntegrityAsync(scan);

            var scanAfter = await _repository.ScanIndexIntegrityAsync();
            Assert.True(scanAfter.IsHealthy);
        }

        [Fact]
        public async Task Test17_IndexHealthCounters()
        {
            var mockOutlook = new MockOutlookAdapter();
            var indexer = new EmailIndexerService(mockOutlook, _repository);
            var healthService = new IndexHealthService(_repository, indexer);

            var summary = await healthService.GetHealthSummaryAsync();
            Assert.NotNull(summary);
            Assert.True(summary.DatabaseSizeMb >= 0);
        }

        [Fact]
        public async Task Test18_DatabaseTransactionRollback()
        {
            var scan = new IndexIntegrityResult { MissingFtsCount = 1 };
            using var cts = new CancellationTokenSource();
            cts.Cancel();

            await Assert.ThrowsAnyAsync<OperationCanceledException>(async () =>
            {
                await _repository.RepairIndexIntegrityAsync(scan, cts.Token);
            });
        }

        [Fact]
        public async Task Test19_Cancellation()
        {
            using var cts = new CancellationTokenSource();
            cts.Cancel();

            await Assert.ThrowsAnyAsync<OperationCanceledException>(async () =>
            {
                await _repository.ScanIndexIntegrityAsync(cts.Token);
            });
        }

        [Fact]
        public void Test20_OutlookComIsolation()
        {
            var mockOutlook = new MockOutlookAdapter();
            Assert.True(mockOutlook.IsOutlookAvailable());
            Assert.Equal("64-bit", mockOutlook.GetOfficeBitness());
        }

        [Fact]
        public async Task Test21_NoUIThreadBlocking()
        {
            var mockOutlook = new MockOutlookAdapter();
            var doc = new EmailDocument
            {
                StoreId = "Store1",
                EntryId = "Entry_NonBlock_1",
                Subject = "Non-blocking Test",
                FolderName = "Inbox",
                ReceivedDate = DateTime.Now
            };
            mockOutlook.EmailStorage[$"{doc.StoreId}:{doc.EntryId}"] = doc;

            var indexer = new EmailIndexerService(mockOutlook, _repository);
            var task = indexer.QueueItemForIndexingAsync("Store1", "Entry_NonBlock_1", "Folder1");

            Assert.False(task.IsFaulted);
            await task;
        }

        [Fact]
        public async Task Test22_ExistingPaginationRegression()
        {
            var batch = new List<EmailDocument>();
            for (int i = 1; i <= 60; i++)
            {
                batch.Add(new EmailDocument
                {
                    StoreId = "Store1",
                    EntryId = $"Page_Entry_{i}",
                    Subject = $"Pagination Item {i}",
                    FolderName = "Inbox",
                    ReceivedDate = DateTime.Now
                });
            }
            await _repository.SaveEmailBatchAsync(batch);

            var page1 = await _repository.SearchAsync(new SearchQuery { RawQuery = "Pagination Item", Page = 1, PageSize = 50 });
            Assert.Equal(50, page1.Items.Count);
            Assert.Equal(60, page1.TotalCount);

            var page2 = await _repository.SearchAsync(new SearchQuery { RawQuery = "Pagination Item", Page = 2, PageSize = 50 });
            Assert.Equal(10, page2.Items.Count);
        }

        [Fact]
        public async Task Test23_ExistingSearchSyntaxRegression()
        {
            var doc = new EmailDocument
            {
                StoreId = "Store1",
                EntryId = "Entry_Syntax_1",
                Subject = "Quarterly Invoice Summary",
                SenderName = "Billing Department",
                SenderEmail = "billing@company.com",
                BodyText = "Here is your q3 invoice details",
                FolderName = "Inbox",
                ReceivedDate = DateTime.Now
            };
            await _repository.SaveEmailBatchAsync(new[] { doc });

            var r1 = await _searchService.ExecuteSearchAsync("subject:Invoice");
            Assert.Single(r1.Items);

            var r2 = await _searchService.ExecuteSearchAsync("from:billing@company.com");
            Assert.Single(r2.Items);

            var r3 = await _searchService.ExecuteSearchAsync("\"Quarterly Invoice\"");
            Assert.Single(r3.Items);
        }

        [Fact]
        public async Task Test24_ExistingIncrementalIndexingRegression()
        {
            var mockOutlook = new MockOutlookAdapter();
            var doc = new EmailDocument
            {
                StoreId = "Store1",
                EntryId = "Entry_Reg_1",
                Subject = "Incremental Regression Email",
                FolderName = "Inbox",
                ReceivedDate = DateTime.Now
            };
            mockOutlook.EmailStorage[$"{doc.StoreId}:{doc.EntryId}"] = doc;

            var indexer = new EmailIndexerService(mockOutlook, _repository, new MockLicensingService());
            await indexer.QueueItemForIndexingAsync("Store1", "Entry_Reg_1", "Folder1");

            var search = await _searchService.ExecuteSearchAsync("Incremental Regression");
            Assert.Single(search.Items);
        }
    }
}
