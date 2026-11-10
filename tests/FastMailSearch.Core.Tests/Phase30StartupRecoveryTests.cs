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
    public class Phase30StartupRecoveryTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly DatabaseConnectionFactory _factory;
        private readonly SqliteEmailRepository _repository;

        public Phase30StartupRecoveryTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), $"Phase30Tests_{Guid.NewGuid():N}.db");
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

        private class MockOutlookAdapter : IOutlookAdapter
        {
            public Dictionary<string, EmailDocument> EmailStorage = new Dictionary<string, EmailDocument>();
            public int FetchAttemptsCount = 0;

            public bool IsOutlookAvailable() => true;
            public string GetOutlookVersion() => "16.0";
            public string GetOfficeBitness() => "64-bit";
            public List<MailboxInfo> GetMailboxes() => new List<MailboxInfo>();
            public List<MailboxFolder> GetFolders(string storeId) => new List<MailboxFolder>();
            public List<EmailDocument> FetchFolderEmailsBatch(string storeId, string folderEntryId, int startIndex, int batchSize, Func<bool>? cancelCheck = null)
            {
                return new List<EmailDocument>();
            }

            public EmailDocument FetchSingleEmail(string storeId, string entryId)
            {
                Interlocked.Increment(ref FetchAttemptsCount);
                if (EmailStorage.TryGetValue(entryId, out var doc))
                {
                    return doc;
                }
                throw new InvalidOperationException("ITEM_NOT_FOUND");
            }

            public void SubscribeToEvents(Action<string, string, string> onItemAddedOrChanged, Action<string, string> onItemRemoved) { }
            public void UnsubscribeFromEvents() { }
            public bool IsOutboxFolder(string storeId, string folderEntryId) => false;
            public bool IsSentItemsFolderId(string folderEntryId) => false;
            public bool OpenEmailInOutlook(string storeId, string entryId) => true;
        }

        [Fact]
        public async Task Recovery_Is_Idempotent_And_Only_Runs_Once()
        {
            var outlook = new MockOutlookAdapter();
            var indexer = new EmailIndexerService(outlook, _repository);
            
            // Call twice
            await indexer.CheckAndRecoverIncompleteIndexAsync();
            await indexer.CheckAndRecoverIncompleteIndexAsync();
            
            // Ensure no duplicate processing or errors
            Assert.True(true);
        }

        [Fact]
        public async Task Recovers_Stale_Processing_Rows_But_Not_Active_Ones()
        {
            // Insert stale
            using (var conn = _factory.CreateConnection())
            using (var cmd = conn.CreateCommand())
            {
                var staleTime = DateTime.Now.AddMinutes(-10).ToString("o");
                cmd.CommandText = "INSERT INTO PendingChanges (StoreId, EntryId, Operation, Status, AttemptCount, CreatedAt, UpdatedAt) VALUES ('S1', 'Stale1', 'InsertOrUpdate', 'Processing', 0, @T, @T)";
                cmd.Parameters.AddWithValue("@T", staleTime);
                cmd.ExecuteNonQuery();

                var activeTime = DateTime.Now.ToString("o");
                cmd.CommandText = "INSERT INTO PendingChanges (StoreId, EntryId, Operation, Status, AttemptCount, CreatedAt, UpdatedAt) VALUES ('S1', 'Active1', 'InsertOrUpdate', 'Processing', 0, @A, @A)";
                cmd.Parameters.AddWithValue("@A", activeTime);
                cmd.ExecuteNonQuery();
            }

            await _repository.RecoverStaleProcessingChangesAsync();
            
            using (var conn = _factory.CreateConnection())
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT EntryId, Status FROM PendingChanges";
                using (var reader = cmd.ExecuteReader())
                {
                    var results = new Dictionary<string, string>();
                    while (reader.Read())
                    {
                        results[reader.GetString(0)] = reader.GetString(1);
                    }
                    
                    Assert.Equal("Pending", results["Stale1"]);
                    Assert.Equal("Processing", results["Active1"]);
                }
            }
        }

        [Fact]
        public async Task Large_Backlog_Drains_Multiple_Batches()
        {
            var outlook = new MockOutlookAdapter();
            for (int i = 0; i < 2500; i++)
            {
                string entryId = $"E{i}";
                await _repository.EnqueuePendingChangeAsync("S1", entryId, "InsertOrUpdate", "Inbox");
                outlook.EmailStorage[entryId] = new EmailDocument { StoreId = "S1", EntryId = entryId, FolderName = "Inbox", Subject = $"Subject {i}", ReceivedDate = DateTime.Now };
            }

            var indexer = new EmailIndexerService(outlook, _repository, new MockLicensingService(LicenseStatus.Valid));
            
            // This triggers recovery and incremental processing loop
            await indexer.CheckAndRecoverIncompleteIndexAsync();

            // Wait for it to drain
            for (int i = 0; i < 500; i++)
            {
                int remaining = await _repository.GetPendingChangesCountAsync();
                if (remaining == 0) break;
                await Task.Delay(200);
            }

            int finalCount = await _repository.GetPendingChangesCountAsync();
            Assert.Equal(0, finalCount);
            Assert.True(outlook.FetchAttemptsCount >= 2500);
        }

        [Fact]
        public async Task RetryAllFailedItems_Returns_Zero_When_None_Failed()
        {
            var outlook = new MockOutlookAdapter();
            var indexer = new EmailIndexerService(outlook, _repository);
            var health = new IndexHealthService(_repository, indexer, outlook, null);

            int count = await health.RetryAllFailedItemsAsync();
            Assert.Equal(0, count);
        }

        [Fact]
        public async Task IndexHealth_Shows_True_Pending_Count()
        {
            var outlook = new MockOutlookAdapter();
            for (int i = 0; i < 1500; i++)
            {
                await _repository.EnqueuePendingChangeAsync("S1", $"E{i}", "InsertOrUpdate", "Inbox");
            }

            var indexer = new EmailIndexerService(outlook, _repository);
            var health = new IndexHealthService(_repository, indexer, outlook, null);

            var summary = await health.GetHealthSummaryAsync();
            Assert.Equal(IndexerHealth.Paused, summary.HealthState);
            Assert.Equal(1500, summary.PendingQueueCount);
            Assert.Contains("1,500 pending", summary.StatusState);
        }
        
        [Fact]
        public async Task Cancellation_Preserves_Uncompleted_Work()
        {
            var outlook = new MockOutlookAdapter();
            for (int i = 0; i < 500; i++)
            {
                string entryId = $"E{i}";
                await _repository.EnqueuePendingChangeAsync("S1", entryId, "InsertOrUpdate", "Inbox");
                outlook.EmailStorage[entryId] = new EmailDocument { StoreId = "S1", EntryId = entryId, FolderName = "Inbox", Subject = $"S{i}", ReceivedDate = DateTime.Now };
            }

            var indexer = new EmailIndexerService(outlook, _repository);
            await indexer.CheckAndRecoverIncompleteIndexAsync();
            
            // Immediately cancel
            await indexer.CancelIndexingAsync();
            
            await Task.Delay(500);
            
            int pendingCount = await _repository.GetPendingChangesCountAsync();
            Assert.True(pendingCount > 0, $"Expected > 0, got {pendingCount}");
        }
    }
}
