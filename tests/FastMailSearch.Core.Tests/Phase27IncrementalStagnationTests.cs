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
    public class Phase27IncrementalStagnationTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly DatabaseConnectionFactory _factory;
        private readonly SqliteEmailRepository _repository;

        public Phase27IncrementalStagnationTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), $"Phase27Tests_{Guid.NewGuid():N}.db");
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
            public string GetOutlookVersion() => "16.0.15000.20000";
            public string GetOfficeBitness() => "64-bit";
            public List<MailboxInfo> GetMailboxes() => new List<MailboxInfo>
            {
                new MailboxInfo { StoreId = "Store1", DisplayName = "Test Mailbox" }
            };
            public List<MailboxFolder> GetFolders(string storeId) => new List<MailboxFolder>
            {
                new MailboxFolder { StoreId = storeId, EntryId = "Folder1", Name = "Inbox", ItemCount = EmailStorage.Count }
            };
            public List<EmailDocument> FetchFolderEmailsBatch(string storeId, string folderEntryId, int startIndex, int batchSize, Func<bool>? cancelCheck = null)
            {
                return EmailStorage.Values.Skip(startIndex - 1).Take(batchSize).ToList();
            }
            public EmailDocument? FetchSingleEmail(string storeId, string entryId)
            {
                FetchAttemptsCount++;
                string key = $"{storeId}:{entryId}";
                return EmailStorage.TryGetValue(key, out var doc) ? doc : null;
            }
            public bool OpenEmailInOutlook(string storeId, string entryId) => true;
            public bool IsOutboxFolder(string storeId, string folderEntryId) => false;
            public void SubscribeToEvents(Action<string, string, string> onItemAddedOrChanged, Action<string, string> onItemRemoved) { }
            public void UnsubscribeFromEvents() { }
        }

        [Fact]
        public async Task IncrementalProcessing_DrainsEntireDatabaseQueue_Automatically()
        {
            // Seed 2500 pending items directly into the DB to simulate a backlog that exceeds the 1000-item memory limit
            for (int i = 0; i < 2500; i++)
            {
                await _repository.EnqueuePendingChangeAsync("Store1", $"Entry_{i}", "Add", "Folder1");
            }

            var mockOutlook = new MockOutlookAdapter();
            for (int i = 0; i < 2500; i++)
            {
                var doc = new EmailDocument
                {
                    StoreId = "Store1",
                    EntryId = $"Entry_{i}",
                    Subject = $"Subject {i}",
                    FolderEntryId = "Folder1",
                    ReceivedDate = DateTime.Now
                };
                mockOutlook.EmailStorage[$"Store1:Entry_{i}"] = doc;
            }

            var indexer = new EmailIndexerService(mockOutlook, _repository, new MockLicensingService(LicenseStatus.Valid));
            // Wait for queue processing (it normally runs asynchronously when checking for incomplete index)
            await indexer.CheckAndRecoverIncompleteIndexAsync();

            // Allow time for the background task to complete all 3 batches (1000, 1000, 500)
            int attempts = 0;
            while (true)
            {
                int remaining = await _repository.GetPendingChangesCountAsync();
                if (remaining == 0) break;
                await Task.Delay(200);
                attempts++;
                if (attempts > 750) break; // timeout after 150 seconds
            }
            
            // Allow processing to catch up
            await Task.Delay(1000);

            var pendingCount = await _repository.GetPendingChangesCountAsync();
            var indexedCount = await _repository.GetIndexedEmailCountAsync();
            var status = indexer.GetStatus();

            Assert.Equal(0, pendingCount);
            Assert.Equal(2500, indexedCount);
            
            // Verifies the UI no longer says "Updating" when DB is empty
            Assert.DoesNotContain("Updating", status.CurrentStatusMessage);
        }
        
        [Fact]
        public async Task IncrementalProcessing_Drains3712Items_WithoutRestart()
        {
            // Seed 3712 pending items directly into the DB
            for (int i = 0; i < 3712; i++)
            {
                await _repository.EnqueuePendingChangeAsync("Store1", $"Entry_{i}", "Add", "Folder1");
            }

            var mockOutlook = new MockOutlookAdapter();
            for (int i = 0; i < 3712; i++)
            {
                var doc = new EmailDocument
                {
                    StoreId = "Store1",
                    EntryId = $"Entry_{i}",
                    Subject = $"Subject {i}",
                    FolderEntryId = "Folder1",
                    ReceivedDate = DateTime.Now
                };
                mockOutlook.EmailStorage[$"Store1:Entry_{i}"] = doc;
            }

            var indexer = new EmailIndexerService(mockOutlook, _repository, new MockLicensingService(LicenseStatus.Valid));
            await indexer.CheckAndRecoverIncompleteIndexAsync();

            int attempts = 0;
            while (true)
            {
                int remaining = await _repository.GetPendingChangesCountAsync();
                if (remaining == 0) break;
                await Task.Delay(200);
                attempts++;
                if (attempts > 750) break; // timeout after 150 seconds
            }
            
            await Task.Delay(1000);

            var pendingCount = await _repository.GetPendingChangesCountAsync();
            var indexedCount = await _repository.GetIndexedEmailCountAsync();

            Assert.Equal(0, pendingCount);
            Assert.Equal(3712, indexedCount);
        }
    }
}
