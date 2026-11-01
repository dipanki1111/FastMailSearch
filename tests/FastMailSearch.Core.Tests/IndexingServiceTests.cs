using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using FastMailSearch.Core.Interfaces;
using FastMailSearch.Core.Models;
using FastMailSearch.Infrastructure.Database;
using FastMailSearch.Infrastructure.Indexing;
using Xunit;

namespace FastMailSearch.Core.Tests
{
    public class FakeOutlookAdapter : IOutlookAdapter
    {
        public bool Available { get; set; } = true;
        public List<MailboxInfo> Mailboxes { get; set; } = new List<MailboxInfo>();
        public Dictionary<string, List<MailboxFolder>> FoldersMap { get; set; } = new Dictionary<string, List<MailboxFolder>>();

        public bool IsOutlookAvailable() => Available;
        public string GetOutlookVersion() => "16.0";
        public string GetOfficeBitness() => "64-bit";
        public List<MailboxInfo> GetMailboxes() => Mailboxes;
        public List<MailboxFolder> GetFolders(string storeId) => FoldersMap.ContainsKey(storeId) ? FoldersMap[storeId] : new List<MailboxFolder>();

        public List<EmailDocument> FetchFolderEmailsBatch(string storeId, string folderEntryId, int startIndex, int batchSize, Func<bool>? cancelCheck = null)
        {
            var list = new List<EmailDocument>();
            for (int i = 0; i < batchSize && (startIndex + i <= 10); i++)
            {
                if (cancelCheck != null && cancelCheck()) break;
                list.Add(new EmailDocument
                {
                    StoreId = storeId,
                    EntryId = $"{folderEntryId}_email_{startIndex + i}",
                    Subject = $"Subject {startIndex + i}",
                    SenderName = "Sender",
                    SenderEmail = "sender@test.com",
                    ReceivedDate = DateTime.Now,
                    FolderName = "Inbox"
                });
            }
            return list;
        }

        public EmailDocument? FetchSingleEmail(string storeId, string entryId) => null;
        public bool OpenEmailInOutlook(string storeId, string entryId) => true;
        public bool IsOutboxFolder(string storeId, string folderEntryId) => false;
        public void SubscribeToEvents(Action<string, string, string> onItemAddedOrChanged, Action<string, string> onItemRemoved) { }
        public void UnsubscribeFromEvents() { }
    }

    public class IndexingServiceTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly DatabaseConnectionFactory _factory;
        private readonly SqliteEmailRepository _repository;
        private readonly FakeOutlookAdapter _outlookAdapter;
        private readonly EmailIndexerService _indexerService;

        public IndexingServiceTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), $"IndexerTest_{Guid.NewGuid():N}.db");
            _factory = new DatabaseConnectionFactory(_dbPath);
            _repository = new SqliteEmailRepository(_factory);
            _repository.InitializeDatabase();

            _outlookAdapter = new FakeOutlookAdapter
            {
                Mailboxes = new List<MailboxInfo>
                {
                    new MailboxInfo { StoreId = "Store1", DisplayName = "Test Mailbox" }
                },
                FoldersMap = new Dictionary<string, List<MailboxFolder>>
                {
                    ["Store1"] = new List<MailboxFolder>
                    {
                        new MailboxFolder { StoreId = "Store1", EntryId = "Folder1", Name = "Inbox", ItemCount = 10 }
                    }
                }
            };

            _indexerService = new EmailIndexerService(_outlookAdapter, _repository, new MockLicensingService(LicenseStatus.Valid));
        }

        public void Dispose()
        {
            try { if (File.Exists(_dbPath)) File.Delete(_dbPath); } catch { }
        }

        [Fact]
        public async Task StartInitialIndexing_RunsAndCompletes()
        {
            bool progressRaised = false;
            _indexerService.ProgressChanged += (s, e) =>
            {
                progressRaised = true;
            };

            var indexingTask = _indexerService.StartInitialIndexingAsync();
            await indexingTask;

            Assert.True(progressRaised);
            Assert.Equal(IndexerStatus.Idle, _indexerService.CurrentStatus);
            Assert.Equal(10, _indexerService.TotalItems);
            Assert.Equal(10, _indexerService.ProcessedItems);
        }

        [Fact]
        public async Task RebuildIndex_GuardsAgainstConcurrentRun()
        {
            var task1 = _indexerService.StartInitialIndexingAsync();
            var task2 = _indexerService.RebuildIndexAsync();

            await Task.WhenAll(task1, task2);

            Assert.Equal(IndexerStatus.Idle, _indexerService.CurrentStatus);
        }

        [Fact]
        public async Task GetStatus_ReturnsCurrentProgress()
        {
            await _indexerService.StartInitialIndexingAsync();
            var status = _indexerService.GetStatus();

            Assert.Equal(10, status.TotalItems);
            Assert.Equal(10, status.ProcessedItems);
            Assert.Equal(100.0, status.Percentage);
            Assert.Equal(IndexerStatus.Idle, status.Status);
        }

        [Fact]
        public async Task Cancellation_StopsIndexerGracefully()
        {
            var cts = new CancellationTokenSource();
            var task = _indexerService.StartInitialIndexingAsync(cts.Token);
            cts.Cancel();
            await task;

            Assert.Equal(IndexerStatus.Idle, _indexerService.CurrentStatus);
        }
    }
}
