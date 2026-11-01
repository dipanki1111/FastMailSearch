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
    public class Phase33IndexingProgressAndDarkButtonsTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly DatabaseConnectionFactory _factory;
        private readonly SqliteEmailRepository _repository;

        public Phase33IndexingProgressAndDarkButtonsTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), $"Phase33Tests_{Guid.NewGuid():N}.db");
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

        private class MockOutlookAdapter33 : IOutlookAdapter
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
        public async Task Test01_Progress_Counter_Includes_Already_Indexed_Folders()
        {
            var outlook = new MockOutlookAdapter33();
            outlook.Mailboxes.Add(new MailboxInfo { StoreId = "store-1", DisplayName = "Primary" });

            outlook.StoreFolders["store-1"] = new List<MailboxFolder>
            {
                new MailboxFolder { StoreId = "store-1", EntryId = "f1", Name = "Archive", ItemCount = 100 },
                new MailboxFolder { StoreId = "store-1", EntryId = "f2", Name = "Inbox", ItemCount = 50 }
            };

            // Folder 1 is already completed
            await _repository.SaveFolderIndexStateAsync("store-1", "f1", "Completed", 100, 100);

            // Folder 2 emails
            var inboxEmails = new List<EmailDocument>();
            for (int i = 1; i <= 50; i++)
            {
                inboxEmails.Add(new EmailDocument
                {
                    StoreId = "store-1",
                    EntryId = $"inbox-{i}",
                    FolderEntryId = "f2",
                    FolderName = "Inbox",
                    Subject = $"Inbox item {i}",
                    ReceivedDate = DateTime.Now
                });
            }
            outlook.FolderEmails["store-1:f2"] = inboxEmails;

            var indexer = new EmailIndexerService(outlook, _repository, new MockLicensingService(LicenseStatus.Valid));
            var progressEvents = new List<IndexerProgressEventArgs>();
            indexer.ProgressChanged += (s, e) => progressEvents.Add(e);

            var summary = await indexer.StartActiveStoreBackfillAsync();
            Assert.False(summary.IsCancelled);

            // Verify progress events recorded the already-indexed folder
            var folder2Events = progressEvents.Where(e => e.CurrentFolder == "Inbox" && e.Status == IndexerStatus.Indexing).ToList();
            Assert.NotEmpty(folder2Events);

            // First event for Inbox should reflect 100 already-indexed items from Archive
            var firstInboxEvent = folder2Events.First();
            Assert.Equal(150, firstInboxEvent.TotalItems);
            Assert.True(firstInboxEvent.ProcessedItems >= 100, $"Expected >= 100 but was {firstInboxEvent.ProcessedItems}");
            Assert.Equal(50, firstInboxEvent.FolderTotalItems);

            // Final event should reach 150/150 (100 Archive + 50 Inbox)
            var lastIndexingEvent = progressEvents.Last(e => e.Status == IndexerStatus.Indexing);
            Assert.Equal(150, lastIndexingEvent.ProcessedItems);
            Assert.Equal(150, lastIndexingEvent.TotalItems);
            Assert.Equal(50, lastIndexingEvent.FolderProcessedItems);
            Assert.Equal(50, lastIndexingEvent.FolderTotalItems);
        }

        [Fact]
        public async Task Test02_Resuming_Interrupted_Folder_Tracks_Folder_And_Overall_Progress()
        {
            var outlook = new MockOutlookAdapter33();
            outlook.Mailboxes.Add(new MailboxInfo { StoreId = "store-resume", DisplayName = "Mailbox" });
            outlook.StoreFolders["store-resume"] = new List<MailboxFolder>
            {
                new MailboxFolder { StoreId = "store-resume", EntryId = "f-large", Name = "LargeFolder", ItemCount = 500 }
            };

            // Previously interrupted at 250 items
            await _repository.SaveFolderIndexStateAsync("store-resume", "f-large", "Indexing", 250, 500);

            var emails = new List<EmailDocument>();
            for (int i = 1; i <= 500; i++)
            {
                emails.Add(new EmailDocument
                {
                    StoreId = "store-resume",
                    EntryId = $"msg-{i}",
                    FolderEntryId = "f-large",
                    FolderName = "LargeFolder",
                    Subject = $"Msg {i}",
                    ReceivedDate = DateTime.Now
                });
            }
            outlook.FolderEmails["store-resume:f-large"] = emails;

            var indexer = new EmailIndexerService(outlook, _repository, new MockLicensingService(LicenseStatus.Valid));
            var progressEvents = new List<IndexerProgressEventArgs>();
            indexer.ProgressChanged += (s, e) => progressEvents.Add(e);

            var summary = await indexer.StartActiveStoreBackfillAsync();
            Assert.False(summary.IsCancelled);

            // Verify initial progress event on resume starts from 250
            var initialResumeEvent = progressEvents.First(e => e.CurrentFolder == "LargeFolder" && e.Status == IndexerStatus.Indexing);
            Assert.Equal(250, initialResumeEvent.ProcessedItems);
            Assert.Equal(250, initialResumeEvent.FolderProcessedItems);
            Assert.Equal(500, initialResumeEvent.FolderTotalItems);
            Assert.Equal(500, initialResumeEvent.TotalItems);

            // Final event reaches 500
            var lastEvent = progressEvents.Last(e => e.Status == IndexerStatus.Indexing);
            Assert.Equal(500, lastEvent.ProcessedItems);
            Assert.Equal(500, lastEvent.FolderProcessedItems);
            Assert.Equal(500, lastEvent.FolderTotalItems);
        }

        [Fact]
        public void Test03_IndexerProgressEventArgs_Exposes_Accurate_Folder_Metrics()
        {
            var args = new IndexerProgressEventArgs
            {
                TotalItems = 1000,
                ProcessedItems = 350,
                CurrentFolder = "Inbox",
                FolderProcessedItems = 150,
                FolderTotalItems = 400,
                Status = IndexerStatus.Indexing
            };

            Assert.Equal(1000, args.TotalItems);
            Assert.Equal(350, args.ProcessedItems);
            Assert.Equal(150, args.FolderProcessedItems);
            Assert.Equal(400, args.FolderTotalItems);
            Assert.Equal("Inbox", args.CurrentFolder);
        }

        [Fact]
        public void Test04_IndexHealthWindow_Xaml_Contains_Dark_Button_Styles()
        {
            // Verify IndexHealthWindow.xaml uses dark action button styles for all buttons
            var dir = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
            string? xamlPath = null;
            while (dir != null)
            {
                string candidate = Path.Combine(dir.FullName, @"src\FastMailSearch.UI\Views\IndexHealthWindow.xaml");
                if (File.Exists(candidate))
                {
                    xamlPath = candidate;
                    break;
                }
                dir = dir.Parent;
            }
            Assert.True(xamlPath != null && File.Exists(xamlPath), $"File should exist at {xamlPath}");

            string content = File.ReadAllText(xamlPath);
            Assert.Contains("GreenActionButtonStyle", content);
            Assert.Contains("DarkActionButtonStyle", content);
            Assert.Contains("BlueActionButtonStyle", content);
            Assert.Contains("CloseActionButtonStyle", content);

            // Ensure disabled triggers specify dark backgrounds and readable foregrounds
            Assert.Contains("IsEnabled", content);
            Assert.Contains("#143820", content); // Green button disabled background
            Assert.Contains("#202124", content); // Dark button disabled background
            Assert.Contains("#6EE7B7", content); // Green button disabled text
            Assert.Contains("#71717A", content); // Dark button disabled text
        }
    }
}
