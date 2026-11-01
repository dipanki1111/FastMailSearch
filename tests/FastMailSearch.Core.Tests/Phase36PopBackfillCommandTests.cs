using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using FastMailSearch.Core.Interfaces;
using FastMailSearch.Core.Models;
using FastMailSearch.Core.Services;
using FastMailSearch.Infrastructure.Database;
using FastMailSearch.Infrastructure.Indexing;
using FastMailSearch.OutlookAddIn;
using FastMailSearch.OutlookAddIn.Ribbon;
using Xunit;

namespace FastMailSearch.Core.Tests
{
    public class Phase36PopBackfillCommandTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly DatabaseConnectionFactory _factory;
        private readonly SqliteEmailRepository _repository;

        public Phase36PopBackfillCommandTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), $"Phase36BackfillTests_{Guid.NewGuid():N}.db");
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

        private class MockOutlookAdapter36 : IOutlookAdapter
        {
            public List<MailboxInfo> Mailboxes { get; set; } = new List<MailboxInfo>();
            public Dictionary<string, List<MailboxFolder>> StoreFolders { get; set; } = new Dictionary<string, List<MailboxFolder>>();
            public Dictionary<string, List<EmailDocument>> FolderEmails { get; set; } = new Dictionary<string, List<EmailDocument>>();

            public bool IsOutlookAvailable() => true;
            public string GetOutlookVersion() => "16.0 (64-bit)";
            public string GetOfficeBitness() => "x64";
            public List<MailboxInfo> GetMailboxes() => Mailboxes;
            public List<MailboxFolder> GetFolders(string storeId) => StoreFolders.TryGetValue(storeId, out var f) ? f : new List<MailboxFolder>();

            public List<EmailDocument> FetchFolderEmailsBatch(string storeId, string folderEntryId, int startIndex, int batchSize, Func<bool>? cancelCheck = null)
            {
                string key = $"{storeId}_{folderEntryId}";
                if (!FolderEmails.TryGetValue(key, out var list)) return new List<EmailDocument>();
                return list.Skip(startIndex - 1).Take(batchSize).ToList();
            }

            public EmailDocument? FetchSingleEmail(string storeId, string entryId) => null;
            public bool OpenEmailInOutlook(string storeId, string entryId) => true;
            public bool IsOutboxFolder(string storeId, string folderEntryId) => false;
            public void SubscribeToEvents(Action<string, string, string> onItemAddedOrChanged, Action<string, string> onItemRemoved) { }
            public void UnsubscribeFromEvents() { }
        }

        [Fact]
        public void Test01_EmailIndexerService_HasTwoParameterConstructor_PreservingBinaryCompatibility()
        {
            // Verify by reflection that .ctor(IOutlookAdapter, IEmailRepository) exists explicitly
            var type = typeof(EmailIndexerService);
            var ctor2 = type.GetConstructor(new[] { typeof(IOutlookAdapter), typeof(IEmailRepository) });
            Assert.NotNull(ctor2);

            var ctor3 = type.GetConstructor(new[] { typeof(IOutlookAdapter), typeof(IEmailRepository), typeof(bool) });
            Assert.NotNull(ctor3);

            // Instantiate using the 2-parameter constructor
            var mock = new MockOutlookAdapter36();
            var instance = (EmailIndexerService)ctor2.Invoke(new object[] { mock, _repository });
            Assert.NotNull(instance);
            Assert.Equal(IndexerStatus.Idle, instance.CurrentStatus);
        }

        [Fact]
        public void Test02_SearchService_HasTwoParameterConstructor_PreservingBinaryCompatibility()
        {
            var type = typeof(SearchService);
            var ctor2 = type.GetConstructor(new[] { typeof(ISearchEngine), typeof(IEmailRepository) });
            Assert.NotNull(ctor2);

            var ctor3 = type.GetConstructor(new[] { typeof(ISearchEngine), typeof(IEmailRepository), typeof(ILicensingService) });
            Assert.NotNull(ctor3);

            var instance = (SearchService)ctor2.Invoke(new object[] { _repository, _repository });
            Assert.NotNull(instance);
        }

        [Fact]
        public void Test03_RibbonXML_ContainsBackfillPopMailboxButton_AndValidCallbacks()
        {
            // Verify embedded XML
            var asm = typeof(FastMailRibbon).Assembly;
            var stream = asm.GetManifestResourceStream("FastMailSearch.OutlookAddIn.Ribbon.FastMailRibbon.xml");
            Assert.NotNull(stream);

            using var reader = new StreamReader(stream!);
            string xmlContent = reader.ReadToEnd();
            var doc = XDocument.Parse(xmlContent);

            XNamespace ns = "http://schemas.microsoft.com/office/2009/07/customui";
            var button = doc.Descendants(ns + "button").FirstOrDefault(b => (string?)b.Attribute("id") == "btnBackfillActiveStore");

            Assert.NotNull(button);
            Assert.Equal("Backfill POP Mailbox", (string?)button.Attribute("label"));
            Assert.Equal("OnBackfillActiveStoreClicked", (string?)button.Attribute("onAction"));

            // Verify both ThisAddIn and FastMailRibbon expose OnBackfillActiveStoreClicked
            var methodOnThisAddIn = typeof(ThisAddIn).GetMethod("OnBackfillActiveStoreClicked");
            Assert.NotNull(methodOnThisAddIn);

            var methodOnRibbon = typeof(FastMailRibbon).GetMethod("OnBackfillActiveStoreClicked");
            Assert.NotNull(methodOnRibbon);
        }

        [Fact]
        public async Task Test04_StoreCoverage_DetectsUnindexedEmails_AndBackfillResumesFromCheckpoint()
        {
            var mock = new MockOutlookAdapter36();
            string storeId = "pst-store-1";
            string folderId = "folder-inbox";

            mock.Mailboxes.Add(new MailboxInfo
            {
                StoreId = storeId,
                DisplayName = "personal-test@example.com",
                IsPst = true,
                FilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Outlook Files", "test.pst")
            });

            mock.StoreFolders[storeId] = new List<MailboxFolder>
            {
                new MailboxFolder { EntryId = folderId, Name = "Inbox", ItemCount = 500 }
            };

            // Seed existing index state indicating 250 items were already processed (offset = 250, total = 500)
            await _repository.SaveFolderIndexStateAsync(storeId, folderId, "Indexing", 250, 500);

            // Populate the mock folder with items 251..500
            var emails = new List<EmailDocument>();
            for (int i = 1; i <= 500; i++)
            {
                emails.Add(new EmailDocument
                {
                    StoreId = storeId,
                    FolderEntryId = folderId,
                    EntryId = $"msg-{i}",
                    Subject = $"Item {i}",
                    SenderName = "Sender",
                    SenderEmail = "sender@test.com",
                    ReceivedDate = new DateTime(2026, 1, 1).AddMinutes(i),
                    FolderName = "Inbox",
                    BodyText = $"Body for item {i}"
                });
            }
            mock.FolderEmails[$"{storeId}_{folderId}"] = emails;

            var indexer = new EmailIndexerService(mock, _repository, new MockLicensingService());

            // Verify store coverage detects unindexed items
            var coverage = await indexer.CheckStoreCoverageAsync();
            Assert.Equal(500, coverage.TotalDiscoveredItems);
            Assert.True(coverage.UnindexedItems > 0);
            Assert.Contains("Backfill recommended", coverage.SummaryText);

            // Run active store backfill: it should resume from offset 251 and index the remaining 250 items
            var summary = await indexer.StartActiveStoreBackfillAsync();
            Assert.Equal(500, summary.TotalDiscovered);
            Assert.Equal(250, summary.NewlyIndexedCount);
            Assert.False(summary.IsCancelled);

            // Check that folder state is now Completed at 500
            var finalState = await _repository.GetFolderIndexStateAsync(storeId, folderId);
            Assert.NotNull(finalState);
            Assert.Equal("Completed", finalState.Value.Status);
            Assert.Equal(500, finalState.Value.ProcessedOffset);
        }

        [Fact]
        public void Test05_ThisAddIn_CreateEmailIndexerService_DynamicallyResolvesConstructor_WithoutMissingMethodException()
        {
            var mock = new MockOutlookAdapter36();
            var instance = ThisAddIn.CreateEmailIndexerService(mock, _repository);
            Assert.NotNull(instance);
            Assert.Equal(IndexerStatus.Idle, instance.CurrentStatus);
        }

        [Fact]
        public async Task Test06_CheckAndRecoverIncompleteIndex_StartProcessingFalse_DoesNotAutoStartFullBackfill()
        {
            var mock = new MockOutlookAdapter36();
            var indexer = new EmailIndexerService(mock, _repository);

            // Simulate startup check with startProcessing: false
            await indexer.CheckAndRecoverIncompleteIndexAsync(startProcessing: false);

            Assert.Equal(IndexerStatus.Idle, indexer.CurrentStatus);
            Assert.False(indexer.IsIndexing);
            Assert.Equal("Idle", indexer.CurrentOperation);
        }

        [Fact]
        public async Task Test07_Backfill_CanExecuteWhenPendingChangesQueueIsPopulated()
        {
            var mock = new MockOutlookAdapter36();
            string storeId = "pst-pending-test";
            string folderId = "inbox-pending-test";

            mock.Mailboxes.Add(new MailboxInfo
            {
                StoreId = storeId,
                DisplayName = "test@domain.com",
                IsPst = true,
                FilePath = "C:\\test.pst"
            });

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
                    FolderEntryId = folderId,
                    EntryId = $"msg-{i}",
                    Subject = $"Item {i}",
                    SenderName = "Sender",
                    SenderEmail = "sender@test.com",
                    ReceivedDate = DateTime.Now,
                    FolderName = "Inbox",
                    BodyText = "Body"
                });
            }
            mock.FolderEmails[$"{storeId}_{folderId}"] = emails;

            // Seed 50 pending changes into PendingChanges to simulate the 3,429 pending changes
            for (int i = 1; i <= 50; i++)
            {
                await _repository.EnqueuePendingChangeAsync(storeId, $"pending-entry-{i}", folderId, "Add");
            }
            int pendingCount = await _repository.GetPendingChangesCountAsync();
            Assert.Equal(50, pendingCount);

            var indexer = new EmailIndexerService(mock, _repository, new MockLicensingService());

            // Active store backfill should execute cleanly even with pending changes
            var summary = await indexer.StartActiveStoreBackfillAsync();
            Assert.Equal(10, summary.TotalDiscovered);
            Assert.Equal(10, summary.NewlyIndexedCount);
            Assert.False(summary.IsCancelled);
            Assert.Null(summary.ErrorMessage);
        }
    }
}
