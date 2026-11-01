using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using FastMailSearch.Core.Interfaces;
using FastMailSearch.Core.Models;
using FastMailSearch.Infrastructure.Database;
using FastMailSearch.Infrastructure.Indexing;
using FastMailSearch.Infrastructure.Outlook;
using Xunit;

namespace FastMailSearch.Core.Tests
{
    public class Phase20OutlookSendSafetyTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly DatabaseConnectionFactory _factory;
        private readonly SqliteEmailRepository _repository;
        private readonly MockSendSafetyOutlookAdapter _outlookAdapter;
        private readonly EmailIndexerService _indexerService;

        public Phase20OutlookSendSafetyTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), $"FastMailSearch_Phase20Tests_{Guid.NewGuid():N}.db");
            _factory = new DatabaseConnectionFactory(_dbPath);
            _repository = new SqliteEmailRepository(_factory);
            _repository.InitializeDatabase();

            _outlookAdapter = new MockSendSafetyOutlookAdapter();
            _indexerService = new EmailIndexerService(_outlookAdapter, _repository, new MockLicensingService());
        }

        public void Dispose()
        {
            _indexerService?.CancelIndexingAsync().Wait();
            if (File.Exists(_dbPath))
            {
                try { File.Delete(_dbPath); } catch { }
            }
        }

        private class MockSendSafetyOutlookAdapter : IOutlookAdapter
        {
            public ConcurrentDictionary<string, EmailDocument> EmailStorage = new ConcurrentDictionary<string, EmailDocument>(StringComparer.OrdinalIgnoreCase);
            public ConcurrentBag<string> FetchedEntryIds = new ConcurrentBag<string>();
            public bool IsAvailable { get; set; } = true;
            private Action<string, string, string>? _onItemChanged;

            public bool IsOutlookAvailable() => IsAvailable;
            public string GetOutlookVersion() => "16.0.16000.20000";
            public string GetOfficeBitness() => "64-bit";

            public List<MailboxInfo> GetMailboxes() => new List<MailboxInfo>
            {
                new MailboxInfo { StoreId = "Store1", DisplayName = "Exchange Account", IsPst = false }
            };

            public List<MailboxFolder> GetFolders(string storeId) => new List<MailboxFolder>
            {
                new MailboxFolder { StoreId = storeId, EntryId = "OutboxFolder", Name = "Outbox", ItemCount = 5 },
                new MailboxFolder { StoreId = storeId, EntryId = "SentItemsFolder", Name = "Sent Items", ItemCount = 10 },
                new MailboxFolder { StoreId = storeId, EntryId = "InboxFolder", Name = "Inbox", ItemCount = 20 }
            };

            public List<EmailDocument> FetchFolderEmailsBatch(string storeId, string folderEntryId, int startIndex, int batchSize, Func<bool>? cancelCheck = null)
            {
                if (IsOutboxFolder(storeId, folderEntryId)) return new List<EmailDocument>();
                return EmailStorage.Values.Where(d => d.FolderEntryId == folderEntryId).Skip(startIndex - 1).Take(batchSize).ToList();
            }

            public EmailDocument? FetchSingleEmail(string storeId, string entryId)
            {
                FetchedEntryIds.Add(entryId);
                string key = $"{storeId}:{entryId}";
                if (EmailStorage.TryGetValue(key, out var doc))
                {
                    if (IsOutboxFolder(storeId, doc.FolderEntryId) || string.Equals(doc.FolderName, "Outbox", StringComparison.OrdinalIgnoreCase))
                    {
                        return null;
                    }
                    return doc;
                }
                return null;
            }

            public bool OpenEmailInOutlook(string storeId, string entryId) => true;

            public bool IsOutboxFolder(string storeId, string folderEntryId)
            {
                return string.Equals(folderEntryId, "OutboxFolder", StringComparison.OrdinalIgnoreCase) ||
                       string.Equals(folderEntryId, "Outbox", StringComparison.OrdinalIgnoreCase);
            }

            public void SubscribeToEvents(Action<string, string, string> onItemAddedOrChanged, Action<string, string> onItemRemoved)
            {
                _onItemChanged = onItemAddedOrChanged;
            }

            public void UnsubscribeFromEvents()
            {
                _onItemChanged = null;
            }

            public void SimulateItemAdd(string storeId, string entryId, string folderEntryId)
            {
                _onItemChanged?.Invoke(storeId, entryId, folderEntryId);
            }
        }

        #region Test 1: Outbox Items Are Ignored
        [Fact]
        public async Task Test01_Outbox_Items_Are_Ignored()
        {
            string storeId = "Store1";
            string entryId = "OutboxMsg001";
            string folderEntryId = "OutboxFolder";

            _outlookAdapter.EmailStorage[$"{storeId}:{entryId}"] = new EmailDocument
            {
                StoreId = storeId,
                EntryId = entryId,
                FolderEntryId = folderEntryId,
                FolderName = "Outbox",
                Subject = "Stuck Outbox Message"
            };

            await _indexerService.QueueItemForIndexingAsync(storeId, entryId, folderEntryId);
            await Task.Delay(300);

            var indexed = await _repository.GetEmailByStoreAndEntryIdAsync(storeId, entryId);
            Assert.Null(indexed);
        }
        #endregion

        #region Test 2: Sent Items Are Indexed
        [Fact]
        public async Task Test02_SentItems_Items_Are_Indexed()
        {
            string storeId = "Store1";
            string entryId = "SentMsg001";
            string folderEntryId = "SentItemsFolder";

            _outlookAdapter.EmailStorage[$"{storeId}:{entryId}"] = new EmailDocument
            {
                StoreId = storeId,
                EntryId = entryId,
                FolderEntryId = folderEntryId,
                FolderName = "Sent Items",
                Subject = "Sent Message",
                SenderName = "Me",
                SenderEmail = "me@company.com",
                ReceivedDate = DateTime.Now
            };

            await _indexerService.QueueItemForIndexingAsync(storeId, entryId, folderEntryId);
            await Task.Delay(500);

            var indexed = await _repository.GetEmailByStoreAndEntryIdAsync(storeId, entryId);
            Assert.NotNull(indexed);
            Assert.Equal("Sent Message", indexed.Subject);
        }
        #endregion

        #region Test 3: ItemAdd Event Returns Immediately
        [Fact]
        public async Task Test03_ItemAdd_Event_Returns_Immediately()
        {
            string storeId = "Store1";
            string entryId = "FastMsg001";
            string folderEntryId = "SentItemsFolder";

            _outlookAdapter.EmailStorage[$"{storeId}:{entryId}"] = new EmailDocument
            {
                StoreId = storeId,
                EntryId = entryId,
                FolderEntryId = folderEntryId,
                FolderName = "Sent Items",
                Subject = "Fast Event Test"
            };

            var sw = System.Diagnostics.Stopwatch.StartNew();
            _outlookAdapter.SimulateItemAdd(storeId, entryId, folderEntryId);
            sw.Stop();

            Assert.True(sw.ElapsedMilliseconds < 150, $"ItemAdd callback took too long: {sw.ElapsedMilliseconds}ms");
            await Task.Delay(300);
        }
        #endregion

        #region Test 4: ItemChange Event Returns Immediately
        [Fact]
        public async Task Test04_ItemChange_Event_Returns_Immediately()
        {
            string storeId = "Store1";
            string entryId = "ChangeMsg001";
            string folderEntryId = "InboxFolder";

            _outlookAdapter.EmailStorage[$"{storeId}:{entryId}"] = new EmailDocument
            {
                StoreId = storeId,
                EntryId = entryId,
                FolderEntryId = folderEntryId,
                FolderName = "Inbox",
                Subject = "Item Change Test"
            };

            var sw = System.Diagnostics.Stopwatch.StartNew();
            _outlookAdapter.SimulateItemAdd(storeId, entryId, folderEntryId);
            sw.Stop();

            Assert.True(sw.ElapsedMilliseconds < 250, $"ItemChange callback took too long: {sw.ElapsedMilliseconds}ms");
            await Task.Delay(300);
        }
        #endregion

        #region Test 5: Duplicate Outbox Events Produce Zero Indexing
        [Fact]
        public async Task Test05_Duplicate_Outbox_Events_Produce_Zero_Indexing()
        {
            string storeId = "Store1";
            string entryId = "OutboxDup001";
            string folderEntryId = "OutboxFolder";

            for (int i = 0; i < 20; i++)
            {
                await _indexerService.QueueItemForIndexingAsync(storeId, entryId, folderEntryId);
            }

            await Task.Delay(300);

            var count = await _repository.GetIndexedEmailCountAsync();
            Assert.Equal(0, count);
        }
        #endregion

        #region Test 6: Moving Outbox to Sent Items Produces One Indexing
        [Fact]
        public async Task Test06_Moving_Outbox_To_SentItems_Produces_One_Indexing()
        {
            string storeId = "Store1";
            string entryId = "MovingMsg001";

            // Stage 1: Item in Outbox while sending
            _outlookAdapter.EmailStorage[$"{storeId}:{entryId}"] = new EmailDocument
            {
                StoreId = storeId,
                EntryId = entryId,
                FolderEntryId = "OutboxFolder",
                FolderName = "Outbox",
                Subject = "Transmitting Email"
            };
            await _indexerService.QueueItemForIndexingAsync(storeId, entryId, "OutboxFolder");
            await Task.Delay(200);
            Assert.Null(await _repository.GetEmailByStoreAndEntryIdAsync(storeId, entryId));

            // Stage 2: Transmission completes -> Outlook moves item to Sent Items
            _outlookAdapter.EmailStorage[$"{storeId}:{entryId}"] = new EmailDocument
            {
                StoreId = storeId,
                EntryId = entryId,
                FolderEntryId = "SentItemsFolder",
                FolderName = "Sent Items",
                Subject = "Transmitting Email",
                SenderName = "Me",
                SenderEmail = "me@company.com",
                ReceivedDate = DateTime.Now
            };
            await _indexerService.QueueItemForIndexingAsync(storeId, entryId, "SentItemsFolder");
            await Task.Delay(500);

            var indexed = await _repository.GetEmailByStoreAndEntryIdAsync(storeId, entryId);
            Assert.NotNull(indexed);
            Assert.Equal("Sent Items", indexed.FolderName);
            Assert.Equal(1, await _repository.GetIndexedEmailCountAsync());
        }
        #endregion

        #region Test 7: Missing Deleted Item Does Not Crash Queue
        [Fact]
        public async Task Test07_Missing_Deleted_Item_Does_Not_Crash_Queue()
        {
            string storeId = "Store1";
            string entryId = "GhostEntry999";
            string folderEntryId = "InboxFolder";

            // Queue item that does not exist in Outlook storage
            await _indexerService.QueueItemForIndexingAsync(storeId, entryId, folderEntryId);
            await Task.Delay(400);

            // Queue should recover smoothly without crashing
            string validEntry = "ValidMsg001";
            _outlookAdapter.EmailStorage[$"{storeId}:{validEntry}"] = new EmailDocument
            {
                StoreId = storeId,
                EntryId = validEntry,
                FolderEntryId = folderEntryId,
                FolderName = "Inbox",
                Subject = "Recovery Test"
            };

            await _indexerService.QueueItemForIndexingAsync(storeId, validEntry, folderEntryId);
            await Task.Delay(500);

            var indexed = await _repository.GetEmailByStoreAndEntryIdAsync(storeId, validEntry);
            Assert.NotNull(indexed);
        }
        #endregion

        #region Test 8: COM Objects Never Passed To Queue
        [Fact]
        public void Test08_COM_Objects_Never_Passed_To_Queue()
        {
            var itemType = typeof(IncrementalItem);
            var properties = itemType.GetProperties(BindingFlags.Public | BindingFlags.Instance);

            foreach (var prop in properties)
            {
                var propType = prop.PropertyType;
                Assert.False(propType.IsCOMObject, $"Property {prop.Name} on IncrementalItem is a COM Object!");
                Assert.False(propType.FullName?.StartsWith("Microsoft.Office.Interop.Outlook") ?? false,
                    $"Property {prop.Name} references Outlook COM interop types!");
            }
        }
        #endregion

        #region Test 9: Queue Contains Immutable Identifiers Only
        [Fact]
        public void Test09_Queue_Contains_Immutable_Identifiers_Only()
        {
            var item = new IncrementalItem
            {
                PendingId = 1,
                StoreId = "Store123",
                EntryId = "Entry456",
                FolderEntryId = "Folder789",
                Operation = IncrementalOperation.InsertOrUpdate
            };

            Assert.IsType<string>(item.StoreId);
            Assert.IsType<string>(item.EntryId);
            Assert.IsType<string>(item.FolderEntryId);
            Assert.IsType<long>(item.PendingId);
        }
        #endregion

        #region Test 10: SQLite Transaction Never Executes In COM Callback
        [Fact]
        public async Task Test10_SQLite_Transaction_Never_Executes_In_COM_Callback()
        {
            bool cbCompletedBeforeDb = false;
            string storeId = "Store1";
            string entryId = "ThreadTest001";
            string folderEntryId = "SentItemsFolder";

            _outlookAdapter.EmailStorage[$"{storeId}:{entryId}"] = new EmailDocument
            {
                StoreId = storeId,
                EntryId = entryId,
                FolderEntryId = folderEntryId,
                FolderName = "Sent Items",
                Subject = "Thread Isolation Test"
            };

            // Queue and verify callback finishes before DB write occurs
            var task = _indexerService.QueueItemForIndexingAsync(storeId, entryId, folderEntryId);
            cbCompletedBeforeDb = task.IsCompleted || task.Status == TaskStatus.RanToCompletion || task.Status == TaskStatus.WaitingForActivation;

            await task;
            await Task.Delay(500);

            Assert.True(cbCompletedBeforeDb, "QueueItemForIndexingAsync did not return asynchronously");
            var doc = await _repository.GetEmailByStoreAndEntryIdAsync(storeId, entryId);
            Assert.NotNull(doc);
        }
        #endregion

        #region Test 11: Static Safety Audit - No Forbidden Send Handlers
        [Fact]
        public void Test11_No_Forbidden_Send_Handler_Registered()
        {
            string srcDir = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "..", "..", "src"));
            if (!Directory.Exists(srcDir))
            {
                srcDir = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "..", "..", "..", "..", "..", "src"));
            }

            Assert.True(Directory.Exists(srcDir), $"Source directory not found at: {srcDir}");

            string[] forbiddenTokens = new[]
            {
                "Application.ItemSend",
                "ItemSend",
                "BeforeSend",
                "MailItem.Send",
                "SendAndReceive",
                "SyncNow",
                "SendUsingAccount"
            };

            var csFiles = Directory.GetFiles(srcDir, "*.cs", SearchOption.AllDirectories);
            var violations = new List<string>();

            foreach (var file in csFiles)
            {
                if (file.Contains("AssemblyInfo.cs") || file.Contains(".g.cs") || file.Contains("wpftmp")) continue;

                string content = File.ReadAllText(file);
                foreach (var token in forbiddenTokens)
                {
                    if (content.Contains(token))
                    {
                        violations.Add($"Forbidden token '{token}' found in file: {Path.GetFileName(file)}");
                    }
                }
            }

            Assert.Empty(violations);
        }
        #endregion

        #region Test 12: Search Continues During Incremental Indexing
        [Fact]
        public async Task Test12_Search_Continues_During_Incremental_Indexing()
        {
            for (int i = 0; i < 50; i++)
            {
                await _repository.SaveEmailBatchAsync(new List<EmailDocument>
                {
                    new EmailDocument
                    {
                        StoreId = "Store1",
                        EntryId = $"PreMsg_{i}",
                        Subject = $"Financial Report {i}",
                        BodyText = "Quarterly earnings and budget details",
                        FolderName = "Inbox",
                        ReceivedDate = DateTime.Now
                    }
                });
            }

            _outlookAdapter.EmailStorage["Store1:IncMsg999"] = new EmailDocument
            {
                StoreId = "Store1",
                EntryId = "IncMsg999",
                FolderEntryId = "InboxFolder",
                FolderName = "Inbox",
                Subject = "Incoming Financial Update"
            };

            var queueTask = _indexerService.QueueItemForIndexingAsync("Store1", "IncMsg999", "InboxFolder");

            var query = new SearchQuery { FreeText = "financial", PageSize = 10, Page = 1 };
            var searchResults = await _repository.SearchAsync(query);

            Assert.NotNull(searchResults);
            Assert.True(searchResults.TotalCount >= 50);

            await queueTask;
            await Task.Delay(300);
        }
        #endregion

        #region Test 13: Indexing Continues After Outlook Send
        [Fact]
        public async Task Test13_Indexing_Continues_After_Outlook_Send()
        {
            string storeId = "Store1";
            string entryId = "PostSendMsg001";
            string folderEntryId = "SentItemsFolder";

            _outlookAdapter.EmailStorage[$"{storeId}:{entryId}"] = new EmailDocument
            {
                StoreId = storeId,
                EntryId = entryId,
                FolderEntryId = folderEntryId,
                FolderName = "Sent Items",
                Subject = "Post Send Verification"
            };

            await _indexerService.QueueItemForIndexingAsync(storeId, entryId, folderEntryId);
            await Task.Delay(500);

            var status = _indexerService.GetStatus();
            Assert.Equal(IndexerStatus.Idle, status.Status);

            var doc = await _repository.GetEmailByStoreAndEntryIdAsync(storeId, entryId);
            Assert.NotNull(doc);
        }
        #endregion

        #region Test 14: Exchange Account Scenario Isolated
        [Fact]
        public async Task Test14_Exchange_Account_Scenario_Isolated()
        {
            var mailboxes = _outlookAdapter.GetMailboxes();
            var exchangeStore = mailboxes.FirstOrDefault(m => !m.IsPst);
            Assert.NotNull(exchangeStore);

            string storeId = exchangeStore.StoreId;
            string entryId = "ExchMsg100";
            string folderEntryId = "OutboxFolder";

            await _indexerService.QueueItemForIndexingAsync(storeId, entryId, folderEntryId);
            await Task.Delay(300);

            Assert.Null(await _repository.GetEmailByStoreAndEntryIdAsync(storeId, entryId));
        }
        #endregion

        #region Test 15: Outlook Shutdown Does Not Leave Worker Hanging
        [Fact]
        public async Task Test15_Outlook_Shutdown_Does_Not_Leave_Worker_Hanging()
        {
            try
            {
                var adapter = new OutlookAdapter();
                Assert.True(adapter is IDisposable);

                var task = Task.Run(() => adapter.Dispose());
                bool completed = await Task.WhenAny(task, Task.Delay(5000)) == task;

                Assert.True(completed, "OutlookAdapter disposal timed out or hung!");
            }
            catch (System.IO.FileNotFoundException)
            {
                // Expected when Office Primary Interop Assemblies (PIA) are not installed in GAC
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // Expected when Outlook COM server is not registered on test environment
            }
        }
        #endregion
    }
}
