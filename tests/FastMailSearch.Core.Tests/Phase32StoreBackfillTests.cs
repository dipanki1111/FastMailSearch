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
    public class Phase32StoreBackfillTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly DatabaseConnectionFactory _factory;
        private readonly SqliteEmailRepository _repository;

        public Phase32StoreBackfillTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), $"Phase32Tests_{Guid.NewGuid():N}.db");
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

        private class BackfillMockOutlookAdapter : IOutlookAdapter
        {
            public List<MailboxInfo> Mailboxes { get; set; } = new List<MailboxInfo>();
            public Dictionary<string, List<MailboxFolder>> StoreFolders { get; set; } = new Dictionary<string, List<MailboxFolder>>();
            public Dictionary<string, List<EmailDocument>> FolderEmails { get; set; } = new Dictionary<string, List<EmailDocument>>();
            public HashSet<string> OutboxEntryIds { get; set; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            public List<int> RequestedOffsets { get; } = new List<int>();

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
                RequestedOffsets.Add(startIndex);
                if (cancelCheck != null && cancelCheck()) return new List<EmailDocument>();

                string key = $"{storeId}:{folderEntryId}";
                if (!FolderEmails.TryGetValue(key, out var allEmails)) return new List<EmailDocument>();

                int zeroIndex = startIndex - 1;
                if (zeroIndex < 0 || zeroIndex >= allEmails.Count) return new List<EmailDocument>();

                int count = Math.Min(batchSize, allEmails.Count - zeroIndex);
                return allEmails.GetRange(zeroIndex, count);
            }

            public EmailDocument? FetchSingleEmail(string storeId, string entryId)
            {
                foreach (var list in FolderEmails.Values)
                {
                    var found = list.FirstOrDefault(e => e.StoreId == storeId && e.EntryId == entryId);
                    if (found != null) return found;
                }
                return null;
            }

            public bool OpenEmailInOutlook(string storeId, string entryId) => true;
            public bool IsOutboxFolder(string storeId, string folderEntryId) => OutboxEntryIds.Contains(folderEntryId);
            public void SubscribeToEvents(Action<string, string, string> onItemAddedOrChanged, Action<string, string> onItemRemoved) { }
            public void UnsubscribeFromEvents() { }
        }

        [Fact]
        public async Task Test01_Duplicate_Prevention_When_Backfilling_Existing_Records()
        {
            // Seed repository with an existing email
            var existing = new EmailDocument
            {
                StoreId = "active-store-1",
                EntryId = "entry-001",
                FolderEntryId = "inbox-folder",
                FolderName = "Inbox",
                Subject = "Original Subject",
                SenderName = "Alice",
                SenderEmail = "alice@example.com",
                BodyText = "Initial body",
                ReceivedDate = new DateTime(2026, 7, 11)
            };
            await _repository.SaveEmailBatchAsync(new[] { existing });

            Assert.Equal(1, await _repository.GetIndexedEmailCountAsync());

            // Outlook has the same email plus a new one
            var outlook = new BackfillMockOutlookAdapter();
            outlook.Mailboxes.Add(new MailboxInfo { StoreId = "active-store-1", DisplayName = "Primary Mailbox" });
            outlook.StoreFolders["active-store-1"] = new List<MailboxFolder>
            {
                new MailboxFolder { StoreId = "active-store-1", EntryId = "inbox-folder", Name = "Inbox", ItemCount = 2 }
            };

            var email2 = new EmailDocument
            {
                StoreId = "active-store-1",
                EntryId = "entry-002",
                FolderEntryId = "inbox-folder",
                FolderName = "Inbox",
                Subject = "Second Subject",
                SenderName = "Bob",
                SenderEmail = "bob@example.com",
                BodyText = "Second body",
                ReceivedDate = new DateTime(2026, 7, 12)
            };

            outlook.FolderEmails["active-store-1:inbox-folder"] = new List<EmailDocument> { existing, email2 };

            var indexer = new EmailIndexerService(outlook, _repository, new MockLicensingService(LicenseStatus.Valid));
            var summary = await indexer.StartActiveStoreBackfillAsync();

            Assert.False(summary.IsCancelled);
            int countAfter = await _repository.GetIndexedEmailCountAsync();
            Assert.Equal(2, countAfter);

            // Verify entry-001 is still present and not duplicated
            var doc1 = await _repository.GetEmailByStoreAndEntryIdAsync("active-store-1", "entry-001");
            Assert.NotNull(doc1);
            Assert.Equal("Original Subject", doc1!.Subject);
        }

        [Fact]
        public async Task Test02_Preservation_Of_Legacy_Pst_Records_During_Active_Store_Backfill()
        {
            // Seed legacy PST emails (simulating 13,898 PST emails)
            var legacyEmails = new List<EmailDocument>();
            for (int i = 1; i <= 5; i++)
            {
                legacyEmails.Add(new EmailDocument
                {
                    StoreId = "pst-archive-001",
                    EntryId = $"pst-entry-{i:D4}",
                    FolderEntryId = "pst-inbox",
                    FolderName = "PST Inbox",
                    Subject = $"Archived email {i}",
                    SenderName = "Legacy Sender",
                    SenderEmail = "legacy@example.com",
                    ReceivedDate = new DateTime(2026, 1, i)
                });
            }
            await _repository.SaveEmailBatchAsync(legacyEmails);
            Assert.Equal(5, await _repository.GetIndexedEmailCountAsync());

            // Active store has 3 separate emails
            var outlook = new BackfillMockOutlookAdapter();
            outlook.Mailboxes.Add(new MailboxInfo { StoreId = "ost-active-001", DisplayName = "Active Exchange Account" });
            outlook.StoreFolders["ost-active-001"] = new List<MailboxFolder>
            {
                new MailboxFolder { StoreId = "ost-active-001", EntryId = "ost-inbox", Name = "Inbox", ItemCount = 3 }
            };

            var activeEmails = new List<EmailDocument>();
            for (int i = 1; i <= 3; i++)
            {
                activeEmails.Add(new EmailDocument
                {
                    StoreId = "ost-active-001",
                    EntryId = $"ost-entry-{i:D4}",
                    FolderEntryId = "ost-inbox",
                    FolderName = "Inbox",
                    Subject = $"Live email {i}",
                    SenderName = "Live Sender",
                    SenderEmail = "live@example.com",
                    ReceivedDate = new DateTime(2026, 7, 11)
                });
            }
            outlook.FolderEmails["ost-active-001:ost-inbox"] = activeEmails;

            var indexer = new EmailIndexerService(outlook, _repository, new MockLicensingService(LicenseStatus.Valid));
            await indexer.StartActiveStoreBackfillAsync();

            // Total is now 5 PST + 3 OST = 8
            Assert.Equal(8, await _repository.GetIndexedEmailCountAsync());

            // All legacy PST records are completely intact
            for (int i = 1; i <= 5; i++)
            {
                var doc = await _repository.GetEmailByStoreAndEntryIdAsync("pst-archive-001", $"pst-entry-{i:D4}");
                Assert.NotNull(doc);
                Assert.Equal($"Archived email {i}", doc!.Subject);
            }
        }

        [Fact]
        public async Task Test03_Interrupted_And_Resumed_Backfill_From_Checkpoint()
        {
            // Folder with 500 items (2 batches of 250)
            var outlook = new BackfillMockOutlookAdapter();
            outlook.Mailboxes.Add(new MailboxInfo { StoreId = "ost-001", DisplayName = "Exchange" });
            outlook.StoreFolders["ost-001"] = new List<MailboxFolder>
            {
                new MailboxFolder { StoreId = "ost-001", EntryId = "folder-large", Name = "Large Folder", ItemCount = 500 }
            };

            var allEmails = new List<EmailDocument>();
            for (int i = 1; i <= 500; i++)
            {
                allEmails.Add(new EmailDocument
                {
                    StoreId = "ost-001",
                    EntryId = $"entry-{i:D4}",
                    FolderEntryId = "folder-large",
                    FolderName = "Large Folder",
                    Subject = $"Email {i}",
                    SenderEmail = $"sender{i}@example.com",
                    ReceivedDate = DateTime.Now
                });
            }
            outlook.FolderEmails["ost-001:folder-large"] = allEmails;

            // 1. First run: cancel after batch 1 (250 items)
            using (var cts = new CancellationTokenSource())
            {
                var indexer1 = new EmailIndexerService(outlook, _repository, new MockLicensingService(LicenseStatus.Valid));
                indexer1.ProgressChanged += (s, e) =>
                {
                    if (e.ProcessedItems >= 250)
                    {
                        cts.Cancel();
                    }
                };

                var res1 = await indexer1.StartActiveStoreBackfillAsync(cts.Token);
                Assert.True(res1.IsCancelled);
            }

            // Verify checkpoint was saved at offset 250
            var state = await _repository.GetFolderIndexStateAsync("ost-001", "folder-large");
            Assert.NotNull(state);
            Assert.Equal(250, state!.Value.ProcessedOffset);
            Assert.Equal("Indexing", state.Value.Status);
            Assert.Equal(250, await _repository.GetIndexedEmailCountAsync());

            // 2. Second run: resume from checkpoint without cancellation
            outlook.RequestedOffsets.Clear();
            var indexer2 = new EmailIndexerService(outlook, _repository, new MockLicensingService(LicenseStatus.Valid));
            var res2 = await indexer2.StartActiveStoreBackfillAsync();

            Assert.False(res2.IsCancelled);
            // Verify it started from offset 251 instead of 1
            Assert.Contains(251, outlook.RequestedOffsets);
            Assert.DoesNotContain(1, outlook.RequestedOffsets);

            // Final count is 500
            Assert.Equal(500, await _repository.GetIndexedEmailCountAsync());

            // Checkpoint status is now Completed
            var finalState = await _repository.GetFolderIndexStateAsync("ost-001", "folder-large");
            Assert.NotNull(finalState);
            Assert.Equal("Completed", finalState!.Value.Status);
            Assert.Equal(500, finalState.Value.ProcessedOffset);
        }

        [Fact]
        public async Task Test04_Folder_Discovery_Includes_SentMail_And_Excludes_Outbox()
        {
            var outlook = new BackfillMockOutlookAdapter();
            outlook.Mailboxes.Add(new MailboxInfo { StoreId = "store-discovery", DisplayName = "User Mailbox" });
            outlook.OutboxEntryIds.Add("folder-outbox");

            outlook.StoreFolders["store-discovery"] = new List<MailboxFolder>
            {
                new MailboxFolder { StoreId = "store-discovery", EntryId = "folder-inbox", Name = "Inbox", ItemCount = 10 },
                new MailboxFolder { StoreId = "store-discovery", EntryId = "folder-sent", Name = "Sent Mail", ItemCount = 5 },
                new MailboxFolder { StoreId = "store-discovery", EntryId = "folder-important", Name = "Important", ItemCount = 3 },
                new MailboxFolder { StoreId = "store-discovery", EntryId = "folder-outbox", Name = "Outbox", ItemCount = 4 }
            };

            // Seed emails for each folder
            void SeedFolder(string folderId, string name, int count)
            {
                var list = new List<EmailDocument>();
                for (int i = 1; i <= count; i++)
                {
                    list.Add(new EmailDocument
                    {
                        StoreId = "store-discovery",
                        EntryId = $"{folderId}-msg-{i}",
                        FolderEntryId = folderId,
                        FolderName = name,
                        Subject = $"{name} item {i}",
                        ReceivedDate = DateTime.Now
                    });
                }
                outlook.FolderEmails[$"store-discovery:{folderId}"] = list;
            }

            SeedFolder("folder-inbox", "Inbox", 10);
            SeedFolder("folder-sent", "Sent Mail", 5);
            SeedFolder("folder-important", "Important", 3);
            SeedFolder("folder-outbox", "Outbox", 4);

            var indexer = new EmailIndexerService(outlook, _repository, new MockLicensingService(LicenseStatus.Valid));

            // Coverage check should discover 10 + 5 + 3 = 18 items (excluding Outbox 4)
            var coverage = await indexer.CheckStoreCoverageAsync();
            Assert.Equal(18, coverage.TotalDiscoveredItems);
            Assert.Equal(18, coverage.UnindexedItems);
            Assert.Equal(3, coverage.UnindexedFolderCount);

            // Run backfill
            var summary = await indexer.StartActiveStoreBackfillAsync();
            Assert.Equal(18, summary.NewlyIndexedCount);
            Assert.Equal(18, await _repository.GetIndexedEmailCountAsync());

            // Outbox was not indexed
            var outboxMsg = await _repository.GetEmailByStoreAndEntryIdAsync("store-discovery", "folder-outbox-msg-1");
            Assert.Null(outboxMsg);

            // Sent Mail was indexed
            var sentMsg = await _repository.GetEmailByStoreAndEntryIdAsync("store-discovery", "folder-sent-msg-1");
            Assert.NotNull(sentMsg);
            Assert.Equal("Sent Mail item 1", sentMsg!.Subject);
        }

        [Fact]
        public async Task Test05_Multiple_Stores_Coverage_And_Discovery()
        {
            var outlook = new BackfillMockOutlookAdapter();
            outlook.Mailboxes.Add(new MailboxInfo { StoreId = "store-1", DisplayName = "Archive Mailbox" });
            outlook.Mailboxes.Add(new MailboxInfo { StoreId = "store-2", DisplayName = "Primary Mailbox" });

            outlook.StoreFolders["store-1"] = new List<MailboxFolder>
            {
                new MailboxFolder { StoreId = "store-1", EntryId = "f1", Name = "Archive Inbox", ItemCount = 10 }
            };
            outlook.StoreFolders["store-2"] = new List<MailboxFolder>
            {
                new MailboxFolder { StoreId = "store-2", EntryId = "f2", Name = "Active Inbox", ItemCount = 20 }
            };

            // Seed store 1 as already indexed in DB
            var s1Emails = new List<EmailDocument>();
            for (int i = 1; i <= 10; i++)
            {
                s1Emails.Add(new EmailDocument
                {
                    StoreId = "store-1",
                    EntryId = $"s1-msg-{i}",
                    FolderEntryId = "f1",
                    FolderName = "Archive Inbox",
                    Subject = $"S1 Item {i}",
                    ReceivedDate = DateTime.Now
                });
            }
            await _repository.SaveEmailBatchAsync(s1Emails);

            var indexer = new EmailIndexerService(outlook, _repository, new MockLicensingService(LicenseStatus.Valid));
            var coverage = await indexer.CheckStoreCoverageAsync();

            Assert.Equal(2, coverage.ActiveStoreCount);
            Assert.Equal(1, coverage.UnindexedStoreCount); // store-2 has 0 in DB
            Assert.Equal(1, coverage.UnindexedFolderCount); // f2 has 20 unindexed
            Assert.Equal(30, coverage.TotalDiscoveredItems); // 10 + 20
            Assert.Equal(20, coverage.UnindexedItems); // 20 from store-2
            Assert.False(coverage.IsComplete);
        }

        [Fact]
        public async Task Test06_Diagnostics_Separates_Database_Integrity_From_Mailbox_Coverage()
        {
            var outlook = new BackfillMockOutlookAdapter();
            outlook.Mailboxes.Add(new MailboxInfo { StoreId = "store-active", DisplayName = "User Active Store" });
            outlook.StoreFolders["store-active"] = new List<MailboxFolder>
            {
                new MailboxFolder { StoreId = "store-active", EntryId = "f-active", Name = "Inbox", ItemCount = 100 }
            };

            var indexer = new EmailIndexerService(outlook, _repository, new MockLicensingService(LicenseStatus.Valid));
            var healthService = new IndexHealthService(_repository, indexer);

            var summary = await healthService.GetHealthSummaryAsync();

            // Database integrity is 100% healthy
            Assert.True(summary.IntegrityResult.IsHealthy);
            // But status state says Database Healthy rather than implying all mail is indexed
            Assert.Contains("Database Healthy", summary.StatusState);

            // Coverage is clearly separated and reports incomplete
            Assert.False(summary.IsCoverageComplete);
            Assert.Equal(100, summary.TotalDiscoveredOutlookItems);
            Assert.Equal(100, summary.UnindexedOutlookItems);
            Assert.Contains("unindexed", summary.MailboxCoverageStatus, StringComparison.OrdinalIgnoreCase);
        }
    }
}
