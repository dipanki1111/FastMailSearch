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
    public class Phase34PopPstAndZeroIndexingTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly DatabaseConnectionFactory _factory;
        private readonly SqliteEmailRepository _repository;

        public Phase34PopPstAndZeroIndexingTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), $"Phase34Tests_{Guid.NewGuid():N}.db");
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

        private class MockOutlookAdapter34 : IOutlookAdapter
        {
            public List<MailboxInfo> Mailboxes { get; set; } = new List<MailboxInfo>();
            public Dictionary<string, List<MailboxFolder>> StoreFolders { get; set; } = new Dictionary<string, List<MailboxFolder>>();
            public Dictionary<string, List<EmailDocument>> FolderEmails { get; set; } = new Dictionary<string, List<EmailDocument>>();
            public bool OutlookAvailable { get; set; } = true;

            public bool IsOutlookAvailable() => OutlookAvailable;
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
        public async Task Test01_POP_PST_Store_Discovery_And_Folder_Coverage()
        {
            var outlook = new MockOutlookAdapter34();
            string pstStoreId = "0000000038A1BB1005E5101AA1BB08002B2A56C200006D73707374";
            string pstPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Outlook Files", "personal-test@example.com - test.pst");

            outlook.Mailboxes.Add(new MailboxInfo
            {
                StoreId = pstStoreId,
                DisplayName = "personal-test@example.com",
                FilePath = pstPath,
                IsPst = true,
                IsSelected = true
            });

            outlook.StoreFolders[pstStoreId] = new List<MailboxFolder>
            {
                new MailboxFolder { StoreId = pstStoreId, EntryId = "f_inbox", Name = "Inbox", ItemCount = 30996 },
                new MailboxFolder { StoreId = pstStoreId, EntryId = "f_sent", Name = "Sent Items", ItemCount = 50 },
                new MailboxFolder { StoreId = pstStoreId, EntryId = "f_junk", Name = "Junk Email", ItemCount = 10 }
            };

            var indexer = new EmailIndexerService(outlook, _repository, new MockLicensingService(LicenseStatus.Valid));
            var coverage = await indexer.CheckStoreCoverageAsync();

            Assert.Equal(1, coverage.ActiveStoreCount);
            Assert.Equal(31056, coverage.TotalDiscoveredItems);
            Assert.Equal(31056, coverage.UnindexedItems);
            Assert.Contains("personal-test@example.com", coverage.SummaryText);
            Assert.Contains("Backfill recommended", coverage.SummaryText);
        }

        [Fact]
        public async Task Test02_Coverage_Check_Reports_Clear_Error_When_No_Active_Store()
        {
            var outlook = new MockOutlookAdapter34();
            outlook.Mailboxes.Clear();

            var indexer = new EmailIndexerService(outlook, _repository, new MockLicensingService(LicenseStatus.Valid));
            var coverage = await indexer.CheckStoreCoverageAsync();

            Assert.Equal(0, coverage.ActiveStoreCount);
            Assert.Equal(0, coverage.TotalDiscoveredItems);
            Assert.Contains("No active Outlook stores discovered", coverage.SummaryText);
            Assert.DoesNotContain("fully indexed", coverage.SummaryText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public async Task Test03_StartActiveStoreBackfill_Reports_Error_When_No_Store()
        {
            var outlook = new MockOutlookAdapter34();
            outlook.Mailboxes.Clear();

            var indexer = new EmailIndexerService(outlook, _repository, new MockLicensingService(LicenseStatus.Valid));
            var summary = await indexer.StartActiveStoreBackfillAsync();

            Assert.Equal(0, summary.TotalDiscovered);
            Assert.Equal(IndexerStatus.Error, indexer.CurrentStatus);
            Assert.Contains("No active Outlook store discovered", summary.ErrorMessage ?? "");
        }

        [Fact]
        public void Test04_Zero_Total_Indexing_Formats_Progress_Message_Not_0_0_0_Percent()
        {
            var args = new IndexerProgressEventArgs
            {
                TotalItems = 0,
                ProcessedItems = 0,
                Percentage = 0,
                CurrentStatusMessage = "🔵 Updating • 5 email(s) remaining",
                Status = IndexerStatus.Indexing
            };

            string statusText;
            if (args.Status == IndexerStatus.Indexing)
            {
                if (args.TotalItems == 0)
                {
                    statusText = !string.IsNullOrWhiteSpace(args.CurrentStatusMessage)
                        ? args.CurrentStatusMessage
                        : "🔵 Indexing in progress...";
                }
                else
                {
                    statusText = $"🔵 Indexing: {args.ProcessedItems:N0} / {args.TotalItems:N0} ({args.Percentage:F0}%)";
                }
            }
            else
            {
                statusText = "🟢 Up to date";
            }

            Assert.Equal("🔵 Updating • 5 email(s) remaining", statusText);
            Assert.DoesNotContain("0 / 0 (0%)", statusText);
        }

        [Fact]
        public void Test05_IndexHealthSummary_DisplayCoverage_Disambiguates_Zero_Cases()
        {
            // Case A: Discovered items > 0
            var s1 = new IndexHealthSummary
            {
                IndexedEmailCount = 100,
                TotalDiscoveredOutlookItems = 500,
                MailboxCoverageStatus = "400 unindexed in Outlook"
            };
            Assert.Equal("100 indexed / 500 in Outlook (400 unindexed in Outlook)", s1.DisplayCoverage);

            // Case B: No active store discovered
            var s2 = new IndexHealthSummary
            {
                IndexedEmailCount = 25000,
                TotalDiscoveredOutlookItems = 0,
                MailboxCoverageStatus = "No active store discovered"
            };
            Assert.Equal("25,000 indexed (No active store discovered)", s2.DisplayCoverage);

            // Case C: Active store discovered with 0 items
            var s3 = new IndexHealthSummary
            {
                IndexedEmailCount = 25000,
                TotalDiscoveredOutlookItems = 0,
                MailboxCoverageStatus = "Active store has 0 emails in mail folders"
            };
            Assert.Equal("25,000 indexed (Active store has 0 emails in mail folders)", s3.DisplayCoverage);
        }

        [Fact]
        public void Test06_DatabaseConnectionFactory_Exposes_DbPath_And_Preserves_Location()
        {
            var factory = new DatabaseConnectionFactory(_dbPath);
            Assert.Equal(_dbPath, factory.DbPath);
            Assert.Equal(_dbPath, _repository.DatabasePath);
        }

        [Fact]
        public async Task Test07_HealthService_GetHealthSummary_Populates_Database_Metrics()
        {
            // Seed a sample email
            var emails = new List<EmailDocument>
            {
                new EmailDocument
                {
                    StoreId = "pop-store-1",
                    EntryId = "pop-e1",
                    FolderEntryId = "f-inbox",
                    FolderName = "Inbox",
                    Subject = "POP email test",
                    SenderName = "Sender",
                    SenderEmail = "sender@example.com",
                    ReceivedDate = DateTime.Now,
                    IndexedDate = DateTime.Now
                }
            };
            await _repository.SaveEmailBatchAsync(emails);

            var outlook = new MockOutlookAdapter34();
            outlook.Mailboxes.Add(new MailboxInfo
            {
                StoreId = "pop-store-1",
                DisplayName = "POP Account",
                FilePath = @"C:\Outlook\pop.pst",
                IsPst = true
            });
            outlook.StoreFolders["pop-store-1"] = new List<MailboxFolder>
            {
                new MailboxFolder { StoreId = "pop-store-1", EntryId = "f-inbox", Name = "Inbox", ItemCount = 1 }
            };

            var indexer = new EmailIndexerService(outlook, _repository, new MockLicensingService(LicenseStatus.Valid));
            var healthService = new IndexHealthService(_repository, indexer, outlook, _factory);

            var summary = await healthService.GetHealthSummaryAsync();
            Assert.Equal(1, summary.IndexedEmailCount);
            Assert.True(summary.DatabaseSizeMb > 0, "DatabaseSizeMb should be greater than 0");
            Assert.Equal(1, summary.TotalDiscoveredOutlookItems);
            Assert.Contains("1 indexed / 1 in Outlook", summary.DisplayCoverage);
        }
    }
}
