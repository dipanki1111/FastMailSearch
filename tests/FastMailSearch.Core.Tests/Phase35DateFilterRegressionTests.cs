using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using FastMailSearch.Core.Models;
using FastMailSearch.Core.Search;
using FastMailSearch.Infrastructure.Database;
using Xunit;

namespace FastMailSearch.Core.Tests
{
    public class Phase35DateFilterRegressionTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly DatabaseConnectionFactory _factory;
        private readonly SqliteEmailRepository _repository;

        public Phase35DateFilterRegressionTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), $"Phase35DateFilterTests_{Guid.NewGuid():N}.db");
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

        [Fact]
        public async Task Test01_DateTo_AtMidnight_IncludesEmailsSentLaterOnTheEndDate()
        {
            // Seed email sent on October 21, 2026 at 14:30:00 (mid-day)
            var email = new EmailDocument
            {
                StoreId = "Store1",
                EntryId = "Entry_Oct21_Midday",
                Subject = "Afternoon Status Report",
                SenderName = "Alice",
                SenderEmail = "alice@example.com",
                ReceivedDate = new DateTime(2026, 10, 21, 14, 30, 0, DateTimeKind.Unspecified),
                FolderName = "Inbox",
                BodyText = "Status update for October 21st."
            };
            await _repository.SaveEmailBatchAsync(new List<EmailDocument> { email });

            // Query with DateTo at midnight (00:00:00), exactly as WPF DatePicker produces
            var query = new SearchQuery
            {
                DateFrom = new DateTime(2026, 10, 21, 0, 0, 0),
                DateTo = new DateTime(2026, 10, 21, 0, 0, 0)
            };

            var result = await _repository.SearchAsync(query);

            Assert.Equal(1, result.TotalCount);
            Assert.Single(result.Items);
            Assert.Equal("Afternoon Status Report", result.Items[0].Subject);
        }

        [Fact]
        public async Task Test02_DateRange2026_IncludesAllEmailsWithin2026_AndExcludes2025()
        {
            var emails = new List<EmailDocument>
            {
                new EmailDocument
                {
                    StoreId = "Store1",
                    EntryId = "Entry_2025_Nov",
                    Subject = "Old Email from November 2025",
                    SenderName = "Narendra",
                    SenderEmail = "naren@example.com",
                    ReceivedDate = new DateTime(2025, 11, 15, 9, 57, 56),
                    FolderName = "Inbox",
                    BodyText = "Mac Mini Setup"
                },
                new EmailDocument
                {
                    StoreId = "Store1",
                    EntryId = "Entry_2026_Jan",
                    Subject = "New Year 2026 Notice",
                    SenderName = "MarketScreener",
                    SenderEmail = "noreply@marketscreener.com",
                    ReceivedDate = new DateTime(2026, 1, 1, 15, 42, 41),
                    FolderName = "Inbox",
                    BodyText = "Market open notice"
                },
                new EmailDocument
                {
                    StoreId = "Store1",
                    EntryId = "Entry_2026_Oct12",
                    Subject = "Order Confirmation",
                    SenderName = "Flipkart",
                    SenderEmail = "no-reply@rmp.flipkart.com",
                    ReceivedDate = new DateTime(2026, 10, 12, 23, 6, 5),
                    FolderName = "Inbox",
                    BodyText = "Your item has shipped."
                },
                new EmailDocument
                {
                    StoreId = "Store1",
                    EntryId = "Entry_2026_Oct21_Evening",
                    Subject = "Evening Invoice",
                    SenderName = "Flipkart",
                    SenderEmail = "no-reply@rmp.flipkart.com",
                    ReceivedDate = new DateTime(2026, 10, 21, 22, 38, 0),
                    FolderName = "Inbox",
                    BodyText = "Invoice copy"
                },
                new EmailDocument
                {
                    StoreId = "Store1",
                    EntryId = "Entry_2026_Oct22",
                    Subject = "Future Email Outside Range",
                    SenderName = "Google",
                    SenderEmail = "no-reply@accounts.google.com",
                    ReceivedDate = new DateTime(2026, 10, 22, 8, 0, 0),
                    FolderName = "Inbox",
                    BodyText = "Security alert"
                }
            };
            await _repository.SaveEmailBatchAsync(emails);

            // Filter for DateFrom = 2026-01-01 and DateTo = 2026-10-21
            var query = new SearchQuery
            {
                DateFrom = new DateTime(2026, 1, 1),
                DateTo = new DateTime(2026, 10, 21)
            };

            var result = await _repository.SearchAsync(query);

            // Should match: Jan 1, Oct 12, and Oct 21 evening (total: 3). Should exclude Nov 2025 and Oct 22.
            Assert.Equal(3, result.TotalCount);
            Assert.DoesNotContain(result.Items, x => x.Subject.Contains("Old Email"));
            Assert.DoesNotContain(result.Items, x => x.Subject.Contains("Future Email"));
            Assert.Contains(result.Items, x => x.Subject == "Evening Invoice");
        }

        [Fact]
        public async Task Test03_RupenSearch_DateFilter2026_ReturnsZeroWhenOnly2025EmailsExist()
        {
            var emails = new List<EmailDocument>
            {
                new EmailDocument
                {
                    StoreId = "Store1",
                    EntryId = "Rupen_2025_1",
                    Subject = "RE: REGARDING SYSTEMS SLOW WORKING",
                    SenderName = "Rupen Sheth",
                    SenderEmail = "rupen@ascominfotech.com",
                    ReceivedDate = new DateTime(2025, 11, 10, 18, 28, 6),
                    FolderName = "Inbox",
                    BodyText = "We checked the network and systems."
                },
                new EmailDocument
                {
                    StoreId = "Store1",
                    EntryId = "Rupen_2025_2",
                    Subject = "Fw: Stylsmyth Mac Mini Setup",
                    SenderName = "Narendra",
                    SenderEmail = "naren@ascominfotech.com",
                    Recipients = "rupen@ascominfotech.com",
                    ReceivedDate = new DateTime(2025, 11, 15, 9, 57, 56),
                    FolderName = "Inbox",
                    BodyText = "Forwarded note regarding Mac setup."
                },
                new EmailDocument
                {
                    StoreId = "Store1",
                    EntryId = "Other_2026",
                    Subject = "Flipkart Order Delivered",
                    SenderName = "Flipkart",
                    SenderEmail = "no-reply@rmp.flipkart.com",
                    ReceivedDate = new DateTime(2026, 5, 20, 11, 0, 0),
                    FolderName = "Inbox",
                    BodyText = "Package delivered successfully."
                }
            };
            await _repository.SaveEmailBatchAsync(emails);

            // 1. Without date filter: "rupen" matches 2 emails
            var qNoDate = SearchQueryParser.Parse("rupen");
            var resNoDate = await _repository.SearchAsync(qNoDate);
            Assert.Equal(2, resNoDate.TotalCount);

            // 2. With 2026 date filter: DateFrom = 2026-01-01, DateTo = 2026-10-21
            var q2026 = SearchQueryParser.Parse("rupen");
            q2026.DateFrom = new DateTime(2026, 1, 1);
            q2026.DateTo = new DateTime(2026, 10, 21);
            var res2026 = await _repository.SearchAsync(q2026);
            // Must return 0 because no rupen emails exist in 2026
            Assert.Equal(0, res2026.TotalCount);

            // 3. With 2025 date filter: returns the 2 emails
            var q2025 = SearchQueryParser.Parse("rupen");
            q2025.DateFrom = new DateTime(2025, 1, 1);
            q2025.DateTo = new DateTime(2025, 12, 31);
            var res2025 = await _repository.SearchAsync(q2025);
            Assert.Equal(2, res2025.TotalCount);
        }

        [Fact]
        public async Task Test04_DateFrom_AtMidnight_IncludesEmailsSentAtAnyTimeOnStartDate()
        {
            var emails = new List<EmailDocument>
            {
                new EmailDocument
                {
                    StoreId = "Store1",
                    EntryId = "Entry_Jan1_Midnight",
                    Subject = "Midnight Email",
                    SenderName = "System",
                    SenderEmail = "system@test.com",
                    ReceivedDate = new DateTime(2026, 1, 1, 0, 0, 0),
                    FolderName = "Inbox",
                    BodyText = "Happy new year at 00:00:00"
                },
                new EmailDocument
                {
                    StoreId = "Store1",
                    EntryId = "Entry_Jan1_Noon",
                    Subject = "Noon Email",
                    SenderName = "System",
                    SenderEmail = "system@test.com",
                    ReceivedDate = new DateTime(2026, 1, 1, 12, 0, 0),
                    FolderName = "Inbox",
                    BodyText = "New year lunch"
                }
            };
            await _repository.SaveEmailBatchAsync(emails);

            var query = new SearchQuery
            {
                DateFrom = new DateTime(2026, 1, 1, 0, 0, 0),
                DateTo = new DateTime(2026, 1, 1, 23, 59, 59)
            };

            var result = await _repository.SearchAsync(query);
            Assert.Equal(2, result.TotalCount);
        }

        [Fact]
        public async Task Test05_Timezone_Kind_Consistency_DoesNotBreakLexicographicalComparison()
        {
            var email = new EmailDocument
            {
                StoreId = "Store1",
                EntryId = "Entry_LocalKind",
                Subject = "Timezone Test Email",
                SenderName = "Tester",
                SenderEmail = "tester@test.com",
                ReceivedDate = new DateTime(2026, 6, 15, 10, 0, 0, DateTimeKind.Local),
                FolderName = "Inbox",
                BodyText = "Testing local vs unspecified kind"
            };
            await _repository.SaveEmailBatchAsync(new List<EmailDocument> { email });

            // Query with DateTimeKind.Local
            var queryLocal = new SearchQuery
            {
                DateFrom = DateTime.SpecifyKind(new DateTime(2026, 6, 15, 0, 0, 0), DateTimeKind.Local),
                DateTo = DateTime.SpecifyKind(new DateTime(2026, 6, 15, 0, 0, 0), DateTimeKind.Local)
            };

            var resultLocal = await _repository.SearchAsync(queryLocal);
            Assert.Equal(1, resultLocal.TotalCount);

            // Query with DateTimeKind.Utc
            var queryUtc = new SearchQuery
            {
                DateFrom = DateTime.SpecifyKind(new DateTime(2026, 6, 15, 0, 0, 0), DateTimeKind.Utc),
                DateTo = DateTime.SpecifyKind(new DateTime(2026, 6, 15, 0, 0, 0), DateTimeKind.Utc)
            };

            var resultUtc = await _repository.SearchAsync(queryUtc);
            Assert.Equal(1, resultUtc.TotalCount);
        }
    }
}
