using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using FastMailSearch.Core.Models;
using FastMailSearch.Core.Search;
using FastMailSearch.Infrastructure.Database;
using Xunit;

namespace FastMailSearch.Database.Tests
{
    public class Stage7XOptimizationTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly DatabaseConnectionFactory _factory;
        private readonly SqliteEmailRepository _repository;

        public Stage7XOptimizationTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), $"Stage7XTests_{Guid.NewGuid():N}.db");
            _factory = new DatabaseConnectionFactory(_dbPath);
            _repository = new SqliteEmailRepository(_factory);
            _repository.InitializeDatabase();

            var emails = new List<EmailDocument>
            {
                new EmailDocument
                {
                    StoreId = "S1",
                    EntryId = "E1",
                    Subject = "Welcome to FastMailSearch",
                    SenderName = "FastMail Team",
                    SenderEmail = "welcome@service.com",
                    Recipients = "user@example.com",
                    BodyText = "Your account is activated.",
                    ReceivedDate = new DateTime(2026, 1, 1),
                    FolderName = "Inbox",
                    HasAttachment = false
                },
                new EmailDocument
                {
                    StoreId = "S1",
                    EntryId = "E2",
                    Subject = "Monthly Invoice 100",
                    SenderName = "Billing Dept",
                    SenderEmail = "billing@service.com",
                    Recipients = "user@example.com",
                    BodyText = "Payment of 100 received. Transaction id paytmqr60gb8d.",
                    ReceivedDate = new DateTime(2026, 1, 10),
                    FolderName = "Important",
                    HasAttachment = true
                },
                new EmailDocument
                {
                    StoreId = "S1",
                    EntryId = "E3",
                    Subject = "Invoice 1000 reminder",
                    SenderName = "Accounts",
                    SenderEmail = "accounts@company.com",
                    Recipients = "user@example.com",
                    BodyText = "Payment of 1000 due. Transaction id 44081756961.",
                    ReceivedDate = new DateTime(2026, 1, 15),
                    FolderName = "Important",
                    HasAttachment = false
                },
                new EmailDocument
                {
                    StoreId = "S1",
                    EntryId = "E4",
                    Subject = "मराठी संदेश १०",
                    SenderName = "मराठी युवा",
                    SenderEmail = "marathi@example.in",
                    Recipients = "user@example.com",
                    BodyText = "हा १० क्रमांकाचा संदेश आहे.",
                    ReceivedDate = new DateTime(2026, 2, 1),
                    FolderName = "Inbox",
                    HasAttachment = false
                },
                new EmailDocument
                {
                    StoreId = "S1",
                    EntryId = "E5",
                    Subject = "मराठी संदेश १००",
                    SenderName = "मराठी युवा",
                    SenderEmail = "marathi@example.in",
                    Recipients = "user@example.com",
                    BodyText = "हा १०० क्रमांकाचा संदेश आहे.",
                    ReceivedDate = new DateTime(2026, 2, 5),
                    FolderName = "Inbox",
                    HasAttachment = false
                },
                new EmailDocument
                {
                    StoreId = "S1",
                    EntryId = "E6",
                    Subject = "Project launch 🚀 announcement",
                    SenderName = "Tech Lead",
                    SenderEmail = "lead@tech.com",
                    Recipients = "team@example.com",
                    BodyText = "The new release is ready 🚀 for deployment.",
                    ReceivedDate = new DateTime(2026, 2, 10),
                    FolderName = "Inbox",
                    HasAttachment = false
                },
                new EmailDocument
                {
                    StoreId = "S1",
                    EntryId = "E7",
                    Subject = "Happy Birthday 🎂 wishes",
                    SenderName = "Friend",
                    SenderEmail = "friend@domain.com",
                    Recipients = "user@example.com",
                    BodyText = "Wishing you a great year ahead 🎂!",
                    ReceivedDate = new DateTime(2026, 2, 15),
                    FolderName = "Inbox",
                    HasAttachment = false
                },
                new EmailDocument
                {
                    StoreId = "S1",
                    EntryId = "E8",
                    Subject = "Café résumé review",
                    SenderName = "HR Manager",
                    SenderEmail = "hr@corp.com",
                    Recipients = "user@example.com",
                    BodyText = "Attached is the résumé with café experience.",
                    ReceivedDate = new DateTime(2026, 2, 20),
                    FolderName = "Sent Mail",
                    HasAttachment = false
                }
            };

            _repository.SaveEmailBatchAsync(emails).GetAwaiter().GetResult();
        }

        public void Dispose()
        {
            try { if (File.Exists(_dbPath)) File.Delete(_dbPath); } catch { }
        }

        #region Scope 1: Field Filter Performance & Behavior

        [Fact]
        public async Task SubjectFilter_MatchesTargetSubject()
        {
            var q = SearchQueryParser.Parse("subject:Welcome");
            var result = await _repository.SearchAsync(q);

            Assert.Equal(1, result.TotalCount);
            Assert.Equal("Welcome to FastMailSearch", result.Items[0].Subject);
        }

        [Fact]
        public async Task SubjectFilter_MultipleMatches()
        {
            var q = SearchQueryParser.Parse("subject:Invoice");
            var result = await _repository.SearchAsync(q);

            Assert.Equal(2, result.TotalCount);
        }

        [Fact]
        public async Task FolderFilter_ExactAndCaseInsensitiveSeek()
        {
            var qUpper = SearchQueryParser.Parse("folder:Important");
            var resUpper = await _repository.SearchAsync(qUpper);
            Assert.Equal(2, resUpper.TotalCount);

            var qLower = SearchQueryParser.Parse("folder:important");
            var resLower = await _repository.SearchAsync(qLower);
            Assert.Equal(2, resLower.TotalCount);
        }

        [Fact]
        public async Task FolderFilter_SentMail_Matches()
        {
            var q = SearchQueryParser.Parse("folder:Sent");
            var result = await _repository.SearchAsync(q);

            Assert.Equal(1, result.TotalCount);
            Assert.Equal("Café résumé review", result.Items[0].Subject);
        }

        [Fact]
        public async Task FromFilter_SenderEmailAndDisplayName_BothMatch()
        {
            var qEmail = SearchQueryParser.Parse("from:welcome@service.com");
            var resEmail = await _repository.SearchAsync(qEmail);
            Assert.Equal(1, resEmail.TotalCount);
            Assert.Equal("Welcome to FastMailSearch", resEmail.Items[0].Subject);

            var qName = SearchQueryParser.Parse("from:\"FastMail Team\"");
            var resName = await _repository.SearchAsync(qName);
            Assert.Equal(1, resName.TotalCount);
            Assert.Equal("Welcome to FastMailSearch", resName.Items[0].Subject);
        }

        [Fact]
        public async Task ToFilter_RecipientMatches()
        {
            var q = SearchQueryParser.Parse("to:team@example.com");
            var result = await _repository.SearchAsync(q);

            Assert.Equal(1, result.TotalCount);
            Assert.Equal("Project launch 🚀 announcement", result.Items[0].Subject);
        }

        #endregion

        #region Scope 2: Numeric Prefix Search

        [Fact]
        public async Task NumericSearch_English_ExactMatch_DoesNotExpandPrefix()
        {
            // Searching "100" must NOT match "1000"
            var q100 = SearchQueryParser.Parse("100");
            var res100 = await _repository.SearchAsync(q100);
            Assert.Equal(1, res100.TotalCount);
            Assert.Equal("Monthly Invoice 100", res100.Items[0].Subject);

            // Searching "1000" must NOT match "100"
            var q1000 = SearchQueryParser.Parse("1000");
            var res1000 = await _repository.SearchAsync(q1000);
            Assert.Equal(1, res1000.TotalCount);
            Assert.Equal("Invoice 1000 reminder", res1000.Items[0].Subject);
        }

        [Fact]
        public async Task NumericSearch_Devanagari_ExactMatch_DoesNotExpandPrefix()
        {
            // Searching Devanagari १० must NOT match १००
            var q10 = SearchQueryParser.Parse("१०");
            var res10 = await _repository.SearchAsync(q10);
            Assert.Equal(1, res10.TotalCount);
            Assert.Equal("मराठी संदेश १०", res10.Items[0].Subject);

            // Searching Devanagari १०० must NOT match १०
            var q100 = SearchQueryParser.Parse("१००");
            var res100 = await _repository.SearchAsync(q100);
            Assert.Equal(1, res100.TotalCount);
            Assert.Equal("मराठी संदेश १००", res100.Items[0].Subject);
        }

        [Fact]
        public async Task NumericSearch_LongNumericId_MatchesExact()
        {
            var q = SearchQueryParser.Parse("44081756961");
            var res = await _repository.SearchAsync(q);

            Assert.Equal(1, res.TotalCount);
            Assert.Equal("Invoice 1000 reminder", res.Items[0].Subject);
        }

        [Fact]
        public async Task AlphanumericToken_TransactionId_PrefixPreserved()
        {
            // Full transaction token
            var qFull = SearchQueryParser.Parse("paytmqr60gb8d");
            var resFull = await _repository.SearchAsync(qFull);
            Assert.Equal(1, resFull.TotalCount);

            // Prefix search preserved for alphanumeric token
            var qPrefix = SearchQueryParser.Parse("paytmqr");
            var resPrefix = await _repository.SearchAsync(qPrefix);
            Assert.Equal(1, resPrefix.TotalCount);
            Assert.Equal("Monthly Invoice 100", resPrefix.Items[0].Subject);
        }

        #endregion

        #region Scope 3: Emoji Search

        [Fact]
        public async Task EmojiSearch_Rocket_StandaloneMatches()
        {
            var q = SearchQueryParser.Parse("🚀");
            var res = await _repository.SearchAsync(q);

            Assert.Equal(1, res.TotalCount);
            Assert.Equal("Project launch 🚀 announcement", res.Items[0].Subject);
        }

        [Fact]
        public async Task EmojiSearch_Cake_StandaloneMatches()
        {
            var q = SearchQueryParser.Parse("🎂");
            var res = await _repository.SearchAsync(q);

            Assert.Equal(1, res.TotalCount);
            Assert.Equal("Happy Birthday 🎂 wishes", res.Items[0].Subject);
        }

        [Fact]
        public async Task EmojiSearch_MixedTextAndEmoji_Matches()
        {
            var q = SearchQueryParser.Parse("launch 🚀");
            var res = await _repository.SearchAsync(q);

            Assert.Equal(1, res.TotalCount);
            Assert.Equal("Project launch 🚀 announcement", res.Items[0].Subject);
        }

        [Fact]
        public async Task EmojiSearch_NonExistentEmoji_ReturnsZero()
        {
            var q = SearchQueryParser.Parse("🔥");
            var res = await _repository.SearchAsync(q);

            Assert.Equal(0, res.TotalCount);
        }

        #endregion

        #region Search Syntax: Boolean, Exclusion, Diacritics, Pagination, Cancellation

        [Fact]
        public async Task BooleanAnd_MatchesBothTerms()
        {
            var q = SearchQueryParser.Parse("Invoice 100");
            var res = await _repository.SearchAsync(q);

            Assert.Equal(1, res.TotalCount);
            Assert.Equal("Monthly Invoice 100", res.Items[0].Subject);
        }

        [Fact]
        public async Task BooleanOr_MatchesEitherTerm()
        {
            var q = SearchQueryParser.Parse("Invoice OR Birthday");
            var res = await _repository.SearchAsync(q);

            Assert.Equal(3, res.TotalCount);
        }

        [Fact]
        public async Task TermExclusion_ExcludesUnwantedEmails()
        {
            var q = SearchQueryParser.Parse("Invoice -1000");
            var res = await _repository.SearchAsync(q);

            Assert.Equal(1, res.TotalCount);
            Assert.Equal("Monthly Invoice 100", res.Items[0].Subject);
        }

        [Fact]
        public async Task DiacriticSearch_PreservesUnicodeBehavior()
        {
            var q = SearchQueryParser.Parse("résumé");
            var res = await _repository.SearchAsync(q);

            Assert.True(res.TotalCount >= 1);
            Assert.Equal("Café résumé review", res.Items[0].Subject);
        }

        [Fact]
        public async Task ZeroResultSearch_ReturnsCleanEmptyContainer()
        {
            var qZero = SearchQueryParser.Parse("NonExistentTermXYZ999");
            var resZero = await _repository.SearchAsync(qZero);
            Assert.Equal(0, resZero.TotalCount);
            Assert.Empty(resZero.Items);
            Assert.False(resZero.HasMoreResults);
        }

        [Fact]
        public async Task EmptyQuery_ReturnsAllIndexedEmails()
        {
            var qEmpty = SearchQueryParser.Parse("");
            var resEmpty = await _repository.SearchAsync(qEmpty);
            Assert.Equal(8, resEmpty.TotalCount);
            Assert.Equal(8, resEmpty.Items.Count);
        }

        [Fact]
        public async Task Pagination_LimitsAndOffsetsCorrectly()
        {
            var q = SearchQueryParser.Parse("Invoice");
            q.Limit = 1;
            q.Offset = 0;
            var page1 = await _repository.SearchAsync(q);

            Assert.Equal(2, page1.TotalCount);
            Assert.Single(page1.Items);
            Assert.True(page1.HasMoreResults);

            q.Offset = 1;
            var page2 = await _repository.SearchAsync(q);
            Assert.Equal(2, page2.TotalCount);
            Assert.Single(page2.Items);
            Assert.False(page2.HasMoreResults);
            Assert.NotEqual(page1.Items[0].EntryId, page2.Items[0].EntryId);
        }

        [Fact]
        public async Task SearchAsync_WhenCancelled_ThrowsOperationCanceledException()
        {
            using var cts = new CancellationTokenSource();
            cts.Cancel();

            var q = SearchQueryParser.Parse("Invoice");
            await Assert.ThrowsAnyAsync<OperationCanceledException>(async () =>
            {
                await _repository.SearchAsync(q, cts.Token);
            });
        }

        #endregion
    }
}
