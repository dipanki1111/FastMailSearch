using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using FastMailSearch.Core.Models;
using FastMailSearch.Core.Search;
using FastMailSearch.Infrastructure.Database;
using Xunit;

namespace FastMailSearch.Database.Tests
{
    public class DatabaseSearchQualityTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly DatabaseConnectionFactory _factory;
        private readonly SqliteEmailRepository _repository;

        public DatabaseSearchQualityTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), $"SearchQualityTest_{Guid.NewGuid():N}.db");
            _factory = new DatabaseConnectionFactory(_dbPath);
            _repository = new SqliteEmailRepository(_factory);
            _repository.InitializeDatabase();

            // Populate test emails
            var emails = new List<EmailDocument>
            {
                new EmailDocument
                {
                    StoreId = "S1",
                    EntryId = "E1",
                    Subject = "Quarterly Invoice Report",
                    SenderName = "John Doe",
                    SenderEmail = "john@example.com",
                    Recipients = "accounts@example.com",
                    Cc = "manager@example.com",
                    BodyText = "Please find the invoice attached for Q1.",
                    ReceivedDate = new DateTime(2026, 2, 1),
                    FolderName = "Inbox",
                    HasAttachment = true
                },
                new EmailDocument
                {
                    StoreId = "S1",
                    EntryId = "E2",
                    Subject = "Project Meeting Notes",
                    SenderName = "Alice Smith",
                    SenderEmail = "alice@example.com",
                    Recipients = "team@example.com",
                    BodyText = "Discussion about the new feature timeline.",
                    ReceivedDate = new DateTime(2026, 3, 15),
                    FolderName = "Inbox",
                    HasAttachment = false
                },
                new EmailDocument
                {
                    StoreId = "S1",
                    EntryId = "E3",
                    Subject = "Spam Promotion Invoice",
                    SenderName = "Spammer",
                    SenderEmail = "spam@junk.com",
                    Recipients = "all@example.com",
                    BodyText = "Buy cheap products spam.",
                    ReceivedDate = new DateTime(2026, 4, 1),
                    FolderName = "Junk",
                    HasAttachment = false
                }
            };

            _repository.SaveEmailBatchAsync(emails).GetAwaiter().GetResult();
        }

        public void Dispose()
        {
            try { if (File.Exists(_dbPath)) File.Delete(_dbPath); } catch { }
        }

        [Fact]
        public async Task Search_FilterBySender_ReturnsMatching()
        {
            var q = SearchQueryParser.Parse("from:john@example.com");
            var result = await _repository.SearchAsync(q);

            Assert.Equal(1, result.TotalCount);
            Assert.Equal("Quarterly Invoice Report", result.Items[0].Subject);
        }

        [Fact]
        public async Task Search_FilterByDateRange_ReturnsMatching()
        {
            var q = SearchQueryParser.Parse("after:2026-03-01 before:2026-03-31");
            var result = await _repository.SearchAsync(q);

            Assert.Equal(1, result.TotalCount);
            Assert.Equal("Project Meeting Notes", result.Items[0].Subject);
        }

        [Fact]
        public async Task Search_Exclusion_FiltersOutSpam()
        {
            var q = SearchQueryParser.Parse("Invoice -spam");
            var result = await _repository.SearchAsync(q);

            Assert.Equal(1, result.TotalCount);
            Assert.Equal("Quarterly Invoice Report", result.Items[0].Subject);
        }

        [Fact]
        public async Task Search_HasAttachment_ReturnsAttachmentEmails()
        {
            var q = SearchQueryParser.Parse("has:attachment");
            var result = await _repository.SearchAsync(q);

            Assert.Equal(1, result.TotalCount);
            Assert.True(result.Items[0].HasAttachment);
        }

        [Fact]
        public async Task Search_FilterBySubject_ReturnsMatching()
        {
            var q = SearchQueryParser.Parse("subject:Invoice");
            var result = await _repository.SearchAsync(q);

            Assert.Equal(2, result.TotalCount);
        }

        [Fact]
        public async Task Search_FilterByRecipientAndCc_ReturnsMatching()
        {
            var qTo = SearchQueryParser.Parse("to:accounts@example.com");
            var resTo = await _repository.SearchAsync(qTo);
            Assert.Equal(1, resTo.TotalCount);

            var qCc = SearchQueryParser.Parse("cc:manager@example.com");
            var resCc = await _repository.SearchAsync(qCc);
            Assert.Equal(1, resCc.TotalCount);
        }

        [Fact]
        public async Task Search_FilterByFolder_ReturnsMatchingFolder()
        {
            var q = SearchQueryParser.Parse("folder:Junk");
            var result = await _repository.SearchAsync(q);

            Assert.Equal(1, result.TotalCount);
            Assert.Equal("Spam Promotion Invoice", result.Items[0].Subject);
        }

        [Fact]
        public async Task Search_SortOrder_OldestFirstAndNewestFirst_Works()
        {
            var qNewest = SearchQueryParser.Parse("Invoice");
            qNewest.SortOrder = SortOption.NewestFirst;
            var resNewest = await _repository.SearchAsync(qNewest);
            Assert.Equal(2, resNewest.TotalCount);
            Assert.Equal("E3", resNewest.Items[0].EntryId); // Apr 2026 > Feb 2026

            var qOldest = SearchQueryParser.Parse("Invoice");
            qOldest.SortOrder = SortOption.OldestFirst;
            var resOldest = await _repository.SearchAsync(qOldest);
            Assert.Equal(2, resOldest.TotalCount);
            Assert.Equal("E1", resOldest.Items[0].EntryId); // Feb 2026 < Apr 2026
        }
    }
}
