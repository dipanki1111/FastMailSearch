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
    public class SqliteEmailRepositoryTests : IDisposable
    {
        private readonly string _testDbPath;
        private readonly DatabaseConnectionFactory _factory;
        private readonly SqliteEmailRepository _repository;

        public SqliteEmailRepositoryTests()
        {
            _testDbPath = Path.Combine(Path.GetTempPath(), $"FastMailTest_{Guid.NewGuid():N}.db");
            _factory = new DatabaseConnectionFactory(_testDbPath);
            _repository = new SqliteEmailRepository(_factory);
            _repository.InitializeDatabase();
        }

        public void Dispose()
        {
            try
            {
                if (File.Exists(_testDbPath)) File.Delete(_testDbPath);
            }
            catch { }
        }

        [Fact]
        public async Task SaveEmailBatch_And_SearchAsync_ReturnsMatchingRecords()
        {
            var emails = new List<EmailDocument>
            {
                new EmailDocument
                {
                    StoreId = "Store1",
                    EntryId = "Entry1",
                    Subject = "Invoice INV-001",
                    SenderName = "ABC Ltd",
                    SenderEmail = "accounts@abc.com",
                    Recipients = "me@company.com",
                    BodyText = "Please find attached invoice for March payment",
                    ReceivedDate = new DateTime(2026, 4, 1),
                    FolderName = "Inbox",
                    HasAttachment = true,
                    Attachments = new List<EmailAttachment>
                    {
                        new EmailAttachment { Name = "invoice.pdf", Extension = ".pdf", Size = 1024 }
                    }
                },
                new EmailDocument
                {
                    StoreId = "Store1",
                    EntryId = "Entry2",
                    Subject = "Weekly Status Report",
                    SenderName = "John Smith",
                    SenderEmail = "john@company.com",
                    Recipients = "team@company.com",
                    BodyText = "All project milestones completed successfully",
                    ReceivedDate = new DateTime(2026, 3, 25),
                    FolderName = "Inbox",
                    HasAttachment = false
                }
            };

            await _repository.SaveEmailBatchAsync(emails);

            var query = SearchQueryParser.Parse("invoice ABC");
            var result = await _repository.SearchAsync(query);

            Assert.True(result.TotalCount >= 1);
            Assert.Equal("Invoice INV-001", result.Items[0].Subject);
        }

        [Fact]
        public async Task SaveEmailBatch_PreventsDuplicates_OnStoreIdAndEntryId()
        {
            var email = new EmailDocument
            {
                StoreId = "Store1",
                EntryId = "Entry1",
                Subject = "Original Subject",
                SenderName = "Original Sender",
                ReceivedDate = DateTime.Now,
                FolderName = "Inbox"
            };

            await _repository.SaveEmailBatchAsync(new[] { email });

            email.Subject = "Updated Subject";
            await _repository.SaveEmailBatchAsync(new[] { email });

            var meta = await _repository.GetIndexMetadataAsync();
            Assert.Equal(1, meta.TotalEmailsIndexed);

            var fetched = await _repository.GetEmailByStoreAndEntryIdAsync("Store1", "Entry1");
            Assert.NotNull(fetched);
            Assert.Equal("Updated Subject", fetched.Subject);
        }

        [Fact]
        public async Task DeleteEmailAsync_RemovesEmailFromIndex()
        {
            var email = new EmailDocument
            {
                StoreId = "Store1",
                EntryId = "EntryToDelete",
                Subject = "Delete Me",
                ReceivedDate = DateTime.Now,
                FolderName = "Inbox"
            };

            await _repository.SaveEmailBatchAsync(new[] { email });
            await _repository.DeleteEmailAsync("Store1", "EntryToDelete");

            var fetched = await _repository.GetEmailByStoreAndEntryIdAsync("Store1", "EntryToDelete");
            Assert.Null(fetched);
        }
    }
}
