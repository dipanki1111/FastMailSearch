using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using FastMailSearch.Core.Models;
using FastMailSearch.Infrastructure.Database;
using Xunit;

namespace FastMailSearch.Core.Tests
{
    public class PaginationAndCancellationTests
    {
        [Fact]
        public void SearchQuery_PaginationMapping_CalculatesLimitAndOffsetCorrectly()
        {
            var query = new SearchQuery
            {
                PageSize = 50,
                Page = 1
            };

            Assert.Equal(50, query.Limit);
            Assert.Equal(0, query.Offset);
            Assert.Equal(1, query.Page);
            Assert.Equal(50, query.PageSize);

            query.Page = 2;
            Assert.Equal(50, query.Offset);
            Assert.Equal(2, query.Page);

            query.Page = 3;
            Assert.Equal(100, query.Offset);
            Assert.Equal(3, query.Page);

            query.PageSize = 25;
            query.Page = 4;
            Assert.Equal(25, query.Limit);
            Assert.Equal(75, query.Offset);
            Assert.Equal(4, query.Page);
        }

        [Fact]
        public void PagedSearchResult_CalculatesTotalPagesAndNavigation()
        {
            var paged = new PagedSearchResult<string>
            {
                TotalCount = 125,
                PageSize = 50,
                Page = 1
            };

            Assert.Equal(3, paged.TotalPages);
            Assert.False(paged.HasPreviousPage);
            Assert.True(paged.HasNextPage);

            paged.Page = 2;
            Assert.True(paged.HasPreviousPage);
            Assert.True(paged.HasNextPage);

            paged.Page = 3;
            Assert.True(paged.HasPreviousPage);
            Assert.False(paged.HasNextPage);
        }

        [Fact]
        public async Task Database_Pagination_ReturnsExactPageItemsAndCorrectTotalCount()
        {
            string dbPath = Path.Combine(Path.GetTempPath(), $"PaginationTest_{Guid.NewGuid():N}.db");
            try
            {
                var factory = new DatabaseConnectionFactory(dbPath);
                var repo = new SqliteEmailRepository(factory);
                repo.InitializeDatabase();

                var batch = new List<EmailDocument>();
                for (int i = 1; i <= 125; i++)
                {
                    batch.Add(new EmailDocument
                    {
                        StoreId = "store1",
                        EntryId = $"entry_{i}",
                        Subject = $"Pagination Invoice Test {i}",
                        SenderName = "Sender",
                        SenderEmail = "sender@example.com",
                        Recipients = "user@example.com",
                        BodyText = $"Body text content for invoice email number {i}",
                        ReceivedDate = DateTime.UtcNow.AddMinutes(-i),
                        FolderName = "Inbox"
                    });
                }

                await repo.SaveEmailBatchAsync(batch);

                // Query Page 1
                var query1 = new SearchQuery { RawQuery = "Invoice", PageSize = 50, Page = 1 };
                var res1 = await repo.SearchAsync(query1);
                Assert.Equal(125, res1.TotalCount);
                Assert.Equal(50, res1.Items.Count);
                Assert.Equal(1, res1.Page);

                // Query Page 2
                var query2 = new SearchQuery { RawQuery = "Invoice", PageSize = 50, Page = 2 };
                var res2 = await repo.SearchAsync(query2);
                Assert.Equal(125, res2.TotalCount);
                Assert.Equal(50, res2.Items.Count);
                Assert.Equal(2, res2.Page);

                // Query Page 3
                var query3 = new SearchQuery { RawQuery = "Invoice", PageSize = 50, Page = 3 };
                var res3 = await repo.SearchAsync(query3);
                Assert.Equal(125, res3.TotalCount);
                Assert.Equal(25, res3.Items.Count);
                Assert.Equal(3, res3.Page);

                // Query Out of Bounds (Page 5)
                var query5 = new SearchQuery { RawQuery = "Invoice", PageSize = 50, Page = 5 };
                var res5 = await repo.SearchAsync(query5);
                Assert.Equal(125, res5.TotalCount);
                Assert.Empty(res5.Items);
            }
            finally
            {
                if (File.Exists(dbPath))
                {
                    try { File.Delete(dbPath); } catch { }
                }
            }
        }

        [Fact]
        public async Task SearchAsync_WithCancelledToken_ThrowsOperationCanceledException()
        {
            string dbPath = Path.Combine(Path.GetTempPath(), $"CancellationTest_{Guid.NewGuid():N}.db");
            try
            {
                var factory = new DatabaseConnectionFactory(dbPath);
                var repo = new SqliteEmailRepository(factory);
                repo.InitializeDatabase();

                var batch = new List<EmailDocument>();
                for (int i = 1; i <= 50; i++)
                {
                    batch.Add(new EmailDocument
                    {
                        StoreId = "store1",
                        EntryId = $"entry_{i}",
                        Subject = $"Cancellation Test Email {i}",
                        SenderName = "Sender",
                        SenderEmail = "sender@example.com",
                        BodyText = $"Body {i}",
                        ReceivedDate = DateTime.UtcNow
                    });
                }
                await repo.SaveEmailBatchAsync(batch);

                var cts = new CancellationTokenSource();
                cts.Cancel(); // Pre-cancelled token

                var query = new SearchQuery { RawQuery = "Cancellation" };

                await Assert.ThrowsAsync<OperationCanceledException>(async () =>
                {
                    await repo.SearchAsync(query, cts.Token);
                });
            }
            finally
            {
                if (File.Exists(dbPath))
                {
                    try { File.Delete(dbPath); } catch { }
                }
            }
        }
    }
}
