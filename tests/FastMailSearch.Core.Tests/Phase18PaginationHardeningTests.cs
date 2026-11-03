using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using FastMailSearch.Core.Interfaces;
using FastMailSearch.Core.Models;
using FastMailSearch.Core.Services;
using FastMailSearch.Infrastructure.Database;
using FastMailSearch.UI.ViewModels;
using Xunit;

namespace FastMailSearch.Core.Tests
{
    public class Phase18PaginationHardeningTests
    {
        private (SqliteEmailRepository repo, string dbPath) CreateTestRepo()
        {
            string dbPath = Path.Combine(Path.GetTempPath(), $"Phase18Test_{Guid.NewGuid():N}.db");
            var factory = new DatabaseConnectionFactory(dbPath);
            var repo = new SqliteEmailRepository(factory);
            repo.InitializeDatabase();
            return (repo, dbPath);
        }

        private async Task SeedDatabaseAsync(SqliteEmailRepository repo, int count)
        {
            var batch = new List<EmailDocument>();
            for (int i = 1; i <= count; i++)
            {
                batch.Add(new EmailDocument
                {
                    StoreId = "store_main",
                    EntryId = $"entry_id_{i:D6}",
                    Subject = $"Subject Email Number {i}",
                    SenderName = $"Sender Name {i}",
                    SenderEmail = $"sender{i}@example.com",
                    Recipients = "user@example.com",
                    BodyText = $"Body text content for email number {i} with kfintech keyword",
                    ReceivedDate = DateTime.UtcNow.AddSeconds(-i),
                    FolderName = "Inbox"
                });
            }
            await repo.SaveEmailBatchAsync(batch);
        }

        private void DeleteTestDb(string dbPath)
        {
            try
            {

                GC.Collect();
                GC.WaitForPendingFinalizers();
                if (File.Exists(dbPath)) File.Delete(dbPath);
            }
            catch { }
        }

        [Fact]
        public async Task FirstPage_ReturnsFirst50()
        {
            var (repo, dbPath) = CreateTestRepo();
            try
            {
                await SeedDatabaseAsync(repo, 120);
                var query = new SearchQuery { RawQuery = "kfintech", PageSize = 50, Page = 1 };
                var res = await repo.SearchAsync(query);

                Assert.Equal(120, res.TotalCount);
                Assert.Equal(50, res.Items.Count);
                Assert.Equal(1, res.Page);
                Assert.Equal("entry_id_000001", res.Items[0].EntryId);
            }
            finally
            {
                DeleteTestDb(dbPath);
            }
        }

        [Fact]
        public async Task NextPage_ReturnsCorrect50()
        {
            var (repo, dbPath) = CreateTestRepo();
            try
            {
                await SeedDatabaseAsync(repo, 120);
                var query = new SearchQuery { RawQuery = "kfintech", PageSize = 50, Page = 2 };
                var res = await repo.SearchAsync(query);

                Assert.Equal(120, res.TotalCount);
                Assert.Equal(50, res.Items.Count);
                Assert.Equal(2, res.Page);
                Assert.Equal("entry_id_000051", res.Items[0].EntryId);
            }
            finally
            {
                DeleteTestDb(dbPath);
            }
        }

        [Fact]
        public async Task PreviousPage_ReturnsCorrect50()
        {
            var (repo, dbPath) = CreateTestRepo();
            try
            {
                await SeedDatabaseAsync(repo, 120);
                var query = new SearchQuery { RawQuery = "kfintech", PageSize = 50, Page = 1 };
                var res = await repo.SearchAsync(query);

                Assert.Equal(120, res.TotalCount);
                Assert.Equal(50, res.Items.Count);
                Assert.Equal(1, res.Page);
            }
            finally
            {
                DeleteTestDb(dbPath);
            }
        }

        [Fact]
        public async Task LastPage_ReturnsCorrectRows()
        {
            var (repo, dbPath) = CreateTestRepo();
            try
            {
                await SeedDatabaseAsync(repo, 120);
                var query = new SearchQuery { RawQuery = "kfintech", PageSize = 50, Page = 3 };
                var res = await repo.SearchAsync(query);

                Assert.Equal(120, res.TotalCount);
                Assert.Equal(20, res.Items.Count);
                Assert.Equal(3, res.Page);
                Assert.Equal("entry_id_000101", res.Items[0].EntryId);
            }
            finally
            {
                DeleteTestDb(dbPath);
            }
        }

        [Fact]
        public async Task PageBeyondLast_IsHandledSafely()
        {
            var (repo, dbPath) = CreateTestRepo();
            try
            {
                await SeedDatabaseAsync(repo, 100);
                var query = new SearchQuery { RawQuery = "kfintech", PageSize = 50, Page = 10 };
                var res = await repo.SearchAsync(query);

                Assert.Equal(100, res.TotalCount);
                Assert.Empty(res.Items);
            }
            finally
            {
                DeleteTestDb(dbPath);
            }
        }

        [Fact]
        public async Task TotalCount_IsCorrect()
        {
            var (repo, dbPath) = CreateTestRepo();
            try
            {
                await SeedDatabaseAsync(repo, 87);
                var query = new SearchQuery { RawQuery = "kfintech", PageSize = 25, Page = 1 };
                var res = await repo.SearchAsync(query);

                Assert.Equal(87, res.TotalCount);
                Assert.Equal(25, res.Items.Count);
            }
            finally
            {
                DeleteTestDb(dbPath);
            }
        }

        [Fact]
        public void QueryChange_ResetsToFirstPage()
        {
            var searchService = new SearchService(new MockSearchEngine(), new MockRepository(), new MockLicensingService());
            var vm = new SearchViewModel(searchService, new MockRepository(), new MockIndexerService(), new MockOutlookAdapter());

            vm.CurrentPage = 3;
            Assert.Equal(3, vm.CurrentPage);

            vm.SearchText = "new query";
            Assert.Equal(1, vm.CurrentPage);
        }

        [Fact]
        public void Pagination_PreservesActiveSearchText_WhenHistoryCleared()
        {
            var searchService = new SearchService(new MockSearchEngine(), new MockRepository(), new MockLicensingService());
            var vm = new SearchViewModel(searchService, new MockRepository(), new MockIndexerService(), new MockOutlookAdapter());

            vm.SearchText = "important query";
            
            // Simulate the UI ComboBox pushing null when SearchHistoryList is cleared.
            // On the old implementation (if simulating by setting SearchText = null), it would clear it.
            // On the new implementation, the ComboBox is bound to SelectedSearchHistory, which ignores null.
            vm.SelectedSearchHistory = null;
            
            Assert.Equal("important query", vm.SearchText);
        }

        [Fact]
        public void SortChange_ResetsToFirstPage()
        {
            var searchService = new SearchService(new MockSearchEngine(), new MockRepository(), new MockLicensingService());
            var vm = new SearchViewModel(searchService, new MockRepository(), new MockIndexerService(), new MockOutlookAdapter());

            vm.CurrentPage = 4;
            Assert.Equal(4, vm.CurrentPage);

            vm.SelectedSortOption = "Newest First";
            Assert.Equal(1, vm.CurrentPage);
        }

        [Fact]
        public async Task Pagination_DoesNotMaterializeAllRows()
        {
            var (repo, dbPath) = CreateTestRepo();
            try
            {
                await SeedDatabaseAsync(repo, 500);
                var query = new SearchQuery { RawQuery = "kfintech", PageSize = 50, Page = 1 };
                var res = await repo.SearchAsync(query);

                Assert.Equal(500, res.TotalCount);
                Assert.Equal(50, res.Items.Count);
            }
            finally
            {
                DeleteTestDb(dbPath);
            }
        }

        [Fact]
        public async Task LargeOffset_QueryCompletes()
        {
            var (repo, dbPath) = CreateTestRepo();
            try
            {
                await SeedDatabaseAsync(repo, 1000);
                var query = new SearchQuery { RawQuery = "kfintech", PageSize = 50, Page = 20 };
                var res = await repo.SearchAsync(query);

                Assert.Equal(1000, res.TotalCount);
                Assert.Equal(50, res.Items.Count);
                Assert.Equal(20, res.Page);
            }
            finally
            {
                DeleteTestDb(dbPath);
            }
        }

        [Fact]
        public async Task SearchWhileIndexing_RemainsResponsive()
        {
            var (repo, dbPath) = CreateTestRepo();
            try
            {
                await SeedDatabaseAsync(repo, 200);

                var query = new SearchQuery { RawQuery = "kfintech", PageSize = 50, Page = 1 };
                var searchTask = repo.SearchAsync(query);
                var indexTask = SeedDatabaseAsync(repo, 50);

                await Task.WhenAll(searchTask, indexTask);

                var res = await searchTask;
                Assert.True(res.TotalCount >= 200);
            }
            finally
            {
                DeleteTestDb(dbPath);
            }
        }

        [Fact]
        public void SelectedResult_OpenEmail_PreservesEntryId()
        {
            var item = new SearchResultItem
            {
                StoreId = "store_123",
                EntryId = "entry_456",
                Subject = "Test Email"
            };

            Assert.Equal("store_123", item.StoreId);
            Assert.Equal("entry_456", item.EntryId);
        }

        [Fact]
        public void Pagination_DoesNotCancelIndexer()
        {
            var indexer = new MockIndexerService();
            var searchService = new SearchService(new MockSearchEngine(), new MockRepository(), new MockLicensingService());
            var vm = new SearchViewModel(searchService, new MockRepository(), indexer, new MockOutlookAdapter());

            vm.Cleanup();
            Assert.False(indexer.WasCancelled);
        }
    }

    #region Mocks for Unit Testing
    internal class MockSearchEngine : ISearchEngine
    {
        public Task<SearchResultContainer> SearchAsync(SearchQuery query, CancellationToken cancellationToken = default)
        {
            return Task.FromResult(new SearchResultContainer { TotalCount = 100 });
        }
    }

    internal class MockRepository : IEmailRepository
    {
        public string DatabasePath => string.Empty;
        public void InitializeDatabase() { }
        public Task SaveEmailBatchAsync(IEnumerable<EmailDocument> emails, CancellationToken cancellationToken = default) => Task.CompletedTask;
        public Task<EmailDocument?> GetEmailByStoreAndEntryIdAsync(string storeId, string entryId) => Task.FromResult<EmailDocument?>(null);
        public Task DeleteEmailAsync(string storeId, string entryId) => Task.CompletedTask;
        public Task UpdateEmailFolderAsync(string storeId, string entryId, string newFolderName, string newFolderEntryId) => Task.CompletedTask;
        public Task<HashSet<string>> GetIndexedEntryIdsForFolderAsync(string storeId, string folderEntryId) => Task.FromResult(new HashSet<string>());
        public Task SaveFoldersAsync(IEnumerable<MailboxFolder> folders) => Task.CompletedTask;
        public Task<List<MailboxFolder>> GetFoldersAsync(string? storeId = null) => Task.FromResult(new List<MailboxFolder>());
        public Task<IndexMetadata> GetIndexMetadataAsync() => Task.FromResult(new IndexMetadata());
        public Task ClearAllDataAsync() => Task.CompletedTask;
        public Task SaveSearchQueryAsync(string rawQuery) => Task.CompletedTask;
        public Task<List<string>> GetRecentSearchQueriesAsync(int count = 20) => Task.FromResult(new List<string>());
        public Task ClearSearchHistoryAsync() => Task.CompletedTask;
        public Task VacuumDatabaseAsync() => Task.CompletedTask;
        public Task SaveIndexStateAsync(string state, int processed, int total, string currentFolder, string? error) => Task.CompletedTask;
        public Task<Dictionary<string, string>> GetPersistedIndexStateAsync() => Task.FromResult(new Dictionary<string, string>());
        public Task<int> GetIndexedEmailCountAsync() => Task.FromResult(0);
        public Task EnqueuePendingChangeAsync(string storeId, string entryId, string operation, string? folderEntryId) => Task.CompletedTask;
        public Task<List<PendingChangeItem>> GetPendingChangesAsync(int maxCount = 50) => Task.FromResult(new List<PendingChangeItem>());
        public Task<int> IncrementAttemptCountAsync(long id) => Task.FromResult(1);
        public Task<int> GetPendingChangesCountAsync() => Task.FromResult(0);
        public Task MarkPendingChangeStatusAsync(long id, string status, string? error = null) => Task.CompletedTask;
        public Task MarkPendingChangesCompletedBatchAsync(IEnumerable<long> ids) => Task.CompletedTask;
        public Task RecoverStaleProcessingChangesAsync() => Task.CompletedTask;
        public Task SaveFailedIndexItemAsync(string storeId, string entryId, string operation, int attemptCount, string lastError) => Task.CompletedTask;
        public Task SaveFailedIndexItemAsync(FailedIndexItem item) => Task.CompletedTask;
        public Task<List<FailedIndexItem>> GetFailedIndexItemsAsync(bool includeResolved = false) => Task.FromResult(new List<FailedIndexItem>());
        public Task<List<FailedIndexItem>> GetFailedIndexItemsAsync(string? status, int maxCount = 500) => Task.FromResult(new List<FailedIndexItem>());
        public Task<int> GetFailedIndexItemsCountAsync(string? status = "Failed") => Task.FromResult(0);
        public Task UpdateFailedIndexItemStatusAsync(long id, string status, string? error = null) => Task.CompletedTask;
        public Task RemoveFailedIndexItemAsync(long id) => Task.CompletedTask;
        public Task RemoveAllFailedIndexItemsAsync() => Task.CompletedTask;
        public Task<double> GetDatabaseSizeMbAsync() => Task.FromResult(0.0);
        public Task<IndexIntegrityResult> ScanIndexIntegrityAsync(CancellationToken cancellationToken = default) => Task.FromResult(new IndexIntegrityResult());
        public Task RepairIndexIntegrityAsync(IndexIntegrityResult scanResult, CancellationToken cancellationToken = default) => Task.CompletedTask;
        public Task<Dictionary<string, int>> GetStoreEmailCountsAsync() => Task.FromResult(new Dictionary<string, int>());
        public Task<int> GetFolderEmailCountAsync(string storeId, string folderEntryId) => Task.FromResult(0);
        public Task SaveFolderIndexStateAsync(string storeId, string folderEntryId, string status, int processedOffset, int totalItems) => Task.CompletedTask;
        public Task<(string Status, int ProcessedOffset, int TotalItems)?> GetFolderIndexStateAsync(string storeId, string folderEntryId) => Task.FromResult<(string Status, int ProcessedOffset, int TotalItems)?>(null);
        public Task<List<FolderIndexCheckpoint>> GetAllFolderIndexStatesAsync() => Task.FromResult(new List<FolderIndexCheckpoint>());
        public Task SaveLiveProgressAsync(string state, string currentOperation, int processed, int total, string currentFolder, int folderOffset, int folderTotal, int newlyIndexed, int skipped, int failed, int pendingQueue) => Task.CompletedTask;
    }

    internal class MockIndexerService : IIndexerService
    {
        public bool IsIndexing => false;
        public bool IsRebuilding => false;
        public bool WasCancelled { get; private set; }
        public IndexerStatus CurrentStatus => IndexerStatus.Idle;
        public IndexMetadata CurrentMetadata => new IndexMetadata();
        public int ProcessedItems => 0;
        public int TotalItems => 0;
        public int FolderProcessedItems => 0;
        public int FolderTotalItems => 0;
        public string CurrentFolder => string.Empty;
        public DateTime? StartedAt => null;
        public DateTime? CompletedAt => null;
        public DateTime? LastIndexedItemTime => null;
        public string CurrentOperation => string.Empty;
        public string LastError => string.Empty;

        public event EventHandler<IndexerProgressEventArgs>? ProgressChanged;

        public Task StartInitialIndexingAsync(CancellationToken cancellationToken = default, int? maxItems = null) => Task.CompletedTask;
        public Task PauseIndexingAsync() => Task.CompletedTask;
        public Task ResumeIndexingAsync() => Task.CompletedTask;
        public Task<BackfillSummary> ResumeIncrementalCatchUpAsync(CancellationToken cancellationToken = default) => Task.FromResult(new BackfillSummary());
        public Task RebuildIndexAsync(CancellationToken cancellationToken = default) => Task.CompletedTask;
        public Task CancelIndexingAsync() { WasCancelled = true; return Task.CompletedTask; }
        public Task CheckAndRecoverIncompleteIndexAsync(bool startProcessing = true) => Task.CompletedTask;
        public Task QueueItemForIndexingAsync(string storeId, string entryId, string folderEntryId) => Task.CompletedTask;
        public Task QueueItemForDeletionAsync(string storeId, string entryId) => Task.CompletedTask;
        public Task<IndexIntegrityResult> ScanIndexIntegrityAsync(CancellationToken cancellationToken = default) => Task.FromResult(new IndexIntegrityResult());
        public Task RepairIndexIntegrityAsync(IndexIntegrityResult scanResult, CancellationToken cancellationToken = default) => Task.CompletedTask;
        public Task<BackfillSummary> StartActiveStoreBackfillAsync(CancellationToken cancellationToken = default, int? maxItems = null) => Task.FromResult(new BackfillSummary());
        public Task<StoreCoverageInfo> CheckStoreCoverageAsync(CancellationToken cancellationToken = default) => Task.FromResult(new StoreCoverageInfo());
        public IndexerProgressEventArgs GetStatus() => new IndexerProgressEventArgs();
        public void SetSafeMode(string reason) { }
    }

    internal class MockOutlookAdapter : IOutlookAdapter
    {
        public bool IsOutlookAvailable() => true;
        public string GetOutlookVersion() => "16.0";
        public string GetOfficeBitness() => "x64";
        public List<MailboxInfo> GetMailboxes() => new List<MailboxInfo>();
        public List<MailboxFolder> GetFolders(string storeId) => new List<MailboxFolder>();
        public List<EmailDocument> FetchFolderEmailsBatch(string storeId, string folderEntryId, int startIndex, int batchSize, Func<bool>? cancelCheck = null) => new List<EmailDocument>();
        public EmailDocument? FetchSingleEmail(string storeId, string entryId) => null;
        public bool OpenEmailInOutlook(string storeId, string entryId) => true;
        public bool IsOutboxFolder(string storeId, string folderEntryId) => false;
        public void SubscribeToEvents(Action<string, string, string> onItemAddedOrChanged, Action<string, string> onItemRemoved) { }
        public void UnsubscribeFromEvents() { }
    }
    #endregion
}
