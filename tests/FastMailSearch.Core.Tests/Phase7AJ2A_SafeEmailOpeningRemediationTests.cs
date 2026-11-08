using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using FastMailSearch.Core.Interfaces;
using FastMailSearch.Core.Models;
using FastMailSearch.Core.Services;
using FastMailSearch.Infrastructure.Outlook;
using FastMailSearch.UI.ViewModels;
using Xunit;

namespace FastMailSearch.Core.Tests
{
    public class Phase7AJ2A_SafeEmailOpeningRemediationTests
    {
        private (SearchViewModel vm, TrackingMockRepository repo, ConfigurableMockOutlookAdapter adapter, List<string> dialogMessages) CreateTestContext()
        {
            var repo = new TrackingMockRepository();
            var adapter = new ConfigurableMockOutlookAdapter();
            var searchEngine = new MockSearchEngine();
            var licensing = new MockLicensingService(LicenseStatus.Valid);
            var searchService = new SearchService(searchEngine, repo, licensing);
            var indexer = new MockIndexerService();

            var vm = new SearchViewModel(searchService, repo, indexer, adapter, licensing);

            var dialogMessages = new List<string>();
            vm.MessageBoxAction = (message, title) =>
            {
                dialogMessages.Add(message);
            };

            return (vm, repo, adapter, dialogMessages);
        }

        private SearchResultItem CreateSampleItem(string subject = "Important System Alert", string storeId = "Store-123", string entryId = "Entry-456")
        {
            return new SearchResultItem
            {
                EmailId = 101,
                StoreId = storeId,
                EntryId = entryId,
                Subject = subject,
                SenderName = "Security Team",
                SenderEmail = "security@example.com",
                ReceivedDate = new DateTime(2026, 11, 1, 10, 0, 0, DateTimeKind.Utc)
            };
        }

        [Fact]
        public void Requirement1_OutlookUnavailable_NoDatabaseDeletion_ItemPreserved_InformsUser()
        {
            var (vm, repo, adapter, dialogs) = CreateTestContext();
            adapter.IsAvailableResult = false;

            var item = CreateSampleItem();
            vm.SearchResults.Add(item);
            vm.TotalResultsCount = 1;

            vm.OpenResultCommand.Execute(item);

            // Assertions
            Assert.Equal(0, repo.DeleteEmailCallCount);
            Assert.Equal(0, adapter.OpenEmailCallCount); // COM not called when Outlook is closed
            Assert.Single(vm.SearchResults);
            Assert.Same(item, vm.SearchResults[0]);
            Assert.Equal(1, vm.TotalResultsCount);
            Assert.Single(dialogs);
            Assert.Contains("Outlook is not currently running", dialogs[0]);
        }

        [Fact]
        public void Requirement2_ComReturnsFalse_NoDatabaseDeletion_ItemPreserved_InformsUser()
        {
            var (vm, repo, adapter, dialogs) = CreateTestContext();
            adapter.IsAvailableResult = true;
            adapter.OpenEmailResult = false; // COM call failed / item not opened

            var item = CreateSampleItem();
            vm.SearchResults.Add(item);
            vm.TotalResultsCount = 1;

            vm.OpenResultCommand.Execute(item);

            // Assertions
            Assert.Equal(0, repo.DeleteEmailCallCount);
            Assert.Equal(1, adapter.OpenEmailCallCount);
            Assert.Single(vm.SearchResults);
            Assert.Same(item, vm.SearchResults[0]);
            Assert.Equal(1, vm.TotalResultsCount);
            Assert.Single(dialogs);
            Assert.Contains("Unable to open this email in Outlook", dialogs[0]);
        }

        [Fact]
        public void Requirement3_ComException_NoDatabaseDeletion_ItemPreserved_SafeHandling()
        {
            var (vm, repo, adapter, dialogs) = CreateTestContext();
            adapter.IsAvailableResult = true;
            adapter.OpenEmailException = new COMException("The RPC server is unavailable.", unchecked((int)0x800706BA));

            var item = CreateSampleItem();
            vm.SearchResults.Add(item);
            vm.TotalResultsCount = 1;

            // Must not throw unhandled exception
            var ex = Record.Exception(() => vm.OpenResultCommand.Execute(item));
            Assert.Null(ex);

            // Assertions
            Assert.Equal(0, repo.DeleteEmailCallCount);
            Assert.Single(vm.SearchResults);
            Assert.Same(item, vm.SearchResults[0]);
            Assert.Equal(1, vm.TotalResultsCount);
        }

        [Fact]
        public void Requirement4_InvalidEntryId_ReturnsFalse_NoDatabaseDeletion()
        {
            var (vm, repo, adapter, dialogs) = CreateTestContext();
            adapter.IsAvailableResult = true;
            adapter.OpenEmailResult = false;

            var item = CreateSampleItem(entryId: "CORRUPTED_ENTRY_ID_000000");
            vm.SearchResults.Add(item);
            vm.TotalResultsCount = 1;

            vm.OpenResultCommand.Execute(item);

            Assert.Equal(0, repo.DeleteEmailCallCount);
            Assert.Single(vm.SearchResults);
            Assert.Equal(1, vm.TotalResultsCount);
            Assert.Equal("CORRUPTED_ENTRY_ID_000000", adapter.LastOpenedEntryId);
        }

        [Fact]
        public void Requirement5_UnsupportedOutlookItemType_SafeFailure_NoDatabaseDeletion()
        {
            var (vm, repo, adapter, dialogs) = CreateTestContext();
            adapter.IsAvailableResult = true;
            // Simulated unsupported item failure (e.g. unknown custom item)
            adapter.OpenEmailResult = false;

            var item = CreateSampleItem(subject: "Custom Task Item");
            vm.SearchResults.Add(item);
            vm.TotalResultsCount = 1;

            vm.OpenResultCommand.Execute(item);

            Assert.Equal(0, repo.DeleteEmailCallCount);
            Assert.Single(vm.SearchResults);
            Assert.Equal(1, vm.TotalResultsCount);
            Assert.Single(dialogs);
            Assert.Contains("Unable to open this email in Outlook", dialogs[0]);
        }

        [Fact]
        public void Requirement6_SuccessfulMailItemOpening_SuccessPreserved_ItemRemains()
        {
            var (vm, repo, adapter, dialogs) = CreateTestContext();
            adapter.IsAvailableResult = true;
            adapter.OpenEmailResult = true;

            var item = CreateSampleItem();
            vm.SearchResults.Add(item);
            vm.TotalResultsCount = 1;

            vm.OpenResultCommand.Execute(item);

            Assert.Equal(0, repo.DeleteEmailCallCount);
            Assert.Equal(1, adapter.OpenEmailCallCount);
            Assert.Single(vm.SearchResults);
            Assert.Same(item, vm.SearchResults[0]);
            Assert.Equal(1, vm.TotalResultsCount);
            Assert.Empty(dialogs); // No error dialog on success
        }

        [Fact]
        public void Requirement7_SearchResultsRemainInUI_AfterOpeningFailure()
        {
            var (vm, repo, adapter, dialogs) = CreateTestContext();
            adapter.IsAvailableResult = true;
            adapter.OpenEmailResult = false;

            var item1 = CreateSampleItem("Email 1", "S1", "E1");
            var item2 = CreateSampleItem("Email 2", "S1", "E2");
            var item3 = CreateSampleItem("Email 3", "S1", "E3");

            vm.SearchResults.Add(item1);
            vm.SearchResults.Add(item2);
            vm.SearchResults.Add(item3);
            vm.TotalResultsCount = 3;

            // Attempt to open the middle item (fails)
            vm.OpenResultCommand.Execute(item2);

            Assert.Equal(3, vm.SearchResults.Count);
            Assert.Same(item1, vm.SearchResults[0]);
            Assert.Same(item2, vm.SearchResults[1]);
            Assert.Same(item3, vm.SearchResults[2]);
            Assert.Equal(0, repo.DeleteEmailCallCount);
        }

        [Fact]
        public void Requirement8_ResultCountRemainsUnchanged_AfterOpeningFailure()
        {
            var (vm, repo, adapter, dialogs) = CreateTestContext();
            adapter.IsAvailableResult = true;
            adapter.OpenEmailResult = false;

            var item = CreateSampleItem();
            vm.SearchResults.Add(item);
            vm.TotalResultsCount = 48173; // Simulated production result count

            vm.OpenResultCommand.Execute(item);

            Assert.Equal(48173, vm.TotalResultsCount);
            Assert.Equal(0, repo.DeleteEmailCallCount);
        }

        [Theory]
        [InlineData("OutlookUnavailable")]
        [InlineData("ComReturnedFalse")]
        [InlineData("ComException")]
        [InlineData("NullItem")]
        public void Requirement9_NoAutomaticDeleteEmailAsync_AcrossAllFailureModes(string mode)
        {
            var (vm, repo, adapter, dialogs) = CreateTestContext();
            var item = CreateSampleItem();
            vm.SearchResults.Add(item);
            vm.TotalResultsCount = 1;

            switch (mode)
            {
                case "OutlookUnavailable":
                    adapter.IsAvailableResult = false;
                    vm.OpenResultCommand.Execute(item);
                    break;
                case "ComReturnedFalse":
                    adapter.IsAvailableResult = true;
                    adapter.OpenEmailResult = false;
                    vm.OpenResultCommand.Execute(item);
                    break;
                case "ComException":
                    adapter.IsAvailableResult = true;
                    adapter.OpenEmailException = new InvalidOperationException("Simulated COM fault");
                    vm.OpenResultCommand.Execute(item);
                    break;
                case "NullItem":
                    vm.OpenResultCommand.Execute(null);
                    break;
            }

            Assert.Equal(0, repo.DeleteEmailCallCount);
        }

        [Fact]
        public void OutlookAdapter_PassiveModeCheck_WhenOutlookClosed_ReturnsFalseWithoutSpawningProcess()
        {
            // Verify real OutlookAdapter closed behavior
            var procs = Process.GetProcessesByName("OUTLOOK");
            if (procs.Length == 0)
            {
                using var adapter = new OutlookAdapter();
                bool available = adapter.IsOutlookAvailable();
                Assert.False(available);

                bool opened = adapter.OpenEmailInOutlook("dummy-store", "dummy-entry");
                Assert.False(opened);

                var procsAfter = Process.GetProcessesByName("OUTLOOK");
                Assert.Empty(procsAfter);
            }
        }
    }

    #region Test Mocks
    internal class TrackingMockRepository : IEmailRepository
    {
        public int DeleteEmailCallCount { get; private set; }
        public string? LastDeletedStoreId { get; private set; }
        public string? LastDeletedEntryId { get; private set; }

        public Task DeleteEmailAsync(string storeId, string entryId)
        {
            DeleteEmailCallCount++;
            LastDeletedStoreId = storeId;
            LastDeletedEntryId = entryId;
            return Task.CompletedTask;
        }

        public string DatabasePath => string.Empty;
        public void InitializeDatabase() { }
        public Task SaveEmailBatchAsync(IEnumerable<EmailDocument> emails, CancellationToken cancellationToken = default) => Task.CompletedTask;
        public Task<EmailDocument?> GetEmailByStoreAndEntryIdAsync(string storeId, string entryId) => Task.FromResult<EmailDocument?>(null);
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

    internal class ConfigurableMockOutlookAdapter : IOutlookAdapter
    {
        public bool IsAvailableResult { get; set; } = true;
        public bool OpenEmailResult { get; set; } = true;
        public Exception? OpenEmailException { get; set; }

        public int OpenEmailCallCount { get; private set; }
        public string? LastOpenedStoreId { get; private set; }
        public string? LastOpenedEntryId { get; private set; }

        public bool IsOutlookAvailable() => IsAvailableResult;
        public string GetOutlookVersion() => IsAvailableResult ? "16.0" : "Unavailable";
        public string GetOfficeBitness() => "x64";
        public List<MailboxInfo> GetMailboxes() => new List<MailboxInfo>();
        public List<MailboxFolder> GetFolders(string storeId) => new List<MailboxFolder>();
        public List<EmailDocument> FetchFolderEmailsBatch(string storeId, string folderEntryId, int startIndex, int batchSize, Func<bool>? cancelCheck = null) => new List<EmailDocument>();
        public EmailDocument? FetchSingleEmail(string storeId, string entryId) => null;

        public bool OpenEmailInOutlook(string storeId, string entryId)
        {
            OpenEmailCallCount++;
            LastOpenedStoreId = storeId;
            LastOpenedEntryId = entryId;

            if (OpenEmailException != null)
                throw OpenEmailException;

            return OpenEmailResult;
        }

        public bool IsOutboxFolder(string storeId, string folderEntryId) => false;
        public void SubscribeToEvents(Action<string, string, string> onItemAddedOrChanged, Action<string, string> onItemRemoved) { }
        public void UnsubscribeFromEvents() { }
    }
    #endregion
}
