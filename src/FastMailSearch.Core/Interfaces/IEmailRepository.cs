using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using FastMailSearch.Core.Models;

namespace FastMailSearch.Core.Interfaces
{
    public interface IEmailRepository
    {
        string DatabasePath { get; }
        void InitializeDatabase();
        Task SaveEmailBatchAsync(IEnumerable<EmailDocument> emails, CancellationToken cancellationToken = default);
        Task<EmailDocument?> GetEmailByStoreAndEntryIdAsync(string storeId, string entryId);
        Task DeleteEmailAsync(string storeId, string entryId);
        Task UpdateEmailFolderAsync(string storeId, string entryId, string newFolderName, string newFolderEntryId);
        Task<HashSet<string>> GetIndexedEntryIdsForFolderAsync(string storeId, string folderEntryId);
        Task SaveFoldersAsync(IEnumerable<MailboxFolder> folders);
        Task<List<MailboxFolder>> GetFoldersAsync(string? storeId = null);
        Task<IndexMetadata> GetIndexMetadataAsync();
        Task ClearAllDataAsync();
        Task SaveSearchQueryAsync(string rawQuery);
        Task<List<string>> GetRecentSearchQueriesAsync(int count = 20);
        Task ClearSearchHistoryAsync();
        Task VacuumDatabaseAsync();
        Task SaveIndexStateAsync(string state, int processed, int total, string currentFolder, string? error);
        Task<Dictionary<string, string>> GetPersistedIndexStateAsync();
        Task<int> GetIndexedEmailCountAsync();

        // Phase 32 & 37: Store & Folder Coverage, Resumable Checkpoints and Live Progress
        Task<Dictionary<string, int>> GetStoreEmailCountsAsync();
        Task<int> GetFolderEmailCountAsync(string storeId, string folderEntryId);
        Task SaveFolderIndexStateAsync(string storeId, string folderEntryId, string status, int processedOffset, int totalItems);
        Task<(string Status, int ProcessedOffset, int TotalItems)?> GetFolderIndexStateAsync(string storeId, string folderEntryId);
        Task<List<FolderIndexCheckpoint>> GetAllFolderIndexStatesAsync();
        Task SaveLiveProgressAsync(string state, string currentOperation, int processed, int total, string currentFolder, int folderOffset, int folderTotal, int newlyIndexed, int skipped, int failed, int pendingQueue);

        // Phase 20: Persistent Queue, Failed Item Tracking & Self-Healing
        Task EnqueuePendingChangeAsync(string storeId, string entryId, string operation, string? folderEntryId);
        Task<List<PendingChangeItem>> GetPendingChangesAsync(int maxCount = 50);
        Task<int> GetPendingChangesCountAsync();
        Task<int> IncrementAttemptCountAsync(long id);
        Task MarkPendingChangeStatusAsync(long id, string status, string? error = null);
        Task MarkPendingChangesCompletedBatchAsync(IEnumerable<long> ids);
        Task RecoverStaleProcessingChangesAsync();
        Task SaveFailedIndexItemAsync(string storeId, string entryId, string operation, int attemptCount, string lastError);
        Task SaveFailedIndexItemAsync(FailedIndexItem item);
        Task<List<FailedIndexItem>> GetFailedIndexItemsAsync(bool includeResolved = false);
        Task<List<FailedIndexItem>> GetFailedIndexItemsAsync(string? status, int maxCount = 500);
        Task<int> GetFailedIndexItemsCountAsync(string? status = "Failed");
        Task UpdateFailedIndexItemStatusAsync(long id, string status, string? error = null);
        Task RemoveFailedIndexItemAsync(long id);
        Task RemoveAllFailedIndexItemsAsync();
        Task<double> GetDatabaseSizeMbAsync();
        Task<IndexIntegrityResult> ScanIndexIntegrityAsync(CancellationToken cancellationToken = default);
        Task RepairIndexIntegrityAsync(IndexIntegrityResult scanResult, CancellationToken cancellationToken = default);
    }

    public interface ISearchEngine
    {
        Task<SearchResultContainer> SearchAsync(SearchQuery query, CancellationToken cancellationToken = default);
    }
}
