using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using FastMailSearch.Core.Models;

namespace FastMailSearch.Core.Interfaces
{
    public enum IndexerStatus
    {
        Idle,
        Indexing,
        Paused,
        Cancelled,
        Error,
        SafeMode
    }

    public class IndexerProgressEventArgs : EventArgs
    {
        public int TotalItems { get; set; }
        public int ProcessedItems { get; set; }
        public int SkippedItems { get; set; }
        public string CurrentFolder { get; set; } = string.Empty;
        public double Percentage { get; set; }
        public string CurrentStatusMessage { get; set; } = string.Empty;
        public IndexerStatus Status { get; set; }
        public bool IsRebuilding { get; set; }
        public DateTime? StartedAt { get; set; }
        public DateTime? CompletedAt { get; set; }
        public DateTime? LastIndexedItemTime { get; set; }
        public string CurrentOperation { get; set; } = string.Empty;
        public string LastError { get; set; } = string.Empty;
        public int FolderProcessedItems { get; set; }
        public int FolderTotalItems { get; set; }

        public int TotalDiscovered { get; set; }
        public int ScannedCount { get; set; }
        public int SuccessfullyIndexed { get; set; }
        public int NewlyIndexedCount { get; set; }
        public int FailedCount { get; set; }
        public int RemainingCount { get; set; }
        public int PendingQueueCount { get; set; }
    }

    public interface IIndexerService
    {
        event EventHandler<IndexerProgressEventArgs>? ProgressChanged;
        IndexerStatus CurrentStatus { get; }
        IndexMetadata CurrentMetadata { get; }
        bool IsIndexing { get; }
        bool IsRebuilding { get; }
        int ProcessedItems { get; }
        int TotalItems { get; }
        int FolderProcessedItems { get; }
        int FolderTotalItems { get; }
        string CurrentFolder { get; }
        DateTime? StartedAt { get; }
        DateTime? CompletedAt { get; }
        DateTime? LastIndexedItemTime { get; }
        string CurrentOperation { get; }
        string LastError { get; }

        Task StartInitialIndexingAsync(CancellationToken cancellationToken = default, int? maxItems = null);
        Task<BackfillSummary> StartActiveStoreBackfillAsync(CancellationToken cancellationToken = default, int? maxItems = null);
        Task<StoreCoverageInfo> CheckStoreCoverageAsync(CancellationToken cancellationToken = default);
        Task PauseIndexingAsync();
        Task ResumeIndexingAsync();
        Task<BackfillSummary> ResumeIncrementalCatchUpAsync(CancellationToken cancellationToken = default);
        Task RebuildIndexAsync(CancellationToken cancellationToken = default);
        Task CancelIndexingAsync();
        Task CheckAndRecoverIncompleteIndexAsync(bool startProcessing = true);
        Task QueueItemForIndexingAsync(string storeId, string entryId, string folderEntryId);
        Task QueueItemForDeletionAsync(string storeId, string entryId);
        Task<IndexIntegrityResult> ScanIndexIntegrityAsync(CancellationToken cancellationToken = default);
        Task RepairIndexIntegrityAsync(IndexIntegrityResult scanResult, CancellationToken cancellationToken = default);
        IndexerProgressEventArgs GetStatus();
        void SetSafeMode(string reason);
    }

    public interface IOutlookAdapter
    {
        bool IsOutlookAvailable();
        string GetOutlookVersion();
        string GetOfficeBitness();
        List<MailboxInfo> GetMailboxes();
        List<MailboxFolder> GetFolders(string storeId);
        List<EmailDocument> FetchFolderEmailsBatch(string storeId, string folderEntryId, int startIndex, int batchSize, Func<bool>? cancelCheck = null);
        EmailDocument? FetchSingleEmail(string storeId, string entryId);
        bool OpenEmailInOutlook(string storeId, string entryId);
        bool IsOutboxFolder(string storeId, string folderEntryId);
        void SubscribeToEvents(Action<string, string, string> onItemAddedOrChanged, Action<string, string> onItemRemoved);
        void UnsubscribeFromEvents();
    }

    public interface IFolderBatchReader
    {
        FolderBatchResult FetchFolderBatch(string storeId, string folderEntryId, int startIndex, int batchSize, Func<bool>? cancelCheck = null);
        void ResetCursorSession();
    }
}
