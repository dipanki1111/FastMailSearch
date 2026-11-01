using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using FastMailSearch.Core.Interfaces;
using FastMailSearch.Core.Models;
using FastMailSearch.Infrastructure.Logging;

namespace FastMailSearch.Infrastructure.Indexing
{
    public enum IncrementalOperation
    {
        InsertOrUpdate,
        Delete
    }

    public class IncrementalItem
    {
        public long PendingId { get; set; }
        public string StoreId { get; set; } = string.Empty;
        public string EntryId { get; set; } = string.Empty;
        public string FolderEntryId { get; set; } = string.Empty;
        public IncrementalOperation Operation { get; set; }
        public int AttemptCount { get; set; }
        public string? LastError { get; set; }
        public DateTime QueuedTime { get; set; } = DateTime.Now;
    }

    public class EmailIndexerService : IIndexerService
    {
        private readonly IOutlookAdapter _outlookAdapter;
        private readonly IEmailRepository _repository;
        private readonly ILicensingService _licensingService;
        private bool _wasIneligible;
        private readonly object _catchUpLock = new object();
        private bool _isCatchUpRunning;

        private readonly ConcurrentQueue<IncrementalItem> _incrementalQueue = new ConcurrentQueue<IncrementalItem>();
        private readonly ConcurrentDictionary<string, IncrementalItem> _pendingItems = new ConcurrentDictionary<string, IncrementalItem>(StringComparer.OrdinalIgnoreCase);
        private readonly ConcurrentQueue<IncrementalItem> _failedQueue = new ConcurrentQueue<IncrementalItem>();

        private readonly SemaphoreSlim _indexingLock = new SemaphoreSlim(1, 1);
        private readonly SemaphoreSlim _incrementalProcessingLock = new SemaphoreSlim(1, 1);

        private IndexerStatus _status = IndexerStatus.Idle;
        private readonly object _statusLock = new object();
        private CancellationTokenSource? _indexerCts;
        private CancellationTokenSource? _incrementalCts;
        private bool _isPaused;

        private int _totalItems;
        private int _processedItems;
        private int _skippedItems;
        private string _currentFolder = string.Empty;
        private int _folderProcessedItems;
        private int _folderTotalItems;
        private bool _isRebuilding;
        private DateTime? _startedAt;
        private DateTime? _completedAt;
        private DateTime? _lastIndexedItemTime;
        private string _currentOperation = "Idle";
        private string _lastError = string.Empty;
        private int _pendingDbCount = 0;
        private bool _hasRecovered = false;
        private Func<bool>? _testCancelCheck;
        public void SetBatchCancelCheck(Func<bool>? check) => _testCancelCheck = check;

        public event EventHandler<IndexerProgressEventArgs>? ProgressChanged;

        [Obsolete("Use EmailIndexerService(IOutlookAdapter, IEmailRepository, ILicensingService) instead. A licensing service is required for indexing execution.", false)]
        public EmailIndexerService(IOutlookAdapter outlookAdapter, IEmailRepository repository)
            : this(outlookAdapter, repository, null!, false)
        {
        }

        [Obsolete("Use EmailIndexerService(IOutlookAdapter, IEmailRepository, ILicensingService, bool) instead. A licensing service is required for indexing execution.", false)]
        public EmailIndexerService(IOutlookAdapter outlookAdapter, IEmailRepository repository, bool autoSubscribeEvents)
            : this(outlookAdapter, repository, null!, autoSubscribeEvents)
        {
        }

        public EmailIndexerService(IOutlookAdapter outlookAdapter, IEmailRepository repository, ILicensingService licensingService)
            : this(outlookAdapter, repository, licensingService, false)
        {
        }

        public EmailIndexerService(IOutlookAdapter outlookAdapter, IEmailRepository repository, ILicensingService licensingService, bool autoSubscribeEvents)
        {
            _outlookAdapter = outlookAdapter ?? throw new ArgumentNullException(nameof(outlookAdapter));
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
            _licensingService = licensingService;

            if (autoSubscribeEvents)
            {
                SubscribeToOutlookEvents();
            }
        }

        public bool IsLicensingEligible()
        {
            if (_licensingService == null)
            {
                _wasIneligible = true;
                return false;
            }

            try
            {
                var status = _licensingService.GetStatus();
                bool eligible = (status == LicenseStatus.Valid || status == LicenseStatus.Trial);
                if (eligible)
                {
                    if (_wasIneligible)
                    {
                        _wasIneligible = false;
                        Logger.Info($"[LICENSING] Indexing license transitioned to eligible ({status}). Scheduling incremental catch-up.");
                        _ = Task.Run(async () => await ResumeIncrementalCatchUpAsync());
                    }
                }
                else
                {
                    _wasIneligible = true;
                }
                return eligible;
            }
            catch (Exception ex)
            {
                Logger.Warn($"[LICENSING] Exception evaluating license eligibility: {ex.Message}");
                _wasIneligible = true;
                return false;
            }
        }

        public async Task<BackfillSummary> ResumeIncrementalCatchUpAsync(CancellationToken cancellationToken = default)
        {
            lock (_catchUpLock)
            {
                if (_isCatchUpRunning)
                {
                    Logger.Info("[CATCHUP] Incremental catch-up already running. Ignoring duplicate request.");
                    return new BackfillSummary();
                }
                _isCatchUpRunning = true;
            }

            try
            {
                if (!IsLicensingEligible())
                {
                    Logger.Info("[CATCHUP] Cannot run catch-up: licensing is ineligible.");
                    return new BackfillSummary { ErrorMessage = "Licensing is ineligible." };
                }

                Logger.Info("[CATCHUP] Resuming incremental catch-up after license eligibility restored.");

                // 1. Drain pending changes queue accumulated during pause
                await TriggerIncrementalProcessingAsync();

                // 2. Resume / discover any missed changes across mail folders
                var summary = await StartActiveStoreBackfillAsync(cancellationToken);
                Logger.Info($"[CATCHUP] Incremental catch-up complete. NewlyIndexed={summary.NewlyIndexedCount}, Skipped={summary.SkippedCount}, Failed={summary.FailedCount}");
                return summary;
            }
            catch (Exception ex)
            {
                Logger.Error("[CATCHUP] Error during incremental catch-up", ex);
                return new BackfillSummary { ErrorMessage = ex.Message };
            }
            finally
            {
                lock (_catchUpLock)
                {
                    _isCatchUpRunning = false;
                }
            }
        }

        public void SubscribeToOutlookEvents()
        {
            try
            {
                _outlookAdapter.SubscribeToEvents(
                    onItemAddedOrChanged: (storeId, entryId, folderEntryId) =>
                    {
                        _ = QueueItemForIndexingAsync(storeId, entryId, folderEntryId);
                    },
                    onItemRemoved: (storeId, entryId) =>
                    {
                        _ = QueueItemForDeletionAsync(storeId, entryId);
                    }
                );
            }
            catch (Exception ex)
            {
                Logger.Error("[INCREMENTAL] Error subscribing indexer service to Outlook events", ex);
            }
        }

        public IndexerStatus CurrentStatus
        {
            get { lock (_statusLock) return _status; }
            private set { lock (_statusLock) _status = value; }
        }

        public IndexMetadata CurrentMetadata { get; private set; } = new IndexMetadata();

        public bool IsIndexing => CurrentStatus == IndexerStatus.Indexing;
        public bool IsRebuilding => _isRebuilding;
        public int ProcessedItems => _processedItems;
        public int TotalItems => _totalItems;
        public int FolderProcessedItems => _folderProcessedItems;
        public int FolderTotalItems => _folderTotalItems;
        public string CurrentFolder => _currentFolder;
        public DateTime? StartedAt => _startedAt;
        public DateTime? CompletedAt => _completedAt;
        public DateTime? LastIndexedItemTime => _lastIndexedItemTime;
        public string CurrentOperation => _currentOperation;
        public string LastError => _lastError;

        public void SetSafeMode(string reason)
        {
            lock (_statusLock)
            {
                _status = IndexerStatus.SafeMode;
                _currentOperation = "Safe Mode";
                _lastError = reason;
            }
            Logger.Warn($"[SAFE MODE ACTIVATED] Reason: {reason}");
        }

        public IndexerProgressEventArgs GetStatus()
        {
            lock (_statusLock)
            {
                int remainingQueue = Math.Max(_incrementalQueue.Count, _pendingDbCount);
                double pct = _totalItems > 0 ? Math.Min(100.0, Math.Round((double)_processedItems / _totalItems * 100.0, 2)) : 0.0;
                string msg;
                if (remainingQueue > 0 && _status != IndexerStatus.Error && _status != IndexerStatus.Cancelled && _status != IndexerStatus.Paused)
                {
                    msg = $"🔵 Updating • {remainingQueue:N0} email(s) remaining";
                }
                else if (_status == IndexerStatus.Indexing)
                {
                    msg = _totalItems > 0
                        ? $"Indexing: {_processedItems:N0} / {_totalItems:N0} (Folder: {_currentFolder})"
                        : "Indexing in progress...";
                }
                else
                {
                    msg = CurrentMetadata.CurrentStatusDetail ?? _status.ToString();
                }

                return new IndexerProgressEventArgs
                {
                    TotalItems = _totalItems,
                    ProcessedItems = _processedItems,
                    SkippedItems = _skippedItems,
                    CurrentFolder = _currentFolder,
                    Percentage = pct,
                    CurrentStatusMessage = msg,
                    Status = _status,
                    IsRebuilding = _isRebuilding,
                    StartedAt = _startedAt,
                    CompletedAt = _completedAt,
                    LastIndexedItemTime = _lastIndexedItemTime,
                    CurrentOperation = _currentOperation,
                    LastError = _lastError,
                    FolderProcessedItems = _folderProcessedItems,
                    FolderTotalItems = _folderTotalItems,
                    TotalDiscovered = _totalItems,
                    ScannedCount = _processedItems,
                    SuccessfullyIndexed = _processedItems,
                    NewlyIndexedCount = _processedItems,
                    FailedCount = 0,
                    RemainingCount = Math.Max(0, _totalItems - _processedItems),
                    PendingQueueCount = remainingQueue
                };
            }
        }

        public async Task CheckAndRecoverIncompleteIndexAsync(bool startProcessing = true)
        {
            try
            {
                var stateDict = await _repository.GetPersistedIndexStateAsync();
                if (stateDict.TryGetValue("IndexState", out string? stateStr) && (stateStr == "Indexing" || stateStr == "Updating"))
                {
                    stateDict.TryGetValue("LastProcessedCount", out string? procStr);
                    stateDict.TryGetValue("LastTotalCount", out string? totStr);
                    stateDict.TryGetValue("LastIndexedFolder", out string? folderStr);

                    Logger.Info("[INDEX] Incomplete indexing operation detected from previous session.");
                    Logger.Info($"[INDEX] Previous scan was interrupted at {procStr}/{totStr} in folder '{folderStr}'. Recovering index state...");
                    await _repository.SaveIndexStateAsync("Incomplete", 0, 0, folderStr ?? "", "Interrupted on shutdown - recovered cleanly");
                }

                // Phase 20: Persistent Change Queue Startup Recovery
                if (!_hasRecovered)
                {
                    _hasRecovered = true;
                    int indexedCount = await _repository.GetIndexedEmailCountAsync();
                    if (_processedItems == 0)
                    {
                        _processedItems = indexedCount;
                        _totalItems = indexedCount;
                    }
                    Logger.Info($"[INDEX] Startup check completed. Current DB indexed count: {indexedCount:N0} emails (Target DB: '{_repository.DatabasePath}').");
                    await _repository.RecoverStaleProcessingChangesAsync();
                    if (startProcessing && IsLicensingEligible())
                    {
                        _ = TriggerIncrementalProcessingAsync();
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("[INDEX] Error checking for incomplete index on startup", ex);
            }
        }

        public Task StartInitialIndexingAsync(CancellationToken cancellationToken = default)
        {
            return StartScanInternalAsync(isRebuild: false, cancellationToken);
        }

        public async Task<StoreCoverageInfo> CheckStoreCoverageAsync(CancellationToken cancellationToken = default)
        {
            var info = new StoreCoverageInfo();
            try
            {
                if (!_outlookAdapter.IsOutlookAvailable())
                {
                    info.SummaryText = "Outlook is unavailable";
                    return info;
                }

                Logger.Info($"[COVERAGE] Checking Outlook store coverage. Target DB: '{_repository.DatabasePath}'");
                var mailboxes = TestModeStoreFilter.FilterMailboxes(_outlookAdapter.GetMailboxes());
                info.ActiveStoreCount = mailboxes.Count;

                if (mailboxes.Count == 0)
                {
                    info.SummaryText = "No active Outlook stores discovered. Please ensure Classic Outlook has an active account profile.";
                    Logger.Warn("[COVERAGE] No active Outlook stores discovered.");
                    return info;
                }

                var storeCounts = await _repository.GetStoreEmailCountsAsync();

                int totalDiscovered = 0;
                int totalUnindexed = 0;
                int unindexedStores = 0;
                int unindexedFolders = 0;

                foreach (var mb in mailboxes)
                {
                    bool hasStoreRecords = storeCounts.TryGetValue(mb.StoreId, out int indexedForStore) && indexedForStore > 0;
                    if (!hasStoreRecords)
                    {
                        unindexedStores++;
                    }

                    Logger.Info($"[COVERAGE] Store: '{mb.DisplayName}' (IsPst={mb.IsPst}, FilePath='{mb.FilePath}', IndexedInDb={indexedForStore:N0}, StoreId='{mb.StoreId}')");

                    var folders = _outlookAdapter.GetFolders(mb.StoreId);
                    foreach (var folder in folders)
                    {
                        if (_outlookAdapter.IsOutboxFolder(mb.StoreId, folder.EntryId) || folder.Name.Equals("Outbox", StringComparison.OrdinalIgnoreCase))
                            continue;

                        totalDiscovered += folder.ItemCount;
                        int inDb = await _repository.GetFolderEmailCountAsync(mb.StoreId, folder.EntryId);
                        Logger.Info($"[COVERAGE] Folder: '{folder.Name}', Store='{mb.DisplayName}', ItemsInOutlook={folder.ItemCount:N0}, IndexedInDb={inDb:N0}");

                        if (inDb < folder.ItemCount)
                        {
                            unindexedFolders++;
                            totalUnindexed += (folder.ItemCount - inDb);
                        }
                    }
                }

                info.TotalDiscoveredItems = totalDiscovered;
                info.TotalIndexedItems = await _repository.GetIndexedEmailCountAsync();
                info.UnindexedItems = totalUnindexed;
                info.UnindexedStoreCount = unindexedStores;
                info.UnindexedFolderCount = unindexedFolders;

                Logger.Info($"[COVERAGE] Completed: TotalDiscovered={info.TotalDiscoveredItems:N0}, TotalIndexedInDb={info.TotalIndexedItems:N0}, UnindexedItems={info.UnindexedItems:N0}, UnindexedStores={info.UnindexedStoreCount}, UnindexedFolders={info.UnindexedFolderCount}");

                if (info.ActiveStoreCount == 0)
                {
                    info.SummaryText = "No active Outlook stores discovered.";
                }
                else if (info.TotalDiscoveredItems == 0)
                {
                    info.SummaryText = $"Active store '{mailboxes.FirstOrDefault()?.DisplayName}' has 0 emails in mail folders.";
                }
                else if (info.UnindexedStoreCount > 0)
                {
                    info.SummaryText = $"Active Outlook store '{mailboxes.FirstOrDefault()?.DisplayName}' has unindexed emails ({info.UnindexedItems:N0} remaining of {info.TotalDiscoveredItems:N0} discovered). Backfill recommended.";
                }
                else if (info.UnindexedFolderCount > 0)
                {
                    info.SummaryText = $"{info.UnindexedFolderCount} folder(s) have unindexed emails ({info.UnindexedItems:N0} remaining). Backfill recommended.";
                }
                else
                {
                    info.SummaryText = "All active Outlook stores and folders are fully indexed.";
                }
            }
            catch (Exception ex)
            {
                Logger.Warn($"[COVERAGE] Error checking store coverage: {ex.Message}");
                info.SummaryText = $"Coverage check error: {ex.Message}";
            }
            return info;
        }

        public async Task<BackfillSummary> StartActiveStoreBackfillAsync(CancellationToken cancellationToken = default)
        {
            var summary = new BackfillSummary();

            if (!IsLicensingEligible())
            {
                Logger.Info("[BACKFILL] Licensing is ineligible. Backfill cannot start.");
                summary.ErrorMessage = "Indexing is paused. A valid license or active trial is required to backfill emails.";
                return summary;
            }

            if (!_indexingLock.Wait(0))
            {
                Logger.Info("[BACKFILL] Indexing or backfill already in progress. Ignoring duplicate request.");
                summary.ErrorMessage = "Indexing operation already in progress.";
                return summary;
            }

            try
            {
                await _incrementalProcessingLock.WaitAsync(cancellationToken);
                try
                {
                    _indexerCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
                    CurrentStatus = IndexerStatus.Indexing;
                    _startedAt = DateTime.Now;
                    _completedAt = null;
                    _lastError = string.Empty;
                    _processedItems = 0;
                    _skippedItems = 0;
                    _folderProcessedItems = 0;
                    _folderTotalItems = 0;
                    _currentOperation = "Backfilling Active Store";

                    Logger.Info($"[BACKFILL] Active store backfill started. Target DB: '{_repository.DatabasePath}'. Discovering Outlook mailstores...");
                    await _repository.SaveLiveProgressAsync("Indexing", _currentOperation, 0, 0, "", 0, 0, 0, 0, 0, _pendingDbCount);
                    ReportProgress(0, 0, 0, "", "Discovering active Outlook mailstores...", IndexerStatus.Indexing);

                    if (!_outlookAdapter.IsOutlookAvailable())
                    {
                        Logger.Warn("[BACKFILL] Outlook Classic is unavailable. Backfill aborted.");
                        CurrentStatus = IndexerStatus.Idle;
                        _currentOperation = "Idle";
                        summary.ErrorMessage = "Outlook is unavailable";
                        await _repository.SaveLiveProgressAsync("Idle", "Idle", 0, 0, "", 0, 0, 0, 0, 0, _pendingDbCount);
                        ReportProgress(0, 0, 0, "", "Outlook is unavailable", IndexerStatus.Error);
                        return summary;
                    }

                    var mailboxes = TestModeStoreFilter.FilterMailboxes(_outlookAdapter.GetMailboxes());
                    if (mailboxes.Count == 0)
                    {
                        Logger.Warn("[BACKFILL] No active Outlook stores discovered. Backfill aborted.");
                        CurrentStatus = IndexerStatus.Error;
                        _lastError = "No active Outlook store discovered. Please verify Classic Outlook account setup.";
                        _currentOperation = "Error: No active store";
                        summary.ErrorMessage = _lastError;
                        await _repository.SaveLiveProgressAsync("Error", _currentOperation, 0, 0, "", 0, 0, 0, 0, 0, _pendingDbCount);
                        ReportProgress(0, 0, 0, "", _lastError, IndexerStatus.Error);
                        return summary;
                    }

                    var folderTasks = new List<(string storeId, MailboxFolder folder)>();
                    int totalItemsInAllFolders = 0;

                    foreach (var mb in mailboxes)
                    {
                        Logger.Info($"[BACKFILL] Store: DisplayName='{mb.DisplayName}', FilePath='{mb.FilePath}', IsPst={mb.IsPst}, StoreId='{mb.StoreId}'");
                        var folders = _outlookAdapter.GetFolders(mb.StoreId);
                        await _repository.SaveFoldersAsync(folders);

                        foreach (var folder in folders)
                        {
                            if (_outlookAdapter.IsOutboxFolder(mb.StoreId, folder.EntryId) || folder.Name.Equals("Outbox", StringComparison.OrdinalIgnoreCase))
                                continue;

                            totalItemsInAllFolders += folder.ItemCount;
                            folderTasks.Add((mb.StoreId, folder));
                            Logger.Info($"[BACKFILL] Store '{mb.DisplayName}' Folder: '{folder.Name}', Items={folder.ItemCount:N0}");
                        }
                    }

                    _totalItems = totalItemsInAllFolders;
                    summary.TotalDiscovered = _totalItems;
                    Logger.Info($"[BACKFILL] Discovered {_totalItems:N0} total items across {folderTasks.Count} folders in {mailboxes.Count} active store(s).");

                    if (_totalItems == 0)
                    {
                        Logger.Info("[BACKFILL] Active store has 0 emails in mail folders. Backfill completed with 0 items.");
                        CurrentStatus = IndexerStatus.Idle;
                        _currentOperation = "Idle";
                        await _repository.SaveLiveProgressAsync("Completed", "Idle", 0, 0, "", 0, 0, 0, 0, 0, _pendingDbCount);
                        await _repository.SaveIndexStateAsync("Completed", 0, 0, "", null);
                        ReportProgress(0, 0, 0, "", "Active mailbox has 0 emails in mail folders. Up to date.", IndexerStatus.Idle);
                        return summary;
                    }

                    _currentOperation = $"Preparing backfill ({_totalItems:N0} total items in active mailbox)...";
                    await _repository.SaveLiveProgressAsync("Indexing", _currentOperation, 0, _totalItems, "", 0, 0, 0, 0, 0, _pendingDbCount);
                    ReportProgress(_totalItems, 0, 0, "", _currentOperation, IndexerStatus.Indexing);

                    int completedPreviousFoldersCount = 0;
                    int totalIndexed = 0;
                    int totalSkipped = 0;
                    int totalFailed = 0;
                    int batchSize = 250;

                    foreach (var (storeId, folder) in folderTasks)
                    {
                        if (_indexerCts.Token.IsCancellationRequested)
                        {
                            summary.IsCancelled = true;
                            break;
                        }

                        _currentFolder = folder.Name;
                        int folderItemCount = folder.ItemCount;
                        if (folderItemCount == 0)
                        {
                            await _repository.SaveFolderIndexStateAsync(storeId, folder.EntryId, "Completed", 0, 0);
                            continue;
                        }

                        // Check per-folder resumable state
                        var folderState = await _repository.GetFolderIndexStateAsync(storeId, folder.EntryId);
                        if (folderState.HasValue && folderState.Value.Status == "Completed" && folderState.Value.TotalItems >= folderItemCount && folderState.Value.ProcessedOffset >= folderItemCount)
                        {
                            Logger.Info($"[BACKFILL] Folder '{folder.Name}' is already fully indexed ({folderItemCount} items). Skipping.");
                            completedPreviousFoldersCount += folderItemCount;
                            _processedItems = completedPreviousFoldersCount;
                            ReportProgress(
                                _totalItems,
                                _processedItems,
                                _skippedItems,
                                folder.Name,
                                $"Folder '{folder.Name}' already indexed ({folderItemCount:N0} emails).",
                                IndexerStatus.Indexing,
                                folderProcessed: folderItemCount,
                                folderTotal: folderItemCount,
                                newlyIndexed: totalIndexed,
                                failed: totalFailed
                            );
                            continue;
                        }

                        int startFrom = 1;
                        int folderProcessed = 0;
                        if (folderState.HasValue && folderState.Value.ProcessedOffset > 0 && folderState.Value.ProcessedOffset < folderItemCount)
                        {
                            startFrom = folderState.Value.ProcessedOffset + 1;
                            folderProcessed = folderState.Value.ProcessedOffset;
                            Logger.Info($"[BACKFILL] Resuming folder '{folder.Name}' from item offset {startFrom}/{folderItemCount}.");
                        }
                        else
                        {
                            Logger.Info($"[BACKFILL] Starting folder '{folder.Name}' ({folderItemCount} items).");
                        }

                        _processedItems = completedPreviousFoldersCount + folderProcessed;
                        _currentOperation = $"Backfilling: {folder.Name} ({folderProcessed:N0}/{folderItemCount:N0})";
                        await _repository.SaveLiveProgressAsync("Indexing", _currentOperation, _processedItems, _totalItems, folder.Name, folderProcessed, folderItemCount, totalIndexed, totalSkipped, totalFailed, _pendingDbCount);
                        ReportProgress(
                            _totalItems,
                            _processedItems,
                            _skippedItems,
                            folder.Name,
                            $"Indexing: {_processedItems:N0} / {_totalItems:N0} (Folder: {folder.Name}: {folderProcessed:N0}/{folderItemCount:N0})",
                            IndexerStatus.Indexing,
                            folderProcessed: folderProcessed,
                            folderTotal: folderItemCount,
                            newlyIndexed: totalIndexed,
                            failed: totalFailed
                        );

                        await _repository.SaveFolderIndexStateAsync(storeId, folder.EntryId, "Indexing", startFrom - 1, folderItemCount);

                        int currentStart = startFrom;
                        while (currentStart <= folderItemCount)
                        {
                            if (!IsLicensingEligible())
                            {
                                _isPaused = true;
                                CurrentStatus = IndexerStatus.Paused;
                                _currentOperation = "Backfill paused: License required";
                                summary.ErrorMessage = "Backfill paused: License required.";
                                summary.IsCancelled = true;
                                (_outlookAdapter as IFolderBatchReader)?.ResetCursorSession();
                                Logger.Warn("[BACKFILL] Pausing backfill at safe checkpoint: License is ineligible.");
                                break;
                            }

                            if (_indexerCts.Token.IsCancellationRequested)
                            {
                                summary.IsCancelled = true;
                                (_outlookAdapter as IFolderBatchReader)?.ResetCursorSession();
                                break;
                            }

                            while (_isPaused)
                            {
                                _currentOperation = "Backfill paused by user";
                                await _repository.SaveLiveProgressAsync("Paused", _currentOperation, _processedItems, _totalItems, folder.Name, folderProcessed, folderItemCount, totalIndexed, totalSkipped, totalFailed, _pendingDbCount);
                                ReportProgress(_totalItems, _processedItems, _skippedItems, folder.Name, "Backfill paused by user", IndexerStatus.Paused, folderProcessed, folderItemCount, totalIndexed, totalFailed);
                                await Task.Delay(500, _indexerCts.Token);
                            }

                            FolderBatchResult batchResult = FolderBatchResult.Empty(currentStart);
                            int retryCount = 0;
                            bool batchSuccess = false;

                            while (retryCount < 3 && !batchSuccess && !_indexerCts.Token.IsCancellationRequested)
                            {
                                try
                                {
                                    if (_outlookAdapter is IFolderBatchReader batchReader)
                                    {
                                        batchResult = batchReader.FetchFolderBatch(
                                            storeId,
                                            folder.EntryId,
                                            currentStart,
                                            batchSize,
                                            () => _indexerCts.Token.IsCancellationRequested || _isPaused || (_testCancelCheck != null && _testCancelCheck())
                                        );
                                    }
                                    else
                                    {
                                        var legacyBatch = _outlookAdapter.FetchFolderEmailsBatch(
                                            storeId,
                                            folder.EntryId,
                                            currentStart,
                                            batchSize,
                                            () => _indexerCts.Token.IsCancellationRequested || _isPaused || (_testCancelCheck != null && _testCancelCheck())
                                        );
                                        bool isCanc = _indexerCts.Token.IsCancellationRequested || _isPaused;
                                        int lastIdx = isCanc
                                            ? (currentStart + legacyBatch.Count - 1)
                                            : Math.Min(currentStart + batchSize - 1, folderItemCount);

                                        batchResult = new FolderBatchResult
                                        {
                                            Items = legacyBatch,
                                            StartIndex = currentStart,
                                            LastProcessedIndex = Math.Max(currentStart - 1, lastIdx),
                                            ProcessedCount = Math.Max(0, lastIdx - currentStart + 1),
                                            IndexedCount = legacyBatch.Count,
                                            WasCancelled = isCanc
                                        };
                                    }
                                    batchSuccess = true;
                                }
                                catch (Exception ex)
                                {
                                    retryCount++;
                                    Logger.Warn($"[BACKFILL RETRY] Error fetching batch ({currentStart}..{currentStart + batchSize - 1}) in '{folder.Name}', attempt {retryCount}/3: {ex.Message}");
                                    if (retryCount < 3)
                                    {
                                        await Task.Delay(500 * retryCount, _indexerCts.Token);
                                    }
                                }
                            }

                            if (!batchSuccess)
                            {
                                Logger.Error($"[BACKFILL ERROR] Failed fetching batch ({currentStart}..{currentStart + batchSize - 1}) in folder '{folder.Name}'. Checkpoint preserved at offset {folderProcessed}.");
                                (_outlookAdapter as IFolderBatchReader)?.ResetCursorSession();
                                break;
                            }

                            bool persistenceSuccess = true;
                            if (batchResult.Items.Count > 0)
                            {
                                int saveRetry = 0;
                                persistenceSuccess = false;
                                while (saveRetry < 3 && !persistenceSuccess && !_indexerCts.Token.IsCancellationRequested)
                                {
                                    try
                                    {
                                        await _repository.SaveEmailBatchAsync(batchResult.Items, _indexerCts.Token);
                                        totalIndexed += batchResult.Items.Count;
                                        _lastIndexedItemTime = DateTime.Now;
                                        persistenceSuccess = true;
                                    }
                                    catch (OperationCanceledException)
                                    {
                                        summary.IsCancelled = true;
                                        (_outlookAdapter as IFolderBatchReader)?.ResetCursorSession();
                                        break;
                                    }
                                    catch (Exception ex)
                                    {
                                        saveRetry++;
                                        Logger.Warn($"[BACKFILL RETRY] Error saving email batch in folder '{folder.Name}', attempt {saveRetry}/3: {ex.Message}");
                                        if (saveRetry < 3)
                                        {
                                            await Task.Delay(500 * saveRetry, _indexerCts.Token);
                                        }
                                    }
                                }

                                if (!persistenceSuccess && !_indexerCts.Token.IsCancellationRequested)
                                {
                                    Logger.Error($"[BACKFILL ERROR] Exhausted retries saving email batch ({currentStart}..{batchResult.LastProcessedIndex}) in folder '{folder.Name}'. Checkpoint preserved at offset {folderProcessed}.");
                                    totalFailed += batchResult.Items.Count;
                                    summary.ErrorMessage = $"Failed saving email batch in folder '{folder.Name}'.";
                                    (_outlookAdapter as IFolderBatchReader)?.ResetCursorSession();
                                    break;
                                }
                            }

                            if (summary.IsCancelled || _indexerCts.Token.IsCancellationRequested)
                            {
                                summary.IsCancelled = true;
                                (_outlookAdapter as IFolderBatchReader)?.ResetCursorSession();
                                break;
                            }

                            totalSkipped += batchResult.SkippedCount;

                            // Checkpoint per-folder progress strictly to the last folder position accounted for
                            int currentOffset = Math.Max(folderProcessed, batchResult.LastProcessedIndex);
                            folderProcessed = currentOffset;
                            _processedItems = completedPreviousFoldersCount + folderProcessed;
                            _skippedItems = totalSkipped;

                            bool isFolderDone = currentOffset >= folderItemCount;
                            await _repository.SaveFolderIndexStateAsync(
                                storeId,
                                folder.EntryId,
                                isFolderDone ? "Completed" : "Indexing",
                                currentOffset,
                                folderItemCount
                            );

                            _currentOperation = $"Backfilling: {folder.Name} ({folderProcessed:N0}/{folderItemCount:N0})";
                            await _repository.SaveLiveProgressAsync("Indexing", _currentOperation, _processedItems, _totalItems, folder.Name, folderProcessed, folderItemCount, totalIndexed, totalSkipped, totalFailed, _pendingDbCount);

                            ReportProgress(
                                _totalItems,
                                _processedItems,
                                _skippedItems,
                                folder.Name,
                                $"Indexing: {_processedItems:N0} / {_totalItems:N0} (Folder: {folder.Name}: {folderProcessed:N0}/{folderItemCount:N0})",
                                IndexerStatus.Indexing,
                                folderProcessed: folderProcessed,
                                folderTotal: folderItemCount,
                                newlyIndexed: totalIndexed,
                                failed: totalFailed
                            );

                            if (batchResult.WasCancelled || _indexerCts.Token.IsCancellationRequested)
                            {
                                summary.IsCancelled = true;
                                (_outlookAdapter as IFolderBatchReader)?.ResetCursorSession();
                                break;
                            }

                            if (batchResult.LastProcessedIndex >= currentStart)
                            {
                                currentStart = batchResult.LastProcessedIndex + 1;
                            }
                            else
                            {
                                // No forward progress made (error/exhausted retries)
                                break;
                            }
                        }

                        if (summary.IsCancelled) break;
                        completedPreviousFoldersCount += folderItemCount;
                        _processedItems = completedPreviousFoldersCount;
                        Logger.Info($"[BACKFILL] Completed folder: {folder.Name}");
                    }

                    if (!summary.IsCancelled && !_isPaused)
                    {
                        // Process any pending changes that accumulated during backfill
                        await ProcessIncrementalQueueInternalAsync(_indexerCts.Token);

                        _completedAt = DateTime.Now;
                        CurrentStatus = IndexerStatus.Idle;
                        _currentOperation = "Idle";
                    }
                    else if (_isPaused)
                    {
                        CurrentStatus = IndexerStatus.Paused;
                    }

                    summary.ProcessedCount = totalIndexed + totalSkipped;
                    summary.NewlyIndexedCount = totalIndexed;
                    summary.SkippedCount = totalSkipped;
                    summary.FailedCount = totalFailed;

                    if (!summary.IsCancelled && !_isPaused)
                    {
                        await _repository.SaveLiveProgressAsync("Completed", "Idle", _processedItems, _totalItems, _currentFolder, _folderProcessedItems, _folderTotalItems, totalIndexed, totalSkipped, totalFailed, _pendingDbCount);
                        await _repository.SaveIndexStateAsync("Completed", totalIndexed, _totalItems, _currentFolder, null);
                        ReportProgress(_totalItems, _processedItems, _skippedItems, "", $"Backfill complete. {_processedItems:N0} emails indexed.", IndexerStatus.Idle, newlyIndexed: totalIndexed, failed: totalFailed);
                        Logger.Info($"[BACKFILL] Active store backfill finished successfully. Indexed: {totalIndexed}, Skipped: {totalSkipped}, Failed: {totalFailed}");
                    }
                    else if (_isPaused)
                    {
                        await _repository.SaveLiveProgressAsync("Paused", _currentOperation, _processedItems, _totalItems, _currentFolder, _folderProcessedItems, _folderTotalItems, totalIndexed, totalSkipped, totalFailed, _pendingDbCount);
                        await _repository.SaveIndexStateAsync("Paused", _processedItems, _totalItems, _currentFolder, summary.ErrorMessage ?? "Backfill paused: License required");
                        ReportProgress(_totalItems, _processedItems, _skippedItems, _currentFolder, summary.ErrorMessage ?? "Indexing paused: License required.", IndexerStatus.Paused, _folderProcessedItems, _folderTotalItems, newlyIndexed: totalIndexed, failed: totalFailed);
                    }
                    else
                    {
                        await _repository.SaveLiveProgressAsync("Cancelled", "Idle", _processedItems, _totalItems, _currentFolder, _folderProcessedItems, _folderTotalItems, totalIndexed, totalSkipped, totalFailed, _pendingDbCount);
                        await _repository.SaveIndexStateAsync("Cancelled", _processedItems, _totalItems, _currentFolder, "Backfill cancelled");
                        ReportProgress(_totalItems, _processedItems, _skippedItems, _currentFolder, "Backfill cancelled.", IndexerStatus.Cancelled, _folderProcessedItems, _folderTotalItems, newlyIndexed: totalIndexed, failed: totalFailed);
                    }
                }
                finally
                {
                    _incrementalProcessingLock.Release();
                }
            }
            catch (Exception ex)
            {
                _lastError = ex.Message;
                CurrentStatus = IndexerStatus.Error;
                _currentOperation = "Idle";
                summary.ErrorMessage = ex.Message;
                Logger.Error("[BACKFILL ERROR] Exception during active store backfill", ex);
                await _repository.SaveLiveProgressAsync("Failed", "Idle", _processedItems, _totalItems, _currentFolder, _folderProcessedItems, _folderTotalItems, 0, 0, 1, _pendingDbCount);
                await _repository.SaveIndexStateAsync("Failed", _processedItems, _totalItems, _currentFolder, ex.Message);
            }
            finally
            {
                _indexingLock.Release();
            }

            return summary;
        }

        public async Task RebuildIndexAsync(CancellationToken cancellationToken = default)
        {
            Logger.Info("[INDEX] Rebuild requested.");
            if (!IsLicensingEligible())
            {
                Logger.Info("[INDEX] Rebuild cannot start: Licensing is ineligible.");
                CurrentStatus = IndexerStatus.Paused;
                _currentOperation = "Paused: License required";
                ReportProgress(0, 0, 0, "", "Indexing is paused: License required", IndexerStatus.Paused);
                return;
            }

            if (CurrentStatus == IndexerStatus.Indexing)
            {
                Logger.Info("[INDEX] Rebuild requested while index rebuild is already in progress.");
                ReportProgress(_totalItems, _processedItems, _skippedItems, _currentFolder, "Index rebuild is already in progress.", IndexerStatus.Indexing);
                return;
            }

            if (_indexerCts != null)
            {
                try { _indexerCts.Cancel(); } catch { }
            }

            await _repository.ClearAllDataAsync();
            await _repository.VacuumDatabaseAsync();
            await StartScanInternalAsync(isRebuild: true, cancellationToken);
        }

        private Task StartScanInternalAsync(bool isRebuild, CancellationToken cancellationToken)
        {
            if (!IsLicensingEligible())
            {
                Logger.Info("[INDEX] Indexing cannot start: Licensing is ineligible.");
                CurrentStatus = IndexerStatus.Paused;
                _currentOperation = "Paused: License required";
                ReportProgress(0, 0, 0, "", "Indexing is paused: License required", IndexerStatus.Paused);
                return Task.CompletedTask;
            }

            if (!_indexingLock.Wait(0))
            {
                Logger.Info("[INDEX] Indexing or rebuild already in progress. Ignoring duplicate request.");
                return Task.CompletedTask;
            }

            _indexerCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            CurrentStatus = IndexerStatus.Indexing;
            _isRebuilding = isRebuild;
            _startedAt = DateTime.Now;
            _completedAt = null;
            _lastError = string.Empty;
            _processedItems = 0;
            _skippedItems = 0;
            _totalItems = 0;
            _currentFolder = string.Empty;
            _currentOperation = isRebuild ? "Rebuilding Index" : "Initial Indexing";

            return Task.Run(async () =>
            {
                try
                {
                    await _incrementalProcessingLock.WaitAsync();
                    try
                    {
                        Logger.Info($"[INDEX] Indexing started (Rebuild: {isRebuild}). Discovering Outlook mailstores...");
                        await _repository.SaveIndexStateAsync("Indexing", 0, 0, "", null);
                        ReportProgress(0, 0, 0, "", "Discovering Outlook mailstores...", IndexerStatus.Indexing);

                        if (!_outlookAdapter.IsOutlookAvailable())
                        {
                            Logger.Warn("[INDEX] Outlook Classic is not running or unavailable. Indexer entering idle mode.");
                            ReportProgress(0, 0, 0, "", "Outlook Classic is unavailable.", IndexerStatus.Idle);
                            CurrentStatus = IndexerStatus.Idle;
                            _currentOperation = "Idle";
                            await _repository.SaveIndexStateAsync("Idle", 0, 0, "", "Outlook unavailable");
                            return;
                        }

                        var mailboxes = TestModeStoreFilter.FilterMailboxes(_outlookAdapter.GetMailboxes());
                        int totalItemsInAllFolders = 0;
                        var folderTasks = new List<(string storeId, MailboxFolder folder)>();

                        foreach (var mb in mailboxes)
                        {
                            var folders = _outlookAdapter.GetFolders(mb.StoreId);
                            await _repository.SaveFoldersAsync(folders);

                            foreach (var folder in folders)
                            {
                                totalItemsInAllFolders += folder.ItemCount;
                                folderTasks.Add((mb.StoreId, folder));
                            }
                        }

                        _totalItems = totalItemsInAllFolders;
                        CurrentMetadata.TotalItemsToIndex = _totalItems;
                        Logger.Info($"[INDEX] Total items discovered: {_totalItems}");
                        ReportProgress(_totalItems, 0, 0, "", $"Indexing your mailbox ({_totalItems:N0} emails found)...", IndexerStatus.Indexing);

                        int completedPreviousFoldersCount = 0;
                        int totalIndexed = 0;
                        int totalSkipped = 0;
                        int batchSize = 250;

                        foreach (var (storeId, folder) in folderTasks)
                        {
                            if (_indexerCts.Token.IsCancellationRequested) break;

                            if (_outlookAdapter.IsOutboxFolder(storeId, folder.EntryId) || folder.Name.Equals("Outbox", StringComparison.OrdinalIgnoreCase))
                            {
                                Logger.Info($"[INDEX] Folder skipped: {folder.Name} (Outbox folder excluded from indexing scan)");
                                continue;
                            }

                            _currentFolder = folder.Name;
                            Logger.Info($"[INDEX] Folder started: {folder.Name}");
                            await _repository.SaveIndexStateAsync("Indexing", totalIndexed, _totalItems, folder.Name, null);

                            int folderItemCount = folder.ItemCount;
                            if (folderItemCount == 0)
                            {
                                Logger.Info($"[INDEX] Folder completed: {folder.Name} (0 items)");
                                continue;
                            }

                            int folderProcessed = 0;
                            _processedItems = completedPreviousFoldersCount + folderProcessed;
                            ReportProgress(
                                _totalItems,
                                _processedItems,
                                _skippedItems,
                                folder.Name,
                                $"Indexing: {_processedItems:N0} / {_totalItems:N0} (Folder: {folder.Name}: {folderProcessed:N0}/{folderItemCount:N0})",
                                IndexerStatus.Indexing,
                                folderProcessed: folderProcessed,
                                folderTotal: folderItemCount
                            );

                            for (int start = 1; start <= folderItemCount; start += batchSize)
                            {
                                if (!IsLicensingEligible())
                                {
                                    _isPaused = true;
                                    CurrentStatus = IndexerStatus.Paused;
                                    _currentOperation = "Indexing paused: License required";
                                    ReportProgress(_totalItems, _processedItems, totalSkipped, folder.Name, "Indexing paused: License required", IndexerStatus.Paused, folderProcessed, folderItemCount);
                                    Logger.Warn("[INDEX] Pausing indexing scan: License is ineligible.");
                                    break;
                                }

                                if (_indexerCts.Token.IsCancellationRequested) break;

                                while (_isPaused)
                                {
                                    ReportProgress(_totalItems, _processedItems, totalSkipped, folder.Name, "Indexing paused by user", IndexerStatus.Paused, folderProcessed, folderItemCount);
                                    await Task.Delay(500, _indexerCts.Token);
                                }

                                try
                                {
                                    var batch = _outlookAdapter.FetchFolderEmailsBatch(
                                        storeId,
                                        folder.EntryId,
                                        start,
                                        batchSize,
                                        () => _indexerCts.Token.IsCancellationRequested || _isPaused || (_testCancelCheck != null && _testCancelCheck())
                                    );

                                    if (batch.Count > 0)
                                    {
                                        await _repository.SaveEmailBatchAsync(batch, _indexerCts.Token);
                                        totalIndexed += batch.Count;
                                        _lastIndexedItemTime = DateTime.Now;
                                    }

                                    int batchAttempted = Math.Min(batchSize, folderItemCount - start + 1);
                                    totalSkipped += Math.Max(0, batchAttempted - batch.Count);

                                    int currentOffset = Math.Min(start + batchSize - 1, folderItemCount);
                                    folderProcessed = currentOffset;
                                    _processedItems = completedPreviousFoldersCount + folderProcessed;
                                    _skippedItems = totalSkipped;

                                    ReportProgress(
                                        _totalItems,
                                        _processedItems,
                                        _skippedItems,
                                        folder.Name,
                                        $"Indexing {_processedItems:N0} / {_totalItems:N0} (Folder: {folder.Name}: {folderProcessed:N0}/{folderItemCount:N0})",
                                        IndexerStatus.Indexing,
                                        folderProcessed: folderProcessed,
                                        folderTotal: folderItemCount
                                    );
                                }
                                catch (Exception ex)
                                {
                                    Logger.Error($"[INDEX ERROR] Error indexing batch in folder {folder.Name}. Processed: {_processedItems}/{_totalItems}", ex);
                                    totalSkipped += batchSize;
                                    _skippedItems = totalSkipped;
                                }
                            }

                            completedPreviousFoldersCount += folderItemCount;
                            _processedItems = completedPreviousFoldersCount;
                            Logger.Info($"[INDEX] Folder completed: {folder.Name}");
                        }

                        // Process remaining incremental items
                        await ProcessIncrementalQueueInternalAsync(_indexerCts.Token);

                        _completedAt = DateTime.Now;
                        CurrentStatus = IndexerStatus.Idle;
                        _currentOperation = "Idle";
                        _isRebuilding = false;

                        if (_indexerCts.Token.IsCancellationRequested)
                        {
                            string cancelMsg = $"Indexing cancelled ({_processedItems:N0} emails indexed)";
                            ReportProgress(_totalItems, _processedItems, _skippedItems, _currentFolder, cancelMsg, IndexerStatus.Cancelled, _folderProcessedItems, _folderTotalItems);
                            await _repository.SaveIndexStateAsync("Cancelled", _processedItems, _totalItems, _currentFolder, "Cancelled by user or shutdown");
                        }
                        else
                        {
                            TimeSpan duration = _completedAt.Value - (_startedAt ?? _completedAt.Value);
                            string finalMsg = totalSkipped > 0
                                ? $"Indexing complete. Indexed: {totalIndexed:N0}, Skipped: {totalSkipped:N0}"
                                : $"Up to date ({totalIndexed:N0} emails indexed)";

                            ReportProgress(_totalItems, totalIndexed, totalSkipped, _currentFolder, finalMsg, IndexerStatus.Idle);
                            await _repository.SaveIndexStateAsync("Completed", totalIndexed, _totalItems, _currentFolder, null);

                            Logger.Info($"[INDEX] Indexing completed successfully.");
                            Logger.Info($"[INDEX] Indexed {totalIndexed:N0} emails.");
                            Logger.Info($"[INDEX] Duration: {duration:hh\\:mm\\:ss}");
                        }
                    }
                    finally
                    {
                        _incrementalProcessingLock.Release();
                    }
                }
                catch (OperationCanceledException)
                {
                    CurrentStatus = IndexerStatus.Idle;
                    _currentOperation = "Idle";
                    _isRebuilding = false;
                    string cancelMsg = $"Indexing cancelled ({_processedItems:N0} emails indexed)";
                    ReportProgress(_totalItems, _processedItems, _skippedItems, _currentFolder, cancelMsg, IndexerStatus.Cancelled, _folderProcessedItems, _folderTotalItems);
                    await _repository.SaveIndexStateAsync("Cancelled", _processedItems, _totalItems, _currentFolder, "Cancelled by user or shutdown");
                    Logger.Info($"[INDEX] Indexing operation cancelled gracefully. Processed: {_processedItems}/{_totalItems}");
                }
                catch (Exception ex)
                {
                    CurrentStatus = IndexerStatus.Error;
                    _currentOperation = "Failed";
                    _isRebuilding = false;
                    _lastError = ex.Message;
                    string errorMsg = $"Indexing failed: {ex.Message}";
                    ReportProgress(_totalItems, _processedItems, _skippedItems, _currentFolder, errorMsg, IndexerStatus.Error, _folderProcessedItems, _folderTotalItems);
                    await _repository.SaveIndexStateAsync("Failed", _processedItems, _totalItems, _currentFolder, ex.Message);
                    Logger.Error($"[INDEX ERROR] Exception during indexing scan. Folder: {_currentFolder}, Processed: {_processedItems}/{_totalItems}", ex);
                }
                finally
                {
                    _indexingLock.Release();
                }
            });
        }

        public Task CancelIndexingAsync()
        {
            Logger.Info("[INDEX] Indexing cancellation requested.");
            if (_indexerCts != null && !_indexerCts.IsCancellationRequested)
            {
                try { _indexerCts.Cancel(); } catch { }
            }
            if (_incrementalCts != null && !_incrementalCts.IsCancellationRequested)
            {
                try { _incrementalCts.Cancel(); } catch { }
            }
            CurrentStatus = IndexerStatus.Cancelled;
            ReportProgress(_totalItems, _processedItems, _skippedItems, _currentFolder, $"Indexing cancelled ({_processedItems:N0} emails indexed)", IndexerStatus.Cancelled);
            return Task.CompletedTask;
        }

        public Task PauseIndexingAsync()
        {
            _isPaused = true;
            CurrentStatus = IndexerStatus.Paused;
            _currentOperation = "Paused";
            ReportProgress(_totalItems, _processedItems, _skippedItems, _currentFolder, "Paused", IndexerStatus.Paused);
            Logger.Info("Background indexer paused.");
            return Task.CompletedTask;
        }

        public Task ResumeIndexingAsync()
        {
            if (!IsLicensingEligible())
            {
                _isPaused = true;
                CurrentStatus = IndexerStatus.Paused;
                Logger.Info("[INDEX] Cannot resume indexing: Licensing is ineligible.");
                return Task.CompletedTask;
            }

            _isPaused = false;
            if (CurrentStatus == IndexerStatus.Paused)
            {
                CurrentStatus = IndexerStatus.Indexing;
                _currentOperation = _isRebuilding ? "Rebuilding Index" : "Indexing";
                ReportProgress(_totalItems, _processedItems, _skippedItems, _currentFolder, "Resuming indexing...", IndexerStatus.Indexing);
                Logger.Info("Background indexer resumed.");
                _ = Task.Run(async () => await ResumeIncrementalCatchUpAsync());
            }
            return Task.CompletedTask;
        }

        public async Task QueueItemForIndexingAsync(string storeId, string entryId, string folderEntryId)
        {
            if (_status == IndexerStatus.SafeMode) return;
            if (string.IsNullOrEmpty(storeId) || string.IsNullOrEmpty(entryId)) return;

            if (TestModeStoreFilter.IsTestModeActive && !TestModeStoreFilter.IsStoreIdApproved(storeId))
            {
                Logger.Warn($"[TEST MODE SAFEGUARD] Dropping queued item for unapproved store '{storeId}'.");
                return;
            }

            if (_outlookAdapter.IsOutboxFolder(storeId, folderEntryId))
            {
                Logger.Info($"[OUTLOOK EVENT] Type=ItemAdd Folder=Outbox EntryId={entryId} Action=SKIPPED_OUTBOX");
                return;
            }

            Logger.Info($"[INCREMENTAL] Event received: StoreId={storeId} EntryId={entryId}");

            // Persist to database queue
            await _repository.EnqueuePendingChangeAsync(storeId, entryId, "Add", folderEntryId);

            string key = $"{storeId}:{entryId}";
            if (_pendingItems.TryGetValue(key, out var existing))
            {
                existing.Operation = IncrementalOperation.InsertOrUpdate;
                existing.FolderEntryId = folderEntryId;
                Logger.Info($"[INCREMENTAL] Duplicate ignored: Merged queue item for EntryId={entryId}");
            }
            else
            {
                var item = new IncrementalItem
                {
                    StoreId = storeId,
                    EntryId = entryId,
                    FolderEntryId = folderEntryId,
                    Operation = IncrementalOperation.InsertOrUpdate
                };
                if (_pendingItems.TryAdd(key, item))
                {
                    _incrementalQueue.Enqueue(item);
                    Logger.Info($"[INCREMENTAL] Queued item EntryId={entryId}");
                }
            }

            _lastIndexedItemTime = DateTime.Now;
            if (!IsLicensingEligible())
            {
                if (CurrentStatus == IndexerStatus.Indexing)
                {
                    CurrentStatus = IndexerStatus.Paused;
                }
                return;
            }
            await TriggerIncrementalProcessingAsync();
        }

        public async Task QueueItemForDeletionAsync(string storeId, string entryId)
        {
            if (string.IsNullOrEmpty(storeId) || string.IsNullOrEmpty(entryId)) return;

            if (TestModeStoreFilter.IsTestModeActive && !TestModeStoreFilter.IsStoreIdApproved(storeId))
            {
                Logger.Warn($"[TEST MODE SAFEGUARD] Dropping queued item for unapproved store '{storeId}'.");
                return;
            }

            Logger.Info($"[INCREMENTAL] Event received: Delete item StoreId={storeId} EntryId={entryId}");

            // Persist to database queue
            await _repository.EnqueuePendingChangeAsync(storeId, entryId, "Delete", null);

            string key = $"{storeId}:{entryId}";
            if (_pendingItems.TryGetValue(key, out var existing))
            {
                existing.Operation = IncrementalOperation.Delete;
                Logger.Info($"[INCREMENTAL] Duplicate ignored: Updated queue item to Delete for EntryId={entryId}");
            }
            else
            {
                var item = new IncrementalItem
                {
                    StoreId = storeId,
                    EntryId = entryId,
                    Operation = IncrementalOperation.Delete
                };
                if (_pendingItems.TryAdd(key, item))
                {
                    _incrementalQueue.Enqueue(item);
                    Logger.Info($"[INCREMENTAL] Queued item EntryId={entryId}");
                }
            }

            if (!IsLicensingEligible())
            {
                if (CurrentStatus == IndexerStatus.Indexing)
                {
                    CurrentStatus = IndexerStatus.Paused;
                }
                return;
            }
            await TriggerIncrementalProcessingAsync();
        }

        public Task<IndexIntegrityResult> ScanIndexIntegrityAsync(CancellationToken cancellationToken = default)
        {
            return _repository.ScanIndexIntegrityAsync(cancellationToken);
        }

        public async Task RepairIndexIntegrityAsync(IndexIntegrityResult scanResult, CancellationToken cancellationToken = default)
        {
            ReportProgress(_totalItems, _processedItems, _skippedItems, _currentFolder, "🟡 Index repair in progress...", IndexerStatus.Indexing);
            await _repository.RepairIndexIntegrityAsync(scanResult, cancellationToken);
            ReportStatusUpdate();
        }

        private async Task TriggerIncrementalProcessingAsync()
        {
            if (!IsLicensingEligible())
            {
                if (CurrentStatus == IndexerStatus.Indexing)
                {
                    CurrentStatus = IndexerStatus.Paused;
                }
                return;
            }

            if (CurrentStatus == IndexerStatus.Indexing || _isRebuilding)
            {
                return;
            }

            if (!_incrementalProcessingLock.Wait(0))
            {
                return;
            }

            try
            {
                _incrementalCts = new CancellationTokenSource();
                CurrentStatus = IndexerStatus.Indexing;
                await ProcessIncrementalQueueInternalAsync(_incrementalCts.Token);
            }
            catch (OperationCanceledException)
            {
                Logger.Info("[INCREMENTAL] Incremental processing cancelled.");
            }
            finally
            {
                if (!_isRebuilding)
                {
                    CurrentStatus = IndexerStatus.Idle;
                }
                _incrementalProcessingLock.Release();
            }
        }

        private async Task ProcessIncrementalQueueInternalAsync(CancellationToken token)
        {
            bool hasMoreWork = true;

            while (hasMoreWork)
            {
                if (token.IsCancellationRequested) break;

                if (!IsLicensingEligible())
                {
                    CurrentStatus = IndexerStatus.Paused;
                    _currentOperation = "Paused: License required";
                    Logger.Warn("[INCREMENTAL] Pausing incremental processing: License is ineligible.");
                    break;
                }

                if (_incrementalQueue.IsEmpty)
                {
                    var pendingDbItems = await _repository.GetPendingChangesAsync(1000);
                    if (pendingDbItems.Count == 0)
                    {
                        hasMoreWork = false;
                        break;
                    }

                    foreach (var p in pendingDbItems)
                    {
                        var item = new IncrementalItem
                        {
                            PendingId = p.Id,
                            StoreId = p.StoreId,
                            EntryId = p.EntryId,
                            FolderEntryId = p.FolderEntryId ?? "",
                            Operation = string.Equals(p.Operation, "Delete", StringComparison.OrdinalIgnoreCase) ? IncrementalOperation.Delete : IncrementalOperation.InsertOrUpdate,
                            AttemptCount = p.AttemptCount
                        };

                        string key = $"{p.StoreId}:{p.EntryId}";
                        if (_pendingItems.TryAdd(key, item))
                        {
                            _incrementalQueue.Enqueue(item);
                        }
                    }

                    if (_incrementalQueue.IsEmpty)
                    {
                        hasMoreWork = false;
                        break;
                    }
                }

                _pendingDbCount = await _repository.GetPendingChangesCountAsync();

                Logger.Info($"[INCREMENTAL] Batch started: Queued items = {_incrementalQueue.Count}");
                var pendingEmails = new List<EmailDocument>();
                var pendingEmailIds = new List<long>();
                var pendingBatchItems = new List<(IncrementalItem Item, string Key)>();
                var otherCompletedIds = new List<long>();
                int processedCount = 0;

                while (_incrementalQueue.TryDequeue(out var item))
                {
                    if (token.IsCancellationRequested) break;

                    string key = $"{item.StoreId}:{item.EntryId}";

                    if (_outlookAdapter.IsOutboxFolder(item.StoreId, item.FolderEntryId))
                    {
                        Logger.Info($"[INDEX] EntryId={item.EntryId} Folder=Outbox Action=SKIPPED_OUTBOX");
                        _pendingItems.TryRemove(key, out _);
                        if (item.PendingId > 0) otherCompletedIds.Add(item.PendingId);
                        processedCount++;
                        continue;
                    }

                    if (item.Operation == IncrementalOperation.Delete)
                    {
                        try
                        {
                            await _repository.DeleteEmailAsync(item.StoreId, item.EntryId);
                            Logger.Info($"[INCREMENTAL] Item deleted: StoreId={item.StoreId} EntryId={item.EntryId}");
                            _pendingItems.TryRemove(key, out _);
                            if (item.PendingId > 0) otherCompletedIds.Add(item.PendingId);
                            processedCount++;
                        }
                        catch (Exception ex)
                        {
                            item.LastError = ex.Message;
                            Logger.Error($"[INCREMENTAL ERROR] Failed to delete item EntryId={item.EntryId}", ex);
                            await HandleRetryAsync(item, key);
                        }
                    }
                    else
                    {
                        try
                        {
                            var swIndex = System.Diagnostics.Stopwatch.StartNew();
                            var doc = _outlookAdapter.FetchSingleEmail(item.StoreId, item.EntryId);
                            if (doc != null)
                            {
                                pendingEmails.Add(doc);
                                _pendingItems.TryRemove(key, out _);
                                if (item.PendingId > 0) pendingEmailIds.Add(item.PendingId);
                                pendingBatchItems.Add((item, key));
                                processedCount++;

                                swIndex.Stop();
                                Logger.Info($"[INDEX] EntryId={item.EntryId} DurationMs={swIndex.ElapsedMilliseconds}");

                                if (pendingEmails.Count >= 50)
                                {
                                    try
                                    {
                                        await _repository.SaveEmailBatchAsync(pendingEmails, token);
                                        await _repository.MarkPendingChangesCompletedBatchAsync(pendingEmailIds);
                                        Logger.Info($"[INCREMENTAL] Batch completed ({pendingEmails.Count} items)");
                                    }
                                    catch (Exception ex)
                                    {
                                        Logger.Error($"[INCREMENTAL ERROR] Failed to save batch of {pendingEmails.Count} emails.", ex);
                                        foreach (var b in pendingBatchItems)
                                        {
                                            b.Item.LastError = ex.Message;
                                            _ = HandleRetryAsync(b.Item, b.Key);
                                        }
                                    }
                                    finally
                                    {
                                        pendingEmails.Clear();
                                        pendingEmailIds.Clear();
                                        pendingBatchItems.Clear();
                                    }
                                }
                            }
                            else
                            {
                                if (_outlookAdapter.IsOutboxFolder(item.StoreId, item.FolderEntryId))
                                {
                                    Logger.Info($"[INDEX] EntryId={item.EntryId} Action=SKIPPED_OUTBOX");
                                    _pendingItems.TryRemove(key, out _);
                                    if (item.PendingId > 0) otherCompletedIds.Add(item.PendingId);
                                }
                                else
                                {
                                    item.LastError = "Item could not be fetched from Outlook Classic MAPI.";
                                    await HandleRetryAsync(item, key);
                                }
                            }
                        }
                        catch (InvalidOperationException ex) when (ex.Message == "ITEM_NOT_FOUND" || ex.Message == "UNSUPPORTED_ITEM_TYPE")
                        {
                            Logger.Info($"[INCREMENTAL] Skipping non-retriable item EntryId={item.EntryId}: {ex.Message}");
                            _pendingItems.TryRemove(key, out _);
                            if (item.PendingId > 0) otherCompletedIds.Add(item.PendingId);
                            processedCount++;
                        }
                        catch (Exception ex)
                        {
                            item.LastError = ex.Message;
                            Logger.Error($"[INCREMENTAL ERROR] Exception fetching item EntryId={item.EntryId}", ex);
                            await HandleRetryAsync(item, key);
                        }
                    }

                    if (otherCompletedIds.Count >= 50)
                    {
                        await _repository.MarkPendingChangesCompletedBatchAsync(otherCompletedIds);
                        otherCompletedIds.Clear();
                    }

                    if (processedCount % 10 == 0)
                    {
                        ReportStatusUpdate();
                    }
                }

                if (pendingEmails.Count > 0)
                {
                    try
                    {
                        await _repository.SaveEmailBatchAsync(pendingEmails, token);
                        await _repository.MarkPendingChangesCompletedBatchAsync(pendingEmailIds);
                        Logger.Info($"[INCREMENTAL] Batch completed ({pendingEmails.Count} items)");
                    }
                    catch (Exception ex)
                    {
                        Logger.Error($"[INCREMENTAL ERROR] Failed to save batch of {pendingEmails.Count} emails.", ex);
                        foreach (var b in pendingBatchItems)
                        {
                            b.Item.LastError = ex.Message;
                            _ = HandleRetryAsync(b.Item, b.Key);
                        }
                    }
                    finally
                    {
                        pendingEmails.Clear();
                        pendingEmailIds.Clear();
                        pendingBatchItems.Clear();
                    }
                }

                if (otherCompletedIds.Count > 0)
                {
                    await _repository.MarkPendingChangesCompletedBatchAsync(otherCompletedIds);
                    otherCompletedIds.Clear();
                }

                Logger.Info($"[INCREMENTAL] Batch completed: Processed {processedCount} items");
                
                // Update pendingDbCount after processing
                _pendingDbCount = await _repository.GetPendingChangesCountAsync();
                ReportStatusUpdate();
            }
            
            _pendingDbCount = await _repository.GetPendingChangesCountAsync();
            ReportStatusUpdate();
        }

        private async Task HandleRetryAsync(IncrementalItem item, string key)
        {
            try
            {
                item.AttemptCount++;
                if (item.PendingId > 0)
                {
                    await _repository.IncrementAttemptCountAsync(item.PendingId);
                }
                if (item.AttemptCount <= 3)
                {
                    int delayMs = item.AttemptCount == 1 ? 100 : item.AttemptCount == 2 ? 300 : 1000;
                    Logger.Warn($"[INCREMENTAL ERROR] Item failed (Attempt {item.AttemptCount}/3): EntryId={item.EntryId}. Exponential backoff delay {delayMs}ms...");
                    _ = Task.Run(async () =>
                    {
                        try
                        {
                            await Task.Delay(delayMs);
                        }
                        catch { }
                        _incrementalQueue.Enqueue(item);
                        _ = TriggerIncrementalProcessingAsync();
                    });
                }
                else
                {
                    Logger.Error($"[INCREMENTAL ERROR] Item permanently failed after 3 attempts. EntryId={item.EntryId}. Moved to failed queue.");
                    _pendingItems.TryRemove(key, out _);
                    _failedQueue.Enqueue(item);

                    var failedItem = new FailedIndexItem
                    {
                        StoreId = item.StoreId,
                        EntryId = item.EntryId,
                        FolderName = item.FolderEntryId,
                        OperationType = item.Operation.ToString(),
                        FailureReason = item.LastError ?? "Failed after 3 retries",
                        ExceptionType = "MapiFetchException",
                        AttemptCount = item.AttemptCount,
                        FirstFailedAt = item.QueuedTime,
                        LastFailedAt = DateTime.Now,
                        Status = "Failed",
                        ErrorMessage = item.LastError ?? "Item could not be fetched from Outlook Classic MAPI."
                    };

                    await _repository.SaveFailedIndexItemAsync(failedItem);
                    if (item.PendingId > 0)
                    {
                        await _repository.MarkPendingChangeStatusAsync(item.PendingId, "Failed", item.LastError);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"[INCREMENTAL ERROR] HandleRetryAsync crashed for EntryId={item.EntryId}", ex);
                // Ensure we don't leak it in _pendingItems if we crashed
                _pendingItems.TryRemove(key, out _);
            }
        }

        private void ReportStatusUpdate()
        {
            int remainingQueue = Math.Max(_pendingDbCount, _incrementalQueue.Count);
            string msg;
            IndexerStatus st;

            if (remainingQueue > 0)
            {
                msg = $"🔵 Updating • {remainingQueue:N0} email(s) remaining";
                st = IndexerStatus.Indexing;
            }
            else if (CurrentStatus == IndexerStatus.Indexing)
            {
                msg = _totalItems > 0
                    ? $"Indexing: {_processedItems:N0} / {_totalItems:N0} (Folder: {_currentFolder})"
                    : "Indexing in progress...";
                st = IndexerStatus.Indexing;
            }
            else
            {
                msg = $"🟢 Index Healthy ({_totalItems:N0} emails indexed)";
                st = IndexerStatus.Idle;
            }

            ReportProgress(_totalItems, _processedItems, _skippedItems, _currentFolder, msg, st, _folderProcessedItems, _folderTotalItems);
        }

        private void ReportProgress(
            int total,
            int processed,
            int skipped,
            string currentFolder,
            string message,
            IndexerStatus status,
            int folderProcessed = 0,
            int folderTotal = 0,
            int newlyIndexed = 0,
            int failed = 0)
        {
            _totalItems = total;
            _processedItems = processed;
            _skippedItems = skipped;
            _currentFolder = currentFolder ?? string.Empty;
            _folderProcessedItems = folderProcessed;
            _folderTotalItems = folderTotal;

            int remainingQueue = Math.Max(_incrementalQueue.Count, _pendingDbCount);
            string displayMessage = message;
            
            if (remainingQueue > 0 && status == IndexerStatus.Idle)
            {
                displayMessage = $"🔵 Updating • {remainingQueue:N0} email(s) remaining";
                status = IndexerStatus.Indexing; // Force indexing status if there's queue
            }

            CurrentMetadata.TotalItemsToIndex = total;
            CurrentMetadata.CurrentlyIndexedCount = processed;
            CurrentMetadata.Status = status.ToString();
            CurrentMetadata.CurrentStatusDetail = displayMessage;
            CurrentMetadata.LastIndexedTime = DateTime.Now;

            double pct = total > 0 ? Math.Min(100.0, Math.Round((double)processed / total * 100.0, 2)) : 0.0;
            int remainingToScan = Math.Max(0, total - processed);

            ProgressChanged?.Invoke(this, new IndexerProgressEventArgs
            {
                TotalItems = total,
                ProcessedItems = processed,
                SkippedItems = skipped,
                CurrentFolder = _currentFolder,
                Percentage = pct,
                CurrentStatusMessage = displayMessage,
                Status = status,
                IsRebuilding = _isRebuilding,
                StartedAt = _startedAt,
                CompletedAt = _completedAt,
                LastIndexedItemTime = _lastIndexedItemTime,
                CurrentOperation = _currentOperation,
                LastError = _lastError,
                FolderProcessedItems = _folderProcessedItems,
                FolderTotalItems = _folderTotalItems,
                TotalDiscovered = total,
                ScannedCount = processed,
                SuccessfullyIndexed = newlyIndexed > 0 ? newlyIndexed : processed,
                NewlyIndexedCount = newlyIndexed,
                FailedCount = failed,
                RemainingCount = remainingToScan,
                PendingQueueCount = remainingQueue
            });
        }
    }
}
