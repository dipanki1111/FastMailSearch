using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using FastMailSearch.Core.Interfaces;
using FastMailSearch.Core.Models;
using FastMailSearch.Infrastructure.Database;
using FastMailSearch.Infrastructure.Diagnostics;

namespace FastMailSearch.Infrastructure.Services
{
    public class IndexHealthService : IIndexHealthService
    {
        private readonly IEmailRepository _repository;
        private readonly IIndexerService _indexerService;
        private readonly IOutlookAdapter? _outlookAdapter;
        private readonly DatabaseConnectionFactory? _dbFactory;

        public IndexHealthService(IEmailRepository repository, IIndexerService indexerService, IOutlookAdapter? outlookAdapter = null, DatabaseConnectionFactory? dbFactory = null)
        {
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
            _indexerService = indexerService ?? throw new ArgumentNullException(nameof(indexerService));
            _outlookAdapter = outlookAdapter;
            _dbFactory = dbFactory ?? (repository as SqliteEmailRepository)?.ConnectionFactory ?? (!string.IsNullOrEmpty(repository.DatabasePath) ? new DatabaseConnectionFactory(repository.DatabasePath) : null);
        }

        public async Task<IndexHealthSummary> GetHealthSummaryAsync(CancellationToken cancellationToken = default)
        {
            var meta = await _repository.GetIndexMetadataAsync();
            var failedItems = await _repository.GetFailedIndexItemsAsync(includeResolved: true);
            double dbSizeMb = await _repository.GetDatabaseSizeMbAsync();
            var integrity = await _repository.ScanIndexIntegrityAsync(cancellationToken);

            int failedCount = failedItems.Count(i => string.Equals(i.Status, "Failed", StringComparison.OrdinalIgnoreCase));
            int recoveredCount = failedItems.Count(i => string.Equals(i.Status, "Recovered", StringComparison.OrdinalIgnoreCase));
            int ignoredCount = failedItems.Count(i => string.Equals(i.Status, "Ignored", StringComparison.OrdinalIgnoreCase));

            int pendingCount = await _repository.GetPendingChangesCountAsync();

            var persistedState = await _repository.GetPersistedIndexStateAsync();
            var folderCheckpoints = await _repository.GetAllFolderIndexStatesAsync();

            var relevantCheckpoints = folderCheckpoints;
            var activeIndexingFolder = folderCheckpoints.FirstOrDefault(f => f.Status == "Indexing");
            if (activeIndexingFolder != null)
            {
                relevantCheckpoints = folderCheckpoints.Where(f => string.Equals(f.StoreId, activeIndexingFolder.StoreId, StringComparison.OrdinalIgnoreCase)).ToList();
            }
            else if (_outlookAdapter != null && _outlookAdapter.IsOutlookAvailable())
            {
                try
                {
                    var mailboxes = _outlookAdapter.GetMailboxes();
                    if (mailboxes != null && mailboxes.Count > 0)
                    {
                        var mailboxStoreIds = new HashSet<string>(mailboxes.Select(m => m.StoreId), StringComparer.OrdinalIgnoreCase);
                        var storeMatched = folderCheckpoints.Where(f => mailboxStoreIds.Contains(f.StoreId)).ToList();
                        if (storeMatched.Count > 0)
                        {
                            relevantCheckpoints = storeMatched;
                        }
                    }
                }
                catch { }
            }

            // Detect active cross-process indexing (e.g. running in Outlook Classic or separate worker)
            bool isCrossProcessIndexing = false;
            string? activeOperation = null;
            int activeProcessed = 0;
            int activeTotal = 0;

            if (!_indexerService.IsIndexing && persistedState.TryGetValue("IndexState", out var stateStr) && stateStr == "Indexing")
            {
                if (persistedState.TryGetValue("LastSync", out var lastSyncStr) && DateTime.TryParse(lastSyncStr, out var lastSyncDt))
                {
                    // If last sync was within 3 minutes, it is actively running
                    if ((DateTime.Now - lastSyncDt).TotalMinutes < 3.0)
                    {
                        isCrossProcessIndexing = true;
                        activeOperation = persistedState.TryGetValue("CurrentOperation", out var op) && !string.IsNullOrWhiteSpace(op) ? op : "Backfilling Active Store";
                        int.TryParse(persistedState.TryGetValue("LastProcessedCount", out var pStr) ? pStr : "0", out activeProcessed);
                        int.TryParse(persistedState.TryGetValue("LastTotalCount", out var tStr) ? tStr : "0", out activeTotal);
                    }
                }
            }

            if (!isCrossProcessIndexing && !_indexerService.IsIndexing && activeIndexingFolder != null)
            {
                if (DateTime.TryParse(activeIndexingFolder.LastSync, out var fSyncDt) && (DateTime.Now - fSyncDt).TotalMinutes < 3.0)
                {
                    isCrossProcessIndexing = true;
                    activeOperation = $"Backfilling ({activeIndexingFolder.ProcessedOffset:N0}/{activeIndexingFolder.TotalItems:N0})";
                    activeProcessed = relevantCheckpoints.Sum(f => f.ProcessedOffset);
                    activeTotal = relevantCheckpoints.Sum(f => f.TotalItems);
                }
            }

            IndexerHealth healthState = IndexerHealth.Healthy;
            string statusState = "🟢 Database Healthy · Ready";
            string currentOperation = "Idle";

            if (_indexerService.CurrentStatus == IndexerStatus.SafeMode)
            {
                healthState = IndexerHealth.SafeMode;
                statusState = "🟡 FastMail Search started in Safe Mode. Outlook functionality is unaffected.";
                currentOperation = "Safe Mode";
            }
            else if (_indexerService.CurrentStatus == IndexerStatus.Error)
            {
                healthState = IndexerHealth.Error;
                statusState = "🔴 Indexer Error";
                currentOperation = _indexerService.CurrentOperation ?? "Error";
            }
            else if (_indexerService.CurrentStatus == IndexerStatus.Paused)
            {
                healthState = IndexerHealth.Paused;
                statusState = "⏸ Indexer Paused";
                currentOperation = "Paused";
            }
            else if (_indexerService.IsIndexing)
            {
                healthState = IndexerHealth.Updating;
                currentOperation = _indexerService.CurrentOperation ?? "Indexing";
                double pct = _indexerService.TotalItems > 0 ? (double)_indexerService.ProcessedItems / _indexerService.TotalItems * 100.0 : 0.0;
                int remaining = Math.Max(0, _indexerService.TotalItems - _indexerService.ProcessedItems);
                statusState = $"🔵 Indexing: {_indexerService.ProcessedItems:N0} / {_indexerService.TotalItems:N0} ({pct:F0}%) • {remaining:N0} remaining";
            }
            else if (isCrossProcessIndexing)
            {
                healthState = IndexerHealth.Updating;
                currentOperation = activeOperation ?? "Backfilling Active Store (in Outlook)";
                int totalDisc = activeTotal > 0 ? activeTotal : relevantCheckpoints.Sum(f => f.TotalItems);
                int proc = activeProcessed > 0 ? activeProcessed : relevantCheckpoints.Sum(f => f.ProcessedOffset);
                double pct = totalDisc > 0 ? (double)proc / totalDisc * 100.0 : 0.0;
                int remaining = Math.Max(0, totalDisc - proc);
                statusState = $"🔵 Backfill in Progress • {proc:N0} / {totalDisc:N0} ({pct:F0}%) • {remaining:N0} remaining";
            }
            else if (pendingCount > 0)
            {
                healthState = IndexerHealth.Updating;
                currentOperation = "Processing Pending Queue";
                statusState = $"🔵 Updating • {pendingCount:N0} email(s) remaining";
            }
            else if (failedCount > 0)
            {
                healthState = IndexerHealth.Degraded;
                currentOperation = "Idle";
                statusState = $"🟡 {failedCount:N0} indexing issue(s)";
            }
            else if (!integrity.IsHealthy)
            {
                healthState = IndexerHealth.Degraded;
                currentOperation = "Idle";
                statusState = "🔴 Index requires attention";
            }

            // Calculate mailbox discovered, processed offset, and remaining to scan from real checkpoints and coverage
            int discoveredItems = 0;
            int processedOffset = 0;

            if (relevantCheckpoints.Count > 0)
            {
                discoveredItems = relevantCheckpoints.Sum(f => f.TotalItems);
                processedOffset = relevantCheckpoints.Sum(f => f.ProcessedOffset);
            }

            StoreCoverageInfo? coverage = null;
            if (!_indexerService.IsIndexing && !isCrossProcessIndexing)
            {
                try
                {
                    var coverageTask = _indexerService.CheckStoreCoverageAsync(cancellationToken);
                    var completed = await Task.WhenAny(coverageTask, Task.Delay(1500, cancellationToken));
                    if (completed == coverageTask)
                    {
                        coverage = await coverageTask;
                    }
                }
                catch { }
            }

            int remainingToScan = 0;
            if (coverage != null && coverage.TotalDiscoveredItems > 0)
            {
                discoveredItems = coverage.TotalDiscoveredItems;
                remainingToScan = coverage.UnindexedItems;
            }
            else
            {
                remainingToScan = Math.Max(0, discoveredItems - (processedOffset > 0 ? processedOffset : meta.TotalEmailsIndexed));
            }

            if (processedOffset == 0 && meta.TotalEmailsIndexed > 0 && meta.TotalEmailsIndexed <= discoveredItems)
            {
                processedOffset = meta.TotalEmailsIndexed;
            }

            string mailboxCoverageStatus = "Up to date";
            if (discoveredItems > 0)
            {
                if (remainingToScan == 0)
                {
                    mailboxCoverageStatus = "Fully indexed";
                }
                else if (_indexerService.IsIndexing || isCrossProcessIndexing)
                {
                    mailboxCoverageStatus = $"Backfill in progress ({remainingToScan:N0} remaining)";
                }
                else
                {
                    mailboxCoverageStatus = $"{remainingToScan:N0} unindexed in Outlook (Backfill recommended)";
                }
            }
            else if (coverage != null)
            {
                mailboxCoverageStatus = coverage.SummaryText;
            }

            return new IndexHealthSummary
            {
                HealthState = healthState,
                StatusState = statusState,
                DatabasePath = _dbFactory?.DbPath ?? _repository.DatabasePath ?? string.Empty,
                IndexedEmailCount = meta.TotalEmailsIndexed,
                PendingQueueCount = pendingCount,
                CurrentlyProcessingCount = (_indexerService.IsIndexing || isCrossProcessIndexing) ? 1 : 0,
                FailedItemCount = failedCount,
                RecoveredItemCount = recoveredCount,
                IgnoredItemCount = ignoredCount,
                LastSuccessfulIndexingTime = meta.LastIndexedTime,
                DatabaseSizeMb = dbSizeMb,
                IntegrityResult = integrity,
                CurrentStatusMessage = statusState,
                CurrentOperation = currentOperation,
                LastError = _indexerService.LastError,
                MailboxCoverageStatus = mailboxCoverageStatus,
                TotalDiscoveredOutlookItems = discoveredItems,
                UnindexedOutlookItems = remainingToScan,
                RemainingToScan = remainingToScan,
                ProcessedOffset = processedOffset
            };
        }

        public Task<StoreCoverageInfo> CheckStoreCoverageAsync(CancellationToken cancellationToken = default)
        {
            return _indexerService.CheckStoreCoverageAsync(cancellationToken);
        }

        public Task<BackfillSummary> StartActiveStoreBackfillAsync(CancellationToken cancellationToken = default)
        {
            return _indexerService.StartActiveStoreBackfillAsync(cancellationToken);
        }

        public Task<EnvironmentDiagnostics> GetEnvironmentDiagnosticsAsync(CancellationToken cancellationToken = default)
        {
            var factory = _dbFactory ?? (_repository as SqliteEmailRepository)?.ConnectionFactory ?? new DatabaseConnectionFactory(!string.IsNullOrEmpty(_repository.DatabasePath) ? _repository.DatabasePath : string.Empty);
            var detector = new EnvironmentDetector(factory, _outlookAdapter);
            return Task.FromResult(detector.DetectEnvironment(includeOutlookStores: false));
        }

        public async Task<string> ExportDiagnosticsZipAsync(string targetFolderPath, CancellationToken cancellationToken = default)
        {
            var factory = _dbFactory ?? new DatabaseConnectionFactory(!string.IsNullOrEmpty(_repository.DatabasePath) ? _repository.DatabasePath : string.Empty);
            var detector = new EnvironmentDetector(factory, _outlookAdapter);
            var exporter = new DiagnosticsExporter(detector, this, _repository);
            return await exporter.ExportZipAsync(targetFolderPath);
        }

        public Task<IndexIntegrityResult> VerifyIndexAsync(CancellationToken cancellationToken = default)
        {
            return _repository.ScanIndexIntegrityAsync(cancellationToken);
        }

        public Task RepairIndexAsync(IndexIntegrityResult scanResult, CancellationToken cancellationToken = default)
        {
            return _repository.RepairIndexIntegrityAsync(scanResult, cancellationToken);
        }

        public Task<List<FailedIndexItem>> GetFailedItemsAsync(string? status = null, CancellationToken cancellationToken = default)
        {
            return _repository.GetFailedIndexItemsAsync(status, maxCount: 1000);
        }

        public Task<int> GetFailedItemsCountAsync(string? status = null, CancellationToken cancellationToken = default)
        {
            return _repository.GetFailedIndexItemsCountAsync(status ?? "Failed");
        }

        public async Task RetryFailedItemsBatchAsync(IEnumerable<long> failedItemIds, CancellationToken cancellationToken = default)
        {
            var idList = failedItemIds.ToList();
            if (!idList.Any()) return;

            var allFailed = await _repository.GetFailedIndexItemsAsync(includeResolved: true);
            var targetItems = allFailed.Where(i => idList.Contains(i.Id)).ToList();

            foreach (var item in targetItems)
            {
                if (cancellationToken.IsCancellationRequested) break;

                await _repository.UpdateFailedIndexItemStatusAsync(item.Id, "Retrying");
                if (string.Equals(item.OperationType, "Delete", StringComparison.OrdinalIgnoreCase))
                {
                    await _indexerService.QueueItemForDeletionAsync(item.StoreId, item.EntryId);
                }
                else
                {
                    await _indexerService.QueueItemForIndexingAsync(item.StoreId, item.EntryId, item.FolderName);
                }
                await _repository.UpdateFailedIndexItemStatusAsync(item.Id, "Recovered");
            }
        }

        public async Task<int> RetryAllFailedItemsAsync(CancellationToken cancellationToken = default)
        {
            var allFailed = await _repository.GetFailedIndexItemsAsync(status: "Failed", maxCount: 5000);
            if (allFailed.Count == 0) return 0;
            await RetryFailedItemsBatchAsync(allFailed.Select(i => i.Id), cancellationToken);
            return allFailed.Count;
        }

        public async Task RemoveFailedItemsBatchAsync(IEnumerable<long> failedItemIds, CancellationToken cancellationToken = default)
        {
            foreach (var id in failedItemIds)
            {
                if (cancellationToken.IsCancellationRequested) break;
                await _repository.RemoveFailedIndexItemAsync(id);
            }
        }

        public Task RemoveAllFailedItemsAsync(CancellationToken cancellationToken = default)
        {
            return _repository.RemoveAllFailedIndexItemsAsync();
        }
    }
}
