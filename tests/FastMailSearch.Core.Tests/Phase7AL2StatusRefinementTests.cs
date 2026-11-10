using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using FastMailSearch.Core.Interfaces;
using FastMailSearch.Core.Models;
using FastMailSearch.Core.Search;
using FastMailSearch.Core.Services;
using FastMailSearch.Infrastructure.Database;
using FastMailSearch.Infrastructure.Indexing;
using FastMailSearch.Infrastructure.Outlook;
using FastMailSearch.Infrastructure.Services;
using FastMailSearch.UI.ViewModels;
using Xunit;

namespace FastMailSearch.Core.Tests
{
    public class Phase7AL2StatusRefinementTests : IDisposable
    {
        private readonly string _tempDbPath;
        private readonly DatabaseConnectionFactory _factory;
        private readonly SqliteEmailRepository _repository;

        public Phase7AL2StatusRefinementTests()
        {
            SearchViewModel.DispatcherOverride = action => action();
            _tempDbPath = Path.Combine(Path.GetTempPath(), $"fastmail_status_{Guid.NewGuid():N}.db");
            _factory = new DatabaseConnectionFactory(_tempDbPath);
            _repository = new SqliteEmailRepository(_factory);
            _repository.InitializeDatabase();
        }

        public void Dispose()
        {
            SearchViewModel.DispatcherOverride = null;
            try
            {
                if (File.Exists(_tempDbPath))
                {
                    File.Delete(_tempDbPath);
                }
            }
            catch { }
        }

        private async Task InvokePollLiveStatusAsync(SearchViewModel vm)
        {
            var pollMethod = typeof(SearchViewModel).GetMethod("PollLiveStatusAsync", BindingFlags.NonPublic | BindingFlags.Instance);
            Assert.NotNull(pollMethod);
            var task = (Task)pollMethod!.Invoke(vm, null)!;
            await task;
        }

        [Fact]
        public async Task Test01_ActiveWorker_With_PendingQueue_Reports_Updating()
        {
            // 1. Enqueue 10 pending items
            for (int i = 0; i < 10; i++)
            {
                await _repository.EnqueuePendingChangeAsync("store-1", $"entry-{i}", "InsertOrUpdate", "Inbox");
            }

            var mockIndexer = new ConfigurableMockIndexerService7AL
            {
                IsIndexing = true,
                CurrentStatus = IndexerStatus.Indexing,
                CurrentOperation = "Processing Pending Queue",
                PendingQueueCount = 10
            };

            var outlook = new MockOutlookAdapter();
            var healthService = new IndexHealthService(_repository, mockIndexer, outlook, _factory);
            var searchService = new SearchService(_repository, _repository, new MockLicensingService(LicenseStatus.Valid));
            var vm = new SearchViewModel(searchService, _repository, mockIndexer, outlook, new MockLicensingService(LicenseStatus.Valid), _factory);

            // IndexHealthService check
            var summary = await healthService.GetHealthSummaryAsync();
            Assert.Equal(IndexerHealth.Updating, summary.HealthState);
            Assert.Equal("Processing Pending Queue", summary.CurrentOperation);
            Assert.Equal(1, summary.CurrentlyProcessingCount);
            Assert.Equal(10, summary.PendingQueueCount);
            Assert.Contains("Updating", summary.StatusState);
            Assert.Contains("10 email(s) remaining", summary.StatusState);

            // SearchViewModel check
            await InvokePollLiveStatusAsync(vm);
            Assert.Contains("Updating", vm.StatusText);
            Assert.Contains("10 email(s) remaining", vm.StatusText);
        }

        [Fact]
        public async Task Test02_PendingQueue_With_InactiveWorker_Reports_QueuePaused()
        {
            // 1. Enqueue 3,429 pending items (reproducing the exact Stage 7AL bug scenario)
            for (int i = 0; i < 3429; i++)
            {
                await _repository.EnqueuePendingChangeAsync("store-1", $"entry-{i}", "InsertOrUpdate", "Inbox");
            }

            var mockIndexer = new ConfigurableMockIndexerService7AL
            {
                IsIndexing = false,
                CurrentStatus = IndexerStatus.Idle,
                CurrentOperation = "Idle",
                PendingQueueCount = 3429
            };

            var outlook = new MockOutlookAdapter();
            var healthService = new IndexHealthService(_repository, mockIndexer, outlook, _factory);
            var searchService = new SearchService(_repository, _repository, new MockLicensingService(LicenseStatus.Valid));
            var vm = new SearchViewModel(searchService, _repository, mockIndexer, outlook, new MockLicensingService(LicenseStatus.Valid), _factory);

            // IndexHealthService check
            var summary = await healthService.GetHealthSummaryAsync();
            Assert.Equal(IndexerHealth.Paused, summary.HealthState);
            Assert.Equal("Idle — Queue Paused", summary.CurrentOperation);
            Assert.Equal(0, summary.CurrentlyProcessingCount);
            Assert.Equal(3429, summary.PendingQueueCount);
            Assert.DoesNotContain("Updating", summary.StatusState);
            Assert.Contains("Idle — Queue Paused", summary.StatusState);
            Assert.Contains("3,429 pending", summary.StatusState);

            // SearchViewModel check
            await InvokePollLiveStatusAsync(vm);
            Assert.DoesNotContain("Updating", vm.StatusText);
            Assert.Contains("Idle — Queue Paused", vm.StatusText);
            Assert.Contains("3,429 pending", vm.StatusText);
        }

        [Fact]
        public async Task Test03_EmptyQueue_With_InactiveWorker_Reports_UpToDate()
        {
            // 1. Save 50 dummy indexed emails
            var emails = new List<EmailDocument>();
            for (int i = 0; i < 50; i++)
            {
                emails.Add(new EmailDocument
                {
                    StoreId = "store-1",
                    EntryId = $"entry-{i}",
                    Subject = $"Subject {i}",
                    BodyText = $"Body {i}",
                    SenderEmail = "test@example.com",
                    FolderName = "Inbox",
                    ReceivedDate = DateTime.Now
                });
            }
            await _repository.SaveEmailBatchAsync(emails);

            var mockIndexer = new ConfigurableMockIndexerService7AL
            {
                IsIndexing = false,
                CurrentStatus = IndexerStatus.Idle,
                CurrentOperation = "Idle",
                PendingQueueCount = 0
            };

            var outlook = new MockOutlookAdapter();
            var healthService = new IndexHealthService(_repository, mockIndexer, outlook, _factory);
            var searchService = new SearchService(_repository, _repository, new MockLicensingService(LicenseStatus.Valid));
            var vm = new SearchViewModel(searchService, _repository, mockIndexer, outlook, new MockLicensingService(LicenseStatus.Valid), _factory);

            // IndexHealthService check
            var summary = await healthService.GetHealthSummaryAsync();
            Assert.Equal(IndexerHealth.Healthy, summary.HealthState);
            Assert.Equal("Idle", summary.CurrentOperation);
            Assert.Equal(0, summary.CurrentlyProcessingCount);
            Assert.Equal(0, summary.PendingQueueCount);
            Assert.Contains("Ready", summary.StatusState);

            // SearchViewModel check
            await InvokePollLiveStatusAsync(vm);
            Assert.Equal("🟢 Up to date (50 emails indexed)", vm.StatusText);
        }

        [Fact]
        public async Task Test04_PersistedProcessingRows_Alone_DoNotCountAsActiveWorker()
        {
            // 1. Simulate process crash leaving 25 rows stuck with Status = 'Processing'
            using (var conn = _factory.CreateConnection())
            using (var tx = conn.BeginTransaction())
            {
                for (int i = 0; i < 25; i++)
                {
                    using var cmd = conn.CreateCommand();
                    cmd.Transaction = tx;
                    cmd.CommandText = @"
                        INSERT INTO PendingChanges (StoreId, EntryId, Operation, Status, AttemptCount, CreatedAt, UpdatedAt)
                        VALUES ('store-1', @EntryId, 'InsertOrUpdate', 'Processing', 1, @Now, @Now);";
                    cmd.Parameters.AddWithValue("@EntryId", $"crash-entry-{i}");
                    cmd.Parameters.AddWithValue("@Now", DateTime.Now.ToString("o"));
                    cmd.ExecuteNonQuery();
                }
                tx.Commit();
            }

            var mockIndexer = new ConfigurableMockIndexerService7AL
            {
                IsIndexing = false,
                CurrentStatus = IndexerStatus.Idle,
                CurrentOperation = "Idle",
                PendingQueueCount = 25
            };

            var outlook = new MockOutlookAdapter();
            var healthService = new IndexHealthService(_repository, mockIndexer, outlook, _factory);
            var searchService = new SearchService(_repository, _repository, new MockLicensingService(LicenseStatus.Valid));
            var vm = new SearchViewModel(searchService, _repository, mockIndexer, outlook, new MockLicensingService(LicenseStatus.Valid), _factory);

            // IndexHealthService check: CurrentlyProcessingCount MUST BE 0 (DB 'Processing' rows alone do not count!)
            var summary = await healthService.GetHealthSummaryAsync();
            Assert.Equal(0, summary.CurrentlyProcessingCount);
            Assert.Equal(25, summary.PendingQueueCount);
            Assert.Equal(IndexerHealth.Paused, summary.HealthState);
            Assert.DoesNotContain("Updating", summary.StatusState);
            Assert.Contains("Idle — Queue Paused (25 pending)", summary.StatusState);

            // SearchViewModel check
            await InvokePollLiveStatusAsync(vm);
            Assert.DoesNotContain("Updating", vm.StatusText);
            Assert.Contains("Idle — Queue Paused (25 pending)", vm.StatusText);
        }

        [Fact]
        public async Task Test05_CrossProcessHeartbeat_Fresh_Vs_Stale()
        {
            var mockIndexer = new ConfigurableMockIndexerService7AL
            {
                IsIndexing = false,
                CurrentStatus = IndexerStatus.Idle,
                PendingQueueCount = 50
            };
            var outlook = new MockOutlookAdapter();
            var healthService = new IndexHealthService(_repository, mockIndexer, outlook, _factory);

            // 1. Enqueue 50 items so the database has a pending queue
            for (int i = 0; i < 50; i++)
            {
                await _repository.EnqueuePendingChangeAsync("store-1", $"entry-{i}", "InsertOrUpdate", "Inbox");
            }

            // 2. Fresh heartbeat (< 3 minutes) simulating OUTLOOK.EXE backfill in progress
            await _repository.SaveLiveProgressAsync(
                "Indexing",
                "Backfilling Active Store",
                100,
                500,
                "Inbox",
                100,
                500,
                100,
                0,
                0,
                50
            );

            var freshSummary = await healthService.GetHealthSummaryAsync();
            Assert.Equal(IndexerHealth.Updating, freshSummary.HealthState);
            Assert.Equal(1, freshSummary.CurrentlyProcessingCount);
            Assert.Contains("Backfill in Progress", freshSummary.StatusState);
            Assert.Contains("400 remaining", freshSummary.StatusState);

            // 2. Stale heartbeat (10 minutes old) simulating killed/abandoned OUTLOOK.EXE process
            using (var conn = _factory.CreateConnection())
            using (var cmd = conn.CreateCommand())
            {
                string staleTime = DateTime.Now.AddMinutes(-10).ToString("o");
                cmd.CommandText = @"
                    UPDATE GlobalState SET Value = @StaleTime, UpdatedAt = @StaleTime WHERE Key = 'LastSync';";
                cmd.Parameters.AddWithValue("@StaleTime", staleTime);
                cmd.ExecuteNonQuery();
            }

            var staleSummary = await healthService.GetHealthSummaryAsync();
            Assert.Equal(0, staleSummary.CurrentlyProcessingCount);
            Assert.NotEqual(IndexerHealth.Updating, staleSummary.HealthState);
            Assert.Equal(IndexerHealth.Paused, staleSummary.HealthState);
            Assert.Contains("Idle — Queue Paused", staleSummary.StatusState);
        }

        [Fact]
        public async Task Test06_IndexingFailure_Reports_ErrorState()
        {
            var mockIndexer = new ConfigurableMockIndexerService7AL
            {
                IsIndexing = false,
                CurrentStatus = IndexerStatus.Error,
                LastError = "Disk I/O failure during commit"
            };

            var outlook = new MockOutlookAdapter();
            var healthService = new IndexHealthService(_repository, mockIndexer, outlook, _factory);
            var searchService = new SearchService(_repository, _repository, new MockLicensingService(LicenseStatus.Valid));
            var vm = new SearchViewModel(searchService, _repository, mockIndexer, outlook, new MockLicensingService(LicenseStatus.Valid), _factory);

            // IndexHealthService check
            var summary = await healthService.GetHealthSummaryAsync();
            Assert.Equal(IndexerHealth.Error, summary.HealthState);
            Assert.Contains("🔴 Indexer Error", summary.StatusState);

            // SearchViewModel check via progress update
            mockIndexer.RaiseProgressChanged(new IndexerProgressEventArgs
            {
                Status = IndexerStatus.Error,
                LastError = "Disk I/O failure during commit"
            });

            Assert.Equal("🔴 Index error: Disk I/O failure during commit", vm.StatusText);
        }

        [Fact]
        public async Task Test07_StatusRefresh_AfterRestart_Lifecycle()
        {
            // 1. Enqueue 50 items in DB
            for (int i = 0; i < 50; i++)
            {
                await _repository.EnqueuePendingChangeAsync("store-1", $"entry-{i}", "InsertOrUpdate", "Inbox");
            }

            // 2. Instantiate real EmailIndexerService
            var outlook = new MockOutlookAdapter();
            var licensing = new MockLicensingService(LicenseStatus.Valid);
            var realIndexer = new EmailIndexerService(outlook, _repository, licensing);

            // 3. Simulate startup recovery with startProcessing = false
            await realIndexer.CheckAndRecoverIncompleteIndexAsync(startProcessing: false);

            // Verify indexer initial status message does NOT say Updating
            var initialIndexerStatus = realIndexer.GetStatus();
            Assert.Equal(IndexerStatus.Idle, initialIndexerStatus.Status);
            Assert.Contains("Idle — Queue Paused (50 pending)", initialIndexerStatus.CurrentStatusMessage);
            Assert.DoesNotContain("Updating", initialIndexerStatus.CurrentStatusMessage);

            // 4. Instantiate SearchViewModel
            var searchService = new SearchService(_repository, _repository, licensing);
            var vm = new SearchViewModel(searchService, _repository, realIndexer, outlook, licensing, _factory);

            // 5. Initial status before poll must be stable and never show Updating
            Assert.DoesNotContain("Updating", vm.StatusText);

            // 6. Execute poll
            await InvokePollLiveStatusAsync(vm);
            Assert.DoesNotContain("Updating", vm.StatusText);
            Assert.Contains("Idle — Queue Paused (50 pending)", vm.StatusText);
        }

        [Fact]
        public async Task Test08_Consistency_Between_SearchViewModel_And_IndexHealthService()
        {
            var outlook = new MockOutlookAdapter();
            var licensing = new MockLicensingService(LicenseStatus.Valid);
            var searchService = new SearchService(_repository, _repository, licensing);

            // State A: Idle with dormant queue
            await _repository.EnqueuePendingChangeAsync("store-1", "entry-1", "InsertOrUpdate", "Inbox");
            var mockIndexerA = new ConfigurableMockIndexerService7AL { IsIndexing = false, CurrentStatus = IndexerStatus.Idle, PendingQueueCount = 1 };
            var healthA = new IndexHealthService(_repository, mockIndexerA, outlook, _factory);
            var vmA = new SearchViewModel(searchService, _repository, mockIndexerA, outlook, licensing, _factory);
            var summaryA = await healthA.GetHealthSummaryAsync();
            await InvokePollLiveStatusAsync(vmA);

            Assert.Equal(summaryA.StatusState, vmA.StatusText);
            Assert.Contains("Idle — Queue Paused", vmA.StatusText);

            // State B: Active worker with queue
            var mockIndexerB = new ConfigurableMockIndexerService7AL
            {
                IsIndexing = true,
                CurrentStatus = IndexerStatus.Indexing,
                CurrentOperation = "Processing Pending Queue",
                PendingQueueCount = 1
            };
            var healthB = new IndexHealthService(_repository, mockIndexerB, outlook, _factory);
            var vmB = new SearchViewModel(searchService, _repository, mockIndexerB, outlook, licensing, _factory);
            var summaryB = await healthB.GetHealthSummaryAsync();
            await InvokePollLiveStatusAsync(vmB);

            Assert.Equal(summaryB.StatusState, vmB.StatusText);
            Assert.Contains("Updating", vmB.StatusText);
        }
    }

    internal class ConfigurableMockIndexerService7AL : IIndexerService
    {
        public bool IsIndexing { get; set; } = false;
        public bool IsRebuilding { get; set; } = false;
        public IndexerStatus CurrentStatus { get; set; } = IndexerStatus.Idle;
        public IndexMetadata CurrentMetadata { get; set; } = new IndexMetadata();
        public int ProcessedItems { get; set; } = 0;
        public int TotalItems { get; set; } = 0;
        public int FolderProcessedItems { get; set; } = 0;
        public int FolderTotalItems { get; set; } = 0;
        public string CurrentFolder { get; set; } = string.Empty;
        public DateTime? StartedAt { get; set; } = null;
        public DateTime? CompletedAt { get; set; } = null;
        public DateTime? LastIndexedItemTime { get; set; } = null;
        public string CurrentOperation { get; set; } = "Idle";
        public string LastError { get; set; } = string.Empty;
        public int PendingQueueCount { get; set; } = 0;

        public event EventHandler<IndexerProgressEventArgs>? ProgressChanged;

        public void RaiseProgressChanged(IndexerProgressEventArgs args)
        {
            ProgressChanged?.Invoke(this, args);
        }

        public Task StartInitialIndexingAsync(CancellationToken cancellationToken = default, int? maxItems = null) => Task.CompletedTask;
        public Task PauseIndexingAsync() => Task.CompletedTask;
        public Task ResumeIndexingAsync() => Task.CompletedTask;
        public Task<BackfillSummary> ResumeIncrementalCatchUpAsync(CancellationToken cancellationToken = default) => Task.FromResult(new BackfillSummary());
        public Task RebuildIndexAsync(CancellationToken cancellationToken = default) => Task.CompletedTask;
        public Task CancelIndexingAsync() => Task.CompletedTask;
        public Task CheckAndRecoverIncompleteIndexAsync(bool startProcessing = true) => Task.CompletedTask;
        public Task QueueItemForIndexingAsync(string storeId, string entryId, string folderEntryId) => Task.CompletedTask;
        public Task QueueItemForDeletionAsync(string storeId, string entryId) => Task.CompletedTask;
        public Task<IndexIntegrityResult> ScanIndexIntegrityAsync(CancellationToken cancellationToken = default) => Task.FromResult(new IndexIntegrityResult());
        public Task RepairIndexIntegrityAsync(IndexIntegrityResult scanResult, CancellationToken cancellationToken = default) => Task.CompletedTask;
        public Task<BackfillSummary> StartActiveStoreBackfillAsync(CancellationToken cancellationToken = default, int? maxItems = null) => Task.FromResult(new BackfillSummary());
        public Task<StoreCoverageInfo> CheckStoreCoverageAsync(CancellationToken cancellationToken = default) => Task.FromResult(new StoreCoverageInfo());
        public IndexerProgressEventArgs GetStatus() => new IndexerProgressEventArgs
        {
            Status = CurrentStatus,
            PendingQueueCount = PendingQueueCount,
            CurrentStatusMessage = CurrentStatus == IndexerStatus.Indexing 
                ? $"🔵 Updating • {PendingQueueCount:N0} email(s) remaining" 
                : $"⏸ Idle — Queue Paused ({PendingQueueCount:N0} pending)"
        };
        public void SetSafeMode(string reason) { }
    }
}
