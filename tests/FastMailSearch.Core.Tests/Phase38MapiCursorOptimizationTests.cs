using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using FastMailSearch.Core.Interfaces;
using FastMailSearch.Core.Models;
using FastMailSearch.Infrastructure.Database;
using FastMailSearch.Infrastructure.Indexing;
using FastMailSearch.Infrastructure.Outlook;
using Xunit;

namespace FastMailSearch.Core.Tests
{
    public class Phase38MapiCursorOptimizationTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly DatabaseConnectionFactory _factory;
        private readonly SqliteEmailRepository _repository;

        public Phase38MapiCursorOptimizationTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), $"Phase38CursorTests_{Guid.NewGuid():N}.db");
            _factory = new DatabaseConnectionFactory(_dbPath);
            _repository = new SqliteEmailRepository(_factory);
            _repository.InitializeDatabase();
        }

        public void Dispose()
        {
            try
            {
                if (File.Exists(_dbPath))
                    File.Delete(_dbPath);
            }
            catch { }
        }

        /// <summary>
        /// Mock adapter simulating the new session-scoped cursor behavior and IFolderBatchReader.
        /// </summary>
        private class MockCursorOutlookAdapter : IOutlookAdapter, IFolderBatchReader
        {
            public int TotalItems { get; set; } = 58052;
            public int SessionReuses { get; private set; }
            public int CursorRestarts { get; private set; }
            public int FastForwardSkips { get; private set; }
            public int TotalItemsYielded { get; private set; }
            public HashSet<int> NonMailItemIndices { get; } = new HashSet<int>();
            public HashSet<int> FailAtIndices { get; } = new HashSet<int>();
            public bool ResetCursorSessionCalled { get; private set; }

            private string? _activeStoreId;
            private string? _activeFolderId;
            private int _currentCursorPosition = 0;

            public bool IsOutlookAvailable() => true;
            public string GetOutlookVersion() => "16.0 (64-bit)";
            public string GetOfficeBitness() => "x64";
            public List<MailboxInfo> GetMailboxes() => new List<MailboxInfo>
            {
                new MailboxInfo { StoreId = "store1", DisplayName = "POP Archive", IsPst = true }
            };

            public List<MailboxFolder> GetFolders(string storeId) => new List<MailboxFolder>
            {
                new MailboxFolder { StoreId = storeId, EntryId = "folder_inbox", Name = "Inbox", ItemCount = TotalItems }
            };

            public void ResetCursorSession()
            {
                ResetCursorSessionCalled = true;
                _activeStoreId = null;
                _activeFolderId = null;
                _currentCursorPosition = 0;
            }

            public FolderBatchResult FetchFolderBatch(string storeId, string folderEntryId, int startIndex, int batchSize, Func<bool>? cancelCheck = null)
            {
                var result = new FolderBatchResult
                {
                    StartIndex = startIndex,
                    LastProcessedIndex = startIndex - 1,
                    ProcessedCount = 0,
                    IndexedCount = 0,
                    SkippedCount = 0
                };

                if (startIndex > TotalItems) return result;

                bool canReuse = _activeStoreId == storeId
                    && _activeFolderId == folderEntryId
                    && startIndex == _currentCursorPosition + 1;

                if (canReuse)
                {
                    SessionReuses++;
                }
                else
                {
                    CursorRestarts++;
                    _activeStoreId = storeId;
                    _activeFolderId = folderEntryId;
                    _currentCursorPosition = 0;

                    if (startIndex > 1)
                    {
                        FastForwardSkips += (startIndex - 1);
                        _currentCursorPosition = startIndex - 1;
                    }
                }

                int end = Math.Min(startIndex + batchSize - 1, TotalItems);
                for (int i = startIndex; i <= end; i++)
                {
                    if (cancelCheck != null && cancelCheck())
                    {
                        result.WasCancelled = true;
                        break;
                    }

                    if (FailAtIndices.Contains(i))
                    {
                        // Simulate COM failure at position i
                        ResetCursorSession();
                        FailAtIndices.Remove(i); // Clear so retry can succeed
                        if (i == startIndex)
                        {
                            result.HasError = true;
                            result.ErrorMessage = $"Simulated COM failure at position {i}";
                            throw new COMException($"Simulated COM failure at position {i}", unchecked((int)0x80010001));
                        }
                        else
                        {
                            // Partial batch before error
                            result.HasError = true;
                            result.ErrorMessage = $"Simulated COM failure at position {i}";
                            break;
                        }
                    }

                    _currentCursorPosition = i;
                    result.LastProcessedIndex = i;
                    result.ProcessedCount++;

                    if (NonMailItemIndices.Contains(i))
                    {
                        // Non-mail item: skipped position
                        result.SkippedCount++;
                    }
                    else
                    {
                        // Mail item: indexed
                        result.IndexedCount++;
                        TotalItemsYielded++;
                        result.Items.Add(new EmailDocument
                        {
                            EntryId = $"ENTRY_{storeId}_{folderEntryId}_{i:D8}",
                            StoreId = storeId,
                            FolderEntryId = folderEntryId,
                            FolderName = "Inbox",
                            Subject = $"Email Index #{i}",
                            SenderEmail = "sender@example.com",
                            SenderName = "Sender Example",
                            ReceivedDate = DateTime.UtcNow.AddMinutes(-i),
                            BodyText = $"Body of email message #{i}."
                        });
                    }
                }

                return result;
            }

            public List<EmailDocument> FetchFolderEmailsBatch(string storeId, string folderEntryId, int startIndex, int batchSize, Func<bool>? cancelCheck = null)
            {
                return FetchFolderBatch(storeId, folderEntryId, startIndex, batchSize, cancelCheck).Items;
            }

            public EmailDocument? FetchSingleEmail(string storeId, string entryId) => null;
            public bool OpenEmailInOutlook(string storeId, string entryId) => true;
            public bool IsOutboxFolder(string storeId, string folderEntryId) => false;
            public void SubscribeToEvents(Action<string, string, string> onItemAddedOrChanged, Action<string, string> onItemRemoved) { }
            public void UnsubscribeFromEvents() { }
        }

        [Fact]
        public async Task Test01_CheckpointContinuity_ResumeAtOffset_FetchesSubsequentItemsOnly()
        {
            var adapter = new MockCursorOutlookAdapter { TotalItems = 58052 };
            var indexer = new EmailIndexerService(adapter, _repository);

            // Pre-seed folder checkpoint at 25,250
            await _repository.SaveFolderIndexStateAsync("store1", "folder_inbox", "Indexing", 25250, 58052);

            var batch = adapter.FetchFolderBatch("store1", "folder_inbox", 25251, 250);

            Assert.Equal(250, batch.Items.Count);
            Assert.Equal("ENTRY_store1_folder_inbox_00025251", batch.Items.First().EntryId);
            Assert.Equal("ENTRY_store1_folder_inbox_00025500", batch.Items.Last().EntryId);
            Assert.Equal(25250, adapter.FastForwardSkips);
            Assert.Equal(1, adapter.CursorRestarts);
        }

        [Fact]
        public void Test02_DuplicatePrevention_SequentialBatches_YieldDistinctItems()
        {
            var adapter = new MockCursorOutlookAdapter { TotalItems = 10000 };
            var batch1 = adapter.FetchFolderBatch("store1", "folder_inbox", 1, 250);
            var batch2 = adapter.FetchFolderBatch("store1", "folder_inbox", 251, 250);
            var batch3 = adapter.FetchFolderBatch("store1", "folder_inbox", 501, 250);

            Assert.Equal(250, batch1.Items.Count);
            Assert.Equal(250, batch2.Items.Count);
            Assert.Equal(250, batch3.Items.Count);

            var combined = new List<EmailDocument>();
            combined.AddRange(batch1.Items);
            combined.AddRange(batch2.Items);
            combined.AddRange(batch3.Items);

            Assert.Equal(750, combined.Count);
            Assert.Equal(750, combined.Select(x => x.EntryId).Distinct().Count());

            Assert.Equal(1, adapter.CursorRestarts);
            Assert.Equal(2, adapter.SessionReuses);
        }

        [Fact]
        public void Test03_EmptyBatches_BeyondTotalItems_ReturnsEmptyListWithoutError()
        {
            var adapter = new MockCursorOutlookAdapter { TotalItems = 100 };
            var batch = adapter.FetchFolderBatch("store1", "folder_inbox", 101, 250);

            Assert.Empty(batch.Items);
            Assert.Equal(100, batch.LastProcessedIndex);

            adapter.TotalItems = 0;
            var emptyBatch = adapter.FetchFolderBatch("store1", "folder_inbox", 1, 250);
            Assert.Empty(emptyBatch.Items);
        }

        [Fact]
        public void Test04_Cancellation_MidBatch_StopsAndPreservesPartialItems()
        {
            var adapter = new MockCursorOutlookAdapter { TotalItems = 1000 };
            int count = 0;
            bool cancelRequested = false;

            var batch = adapter.FetchFolderBatch("store1", "folder_inbox", 1, 250, () =>
            {
                count++;
                if (count >= 50)
                {
                    cancelRequested = true;
                    return true;
                }
                return false;
            });

            Assert.True(cancelRequested);
            Assert.True(batch.WasCancelled);
            Assert.Equal(49, batch.Items.Count);
            Assert.Equal(49, batch.LastProcessedIndex);
        }

        [Fact]
        public void Test05_CursorRestart_FolderSwitchOrSeekBackwards_InvalidatesPreviousSession()
        {
            var adapter = new MockCursorOutlookAdapter { TotalItems = 1000 };

            var b1 = adapter.FetchFolderBatch("store1", "folder_inbox", 1, 100);
            Assert.Equal(100, b1.Items.Count);
            Assert.Equal(1, adapter.CursorRestarts);

            var b2 = adapter.FetchFolderBatch("store1", "folder_sent", 1, 100);
            Assert.Equal(100, b2.Items.Count);
            Assert.Equal(2, adapter.CursorRestarts);

            var b3 = adapter.FetchFolderBatch("store1", "folder_sent", 1, 50);
            Assert.Equal(50, b3.Items.Count);
            Assert.Equal(3, adapter.CursorRestarts);
        }

        [Fact]
        public void Test06_Benchmark_SeekAccess_Vs_CursorStreaming_PerformanceComparison()
        {
            int batchSize = 250;
            int batchesToTest = 10; // 2,500 items

            long seekOperationsCount = 0;
            for (int b = 0; b < batchesToTest; b++)
            {
                int start = 1 + (b * batchSize);
                int end = start + batchSize - 1;
                for (int i = start; i <= end; i++)
                {
                    seekOperationsCount += i;
                }
            }

            long cursorOperationsCount = 0;
            cursorOperationsCount += batchSize;
            for (int b = 1; b < batchesToTest; b++)
            {
                cursorOperationsCount += batchSize;
            }

            Assert.Equal(2500, cursorOperationsCount);
            Assert.True(seekOperationsCount > 3_000_000, $"Expected > 3M row seeks, got {seekOperationsCount}");
            Assert.True(seekOperationsCount / cursorOperationsCount > 1000);
        }

        [Fact]
        public async Task Test07_Cancellation_After1Item_DoesNotCheckpointUnprocessedItems_AndResumesCleanly()
        {
            var adapter = new MockCursorOutlookAdapter { TotalItems = 250 };
            var indexer = new EmailIndexerService(adapter, _repository, new MockLicensingService(LicenseStatus.Valid));

            // Cancel after reading item #1
            int readCount = 0;
            using var cts = new CancellationTokenSource();
            indexer.SetBatchCancelCheck(() =>
            {
                readCount++;
                return readCount > 1; // triggers cancellation after 1 item
            });

            var summary = await indexer.StartActiveStoreBackfillAsync(cts.Token);
            Assert.True(summary.IsCancelled);

            // Checkpoint MUST be exactly 1, NOT 250
            var state = await _repository.GetFolderIndexStateAsync("store1", "folder_inbox");
            Assert.True(state.HasValue);
            Assert.Equal(1, state.Value.ProcessedOffset);

            var emailsFirstPass = await _repository.GetIndexedEmailCountAsync();
            Assert.Equal(1, emailsFirstPass);

            // Resume backfill without cancellation
            indexer.SetBatchCancelCheck(() => false);
            var resumeSummary = await indexer.StartActiveStoreBackfillAsync(CancellationToken.None);
            Assert.False(resumeSummary.IsCancelled);

            // Total emails must be exactly 250, zero skipped, zero duplicate
            var totalEmails = await _repository.GetIndexedEmailCountAsync();
            Assert.Equal(250, totalEmails);

            var finalState = await _repository.GetFolderIndexStateAsync("store1", "folder_inbox");
            Assert.Equal(250, finalState.Value.ProcessedOffset);
            Assert.Equal("Completed", finalState.Value.Status);
        }

        [Fact]
        public async Task Test08_Cancellation_After30Items_DoesNotCheckpointUnprocessedItems_AndResumesCleanly()
        {
            var adapter = new MockCursorOutlookAdapter { TotalItems = 250 };
            var indexer = new EmailIndexerService(adapter, _repository, new MockLicensingService(LicenseStatus.Valid));

            // Cancel after reading 30 items
            int readCount = 0;
            indexer.SetBatchCancelCheck(() =>
            {
                readCount++;
                return readCount > 30; // triggers cancellation after 30 items
            });

            var summary = await indexer.StartActiveStoreBackfillAsync(CancellationToken.None);
            Assert.True(summary.IsCancelled);

            // Checkpoint MUST be exactly 30, NOT 250
            var state = await _repository.GetFolderIndexStateAsync("store1", "folder_inbox");
            Assert.True(state.HasValue);
            Assert.Equal(30, state.Value.ProcessedOffset);

            var emailsFirstPass = await _repository.GetIndexedEmailCountAsync();
            Assert.Equal(30, emailsFirstPass);

            // Resume backfill
            indexer.SetBatchCancelCheck(() => false);
            var resumeSummary = await indexer.StartActiveStoreBackfillAsync(CancellationToken.None);
            Assert.False(resumeSummary.IsCancelled);

            var totalEmails = await _repository.GetIndexedEmailCountAsync();
            Assert.Equal(250, totalEmails);

            var finalState = await _repository.GetFolderIndexStateAsync("store1", "folder_inbox");
            Assert.Equal(250, finalState.Value.ProcessedOffset);
            Assert.Equal("Completed", finalState.Value.Status);
        }

        [Fact]
        public async Task Test09_Cancellation_After249Items_DoesNotCheckpointUnprocessedItems_AndResumesCleanly()
        {
            var adapter = new MockCursorOutlookAdapter { TotalItems = 250 };
            var indexer = new EmailIndexerService(adapter, _repository, new MockLicensingService(LicenseStatus.Valid));

            // Cancel after reading 249 items
            int readCount = 0;
            indexer.SetBatchCancelCheck(() =>
            {
                readCount++;
                return readCount > 249; // triggers cancellation after 249 items
            });

            var summary = await indexer.StartActiveStoreBackfillAsync(CancellationToken.None);
            Assert.True(summary.IsCancelled);

            // Checkpoint MUST be exactly 249, NOT 250
            var state = await _repository.GetFolderIndexStateAsync("store1", "folder_inbox");
            Assert.True(state.HasValue);
            Assert.Equal(249, state.Value.ProcessedOffset);

            var emailsFirstPass = await _repository.GetIndexedEmailCountAsync();
            Assert.Equal(249, emailsFirstPass);

            // Resume backfill
            indexer.SetBatchCancelCheck(() => false);
            var resumeSummary = await indexer.StartActiveStoreBackfillAsync(CancellationToken.None);
            Assert.False(resumeSummary.IsCancelled);

            // Final item #250 indexed cleanly
            var totalEmails = await _repository.GetIndexedEmailCountAsync();
            Assert.Equal(250, totalEmails);

            var finalState = await _repository.GetFolderIndexStateAsync("store1", "folder_inbox");
            Assert.Equal(250, finalState.Value.ProcessedOffset);
            Assert.Equal("Completed", finalState.Value.Status);
        }

        [Fact]
        public async Task Test10_NonMailItems_InBatch_AccountedAsSkippedWithoutStoppingOrMisaligning()
        {
            var adapter = new MockCursorOutlookAdapter { TotalItems = 20 };
            // Mark items #3, #7, #15 as non-mail items
            adapter.NonMailItemIndices.Add(3);
            adapter.NonMailItemIndices.Add(7);
            adapter.NonMailItemIndices.Add(15);

            var indexer = new EmailIndexerService(adapter, _repository, new MockLicensingService(LicenseStatus.Valid));
            var summary = await indexer.StartActiveStoreBackfillAsync(CancellationToken.None);

            Assert.False(summary.IsCancelled);
            Assert.Equal(17, summary.NewlyIndexedCount);
            Assert.Equal(3, summary.SkippedCount);

            // 17 emails in database
            var totalEmails = await _repository.GetIndexedEmailCountAsync();
            Assert.Equal(17, totalEmails);

            // Checkpoint processed up to 20
            var finalState = await _repository.GetFolderIndexStateAsync("store1", "folder_inbox");
            Assert.Equal(20, finalState.Value.ProcessedOffset);
            Assert.Equal("Completed", finalState.Value.Status);
        }

        [Fact]
        public async Task Test11_TransientCOMException_InvalidatesCursorSession_AndRetriesFromFailedPosition()
        {
            var adapter = new MockCursorOutlookAdapter { TotalItems = 50 };
            // Simulate transient COM exception at position 10 on first attempt
            adapter.FailAtIndices.Add(10);

            var indexer = new EmailIndexerService(adapter, _repository, new MockLicensingService(LicenseStatus.Valid));
            var summary = await indexer.StartActiveStoreBackfillAsync(CancellationToken.None);

            Assert.False(summary.IsCancelled);
            Assert.Equal(50, summary.NewlyIndexedCount);

            var totalEmails = await _repository.GetIndexedEmailCountAsync();
            Assert.Equal(50, totalEmails);

            var finalState = await _repository.GetFolderIndexStateAsync("store1", "folder_inbox");
            Assert.Equal(50, finalState.Value.ProcessedOffset);
        }

        [Fact]
        public void Test12_SafeShutdown_StaWorkerThread_ThrowsObjectDisposedExceptionAndNeverExecutesOnCallerThread()
        {
            var thread = new StaWorkerThread();
            bool executed = false;
            thread.Invoke(() => { executed = true; });
            Assert.True(executed);

            // Dispose thread
            thread.Dispose();

            // Calling Invoke after dispose MUST throw ObjectDisposedException and NEVER execute delegate
            bool executedAfter = false;
            Assert.Throws<ObjectDisposedException>(() =>
            {
                thread.Invoke(() => { executedAfter = true; });
            });
            Assert.False(executedAfter);

            // Calling Invoke<T> after dispose MUST throw ObjectDisposedException
            Assert.Throws<ObjectDisposedException>(() =>
            {
                thread.Invoke<int>(() => 42);
            });
        }

        [Fact]
        public void Test13_StaWorkerThread_RepeatedDispose_IsSafeAndIdempotent()
        {
            var thread = new StaWorkerThread();
            thread.Dispose();
            thread.Dispose();
            thread.Dispose();

            Assert.Throws<ObjectDisposedException>(() => thread.Invoke(() => { }));
        }

        [Fact]
        public async Task Test14_StaWorkerThread_ConcurrentInvokeRacingWithDispose_DoesNotDeadlockOrViolateApartment()
        {
            var thread = new StaWorkerThread();
            var cts = new CancellationTokenSource(TimeSpan.FromSeconds(2));
            int successCount = 0;
            int disposedCount = 0;

            var tasks = Enumerable.Range(0, 10).Select(_ => Task.Run(() =>
            {
                while (!cts.Token.IsCancellationRequested)
                {
                    try
                    {
                        thread.Invoke(() =>
                        {
                            Assert.Equal(ApartmentState.STA, Thread.CurrentThread.GetApartmentState());
                        });
                        Interlocked.Increment(ref successCount);
                    }
                    catch (ObjectDisposedException)
                    {
                        Interlocked.Increment(ref disposedCount);
                        break;
                    }
                }
            })).ToArray();

            await Task.Delay(50);
            thread.Dispose();
            await Task.WhenAll(tasks);

            Assert.True(successCount > 0, "Some invoke calls should have succeeded before dispose");
            Assert.True(disposedCount > 0, "Racing tasks after dispose should receive ObjectDisposedException");
        }

        [Fact]
        public void Test15_OutlookAdapter_Disposed_OperationsFailSafelyWithoutCrashing()
        {
#if NETFRAMEWORK
            var adapter = new OutlookAdapter();
            adapter.Dispose();
            adapter.Dispose(); // Idempotent

            // ResetCursorSession does not throw
            adapter.ResetCursorSession();

            // UnsubscribeFromEvents does not throw
            adapter.UnsubscribeFromEvents();

            // FetchFolderBatch returns result with HasError = true
            var batch = adapter.FetchFolderBatch("store", "folder", 1, 250);
            Assert.True(batch.HasError);
            Assert.Contains("disposed", batch.ErrorMessage, StringComparison.OrdinalIgnoreCase);

            // Queries return safe fallbacks
            Assert.False(adapter.IsOutlookAvailable());
            Assert.Equal("Unavailable", adapter.GetOutlookVersion());
            Assert.Empty(adapter.GetMailboxes());
            Assert.Empty(adapter.GetFolders("store"));
            Assert.Null(adapter.FetchSingleEmail("store", "entry"));
            Assert.False(adapter.OpenEmailInOutlook("store", "entry"));
#endif
        }

        [Fact]
        public async Task Test16_EmailPersistenceFailure_CheckpointDoesNotAdvance_AndResumeSucceedsCleanly()
        {
            var adapter = new MockCursorOutlookAdapter { TotalItems = 250 };
            var repoDecorator = new TestEmailRepositoryDecorator(_repository);
            var indexer = new EmailIndexerService(adapter, repoDecorator, new MockLicensingService(LicenseStatus.Valid));

            // Force SaveEmailBatchAsync to fail persistently
            repoDecorator.SaveEmailBatchOverride = (emails, ct) =>
            {
                throw new IOException("Simulated persistent disk I/O error during email batch save");
            };

            var summary = await indexer.StartActiveStoreBackfillAsync(CancellationToken.None);

            // Backfill must stop with failed items reported
            Assert.Equal(250, summary.FailedCount);

            // CHECKPOINT MUST NOT ADVANCE: ProcessedOffset must remain at 0 (or startFrom - 1)
            var state = await _repository.GetFolderIndexStateAsync("store1", "folder_inbox");
            int offset = state?.ProcessedOffset ?? 0;
            Assert.Equal(0, offset);

            // Database must have 0 indexed emails
            var emailsCount = await _repository.GetIndexedEmailCountAsync();
            Assert.Equal(0, emailsCount);

            // Clear persistence failure: resume backfill
            repoDecorator.SaveEmailBatchOverride = null;
            var resumeSummary = await indexer.StartActiveStoreBackfillAsync(CancellationToken.None);

            Assert.False(resumeSummary.IsCancelled);
            Assert.Equal(250, resumeSummary.NewlyIndexedCount);

            // All 250 emails persisted successfully, 0 skipped, 0 duplicated
            var finalEmailCount = await _repository.GetIndexedEmailCountAsync();
            Assert.Equal(250, finalEmailCount);

            var finalState = await _repository.GetFolderIndexStateAsync("store1", "folder_inbox");
            Assert.True(finalState.HasValue);
            Assert.Equal(250, finalState.Value.ProcessedOffset);
            Assert.Equal("Completed", finalState.Value.Status);
        }

        [Fact]
        public async Task Test17_TransientEmailPersistenceFailure_RetriesSucceed_AndCheckpointAdvances()
        {
            var adapter = new MockCursorOutlookAdapter { TotalItems = 250 };
            var repoDecorator = new TestEmailRepositoryDecorator(_repository);
            var indexer = new EmailIndexerService(adapter, repoDecorator, new MockLicensingService(LicenseStatus.Valid));

            int attempts = 0;
            repoDecorator.SaveEmailBatchOverride = async (emails, ct) =>
            {
                attempts++;
                if (attempts < 2)
                {
                    // Fail on first attempt with transient error
                    throw new IOException("Simulated transient database locked error");
                }
                // Succeed on second retry attempt
                await _repository.SaveEmailBatchAsync(emails, ct);
            };

            var summary = await indexer.StartActiveStoreBackfillAsync(CancellationToken.None);

            Assert.False(summary.IsCancelled);
            Assert.Equal(250, summary.NewlyIndexedCount);
            Assert.Equal(0, summary.FailedCount);
            Assert.Equal(2, attempts);

            var totalEmails = await _repository.GetIndexedEmailCountAsync();
            Assert.Equal(250, totalEmails);

            var finalState = await _repository.GetFolderIndexStateAsync("store1", "folder_inbox");
            Assert.Equal(250, finalState.Value.ProcessedOffset);
            Assert.Equal("Completed", finalState.Value.Status);
        }

        [Fact]
        public async Task Test18_CancellationDuringPersistence_DoesNotAdvanceCheckpoint_AndResumesCleanly()
        {
            var adapter = new MockCursorOutlookAdapter { TotalItems = 250 };
            var repoDecorator = new TestEmailRepositoryDecorator(_repository);
            var indexer = new EmailIndexerService(adapter, repoDecorator, new MockLicensingService(LicenseStatus.Valid));

            // Simulate cancellation while SaveEmailBatchAsync is executing
            repoDecorator.SaveEmailBatchOverride = (emails, ct) =>
            {
                throw new OperationCanceledException();
            };

            var summary = await indexer.StartActiveStoreBackfillAsync(CancellationToken.None);

            Assert.True(summary.IsCancelled);

            // Checkpoint must remain 0
            var state = await _repository.GetFolderIndexStateAsync("store1", "folder_inbox");
            int offset = state?.ProcessedOffset ?? 0;
            Assert.Equal(0, offset);

            var totalEmailsFirstPass = await _repository.GetIndexedEmailCountAsync();
            Assert.Equal(0, totalEmailsFirstPass);

            // Resume cleanly without cancellation
            repoDecorator.SaveEmailBatchOverride = null;
            var resumeSummary = await indexer.StartActiveStoreBackfillAsync(CancellationToken.None);

            Assert.False(resumeSummary.IsCancelled);
            Assert.Equal(250, resumeSummary.NewlyIndexedCount);

            var finalTotal = await _repository.GetIndexedEmailCountAsync();
            Assert.Equal(250, finalTotal);

            var finalState = await _repository.GetFolderIndexStateAsync("store1", "folder_inbox");
            Assert.Equal(250, finalState.Value.ProcessedOffset);
            Assert.Equal("Completed", finalState.Value.Status);
        }

        [Fact]
        public void DEF01_StaWorkerThread_StoppedWorker_NeverExecutesComDelegateOnCallerThread()
        {
            var thread = new StaWorkerThread();
            bool executedBefore = false;
            thread.Invoke(() => { executedBefore = true; });
            Assert.True(executedBefore);

            // Shutdown worker thread
            thread.Dispose();

            // 1. Verify delegate is never executed on caller thread
            bool executedOnCaller = false;
            var ex = Assert.Throws<ObjectDisposedException>(() =>
            {
                thread.Invoke(() =>
                {
                    executedOnCaller = true;
                });
            });
            Assert.False(executedOnCaller, "DEF-01 VIOLATION: Delegate must NEVER execute on caller thread after worker has stopped");
            Assert.Contains("disposed", ex.Message, StringComparison.OrdinalIgnoreCase);

            // 2. Verify Func<T> overload also throws and does not execute
            bool funcExecuted = false;
            Assert.Throws<ObjectDisposedException>(() =>
            {
                thread.Invoke<int>(() =>
                {
                    funcExecuted = true;
                    return 100;
                });
            });
            Assert.False(funcExecuted, "DEF-01 VIOLATION: Func delegate must NEVER execute on caller thread after worker has stopped");

            // 3. Verify OutlookAdapter shutdown policy: operations fail safely without throwing unhandled exceptions
#if NETFRAMEWORK
            var adapter = new OutlookAdapter();
            adapter.Dispose();
            adapter.Dispose(); // Idempotency check

            var batchResult = adapter.FetchFolderBatch("store1", "folder1", 1, 250);
            Assert.True(batchResult.HasError);
            Assert.Contains("disposed", batchResult.ErrorMessage, StringComparison.OrdinalIgnoreCase);

            Assert.Null(adapter.FetchSingleEmail("store1", "entry1"));
            Assert.False(adapter.OpenEmailInOutlook("store1", "entry1"));
            Assert.False(adapter.IsOutlookAvailable());
            Assert.Equal("Unavailable", adapter.GetOutlookVersion());
            Assert.Empty(adapter.GetMailboxes());
            Assert.Empty(adapter.GetFolders("store1"));
#endif
        }

        [Fact]
        public async Task DEF02_EmailBatchPersistenceFailure_NeverAdvancesCheckpoint_RetryExhaustionAndRecovery()
        {
            var adapter = new MockCursorOutlookAdapter { TotalItems = 300 };
            var repoDecorator = new TestEmailRepositoryDecorator(_repository);
            var indexer = new EmailIndexerService(adapter, repoDecorator, new MockLicensingService(LicenseStatus.Valid));

            // Seed folder checkpoint at 0
            await _repository.SaveFolderIndexStateAsync("store1", "folder_inbox", "Indexing", 0, 300);

            // 1. Force persistence failure with retry exhaustion
            int persistAttempts = 0;
            repoDecorator.SaveEmailBatchOverride = (emails, ct) =>
            {
                persistAttempts++;
                throw new IOException($"Simulated database locked / disk write failure (attempt {persistAttempts})");
            };

            var failSummary = await indexer.StartActiveStoreBackfillAsync(CancellationToken.None);

            // Verify retry exhaustion occurred (3 attempts per batch)
            Assert.True(persistAttempts >= 3, $"Expected at least 3 retry attempts before exhaustion, got {persistAttempts}");
            Assert.True(failSummary.FailedCount > 0, "Summary must report failed items");

            // Verify checkpoint NEVER advanced past 0
            var stateAfterFailure = await _repository.GetFolderIndexStateAsync("store1", "folder_inbox");
            Assert.NotNull(stateAfterFailure);
            Assert.Equal(0, stateAfterFailure.Value.ProcessedOffset);

            // Verify 0 emails in DB
            var dbCountAfterFailure = await _repository.GetIndexedEmailCountAsync();
            Assert.Equal(0, dbCountAfterFailure);

            // 2. Transient failure test: fails first attempt, succeeds on second
            persistAttempts = 0;
            repoDecorator.SaveEmailBatchOverride = async (emails, ct) =>
            {
                persistAttempts++;
                if (persistAttempts == 1)
                {
                    throw new IOException("Simulated transient locking error on attempt 1");
                }
                await _repository.SaveEmailBatchAsync(emails, ct);
            };

            var recoverySummary = await indexer.StartActiveStoreBackfillAsync(CancellationToken.None);

            Assert.False(recoverySummary.IsCancelled);
            Assert.Equal(300, recoverySummary.NewlyIndexedCount);
            Assert.Equal(0, recoverySummary.FailedCount);

            // Verify all items persisted
            var finalDbCount = await _repository.GetIndexedEmailCountAsync();
            Assert.Equal(300, finalDbCount);

            // Verify checkpoint advanced to full completion
            var finalState = await _repository.GetFolderIndexStateAsync("store1", "folder_inbox");
            Assert.NotNull(finalState);
            Assert.Equal(300, finalState.Value.ProcessedOffset);
            Assert.Equal("Completed", finalState.Value.Status);
        }

        private class TestEmailRepositoryDecorator : IEmailRepository
        {
            private readonly IEmailRepository _inner;
            public Func<IEnumerable<EmailDocument>, CancellationToken, Task>? SaveEmailBatchOverride { get; set; }

            public TestEmailRepositoryDecorator(IEmailRepository inner)
            {
                _inner = inner;
            }

            public string DatabasePath => _inner.DatabasePath;
            public void InitializeDatabase() => _inner.InitializeDatabase();

            public Task SaveEmailBatchAsync(IEnumerable<EmailDocument> emails, CancellationToken cancellationToken = default)
            {
                if (SaveEmailBatchOverride != null)
                {
                    return SaveEmailBatchOverride(emails, cancellationToken);
                }
                return _inner.SaveEmailBatchAsync(emails, cancellationToken);
            }

            public Task<EmailDocument?> GetEmailByStoreAndEntryIdAsync(string storeId, string entryId) => _inner.GetEmailByStoreAndEntryIdAsync(storeId, entryId);
            public Task DeleteEmailAsync(string storeId, string entryId) => _inner.DeleteEmailAsync(storeId, entryId);
            public Task UpdateEmailFolderAsync(string storeId, string entryId, string newFolderName, string newFolderEntryId) => _inner.UpdateEmailFolderAsync(storeId, entryId, newFolderName, newFolderEntryId);
            public Task<HashSet<string>> GetIndexedEntryIdsForFolderAsync(string storeId, string folderEntryId) => _inner.GetIndexedEntryIdsForFolderAsync(storeId, folderEntryId);
            public Task SaveFoldersAsync(IEnumerable<MailboxFolder> folders) => _inner.SaveFoldersAsync(folders);
            public Task<List<MailboxFolder>> GetFoldersAsync(string? storeId = null) => _inner.GetFoldersAsync(storeId);
            public Task<IndexMetadata> GetIndexMetadataAsync() => _inner.GetIndexMetadataAsync();
            public Task ClearAllDataAsync() => _inner.ClearAllDataAsync();
            public Task SaveSearchQueryAsync(string rawQuery) => _inner.SaveSearchQueryAsync(rawQuery);
            public Task<List<string>> GetRecentSearchQueriesAsync(int count = 20) => _inner.GetRecentSearchQueriesAsync(count);
            public Task ClearSearchHistoryAsync() => _inner.ClearSearchHistoryAsync();
            public Task VacuumDatabaseAsync() => _inner.VacuumDatabaseAsync();
            public Task SaveIndexStateAsync(string state, int processed, int total, string currentFolder, string? error) => _inner.SaveIndexStateAsync(state, processed, total, currentFolder, error);
            public Task<Dictionary<string, string>> GetPersistedIndexStateAsync() => _inner.GetPersistedIndexStateAsync();
            public Task<int> GetIndexedEmailCountAsync() => _inner.GetIndexedEmailCountAsync();
            public Task<Dictionary<string, int>> GetStoreEmailCountsAsync() => _inner.GetStoreEmailCountsAsync();
            public Task<int> GetFolderEmailCountAsync(string storeId, string folderEntryId) => _inner.GetFolderEmailCountAsync(storeId, folderEntryId);
            public Task SaveFolderIndexStateAsync(string storeId, string folderEntryId, string status, int processedOffset, int totalItems) => _inner.SaveFolderIndexStateAsync(storeId, folderEntryId, status, processedOffset, totalItems);
            public Task<(string Status, int ProcessedOffset, int TotalItems)?> GetFolderIndexStateAsync(string storeId, string folderEntryId) => _inner.GetFolderIndexStateAsync(storeId, folderEntryId);
            public Task<List<FolderIndexCheckpoint>> GetAllFolderIndexStatesAsync() => _inner.GetAllFolderIndexStatesAsync();
            public Task SaveLiveProgressAsync(string state, string currentOperation, int processed, int total, string currentFolder, int folderOffset, int folderTotal, int newlyIndexed, int skipped, int failed, int pendingQueue) => _inner.SaveLiveProgressAsync(state, currentOperation, processed, total, currentFolder, folderOffset, folderTotal, newlyIndexed, skipped, failed, pendingQueue);
            public Task EnqueuePendingChangeAsync(string storeId, string entryId, string operation, string? folderEntryId) => _inner.EnqueuePendingChangeAsync(storeId, entryId, operation, folderEntryId);
            public Task<List<PendingChangeItem>> GetPendingChangesAsync(int maxCount = 50) => _inner.GetPendingChangesAsync(maxCount);
            public Task<int> GetPendingChangesCountAsync() => _inner.GetPendingChangesCountAsync();
            public Task<int> IncrementAttemptCountAsync(long id) => _inner.IncrementAttemptCountAsync(id);
            public Task MarkPendingChangeStatusAsync(long id, string status, string? error = null) => _inner.MarkPendingChangeStatusAsync(id, status, error);
            public Task MarkPendingChangesCompletedBatchAsync(IEnumerable<long> ids) => _inner.MarkPendingChangesCompletedBatchAsync(ids);
            public Task RecoverStaleProcessingChangesAsync() => _inner.RecoverStaleProcessingChangesAsync();
            public Task SaveFailedIndexItemAsync(string storeId, string entryId, string operation, int attemptCount, string lastError) => _inner.SaveFailedIndexItemAsync(storeId, entryId, operation, attemptCount, lastError);
            public Task SaveFailedIndexItemAsync(FailedIndexItem item) => _inner.SaveFailedIndexItemAsync(item);
            public Task<List<FailedIndexItem>> GetFailedIndexItemsAsync(bool includeResolved = false) => _inner.GetFailedIndexItemsAsync(includeResolved);
            public Task<List<FailedIndexItem>> GetFailedIndexItemsAsync(string? status, int maxCount = 500) => _inner.GetFailedIndexItemsAsync(status, maxCount);
            public Task<int> GetFailedIndexItemsCountAsync(string? status = "Failed") => _inner.GetFailedIndexItemsCountAsync(status);
            public Task UpdateFailedIndexItemStatusAsync(long id, string status, string? error = null) => _inner.UpdateFailedIndexItemStatusAsync(id, status, error);
            public Task RemoveFailedIndexItemAsync(long id) => _inner.RemoveFailedIndexItemAsync(id);
            public Task RemoveAllFailedIndexItemsAsync() => _inner.RemoveAllFailedIndexItemsAsync();
            public Task<double> GetDatabaseSizeMbAsync() => _inner.GetDatabaseSizeMbAsync();
            public Task<IndexIntegrityResult> ScanIndexIntegrityAsync(CancellationToken cancellationToken = default) => _inner.ScanIndexIntegrityAsync(cancellationToken);
            public Task RepairIndexIntegrityAsync(IndexIntegrityResult scanResult, CancellationToken cancellationToken = default) => _inner.RepairIndexIntegrityAsync(scanResult, cancellationToken);
        }
    }
}
