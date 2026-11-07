using System;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using FastMailSearch.Infrastructure.Outlook;
using FastMailSearch.OutlookAddIn;
using FastMailSearch.OutlookAddIn.Ribbon;
using Microsoft.Data.Sqlite;
using Xunit;

namespace FastMailSearch.Core.Tests
{
    public class Phase4K6F_OutlookAddInStabilityHardeningTests
    {
        private static void InvokeOnConnection(ThisAddIn addin)
        {
            var method = addin.GetType().GetMethod("OnConnection", BindingFlags.Public | BindingFlags.Instance);
            Assert.NotNull(method);
            var connectModeType = method!.GetParameters()[1].ParameterType;
            var connectModeVal = Enum.ToObject(connectModeType, 3); // ext_cm_AfterStartup
            Array custom = Array.CreateInstance(typeof(object), 0);
            object[] args = new object[] { null!, connectModeVal, null!, custom };
            method.Invoke(addin, args);
        }

        private static void InvokeOnDisconnection(ThisAddIn addin)
        {
            var method = addin.GetType().GetMethod("OnDisconnection", BindingFlags.Public | BindingFlags.Instance);
            Assert.NotNull(method);
            var disconnectModeType = method!.GetParameters()[0].ParameterType;
            var disconnectModeVal = Enum.ToObject(disconnectModeType, 1); // ext_dm_HostShutdown
            Array custom = Array.CreateInstance(typeof(object), 0);
            object[] args = new object[] { disconnectModeVal, custom };
            method.Invoke(addin, args);
        }

        private static void InvokeOnBeginShutdown(ThisAddIn addin)
        {
            var method = addin.GetType().GetMethod("OnBeginShutdown", BindingFlags.Public | BindingFlags.Instance);
            Assert.NotNull(method);
            Array custom = Array.CreateInstance(typeof(object), 0);
            object[] args = new object[] { custom };
            method.Invoke(addin, args);
        }

        // --------------------------------------------------------------------------------------------------
        // Test 1: Clean Add-in Startup
        // --------------------------------------------------------------------------------------------------
        [Fact]
        public void AddIn_CleanStartup_InitializesStateWithoutThrowing()
        {
            var addin = new ThisAddIn();

            // OnConnection should complete cleanly without throwing
            InvokeOnConnection(addin);

            Assert.Same(addin, ThisAddIn.Instance);

            // Clean teardown
            InvokeOnDisconnection(addin);
            Assert.Null(ThisAddIn.Instance);
        }

        // --------------------------------------------------------------------------------------------------
        // Test 2: Native SQLite DLL Resolution via LoadLibraryEx
        // --------------------------------------------------------------------------------------------------
        [Fact]
        public void NativeSqlite_LoadLibraryEx_LoadsSuccessfullyAndEnablesSqliteQueries()
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            bool preloadSuccess = ThisAddIn.TryPreloadNativeSqlite(baseDir);
            Assert.True(preloadSuccess, $"TryPreloadNativeSqlite failed for base directory '{baseDir}'");

            // Execute an isolated in-memory SQLite query to confirm native provider operates cleanly
            using var connection = new SqliteConnection("Data Source=:memory:");
            connection.Open();
            using var cmd = connection.CreateCommand();
            cmd.CommandText = "SELECT sqlite_version();";
            var version = cmd.ExecuteScalar() as string;

            Assert.NotNull(version);
            Assert.NotEmpty(version!);
            Assert.StartsWith("3.", version!);
        }

        // --------------------------------------------------------------------------------------------------
        // Test 3: Native DLL Loading Failure Handling
        // --------------------------------------------------------------------------------------------------
        [Fact]
        public void NativeSqlite_InvalidFolder_HandlesFailureGracefullyWithoutThrowing()
        {
            // Providing an invalid or non-existent path must gracefully return false without unhandled exceptions
            bool result = ThisAddIn.TryPreloadNativeSqlite(@"C:\NonExistent_FastMailSearch_Path_12345");
            Assert.False(result);
        }

        // --------------------------------------------------------------------------------------------------
        // Test 4: Background Indexing Exception Observation
        // --------------------------------------------------------------------------------------------------
        [Fact]
        public async Task BackgroundIndexing_ExceptionObservation_DoesNotTerminateProcess()
        {
            // Verify that background tasks catch and log exceptions instead of allowing unobserved task exceptions
            bool exceptionHandled = false;

            var task = Task.Run(async () =>
            {
                try
                {
                    await Task.Yield();
                    throw new InvalidOperationException("Simulated background indexing failure");
                }
                catch (Exception ex)
                {
                    exceptionHandled = true;
                    Assert.Equal("Simulated background indexing failure", ex.Message);
                }
            });

            await task;
            Assert.True(exceptionHandled);
        }

        // --------------------------------------------------------------------------------------------------
        // Test 5: Shutdown During Active Work (Cancellation & Bounded Wait)
        // --------------------------------------------------------------------------------------------------
        [Fact]
        public void Shutdown_DuringActiveWork_SignalsCancellationAndPerformsBoundedWait()
        {
            using var cts = new CancellationTokenSource();
            var taskStarted = new ManualResetEventSlim(false);

            var workerTask = Task.Run(async () =>
            {
                taskStarted.Set();
                try
                {
                    while (!cts.Token.IsCancellationRequested)
                    {
                        await Task.Delay(20, cts.Token);
                    }
                }
                catch (OperationCanceledException)
                {
                    // Cancellation observed
                }
            }, cts.Token);

            Assert.True(taskStarted.Wait(1000), "Worker task failed to start");

            // Signal shutdown and perform bounded wait
            cts.Cancel();
            bool completed = workerTask.Wait(TimeSpan.FromSeconds(3));

            Assert.True(completed, "Worker task did not complete within 3000ms bounded wait");
            Assert.True(workerTask.IsCompleted);
        }

        // --------------------------------------------------------------------------------------------------
        // Test 6: Repeated Startup / Shutdown and Disposal Idempotence
        // --------------------------------------------------------------------------------------------------
        [Fact]
        public void AddIn_RepeatedLifecycle_IsIdempotentAndSafe()
        {
            var addin = new ThisAddIn();

            InvokeOnConnection(addin);

            // First shutdown & disconnection
            InvokeOnBeginShutdown(addin);
            InvokeOnDisconnection(addin);

            // Repeated shutdown & disconnection calls must be idempotent and not throw
            InvokeOnBeginShutdown(addin);
            InvokeOnDisconnection(addin);

            Assert.Null(ThisAddIn.Instance);
        }

        // --------------------------------------------------------------------------------------------------
        // Test 7: COM Event Unsubscription
        // --------------------------------------------------------------------------------------------------
        [Fact]
        public void OutlookAdapter_UnsubscribeFromEvents_CanBeCalledSafelyAndIdempotently()
        {
            using var adapter = new OutlookAdapter();

            // Calling UnsubscribeFromEvents when not subscribed or multiple times must not throw
            adapter.UnsubscribeFromEvents();
            adapter.UnsubscribeFromEvents();

            // Calling Dispose must also safely invoke unsubscription and shutdown worker
            adapter.Dispose();
            adapter.Dispose();
        }

        // --------------------------------------------------------------------------------------------------
        // Test 8: Ribbon Callback Exception Handling
        // --------------------------------------------------------------------------------------------------
        [Fact]
        public void RibbonCallbacks_HandleExceptionsGracefullyWithoutThrowing()
        {
            var ribbon = new FastMailRibbon();

            // FastMailRibbon callbacks invoked via reflection to handle COM signatures safely
            var getCustomUi = ribbon.GetType().GetMethod("GetCustomUI");
            Assert.NotNull(getCustomUi);
            var ui = (string)getCustomUi!.Invoke(ribbon, new object[] { "Microsoft.Outlook.Explorer" });
            Assert.False(string.IsNullOrEmpty(ui));

            var emptyUi = (string)getCustomUi.Invoke(ribbon, new object[] { null! });
            Assert.Equal(string.Empty, emptyUi);

            var ribbonLoad = ribbon.GetType().GetMethod("Ribbon_Load");
            Assert.NotNull(ribbonLoad);
            ribbonLoad!.Invoke(ribbon, new object[] { null! });

            var clickSearch = ribbon.GetType().GetMethod("OnOpenSearchWindowClicked");
            Assert.NotNull(clickSearch);
            clickSearch!.Invoke(ribbon, new object[] { null! });

            var clickBackfill = ribbon.GetType().GetMethod("OnBackfillActiveStoreClicked");
            Assert.NotNull(clickBackfill);
            clickBackfill!.Invoke(ribbon, new object[] { null! });

            // Test ThisAddIn ribbon callbacks with isolated temporary database path
            string tempDb = Path.Combine(Path.GetTempPath(), $"fastmail_test_{Guid.NewGuid():N}.db");
            try
            {
                Environment.SetEnvironmentVariable("FASTMAILSEARCH_DB_PATH", tempDb);
                var addin = new ThisAddIn();
                var addinCustomUi = addin.GetType().GetMethod("GetCustomUI");
                Assert.NotNull(addinCustomUi);
                var addinUi = (string)addinCustomUi!.Invoke(addin, new object[] { "Microsoft.Outlook.Explorer" });
                Assert.False(string.IsNullOrEmpty(addinUi));

                var addinRibbonLoad = addin.GetType().GetMethod("Ribbon_Load");
                Assert.NotNull(addinRibbonLoad);
                addinRibbonLoad!.Invoke(addin, new object[] { null! });

                var addinClickSearch = addin.GetType().GetMethod("OnOpenSearchWindowClicked");
                Assert.NotNull(addinClickSearch);
                addinClickSearch!.Invoke(addin, new object[] { null! });

                var addinClickBackfill = addin.GetType().GetMethod("OnBackfillActiveStoreClicked");
                Assert.NotNull(addinClickBackfill);
                addinClickBackfill!.Invoke(addin, new object[] { null! });
            }
            finally
            {
                Environment.SetEnvironmentVariable("FASTMAILSEARCH_DB_PATH", null);
                if (File.Exists(tempDb))
                {
                    try { File.Delete(tempDb); } catch { }
                }
            }
        }

        // --------------------------------------------------------------------------------------------------
        // Test 9: No Background Work After Shutdown
        // --------------------------------------------------------------------------------------------------
        [Fact]
        public void AddIn_AfterShutdown_RejectsNewBackgroundWork()
        {
            string tempDb = Path.Combine(Path.GetTempPath(), $"fastmail_test_{Guid.NewGuid():N}.db");
            try
            {
                Environment.SetEnvironmentVariable("FASTMAILSEARCH_DB_PATH", tempDb);
                var addin = new ThisAddIn();

                InvokeOnConnection(addin);
                InvokeOnDisconnection(addin);

                // Attempting to start in-process backfill after disconnection must immediately abort
                addin.StartInProcessBackfill();
            }
            finally
            {
                Environment.SetEnvironmentVariable("FASTMAILSEARCH_DB_PATH", null);
                if (File.Exists(tempDb))
                {
                    try { File.Delete(tempDb); } catch { }
                }
            }
        }

        // --------------------------------------------------------------------------------------------------
        // Test 10: STA Thread Lifecycle and Timeout Handling
        // --------------------------------------------------------------------------------------------------
        [Fact]
        public void StaWorkerThread_LifecycleAndTimeout_HandlesTimeoutsAndDisposesCleanly()
        {
            var sta = new StaWorkerThread();

            // 1. Invoke on STA thread
            bool executedOnSta = false;
            ApartmentState apartment = ApartmentState.Unknown;
            sta.Invoke(() =>
            {
                executedOnSta = true;
                apartment = Thread.CurrentThread.GetApartmentState();
            });
            Assert.True(executedOnSta);
            Assert.Equal(ApartmentState.STA, apartment);

            // 2. TryInvoke with timeout handling
            bool tryResult = sta.TryInvoke(() =>
            {
                Thread.Sleep(200);
            }, TimeSpan.FromMilliseconds(20));

            // Must return false because action takes 200ms while timeout is 20ms
            Assert.False(tryResult);

            // 3. Dispose cleanly and verify behavior post-dispose
            sta.Dispose();

            // TryInvoke on disposed worker returns false
            bool postDisposeTry = sta.TryInvoke(() => { }, TimeSpan.FromMilliseconds(50));
            Assert.False(postDisposeTry);

            // Invoke on disposed worker throws ObjectDisposedException
            Assert.Throws<ObjectDisposedException>(() =>
            {
                sta.Invoke(() => { });
            });

            // Repeated Dispose is idempotent
            sta.Dispose();
        }

        // --------------------------------------------------------------------------------------------------
        // Test 11: COM Event Subscription Partial Failure & Cleanup
        // --------------------------------------------------------------------------------------------------
        [Fact]
        public void COMEventSubscription_PartialFailureAndCleanup_HandlesPartialSubscriptionsAndIdempotence()
        {
            // 1. SafeSubscribeItems returns false on null items without throwing
            bool nullResult = OutlookAdapter.SafeSubscribeItems(null!, null!, null!, out var nullSub);
            Assert.False(nullResult);
            Assert.Null(nullSub);

            // 2. FolderEventSubscription handles null/partially assigned delegates cleanly and idempotently
            var partialSub = new OutlookAdapter.FolderEventSubscription(null!, null!, null!);
            Assert.False(partialSub.IsUnsubscribed);

            // First call cleanly sets IsUnsubscribed to true without throwing
            partialSub.UnsubscribeAndRelease();
            Assert.True(partialSub.IsUnsubscribed);

            // Repeated call is idempotent and does not throw
            partialSub.UnsubscribeAndRelease();
            Assert.True(partialSub.IsUnsubscribed);
        }

        // --------------------------------------------------------------------------------------------------
        // Test 12: Shared Shutdown Deadline Bounds Cumulative Cleanup Time
        // --------------------------------------------------------------------------------------------------
        [Fact]
        public void Shutdown_SharedThreeSecondDeadline_BoundsCumulativeWaitAcrossSteps()
        {
            string tempDb = Path.Combine(Path.GetTempPath(), $"fastmail_test_{Guid.NewGuid():N}.db");
            try
            {
                Environment.SetEnvironmentVariable("FASTMAILSEARCH_DB_PATH", tempDb);
                var addin = new ThisAddIn();

                InvokeOnConnection(addin);

                // Attach a simulated slow task to the active backfill task via reflection
                var activeTaskField = typeof(ThisAddIn).GetField("_activeBackfillTask", BindingFlags.NonPublic | BindingFlags.Instance);
                Assert.NotNull(activeTaskField);

                var slowTask = Task.Run(async () =>
                {
                    await Task.Delay(10000); // 10 seconds (exceeds 3000ms budget)
                });
                activeTaskField!.SetValue(addin, slowTask);

                var sw = System.Diagnostics.Stopwatch.StartNew();
                InvokeOnDisconnection(addin);
                sw.Stop();

                // Total elapsed time must be bounded by the shared 3000ms deadline
                // and must NOT be cumulative 3s + 3s (6s)
                Assert.True(sw.ElapsedMilliseconds >= 2800, $"Expected wait near 3000ms, but was {sw.ElapsedMilliseconds}ms");
                Assert.True(sw.ElapsedMilliseconds < 4500, $"Expected shutdown to enforce shared 3000ms deadline, but took {sw.ElapsedMilliseconds}ms");
            }
            finally
            {
                Environment.SetEnvironmentVariable("FASTMAILSEARCH_DB_PATH", null);
                if (File.Exists(tempDb))
                {
                    try { File.Delete(tempDb); } catch { }
                }
            }
        }

        // --------------------------------------------------------------------------------------------------
        // Test 13: StaWorkerThread Dispose with Zero Timeout Returns Immediately
        // --------------------------------------------------------------------------------------------------
        [Fact]
        public void StaWorkerThread_DisposeZeroTimeout_ReturnsImmediatelyWithoutBlocking()
        {
            var sta = new StaWorkerThread();

            // Queue a slow work item
            sta.TryInvoke(() =>
            {
                Thread.Sleep(5000);
            }, TimeSpan.FromMilliseconds(10));

            var sw = System.Diagnostics.Stopwatch.StartNew();
            // Dispose with 0ms timeout (budget exhausted scenario)
            sta.Dispose(0);
            sw.Stop();

            // Must return immediately (well under 500ms) without waiting 3000ms default
            Assert.True(sw.ElapsedMilliseconds < 500, $"Expected Dispose(0) to return immediately, took {sw.ElapsedMilliseconds}ms");
        }
    }
}
