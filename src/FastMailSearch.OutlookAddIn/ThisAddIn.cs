using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using Extensibility;
using FastMailSearch.Core.Models;
using FastMailSearch.Core.Services;
using FastMailSearch.Infrastructure.Database;
using FastMailSearch.Infrastructure.Indexing;
using FastMailSearch.Infrastructure.Logging;
using FastMailSearch.Infrastructure.Outlook;
using FastMailSearch.Infrastructure.Services;
using FastMailSearch.OutlookAddIn.Ribbon;
using OfficeCore = Microsoft.Office.Core;
using OutlookInterop = Microsoft.Office.Interop.Outlook;

namespace FastMailSearch.OutlookAddIn
{
    [ComVisible(true)]
    [Guid("4A10F4D0-7E55-442C-B23B-983F78BA48D2")]
    [ProgId("FastMailSearch.AddIn")]
    public class ThisAddIn : IDTExtensibility2, OfficeCore.IRibbonExtensibility
    {
        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr LoadLibraryEx(string lpFileName, IntPtr hFile, uint dwFlags);
        private const uint LOAD_WITH_ALTERED_SEARCH_PATH = 0x00000008;

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool FreeLibrary(IntPtr hModule);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        private const int SW_RESTORE = 9;

        private static IntPtr _nativeSqliteHandle = IntPtr.Zero;

        public static bool TryPreloadNativeSqlite(string addinFolder)
        {
            try
            {
                if (string.IsNullOrEmpty(addinFolder) || !Directory.Exists(addinFolder))
                {
                    Logger.Warn($"[NATIVE LOAD WARN] Add-in folder invalid or does not exist: '{addinFolder}'");
                    return false;
                }

                if (_nativeSqliteHandle != IntPtr.Zero) return true;

                string arch = Environment.Is64BitProcess ? "win-x64" : "win-x86";
                string[] candidatePaths = new[]
                {
                    Path.Combine(addinFolder, "runtimes", arch, "native", "e_sqlite3.dll"),
                    Path.Combine(addinFolder, arch, "e_sqlite3.dll"),
                    Path.Combine(addinFolder, "e_sqlite3.dll")
                };

                foreach (string candidate in candidatePaths)
                {
                    if (File.Exists(candidate))
                    {
                        IntPtr h = LoadLibraryEx(candidate, IntPtr.Zero, LOAD_WITH_ALTERED_SEARCH_PATH);
                        if (h != IntPtr.Zero)
                        {
                            _nativeSqliteHandle = h;
                            Logger.Info($"[NATIVE LOAD] Successfully pre-loaded native SQLite module from '{candidate}' (x64={Environment.Is64BitProcess}, Handle=0x{h.ToInt64():X}).");
                            return true;
                        }
                        else
                        {
                            int err = Marshal.GetLastWin32Error();
                            Logger.Warn($"[NATIVE LOAD WARN] LoadLibraryEx failed for candidate '{candidate}' with Win32 error {err}.");
                        }
                    }
                }

                Logger.Warn($"[NATIVE LOAD WARN] Could not locate or load e_sqlite3.dll under '{addinFolder}' for arch={arch}.");
                return false;
            }
            catch (Exception ex)
            {
                Logger.Error("[NATIVE LOAD ERROR] Exception attempting to pre-load native SQLite module", ex);
                return false;
            }
        }

        static ThisAddIn()
        {
            try
            {
                string addinFolder = Path.GetDirectoryName(typeof(ThisAddIn).Assembly.Location) ?? string.Empty;
                if (!string.IsNullOrEmpty(addinFolder) && Directory.Exists(addinFolder))
                {
                    TryPreloadNativeSqlite(addinFolder);
                }

                AppDomain.CurrentDomain.AssemblyResolve += (sender, args) =>
                {
                    try
                    {
                        string folder = Path.GetDirectoryName(typeof(ThisAddIn).Assembly.Location) ?? string.Empty;
                        string asmName = new AssemblyName(args.Name).Name + ".dll";
                        string asmPath = Path.Combine(folder, asmName);
                        if (File.Exists(asmPath))
                        {
                            return Assembly.LoadFrom(asmPath);
                        }
                    }
                    catch { }
                    return null;
                };

                Logger.Info("ThisAddIn static constructor executed. Probing hooks configured.");
            }
            catch (Exception ex)
            {
                try { Logger.Error("Exception in ThisAddIn static constructor", ex); } catch { }
            }
        }

        public ThisAddIn()
        {
            try
            {
                Logger.Info("ThisAddIn parameterless constructor invoked via COM.");
            }
            catch { }
        }

        public static ThisAddIn? Instance { get; private set; }
        public static bool IsSafeMode { get; private set; }
        public static string SafeModeReason { get; private set; } = string.Empty;

        private readonly CancellationTokenSource _shutdownCts = new CancellationTokenSource();
        private Task? _activeBackfillTask;
        private readonly object _lifecycleLock = new object();
        private bool _isDisposed;

        private OutlookInterop.Application? _application;
        private DatabaseConnectionFactory? _dbFactory;
        private SqliteEmailRepository? _repository;
        private SearchService? _searchService;
        private OutlookAdapter? _outlookAdapter;
        private EmailIndexerService? _indexerService;
        private FastMailSearch.Core.Interfaces.ILicensingService? _licensingService;
        private OfficeCore.IRibbonUI? _ribbon;
        private bool _isEngineInitialized;

        public OutlookInterop.Application Application => _application ?? throw new InvalidOperationException("Outlook Application is not initialized.");

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                Logger.Info($"ThisAddIn OnConnection invoked by Outlook COM (ConnectMode={ConnectMode}).");
                Instance = this;
                _application = Application as OutlookInterop.Application;
                _isDisposed = false;
                Logger.Info("ThisAddIn OnConnection completed successfully.");
            }
            catch (Exception ex)
            {
                Logger.Error("Exception during ThisAddIn OnConnection", ex);
            }
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            lock (_lifecycleLock)
            {
                if (_isDisposed) return;
                _isDisposed = true;
            }

            try
            {
                Logger.Info($"ThisAddIn OnDisconnection invoked (RemoveMode={RemoveMode}).");

                const int ShutdownDeadlineMs = 3000;
                var stopwatch = Stopwatch.StartNew();

                int GetRemainingMs()
                {
                    int remaining = ShutdownDeadlineMs - (int)stopwatch.ElapsedMilliseconds;
                    return remaining > 0 ? remaining : 0;
                }

                if (!_shutdownCts.IsCancellationRequested)
                {
                    _shutdownCts.Cancel();
                }

                // 1. Unsubscribe from events first within remaining budget
                try
                {
                    int remMs = GetRemainingMs();
                    _outlookAdapter?.UnsubscribeFromEvents(TimeSpan.FromMilliseconds(remMs));
                }
                catch (Exception ex)
                {
                    Logger.Warn($"Error unsubscribing from Outlook events during shutdown: {ex.Message}");
                }

                // 2. Request indexing cancellation
                try
                {
                    _indexerService?.CancelIndexingAsync();
                }
                catch (Exception ex)
                {
                    Logger.Warn($"Error requesting indexing cancellation during shutdown: {ex.Message}");
                }

                // 3. Bounded wait for active background work to conclude within remaining budget
                Task? runningTask = _activeBackfillTask;
                if (runningTask != null && !runningTask.IsCompleted)
                {
                    int remMs = GetRemainingMs();
                    if (remMs > 0)
                    {
                        Logger.Info($"[SHUTDOWN] Waiting up to {remMs}ms (shared deadline) for active background task to finish...");
                        try
                        {
                            bool finished = runningTask.Wait(TimeSpan.FromMilliseconds(remMs));
                            if (!finished)
                            {
                                Logger.Warn($"[SHUTDOWN] Active background task did not finish within remaining {remMs}ms budget.");
                            }
                        }
                        catch (AggregateException aEx)
                        {
                            Logger.Info($"[SHUTDOWN] Active background task finished: {aEx.InnerException?.Message ?? aEx.Message}");
                        }
                        catch (Exception ex)
                        {
                            Logger.Warn($"[SHUTDOWN] Exception awaiting background task during shutdown: {ex.Message}");
                        }
                    }
                    else
                    {
                        Logger.Warn("[SHUTDOWN] Shared deadline exhausted before awaiting background task.");
                    }
                }

                // 4. Cleanly dispose OutlookAdapter and STA worker thread within remaining budget
                try
                {
                    int remMs = GetRemainingMs();
                    _outlookAdapter?.Dispose(TimeSpan.FromMilliseconds(remMs));
                }
                catch (Exception ex)
                {
                    Logger.Warn($"Error disposing Outlook adapter during shutdown: {ex.Message}");
                }

                // 5. Cleanly reset state
                _outlookAdapter = null;
                _indexerService = null;
                _repository = null;
                _dbFactory = null;
                _searchService = null;
                _licensingService = null;
                _application = null;
                _isEngineInitialized = false;
                Instance = null;

                stopwatch.Stop();
                if (stopwatch.ElapsedMilliseconds >= ShutdownDeadlineMs)
                {
                    Logger.Warn($"[SHUTDOWN] Cleanup did not complete within the shared {ShutdownDeadlineMs}ms deadline (elapsed: {stopwatch.ElapsedMilliseconds}ms). Continuing shutdown to avoid blocking Outlook STA thread.");
                }
                else
                {
                    Logger.Info($"[SHUTDOWN] ThisAddIn OnDisconnection cleanup completed cleanly in {stopwatch.ElapsedMilliseconds}ms within shared {ShutdownDeadlineMs}ms deadline.");
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Exception during ThisAddIn OnDisconnection", ex);
            }
        }

        public void OnAddInsUpdate(ref Array custom)
        {
            Logger.Info("ThisAddIn OnAddInsUpdate invoked.");
        }

        public void OnStartupComplete(ref Array custom)
        {
            try
            {
                Logger.Info("ThisAddIn OnStartupComplete invoked.");
            }
            catch (Exception ex)
            {
                Logger.Error("Exception during ThisAddIn OnStartupComplete", ex);
            }
        }

        public void OnBeginShutdown(ref Array custom)
        {
            try
            {
                Logger.Info("ThisAddIn OnBeginShutdown invoked.");
                if (!_shutdownCts.IsCancellationRequested)
                {
                    _shutdownCts.Cancel();
                }
                _indexerService?.CancelIndexingAsync();
                _outlookAdapter?.UnsubscribeFromEvents();
            }
            catch (Exception ex)
            {
                Logger.Error("Exception during ThisAddIn OnBeginShutdown", ex);
            }
        }

        public string GetCustomUI(string RibbonID)
        {
            try
            {
                Logger.Info($"GetCustomUI requested for RibbonID='{RibbonID}'.");
                if (RibbonID == "Microsoft.Outlook.Explorer")
                {
                    return new FastMailRibbon().GetCustomUI(RibbonID);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Exception in GetCustomUI", ex);
            }
            return string.Empty;
        }

        [ComVisible(true)]
        public void Ribbon_Load(OfficeCore.IRibbonUI ribbonUI)
        {
            try
            {
                Logger.Info("[RIBBON] Ribbon_Load callback invoked by Outlook.");
                _ribbon = ribbonUI;
            }
            catch (Exception ex)
            {
                Logger.Error("[RIBBON ERROR] Ribbon_Load failed", ex);
            }
        }

        [ComVisible(true)]
        public void OnOpenSearchWindowClicked(OfficeCore.IRibbonControl control)
        {
            try
            {
                Logger.Info("[RIBBON] FastMail Search button clicked");
                LaunchOrActivateUIProcess();
            }
            catch (Exception ex)
            {
                Logger.Error("[RIBBON ERROR] OnOpenSearchWindowClicked failed", ex);
            }
        }

        [ComVisible(true)]
        public void OnBackfillActiveStoreClicked(OfficeCore.IRibbonControl control)
        {
            try
            {
                Logger.Info("[RIBBON] Backfill POP Mailbox button clicked");
                StartInProcessBackfill();
            }
            catch (Exception ex)
            {
                Logger.Error("[RIBBON ERROR] OnBackfillActiveStoreClicked failed", ex);
            }
        }

        private readonly object _backfillLock = new object();
        private bool _isBackfillRunning;

        public void StartInProcessBackfill()
        {
            if (_shutdownCts.IsCancellationRequested)
            {
                Logger.Info("[IN-PROCESS BACKFILL] Aborted: add-in shutdown is in progress.");
                return;
            }

            try
            {
                EnsureEngineInitialized();
                var status = _licensingService?.GetStatus() ?? FastMailSearch.Core.Interfaces.LicenseStatus.InvalidKey;
                if (status != FastMailSearch.Core.Interfaces.LicenseStatus.Valid && status != FastMailSearch.Core.Interfaces.LicenseStatus.Trial)
                {
                    Logger.Info($"[IN-PROCESS BACKFILL] Backfill blocked by licensing gate. Current status: {status}");
                    System.Windows.MessageBox.Show(
                        "Email indexing is paused because your 30-day trial has expired or a valid license is required.\n\nPlease activate FastMail Search to resume indexing.",
                        "FastMail Search — License Required",
                        System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Information
                    );
                    return;
                }
            }
            catch (Exception ex)
            {
                Logger.Error("[IN-PROCESS BACKFILL] Exception checking license eligibility", ex);
                return;
            }

            lock (_backfillLock)
            {
                if (_isBackfillRunning)
                {
                    Logger.Info("[IN-PROCESS BACKFILL] Active store backfill is already running.");
                    return;
                }
                _isBackfillRunning = true;
            }

            _activeBackfillTask = Task.Run(async () =>
            {
                try
                {
                    if (_shutdownCts.IsCancellationRequested)
                    {
                        Logger.Info("[IN-PROCESS BACKFILL] Aborted: add-in shutdown is in progress.");
                        return;
                    }

                    Logger.Info("[IN-PROCESS BACKFILL] Initializing FastMail Search engine...");
                    EnsureEngineInitialized();

                    if (_indexerService == null)
                    {
                        Logger.Error("[IN-PROCESS BACKFILL] Indexer service not initialized.");
                        return;
                    }

                    if (_indexerService.IsRebuilding)
                    {
                        Logger.Info("[IN-PROCESS BACKFILL] Indexer is currently rebuilding.");
                        return;
                    }

                    string dbPath = _repository?.DatabasePath ?? string.Empty;
                    Logger.Info($"[IN-PROCESS BACKFILL] Starting active store backfill on confirmed DB: '{dbPath}'");

                    var summary = await _indexerService.StartActiveStoreBackfillAsync();
                    Logger.Info($"[IN-PROCESS BACKFILL] Complete. Discovered={summary.TotalDiscovered}, NewlyIndexed={summary.NewlyIndexedCount}, Skipped={summary.SkippedCount}, Failed={summary.FailedCount}, Cancelled={summary.IsCancelled}");
                }
                catch (OperationCanceledException)
                {
                    Logger.Info("[IN-PROCESS BACKFILL] Backfill cancelled successfully.");
                }
                catch (Exception ex)
                {
                    Logger.Error("[IN-PROCESS BACKFILL] Exception during active store backfill", ex);
                }
                finally
                {
                    lock (_backfillLock) { _isBackfillRunning = false; }
                }
            }, _shutdownCts.Token);
        }

        public static void OpenSearchWindow()
        {
            Logger.Info("[RIBBON] FastMail Search button clicked");
            LaunchOrActivateUIProcess();
        }

        public static bool LaunchOrActivateUIProcess()
        {
            try
            {
                Logger.Info("[UI LAUNCH] Checking existing UI instance");
                string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                string installDir = Path.Combine(localAppData, "Programs", "FastMailSearch");
                if (!Directory.Exists(installDir))
                {
                    installDir = AppContext.BaseDirectory;
                }

                string uiExe = Path.Combine(installDir, "FastMailSearch.UI.exe");
                if (!File.Exists(uiExe))
                {
                    uiExe = Path.Combine(AppContext.BaseDirectory, "FastMailSearch.UI.exe");
                }

                if (!File.Exists(uiExe))
                {
                    Logger.Error($"[UI ERROR] FastMailSearch.UI.exe missing at path {uiExe}");
                    return false;
                }

                var existingProcesses = Process.GetProcessesByName("FastMailSearch.UI");
                if (existingProcesses.Length > 0)
                {
                    Logger.Info("[UI LAUNCH] Existing UI instance found");
                    foreach (var proc in existingProcesses)
                    {
                        IntPtr hWnd = proc.MainWindowHandle;
                        if (hWnd != IntPtr.Zero)
                        {
                            Logger.Info($"[UI ACTIVATE] Activating existing UI window handle: {hWnd}");
                            ShowWindow(hWnd, SW_RESTORE);
                            SetForegroundWindow(hWnd);
                            return true;
                        }
                    }

                    // Retry up to 5 times for main window handle if process is mid-startup
                    for (int retry = 0; retry < 5; retry++)
                    {
                        System.Threading.Thread.Sleep(100);
                        foreach (var proc in Process.GetProcessesByName("FastMailSearch.UI"))
                        {
                            IntPtr hWnd = proc.MainWindowHandle;
                            if (hWnd != IntPtr.Zero)
                            {
                                Logger.Info($"[UI ACTIVATE] Activating existing UI window handle: {hWnd}");
                                ShowWindow(hWnd, SW_RESTORE);
                                SetForegroundWindow(hWnd);
                                return true;
                            }
                        }
                    }

                    Logger.Warn("[UI LAUNCH] Existing process had no main window handle; terminating orphaned process for recovery.");
                    foreach (var proc in existingProcesses)
                    {
                        try { proc.Kill(); } catch { }
                    }
                }

                Logger.Info("[UI LAUNCH] No existing UI instance");
                Logger.Info($"[UI LAUNCH] Starting FastMailSearch.UI.exe at {uiExe}");

                string resolvedDb = Instance?._dbFactory?.DbPath ?? AppSettings.GetDatabasePath();
                ProcessStartInfo psi = CreateUIProcessStartInfo(uiExe, installDir, resolvedDb, Logger.LogDirectory);

                Process? newProc = Process.Start(psi);
                if (newProc != null)
                {
                    Logger.Info($"[UI LAUNCH] Process started PID={newProc.Id}");
                    return true;
                }
                else
                {
                    Logger.Error("[UI ERROR] Process.Start returned null");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Logger.Error("[UI ERROR] LaunchOrActivateUIProcess failed", ex);
                return false;
            }
        }

        public static ProcessStartInfo CreateUIProcessStartInfo(string uiExe, string installDir, string dbPath, string? logDir = null)
        {
            var psi = new ProcessStartInfo
            {
                FileName = uiExe,
                WorkingDirectory = installDir,
                UseShellExecute = false
            };

            // Safely quote arguments to protect paths with spaces
            string safeDbPath = (dbPath ?? string.Empty).Replace("\"", "\\\"");
            psi.Arguments = $"--db-path \"{safeDbPath}\"";

            if (!string.IsNullOrWhiteSpace(dbPath))
            {
                psi.EnvironmentVariables[AppSettings.DatabasePathEnvironmentVariable] = dbPath;
            }

            string? effectiveLogDir = !string.IsNullOrWhiteSpace(logDir)
                ? logDir
                : Environment.GetEnvironmentVariable(Logger.LogDirectoryEnvironmentVariable);

            if (!string.IsNullOrWhiteSpace(effectiveLogDir))
            {
                string safeLogDir = effectiveLogDir!.Trim().Replace("\"", "\\\"");
                psi.Arguments += $" --log-dir \"{safeLogDir}\"";
                psi.EnvironmentVariables[Logger.LogDirectoryEnvironmentVariable] = effectiveLogDir.Trim();
            }

            string? reqOverride = Environment.GetEnvironmentVariable(AppSettings.RequireOverrideEnvironmentVariable);
            if (!string.IsNullOrWhiteSpace(reqOverride))
            {
                psi.EnvironmentVariables[AppSettings.RequireOverrideEnvironmentVariable] = reqOverride.Trim();
            }

            return psi;
        }

        private void OpenSearchWindowInternal()
        {
            LaunchOrActivateUIProcess();
        }

        private void EnsureEngineInitialized()
        {
            if (_isEngineInitialized && _indexerService != null) return;

            lock (this)
            {
                if (_isEngineInitialized && _indexerService != null) return;

                try
                {
                    Logger.Info("Lazy-initializing FastMail Search engine & database asynchronously...");

                    string dbPath = AppSettings.GetDatabasePath();
                    _dbFactory = new DatabaseConnectionFactory(dbPath);
                    _repository = new SqliteEmailRepository(_dbFactory);
                    _repository.InitializeDatabase();

                    _outlookAdapter = new OutlookAdapter(_application);
                    _licensingService = new LicensingService();
                    _searchService = new SearchService(_repository, _repository, _licensingService);
                    _indexerService = CreateEmailIndexerService(_outlookAdapter, _repository, _licensingService);

                    // Subscribe to real-time Outlook item events for incremental indexing
                    _outlookAdapter.SubscribeToEvents(
                        onItemAddedOrChanged: (storeId, entryId, folderEntryId) =>
                        {
                            if (_shutdownCts.IsCancellationRequested) return;
                            Task.Run(async () =>
                            {
                                try
                                {
                                    if (_shutdownCts.IsCancellationRequested) return;
                                    var svc = _indexerService;
                                    if (svc != null)
                                    {
                                        await svc.QueueItemForIndexingAsync(storeId, entryId, folderEntryId).ConfigureAwait(false);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Logger.Error($"[EVENT ERROR] Failed queuing item {entryId} for indexing: {ex.Message}", ex);
                                }
                            });
                        },
                        onItemRemoved: (storeId, entryId) =>
                        {
                            if (_shutdownCts.IsCancellationRequested) return;
                            Task.Run(async () =>
                            {
                                try
                                {
                                    if (_shutdownCts.IsCancellationRequested) return;
                                    var svc = _indexerService;
                                    if (svc != null)
                                    {
                                        await svc.QueueItemForDeletionAsync(storeId, entryId).ConfigureAwait(false);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Logger.Error($"[EVENT ERROR] Failed queuing item {entryId} for deletion: {ex.Message}", ex);
                                }
                            });
                        }
                    );

                    _isEngineInitialized = true;
                    IsSafeMode = false;
                    SafeModeReason = string.Empty;
                    Logger.Info("FastMail Search engine initialized successfully.");
                    Task.Run(async () =>
                    {
                        try
                        {
                            if (_indexerService != null)
                            {
                                await _indexerService.CheckAndRecoverIncompleteIndexAsync(startProcessing: false).ConfigureAwait(false);
                            }
                        }
                        catch (Exception ex)
                        {
                            Logger.Error("[INIT ERROR] Exception checking/recovering incomplete index: " + ex.Message, ex);
                        }
                    });
                }
                catch (Exception ex)
                {
                    Logger.Critical("FastMail Search started in Safe Mode due to initialization failure.", ex);
                    IsSafeMode = true;
                    SafeModeReason = ex.Message;

                    try
                    {
                        if (_dbFactory == null)
                        {
                            string dbPath = AppSettings.GetDatabasePath();
                            _dbFactory = new DatabaseConnectionFactory(dbPath);
                        }
                        if (_repository == null && _dbFactory != null)
                        {
                            _repository = new SqliteEmailRepository(_dbFactory);
                        }
                        if (_licensingService == null)
                        {
                            _licensingService = new LicensingService();
                        }
                        if (_searchService == null && _repository != null)
                        {
                            _searchService = new SearchService(_repository, _repository, _licensingService);
                        }
                        if (_indexerService != null)
                        {
                            _indexerService.SetSafeMode(ex.Message);
                        }
                    }
                    catch { }

                    _isEngineInitialized = (_indexerService != null);
                }
            }
        }

        public static EmailIndexerService CreateEmailIndexerService(FastMailSearch.Core.Interfaces.IOutlookAdapter outlookAdapter, FastMailSearch.Core.Interfaces.IEmailRepository repository, FastMailSearch.Core.Interfaces.ILicensingService? licensingService = null)
        {
            var type = typeof(EmailIndexerService);

            if (licensingService != null)
            {
                var ctorWithLicense = type.GetConstructor(new[] { typeof(FastMailSearch.Core.Interfaces.IOutlookAdapter), typeof(FastMailSearch.Core.Interfaces.IEmailRepository), typeof(FastMailSearch.Core.Interfaces.ILicensingService) });
                if (ctorWithLicense != null)
                {
                    return (EmailIndexerService)ctorWithLicense.Invoke(new object[] { outlookAdapter, repository, licensingService });
                }
            }

            // Attempt 2-parameter constructor first, fallback to 3-parameter constructor for binary compatibility
            var ctor2 = type.GetConstructor(new[] { typeof(FastMailSearch.Core.Interfaces.IOutlookAdapter), typeof(FastMailSearch.Core.Interfaces.IEmailRepository) });
            if (ctor2 != null)
            {
                return (EmailIndexerService)ctor2.Invoke(new object[] { outlookAdapter, repository });
            }

            var ctor3 = type.GetConstructor(new[] { typeof(FastMailSearch.Core.Interfaces.IOutlookAdapter), typeof(FastMailSearch.Core.Interfaces.IEmailRepository), typeof(bool) });
            if (ctor3 != null)
            {
                return (EmailIndexerService)ctor3.Invoke(new object[] { outlookAdapter, repository, false });
            }

            throw new MissingMethodException("EmailIndexerService constructor not found.");
        }
    }
}
