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
        private static extern bool SetDllDirectory(string lpPathName);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        private const int SW_RESTORE = 9;

        static ThisAddIn()
        {
            try
            {
                string addinFolder = Path.GetDirectoryName(typeof(ThisAddIn).Assembly.Location) ?? string.Empty;
                if (!string.IsNullOrEmpty(addinFolder) && Directory.Exists(addinFolder))
                {
                    SetDllDirectory(addinFolder);
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
                Logger.Info("ThisAddIn OnConnection completed successfully.");
            }
            catch (Exception ex)
            {
                Logger.Error("Exception during ThisAddIn OnConnection", ex);
            }
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            try
            {
                Logger.Info($"ThisAddIn OnDisconnection invoked (RemoveMode={RemoveMode}).");
                _outlookAdapter?.UnsubscribeFromEvents();
                _indexerService?.CancelIndexingAsync();
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
            Logger.Info("ThisAddIn OnBeginShutdown invoked.");
            _indexerService?.CancelIndexingAsync();
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
            Logger.Info("[RIBBON] FastMail Search button clicked");
            LaunchOrActivateUIProcess();
        }

        [ComVisible(true)]
        public void OnBackfillActiveStoreClicked(OfficeCore.IRibbonControl control)
        {
            Logger.Info("[RIBBON] Backfill POP Mailbox button clicked");
            StartInProcessBackfill();
        }

        private readonly object _backfillLock = new object();
        private bool _isBackfillRunning;

        public void StartInProcessBackfill()
        {
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

            Task.Run(async () =>
            {
                try
                {
                    Logger.Info("[IN-PROCESS BACKFILL] Initializing FastMail Search engine...");
                    EnsureEngineInitialized();

                    if (_indexerService == null)
                    {
                        Logger.Error("[IN-PROCESS BACKFILL] Indexer service not initialized.");
                        lock (_backfillLock) { _isBackfillRunning = false; }
                        return;
                    }

                    if (_indexerService.IsRebuilding)
                    {
                        Logger.Info("[IN-PROCESS BACKFILL] Indexer is currently rebuilding.");
                        lock (_backfillLock) { _isBackfillRunning = false; }
                        return;
                    }

                    string dbPath = _repository?.DatabasePath ?? string.Empty;
                    Logger.Info($"[IN-PROCESS BACKFILL] Starting active store backfill on confirmed DB: '{dbPath}'");

                    var summary = await _indexerService.StartActiveStoreBackfillAsync();
                    Logger.Info($"[IN-PROCESS BACKFILL] Complete. Discovered={summary.TotalDiscovered}, NewlyIndexed={summary.NewlyIndexedCount}, Skipped={summary.SkippedCount}, Failed={summary.FailedCount}, Cancelled={summary.IsCancelled}");
                }
                catch (Exception ex)
                {
                    Logger.Error("[IN-PROCESS BACKFILL] Exception during active store backfill", ex);
                }
                finally
                {
                    lock (_backfillLock) { _isBackfillRunning = false; }
                }
            });
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
                            _indexerService.QueueItemForIndexingAsync(storeId, entryId, folderEntryId).ConfigureAwait(false);
                        },
                        onItemRemoved: (storeId, entryId) =>
                        {
                            _indexerService.QueueItemForDeletionAsync(storeId, entryId).ConfigureAwait(false);
                        }
                    );

                    _isEngineInitialized = true;
                    IsSafeMode = false;
                    SafeModeReason = string.Empty;
                    Logger.Info("FastMail Search engine initialized successfully.");
                    _indexerService.CheckAndRecoverIncompleteIndexAsync(startProcessing: false).ConfigureAwait(false);
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
