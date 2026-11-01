using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;
using FastMailSearch.Core.Models;
using FastMailSearch.Core.Services;
using FastMailSearch.Infrastructure.Database;
using FastMailSearch.Infrastructure.Indexing;
using FastMailSearch.Infrastructure.Logging;
using FastMailSearch.Infrastructure.Outlook;
using FastMailSearch.Infrastructure.Services;
using FastMailSearch.UI.ViewModels;
using FastMailSearch.UI.Views;

namespace FastMailSearch.UI
{
    public class StartupArguments
    {
        public string? ExplicitDbPath { get; set; }
        public string? ExplicitLogDir { get; set; }
        public bool HasInvalidArguments { get; set; }
        public string? ErrorMessage { get; set; }
    }

    public static class CommandLineParser
    {
        public static StartupArguments ParseCommandLineArgs(string[] args)
        {
            var result = new StartupArguments();
            if (args == null || args.Length == 0) return result;

            for (int i = 0; i < args.Length; i++)
            {
                string arg = args[i];
                if (string.Equals(arg, "--db-path", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(arg, "-d", StringComparison.OrdinalIgnoreCase))
                {
                    if (i + 1 < args.Length && !args[i + 1].StartsWith("-"))
                    {
                        result.ExplicitDbPath = args[++i].Trim();
                    }
                    else
                    {
                        result.HasInvalidArguments = true;
                        result.ErrorMessage = "Missing path value after --db-path argument.";
                        return result;
                    }
                }
                else if (string.Equals(arg, "--log-dir", StringComparison.OrdinalIgnoreCase) ||
                         string.Equals(arg, "-l", StringComparison.OrdinalIgnoreCase))
                {
                    if (i + 1 < args.Length && !args[i + 1].StartsWith("-"))
                    {
                        result.ExplicitLogDir = args[++i].Trim();
                    }
                    else
                    {
                        result.HasInvalidArguments = true;
                        result.ErrorMessage = "Missing directory value after --log-dir argument.";
                        return result;
                    }
                }
                else
                {
                    result.HasInvalidArguments = true;
                    result.ErrorMessage = $"Unrecognized argument: '{arg}'.";
                    return result;
                }
            }

            return result;
        }

        public static bool ValidateStartupDatabasePath(string? explicitDbPath, out string resolvedDbPath, out string? errorMessage)
        {
            errorMessage = null;
            try
            {
                resolvedDbPath = AppSettings.GetDatabasePath();
            }
            catch (Exception ex)
            {
                resolvedDbPath = string.Empty;
                errorMessage = ex.Message;
                return false;
            }

            if (!string.IsNullOrWhiteSpace(explicitDbPath))
            {
                string normExplicit = AppSettings.NormalizePath(explicitDbPath!);
                string normResolved = AppSettings.NormalizePath(resolvedDbPath);

                if (!string.Equals(normExplicit, normResolved, StringComparison.OrdinalIgnoreCase))
                {
                    errorMessage = $"Database path mismatch! Explicit argument is '{explicitDbPath}', but resolved canonical path is '{resolvedDbPath}'.";
                    return false;
                }
            }

            return true;
        }
    }

    public partial class App : Application
    {
        public static StartupArguments ParseCommandLineArgs(string[] args) => CommandLineParser.ParseCommandLineArgs(args);

        public static bool ValidateStartupDatabasePath(string? explicitDbPath, out string resolvedDbPath, out string? errorMessage) =>
            CommandLineParser.ValidateStartupDatabasePath(explicitDbPath, out resolvedDbPath, out errorMessage);

        static App()
        {
            // Register AssemblyResolve to load dependencies from AppContext.BaseDirectory
            AppDomain.CurrentDomain.AssemblyResolve += (sender, args) =>
            {
                try
                {
                    string asmName = new AssemblyName(args.Name).Name + ".dll";
                    string appFolder = AppContext.BaseDirectory;
                    string asmPath = Path.Combine(appFolder, asmName);
                    if (File.Exists(asmPath))
                    {
                        return Assembly.LoadFrom(asmPath);
                    }
                }
                catch { }
                return null;
            };
        }

        private static Mutex? _singleInstanceMutex;
        private const string MutexName = @"Global\FastMailSearch_UI_SingleInstance";

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool IsIconic(IntPtr hWnd);

        private const int SW_RESTORE = 9;
        private const int SW_SHOW = 5;

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            try
            {
                var startupArgs = ParseCommandLineArgs(e.Args);
                if (startupArgs.HasInvalidArguments)
                {
                    Logger.Critical($"[UI STARTUP REJECTED] Invalid arguments: {startupArgs.ErrorMessage}");
                    MessageBox.Show($"FastMail Search Startup Error:\n{startupArgs.ErrorMessage}", "FastMail Search", MessageBoxButton.OK, MessageBoxImage.Error);
                    Shutdown(1);
                    return;
                }

                if (!string.IsNullOrWhiteSpace(startupArgs.ExplicitLogDir))
                {
                    Logger.Initialize(startupArgs.ExplicitLogDir);
                }

                if (!string.IsNullOrWhiteSpace(startupArgs.ExplicitDbPath))
                {
                    Environment.SetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable, startupArgs.ExplicitDbPath);
                }

                if (!ValidateStartupDatabasePath(startupArgs.ExplicitDbPath, out string dbPath, out string? pathError))
                {
                    Logger.Critical($"[UI STARTUP REJECTED] {pathError}");
                    MessageBox.Show($"FastMail Search Startup Error:\n{pathError}\nStartup aborted for data safety.", "FastMail Search", MessageBoxButton.OK, MessageBoxImage.Error);
                    Shutdown(1);
                    return;
                }

                Logger.LogStage(14, "WPF Application initialized");
                Logger.Info("[UI LAUNCH] FastMail Search UI Application starting...");

                bool createdNew = true;
                try
                {
                    _singleInstanceMutex = new Mutex(true, MutexName, out createdNew);
                }
                catch (Exception mutexEx)
                {
                    Logger.Warn($"[UI SINGLE INSTANCE] Exception creating mutex: {mutexEx.Message}");
                    var procs = Process.GetProcessesByName("FastMailSearch.UI");
                    if (procs.Length > 1) createdNew = false;
                }

                if (!createdNew)
                {
                    Logger.Info("[UI SINGLE INSTANCE] Existing instance detected");
                    Logger.Info("[UI SINGLE INSTANCE] Activation signal sent");
                    ActivateExistingInstance();
                    Logger.Info("[UI SINGLE INSTANCE] Exiting secondary process");
                    Shutdown(0);
                    return;
                }

                Logger.Info("[UI LAUNCH] No existing UI instance");
                Logger.Info($"[UI LAUNCH] Starting FastMailSearch.UI.exe PID={Process.GetCurrentProcess().Id}");

                Logger.Info($"[UI LAUNCH] Target database path: '{dbPath}' (Exists={File.Exists(dbPath)})");
                var factory = new DatabaseConnectionFactory(dbPath);
                var repo = new SqliteEmailRepository(factory);
                repo.InitializeDatabase();

                var licensingService = new LicensingService();
                var searchService = new SearchService(repo, repo, licensingService);
                var outlookAdapter = new OutlookAdapter();
                var indexerService = new EmailIndexerService(outlookAdapter, repo, licensingService);
                
                // Asynchronously recover any stale processing items without spawning background queue worker in UI process
                _ = Task.Run(async () => await indexerService.CheckAndRecoverIncompleteIndexAsync(startProcessing: false));

                var vm = new SearchViewModel(searchService, repo, indexerService, outlookAdapter, licensingService, factory);

                Logger.LogStage(18, "SearchWindow.Show entered");
                var mainWin = new SearchWindow(vm);

                MainWindow = mainWin;
                mainWin.Show();
                Logger.LogStage(19, "SearchWindow.Show completed");
                Logger.Info("[UI LAUNCH] UI startup completed");

                IntPtr hWnd = new WindowInteropHelper(mainWin).Handle;
                var workArea = SystemParameters.WorkArea;

                Logger.LogStage(22, $"Window IsVisible: {mainWin.IsVisible}");
                Logger.LogStage(23, $"Window Visibility: {mainWin.Visibility}");
                Logger.LogStage(24, $"Window WindowState: {mainWin.WindowState}");
                Logger.LogStage(25, $"Window Handle: {hWnd}");
                Logger.Info($"Window Bounds: Left={mainWin.Left}, Top={mainWin.Top}, Width={mainWin.Width}, Height={mainWin.Height}, ScreenWorkArea={workArea}");

                // Explicit window visibility and activation sequence
                mainWin.Activate();
                mainWin.BringIntoView();
                mainWin.Topmost = true;
                mainWin.Topmost = false;
                mainWin.Focus();

                Logger.LogStage(20, "Dispatcher.Run entered");
            }
            catch (Exception ex)
            {
                Logger.LogStage(99, "UI Application startup exception", ex);
                Logger.Error("[UI ERROR] FastMailSearch.UI App.OnStartup failed", ex);
                MessageBox.Show($"FastMail Search UI Startup Error:\n{ex.Message}", "FastMail Search", MessageBoxButton.OK, MessageBoxImage.Error);
                Shutdown(1);
            }
        }

        public static bool ActivateExistingInstance()
        {
            try
            {
                int currentPid = Process.GetCurrentProcess().Id;
                for (int attempt = 0; attempt < 5; attempt++)
                {
                    var processes = Process.GetProcessesByName("FastMailSearch.UI");
                    foreach (var proc in processes)
                    {
                        if (proc.Id != currentPid)
                        {
                            IntPtr hWnd = proc.MainWindowHandle;
                            if (hWnd != IntPtr.Zero)
                            {
                                Logger.Info($"[UI ACTIVATE] Activating existing UI window handle: {hWnd}");
                                if (IsIconic(hWnd))
                                {
                                    ShowWindow(hWnd, SW_RESTORE);
                                }
                                else
                                {
                                    ShowWindow(hWnd, SW_SHOW);
                                }
                                SetForegroundWindow(hWnd);
                                return true;
                            }
                        }
                    }
                    System.Threading.Thread.Sleep(100);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("[UI ACTIVATE] Failed activating existing instance window", ex);
            }
            return false;
        }

        protected override void OnExit(ExitEventArgs e)
        {
            try
            {
                if (_singleInstanceMutex != null)
                {
                    _singleInstanceMutex.ReleaseMutex();
                    _singleInstanceMutex.Dispose();
                    _singleInstanceMutex = null;
                }
            }
            catch { }
            base.OnExit(e);
        }
    }
}
