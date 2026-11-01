using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using FastMailSearch.Core.Interfaces;
using FastMailSearch.Core.Models;
using FastMailSearch.Infrastructure.Logging;

namespace FastMailSearch.UI.Views
{
    public partial class DiagnosticsWindow : Window
    {
        private readonly IIndexHealthService? _healthService;
        private DispatcherTimer? _refreshTimer;
        private bool _isRefreshing = false;

        public DiagnosticsWindow(IIndexHealthService? healthService = null)
        {
            InitializeComponent();
            _healthService = healthService;
            ChkBetaMode.IsChecked = Logger.IsBetaModeEnabled;
            Loaded += DiagnosticsWindow_Loaded;
            Unloaded += DiagnosticsWindow_Unloaded;
            Closed += DiagnosticsWindow_Closed;
        }

        private void DiagnosticsWindow_Unloaded(object sender, RoutedEventArgs e)
        {
            StopTimer();
        }

        private void DiagnosticsWindow_Closed(object? sender, EventArgs e)
        {
            StopTimer();
        }

        private void StopTimer()
        {
            if (_refreshTimer != null)
            {
                _refreshTimer.Stop();
                _refreshTimer = null;
            }
        }

        private async void DiagnosticsWindow_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                string diagIcon = this.Icon == null ? "NULL" : this.Icon.ToString();
                string mainIcon = Application.Current.MainWindow?.Icon == null ? "NULL" : Application.Current.MainWindow.Icon.ToString();
                Logger.Info($"[UI ICON DEBUG] DiagnosticsWindow Icon: {diagIcon} | Main Window Icon: {mainIcon}");

                if (_healthService != null)
                {
                    // 1. Fetch static environment diagnostics
                    try
                    {
                        var env = await _healthService.GetEnvironmentDiagnosticsAsync();
                        if (!string.IsNullOrEmpty(env.DatabasePath))
                        {
                            TxtDbPath.Text = env.DatabasePath;
                        }
                        TxtDbSize.Text = $"{env.DatabaseSizeMb:F2} MB";
                        TxtSqliteVer.Text = env.SqliteVersion;
                        TxtFtsStatus.Text = env.IsFts5Available ? "Available" : "Not Available";
                        TxtAboutBitness.Text = $"{env.OutlookBitness} (Office) / {env.OsArchitecture} (OS)";
                        TxtAboutSqliteVer.Text = env.SqliteVersion;
                        TxtAboutVersion.Text = typeof(DiagnosticsWindow).Assembly.GetName().Version?.ToString(3) ?? "1.0.6";

                        TxtEnvInfo.Text = $"OS: {env.OsVersion} ({env.OsArchitecture})\n" +
                                          $"Database: {env.DatabasePath} ({env.DatabaseSizeMb:F2} MB)\n" +
                                          $"SQLite: {env.SqliteVersion} (FTS5: {(env.IsFts5Available ? "Available" : "Not Available")})\n" +
                                          $"Add-in Version: {env.FastMailSearchVersion}";
                    }
                    catch (Exception envEx)
                    {
                        Logger.Warn($"Failed to load environment diagnostics: {envEx.Message}");
                    }

                    // 2. Perform initial live refresh
                    await RefreshLiveDiagnosticsAsync();

                    // 3. Start auto-refresh timer (continuous live update every 1.5s)
                    StopTimer();
                    _refreshTimer = new DispatcherTimer
                    {
                        Interval = TimeSpan.FromSeconds(1.5)
                    };
                    _refreshTimer.Tick += async (s, args) => await RefreshLiveDiagnosticsAsync();
                    _refreshTimer.Start();
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Error loading DiagnosticsWindow", ex);
            }
        }

        private async Task RefreshLiveDiagnosticsAsync()
        {
            if (_healthService == null || _isRefreshing) return;
            _isRefreshing = true;

            try
            {
                var summary = await _healthService.GetHealthSummaryAsync();

                if (!string.IsNullOrEmpty(summary.DatabasePath))
                {
                    TxtDbPath.Text = summary.DatabasePath;
                }
                TxtDbSize.Text = summary.DisplayDbSize;
                TxtIntegrity.Text = summary.IntegrityResult.IsHealthy ? "Passed (100% Healthy)" : summary.IntegrityResult.SummaryText;

                TxtHealthState.Text = summary.StatusState;
                TxtCurrentOp.Text = summary.CurrentOperation ?? "Idle";

                TxtDiscoveredCount.Text = summary.TotalDiscoveredOutlookItems.ToString("N0");
                TxtIndexedCount.Text = summary.IndexedEmailCount.ToString("N0");
                TxtRemainingCount.Text = summary.RemainingToScan.ToString("N0");
                TxtQueueLength.Text = summary.PendingQueueCount.ToString("N0");
                TxtFailedCount.Text = summary.FailedItemCount.ToString("N0");
                TxtLastIndexTime.Text = summary.DisplayLastUpdate;
                TxtMailboxCoverage.Text = summary.DisplayCoverage;

                // Live continuous progress card
                if (summary.TotalDiscoveredOutlookItems > 0)
                {
                    PrgIndexing.Maximum = summary.TotalDiscoveredOutlookItems;
                    PrgIndexing.Value = summary.ScannedCount;
                    TxtProgressPercentage.Text = $"{summary.ProgressPercentage:F0}%";
                    TxtProgressDetail.Text = $"{summary.CurrentOperation} • {summary.ScannedCount:N0} / {summary.TotalDiscoveredOutlookItems:N0} scanned ({summary.RemainingToScan:N0} remaining)";
                }
                else
                {
                    PrgIndexing.Maximum = 100;
                    PrgIndexing.Value = 0;
                    TxtProgressPercentage.Text = "0%";
                    TxtProgressDetail.Text = summary.CurrentOperation ?? "Idle";
                }

                if (summary.IsSafeMode)
                {
                    SafeModeBanner.Visibility = Visibility.Visible;
                    SafeModeReasonText.Text = $"FastMail Search started in Safe Mode.\nReason: {summary.SafeModeReason}";
                }
                else
                {
                    SafeModeBanner.Visibility = Visibility.Collapsed;
                }
            }
            catch (Exception ex)
            {
                Logger.Warn($"Error during live diagnostics refresh: {ex.Message}");
            }
            finally
            {
                _isRefreshing = false;
            }
        }

        private void BtnCheckUpdates_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show(
                "Update checking will be available in a future release.",
                "Check for Updates",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }

        private async void BtnRepairIndex_Click(object sender, RoutedEventArgs e)
        {
            if (_healthService == null) return;

            var result = MessageBox.Show(
                "This will verify database integrity and safely repair search index records. Proceed?",
                "Repair Search Index",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    var scan = await _healthService.VerifyIndexAsync();
                    await _healthService.RepairIndexAsync(scan);
                    MessageBox.Show("Search index repair completed successfully.", "Repair Complete", MessageBoxButton.OK, MessageBoxImage.Information);
                    DiagnosticsWindow_Loaded(this, new RoutedEventArgs());
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Index repair failed: {ex.Message}", "Repair Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private async void BtnExportZip_Click(object sender, RoutedEventArgs e)
        {
            if (_healthService == null) return;

            try
            {
                string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string zipPath = await _healthService.ExportDiagnosticsZipAsync(desktop);
                MessageBox.Show($"Diagnostics package exported successfully:\n\n{zipPath}\n\nNo email content or personal data is included.", "Export Diagnostics", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to export diagnostics: {ex.Message}", "Export Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnOpenLogs_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string logDir = Logger.LogDirectory;
                if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                Process.Start("explorer.exe", logDir);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Could not open logs directory: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ChkBetaMode_Changed(object sender, RoutedEventArgs e)
        {
            Logger.IsBetaModeEnabled = ChkBetaMode.IsChecked == true;
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
