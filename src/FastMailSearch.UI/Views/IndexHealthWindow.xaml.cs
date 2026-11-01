using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using FastMailSearch.Core.Interfaces;
using FastMailSearch.Core.Models;
using FastMailSearch.UI.ViewModels;

namespace FastMailSearch.UI.Views
{
    public partial class IndexHealthWindow : Window, INotifyPropertyChanged
    {
        private readonly IIndexHealthService _healthService;
        private readonly IEmailRepository _repository;
        private readonly IIndexerService _indexerService;
        private DispatcherTimer? _refreshTimer;

        private IndexHealthSummary _summary = new IndexHealthSummary();
        private string _statusMessage = "Ready";
        private bool _isBusy;

        public event PropertyChangedEventHandler? PropertyChanged;

        public IndexHealthWindow(IIndexHealthService healthService, IEmailRepository repository, IIndexerService indexerService)
        {
            InitializeComponent();
            _healthService = healthService ?? throw new ArgumentNullException(nameof(healthService));
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
            _indexerService = indexerService ?? throw new ArgumentNullException(nameof(indexerService));

            DataContext = this;

            BackfillStoreCommand = new RelayCommand(async _ => await ExecuteBackfillStoreAsync(), _ => !_isBusy && !_indexerService.IsIndexing);
            VerifyIndexCommand = new RelayCommand(async _ => await ExecuteVerifyIndexAsync(), _ => !_isBusy);
            RetryFailedCommand = new RelayCommand(async _ => await ExecuteRetryFailedAsync(), _ => !_isBusy);
            ViewFailedCommand = new RelayCommand(_ => ExecuteViewFailed());

            _indexerService.ProgressChanged += (s, e) =>
            {
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    if (IsBusy)
                    {
                        StatusMessage = e.CurrentStatusMessage;
                    }
                }));
            };

            Loaded += async (s, e) =>
            {
                await RefreshSummaryAsync();
                _refreshTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(2.0) };
                _refreshTimer.Tick += async (timerSender, timerArgs) =>
                {
                    if (!_isBusy)
                    {
                        await RefreshSummaryAsync();
                    }
                };
                _refreshTimer.Start();
            };

            Closed += (s, e) =>
            {
                _refreshTimer?.Stop();
                _refreshTimer = null;
            };
        }

        public IndexHealthSummary Summary
        {
            get => _summary;
            set { _summary = value; OnPropertyChanged(); }
        }

        public string StatusMessage
        {
            get => _statusMessage;
            set { _statusMessage = value; OnPropertyChanged(); }
        }

        public bool IsBusy
        {
            get => _isBusy;
            set { _isBusy = value; OnPropertyChanged(); CommandManagerInvalidate(); }
        }

        public RelayCommand BackfillStoreCommand { get; }
        public RelayCommand VerifyIndexCommand { get; }
        public RelayCommand RetryFailedCommand { get; }
        public RelayCommand ViewFailedCommand { get; }

        private async Task RefreshSummaryAsync()
        {
            try
            {
                Summary = await _healthService.GetHealthSummaryAsync();
            }
            catch (Exception ex)
            {
                StatusMessage = $"Error refreshing health summary: {ex.Message}";
            }
        }

        private async Task ExecuteBackfillStoreAsync()
        {
            IsBusy = true;
            StatusMessage = "🔵 Initiating Outlook mailbox backfill in background...";
            try
            {
                var summary = await Task.Run(async () => await _healthService.StartActiveStoreBackfillAsync());
                if (summary.IsCancelled)
                {
                    StatusMessage = "🟡 Backfill was cancelled.";
                }
                else if (!string.IsNullOrEmpty(summary.ErrorMessage))
                {
                    StatusMessage = $"🔴 Backfill error: {summary.ErrorMessage}";
                }
                else
                {
                    StatusMessage = $"🟢 Backfill Complete: {summary.NewlyIndexedCount:N0} newly indexed emails.";
                }
                await RefreshSummaryAsync();
            }
            catch (Exception ex)
            {
                StatusMessage = $"🔴 Backfill error: {ex.Message}";
            }
            finally
            {
                IsBusy = false;
            }
        }

        private async Task ExecuteVerifyIndexAsync()
        {
            IsBusy = true;
            StatusMessage = "🟡 Verification in progress...";
            try
            {
                var result = await _healthService.VerifyIndexAsync();
                if (result.IsHealthy)
                {
                    StatusMessage = "🟢 Verification Complete: Index is healthy.";
                }
                else
                {
                    StatusMessage = $"🟡 Self-Healing Repairing: {result.MissingFtsCount} missing FTS, {result.OrphanFtsCount} orphan FTS, {result.DuplicateCount} duplicates";
                    await _healthService.RepairIndexAsync(result);
                    StatusMessage = "🟢 Repair Completed Successfully.";
                }
                await RefreshSummaryAsync();
            }
            catch (Exception ex)
            {
                StatusMessage = $"🔴 Verification error: {ex.Message}";
            }
            finally
            {
                IsBusy = false;
            }
        }

        private async Task ExecuteRetryFailedAsync()
        {
            IsBusy = true;
            StatusMessage = "🔵 Retrying failed items in background...";
            try
            {
                int retriedCount = await _healthService.RetryAllFailedItemsAsync();
                if (retriedCount > 0)
                {
                    StatusMessage = "🟢 Retry processing initiated.";
                }
                else
                {
                    StatusMessage = "🟢 No failed items to retry.";
                }
                await RefreshSummaryAsync();
            }
            catch (Exception ex)
            {
                StatusMessage = $"🔴 Retry error: {ex.Message}";
            }
            finally
            {
                IsBusy = false;
            }
        }

        private void ExecuteViewFailed()
        {
            var window = new FailedItemsWindow(_healthService, _repository, _indexerService)
            {
                Owner = this
            };
            window.ShowDialog();
            _ = RefreshSummaryAsync();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private void CommandManagerInvalidate()
        {
            System.Windows.Input.CommandManager.InvalidateRequerySuggested();
        }
    }
}
