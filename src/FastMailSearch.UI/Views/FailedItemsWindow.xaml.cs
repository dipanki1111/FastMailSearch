using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows;
using FastMailSearch.Core.Interfaces;
using FastMailSearch.Core.Models;
using FastMailSearch.UI.ViewModels;

namespace FastMailSearch.UI.Views
{
    public partial class FailedItemsWindow : Window, INotifyPropertyChanged
    {
        private readonly IIndexHealthService _healthService;
        private readonly IEmailRepository _repository;
        private readonly IIndexerService _indexerService;

        private ObservableCollection<FailedIndexItem> _failedItems = new ObservableCollection<FailedIndexItem>();
        private string _statusMessage = "Ready";
        private bool _isBusy;

        public event PropertyChangedEventHandler? PropertyChanged;

        public FailedItemsWindow(IIndexHealthService healthService, IEmailRepository repository, IIndexerService indexerService)
        {
            InitializeComponent();
            _healthService = healthService ?? throw new ArgumentNullException(nameof(healthService));
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
            _indexerService = indexerService ?? throw new ArgumentNullException(nameof(indexerService));

            DataContext = this;

            SelectAllCommand = new RelayCommand(_ => ExecuteSelectAll());
            RetrySelectedCommand = new RelayCommand(async _ => await ExecuteRetrySelectedAsync(), _ => !_isBusy && FailedDataGrid.SelectedItems.Count > 0);
            RemoveSelectedCommand = new RelayCommand(async _ => await ExecuteRemoveSelectedAsync(), _ => !_isBusy && FailedDataGrid.SelectedItems.Count > 0);
            RetryAllCommand = new RelayCommand(async _ => await ExecuteRetryAllAsync(), _ => !_isBusy && FailedItems.Count > 0);
            ClearAllCommand = new RelayCommand(async _ => await ExecuteClearAllAsync(), _ => !_isBusy && FailedItems.Count > 0);

            Loaded += async (s, e) => await RefreshItemsAsync();
        }

        public ObservableCollection<FailedIndexItem> FailedItems
        {
            get => _failedItems;
            set { _failedItems = value; OnPropertyChanged(); }
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

        public RelayCommand SelectAllCommand { get; }
        public RelayCommand RetrySelectedCommand { get; }
        public RelayCommand RemoveSelectedCommand { get; }
        public RelayCommand RetryAllCommand { get; }
        public RelayCommand ClearAllCommand { get; }

        private async Task RefreshItemsAsync()
        {
            try
            {
                var items = await _healthService.GetFailedItemsAsync();
                int totalFailedCount = await _healthService.GetFailedItemsCountAsync();
                FailedItems = new ObservableCollection<FailedIndexItem>(items);
                if (FailedItems.Count < totalFailedCount)
                {
                    StatusMessage = $"Showing {FailedItems.Count:N0} of {totalFailedCount:N0} failed item(s).";
                }
                else
                {
                    StatusMessage = $"{FailedItems.Count:N0} failed item(s) found.";
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"Error loading failed items: {ex.Message}";
            }
        }

        private void ExecuteSelectAll()
        {
            FailedDataGrid.SelectAll();
        }

        private async Task ExecuteRetrySelectedAsync()
        {
            var selected = FailedDataGrid.SelectedItems.Cast<FailedIndexItem>().Select(i => i.Id).ToList();
            if (!selected.Any()) return;

            IsBusy = true;
            StatusMessage = $"Retrying {selected.Count} selected item(s)...";
            try
            {
                await _healthService.RetryFailedItemsBatchAsync(selected);
                StatusMessage = $"Re-enqueued {selected.Count} item(s) for indexing.";
                await RefreshItemsAsync();
            }
            catch (Exception ex)
            {
                StatusMessage = $"Error retrying selected items: {ex.Message}";
            }
            finally
            {
                IsBusy = false;
            }
        }

        private async Task ExecuteRemoveSelectedAsync()
        {
            var selected = FailedDataGrid.SelectedItems.Cast<FailedIndexItem>().Select(i => i.Id).ToList();
            if (!selected.Any()) return;

            IsBusy = true;
            StatusMessage = $"Removing {selected.Count} selected failure record(s)...";
            try
            {
                await _healthService.RemoveFailedItemsBatchAsync(selected);
                StatusMessage = $"Removed {selected.Count} item(s).";
                await RefreshItemsAsync();
            }
            catch (Exception ex)
            {
                StatusMessage = $"Error removing selected items: {ex.Message}";
            }
            finally
            {
                IsBusy = false;
            }
        }

        private async Task ExecuteRetryAllAsync()
        {
            IsBusy = true;
            StatusMessage = "Retrying all failed items...";
            try
            {
                await _healthService.RetryAllFailedItemsAsync();
                StatusMessage = "Re-enqueued all failed items for indexing.";
                await RefreshItemsAsync();
            }
            catch (Exception ex)
            {
                StatusMessage = $"Error retrying all items: {ex.Message}";
            }
            finally
            {
                IsBusy = false;
            }
        }

        private async Task ExecuteClearAllAsync()
        {
            IsBusy = true;
            StatusMessage = "Clearing all failed items...";
            try
            {
                await _healthService.RemoveAllFailedItemsAsync();
                StatusMessage = "Cleared all failed item records.";
                await RefreshItemsAsync();
            }
            catch (Exception ex)
            {
                StatusMessage = $"Error clearing all items: {ex.Message}";
            }
            finally
            {
                IsBusy = false;
            }
        }

        private void CopyEntryId_Click(object sender, RoutedEventArgs e)
        {
            if (FailedDataGrid.SelectedItem is FailedIndexItem item && !string.IsNullOrEmpty(item.EntryId))
            {
                Clipboard.SetText(item.EntryId);
                StatusMessage = "Copied Entry ID to clipboard.";
            }
        }

        private void CopyError_Click(object sender, RoutedEventArgs e)
        {
            if (FailedDataGrid.SelectedItem is FailedIndexItem item && !string.IsNullOrEmpty(item.ErrorMessage))
            {
                Clipboard.SetText(item.ErrorMessage);
                StatusMessage = "Copied error details to clipboard.";
            }
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
