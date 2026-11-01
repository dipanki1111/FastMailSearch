using System;
using System.Windows;
using FastMailSearch.Core.Interfaces;
using FastMailSearch.Core.Models;
using FastMailSearch.Infrastructure.Services;

namespace FastMailSearch.UI.Views
{
    public partial class FirstRunWindow : Window
    {
        private readonly IIndexerService? _indexerService;
        private readonly IOutlookAdapter? _outlookAdapter;
        private bool _isStarted = false;

        public FirstRunWindow()
        {
            InitializeComponent();
        }

        public FirstRunWindow(IIndexerService indexerService, IOutlookAdapter outlookAdapter) : this()
        {
            _indexerService = indexerService;
            _outlookAdapter = outlookAdapter;

            if (_indexerService != null)
            {
                _indexerService.ProgressChanged += IndexerService_ProgressChanged;
            }
        }

        private async void BtnStartIndexing_Click(object sender, RoutedEventArgs e)
        {
            if (_isStarted) return;

            _isStarted = true;
            IntroPanel.Visibility = Visibility.Collapsed;
            ProgressPanel.Visibility = Visibility.Visible;
            BtnStartIndexing.Visibility = Visibility.Collapsed;
            BtnPause.Visibility = Visibility.Visible;

            if (_outlookAdapter != null)
            {
                var mailboxes = _outlookAdapter.GetMailboxes();
                TxtMailboxes.Text = $"Mailboxes: {mailboxes.Count} connected";
            }

            if (_indexerService != null)
            {
                try
                {
                    await _indexerService.StartInitialIndexingAsync();
                }
                catch (Exception ex)
                {
                    TxtStatus.Text = $"Setup indexing notice: {ex.Message}";
                }
            }

            BtnClose.Visibility = Visibility.Visible;
            BtnPause.Visibility = Visibility.Collapsed;
            TxtStatus.Text = "Initial indexing step completed. Background indexing is active.";
        }

        private async void BtnPause_Click(object sender, RoutedEventArgs e)
        {
            if (_indexerService != null)
            {
                if (_indexerService.CurrentStatus == IndexerStatus.Paused)
                {
                    await _indexerService.ResumeIndexingAsync();
                    BtnPause.Content = "Pause";
                }
                else
                {
                    await _indexerService.PauseIndexingAsync();
                    BtnPause.Content = "Resume";
                }
            }
        }

        private void IndexerService_ProgressChanged(object? sender, IndexerProgressEventArgs e)
        {
            Dispatcher.Invoke(() =>
            {
                ProgressInitial.Value = e.Percentage;
                TxtStatus.Text = $"Indexing... {e.Percentage:F0}% ({e.ProcessedItems:N0} / {e.TotalItems:N0})";
                TxtCount.Text = $"Emails: {e.ProcessedItems:N0} / {e.TotalItems:N0}";
                TxtCurrentFolder.Text = $"Folder: {e.CurrentFolder}";
            });
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
