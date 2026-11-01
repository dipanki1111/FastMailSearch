using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using FastMailSearch.Core.Interfaces;
using FastMailSearch.Core.Models;
using FastMailSearch.Infrastructure.Logging;

namespace FastMailSearch.UI.ViewModels
{
    public class DiagnosticsViewModel : ObservableObject
    {
        private readonly IOutlookAdapter _outlookAdapter;
        private readonly IEmailRepository _repository;
        private DiagnosticsInfo _diagnostics = new DiagnosticsInfo();

        public DiagnosticsInfo Diagnostics { get => _diagnostics; set => SetProperty(ref _diagnostics, value); }
        public string DisplayDbSize => $"{(_diagnostics.DatabaseSizeBytes / (1024.0 * 1024.0)):F2} MB";
        public string DisplayMailboxCounts => $"{_diagnostics.MailboxCount} mailboxes / {_diagnostics.FolderCount} folders";

        public ICommand TestOutlookCommand { get; }
        public ICommand TestDbCommand { get; }
        public ICommand OpenLogsCommand { get; }

        public DiagnosticsViewModel(IOutlookAdapter outlookAdapter, IEmailRepository repository)
        {
            _outlookAdapter = outlookAdapter ?? throw new ArgumentNullException(nameof(outlookAdapter));
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));

            TestOutlookCommand = new RelayCommand(_ => TestOutlook());
            TestDbCommand = new RelayCommand(async _ => await TestDatabaseAsync());
            OpenLogsCommand = new RelayCommand(_ => OpenLogsFolder());

            _ = LoadDiagnosticsAsync();
        }

        private async Task LoadDiagnosticsAsync()
        {
            var diag = new DiagnosticsInfo
            {
                OutlookVersion = _outlookAdapter.GetOutlookVersion(),
                OfficeBitness = _outlookAdapter.GetOfficeBitness(),
                WindowsVersion = Environment.OSVersion.ToString()
            };

            try
            {
                var meta = await _repository.GetIndexMetadataAsync();
                diag.IndexedEmailCount = meta.TotalEmailsIndexed;
                diag.LastIndexingTime = meta.LastIndexedTime;
                diag.IndexStatus = meta.Status;

                string dbPath = !string.IsNullOrWhiteSpace(_repository?.DatabasePath)
                    ? _repository.DatabasePath
                    : AppSettings.GetDatabasePath();
                diag.DatabasePath = dbPath;
                if (File.Exists(dbPath))
                {
                    diag.DatabaseSizeBytes = new FileInfo(dbPath).Length;
                }

                var mailboxes = _outlookAdapter.GetMailboxes();
                diag.MailboxCount = mailboxes.Count;
                var folders = await _repository.GetFoldersAsync();
                diag.FolderCount = folders.Count;
            }
            catch { }

            Diagnostics = diag;
            OnPropertyChanged(nameof(DisplayDbSize));
            OnPropertyChanged(nameof(DisplayMailboxCounts));
        }

        private void TestOutlook()
        {
            bool avail = _outlookAdapter.IsOutlookAvailable();
            string msg = avail
                ? $"Classic Outlook Connection Successful.\nVersion: {_outlookAdapter.GetOutlookVersion()}"
                : "FastMail Search requires Classic Outlook for Windows. Classic Outlook COM is unavailable.";
            MessageBox.Show(msg, "Outlook Diagnostic Test", MessageBoxButton.OK, avail ? MessageBoxImage.Information : MessageBoxImage.Warning);
        }

        private async Task TestDatabaseAsync()
        {
            try
            {
                var meta = await _repository.GetIndexMetadataAsync();
                MessageBox.Show($"Database Test Passed.\nIndexed Emails: {meta.TotalEmailsIndexed:N0}", "Database Diagnostic Test", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Database Test Failed: {ex.Message}", "Database Diagnostic Test", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OpenLogsFolder()
        {
            try
            {
                string dir = Logger.LogDirectory;
                if (Directory.Exists(dir))
                {
                    Process.Start("explorer.exe", dir);
                }
                else
                {
                    MessageBox.Show($"Log directory not found: {dir}", "Diagnostics", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Could not open logs folder: {ex.Message}", "Diagnostics", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
