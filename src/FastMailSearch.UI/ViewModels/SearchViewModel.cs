using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using FastMailSearch.Core.Interfaces;
using FastMailSearch.Core.Models;
using FastMailSearch.Core.Search;
using FastMailSearch.Core.Services;
using FastMailSearch.Infrastructure.Database;
using FastMailSearch.Infrastructure.Services;
using FastMailSearch.Infrastructure.Logging;

namespace FastMailSearch.UI.ViewModels
{
    public class SearchViewModel : ObservableObject
    {
        private readonly SearchService _searchService;
        private readonly IEmailRepository _repository;
        private readonly IIndexerService _indexerService;
        private readonly IOutlookAdapter _outlookAdapter;
        private readonly ILicensingService? _licensingService;

        private string _searchText = string.Empty;
        private string _selectedSearchHistory = string.Empty;
        private string _fromFilter = string.Empty;
        private string _toFilter = string.Empty;
        private string _subjectFilter = string.Empty;
        private DateTime? _dateFrom;
        private DateTime? _dateTo;
        private string _selectedFolder = "All Folders";
        private string _selectedAttachmentOption = "All";
        private bool _isSearching;

        private int _totalResultsCount;
        private string _searchTimeText = string.Empty;
        private string _statusText = "Up to date";
        private bool _isIndexing;

        private int _indexProcessedItems;
        private int _indexTotalItems;
        private int _indexFolderProcessedItems;
        private int _indexFolderTotalItems;
        private string _indexCurrentFolder = string.Empty;
        private double _indexProgressPercentage;

        private int _currentPage = 1;
        private int _pageSize = 50;
        private CancellationTokenSource? _searchCts;
        private readonly object _lock = new object();

        private bool _isActivationOverlayVisible;
        public bool IsActivationOverlayVisible { get => _isActivationOverlayVisible; set => SetProperty(ref _isActivationOverlayVisible, value); }

        private bool _isTrialWarningVisible;
        public bool IsTrialWarningVisible { get => _isTrialWarningVisible; set => SetProperty(ref _isTrialWarningVisible, value); }

        private bool _isRenewalWarningVisible;
        public bool IsRenewalWarningVisible { get => _isRenewalWarningVisible; set => SetProperty(ref _isRenewalWarningVisible, value); }

        private string _trialDaysRemainingText = string.Empty;
        public string TrialDaysRemainingText { get => _trialDaysRemainingText; set => SetProperty(ref _trialDaysRemainingText, value); }

        private string _renewalWarningText = string.Empty;
        public string RenewalWarningText { get => _renewalWarningText; set => SetProperty(ref _renewalWarningText, value); }

        private string _licenseKeyInput = string.Empty;
        public string LicenseKeyInput { get => _licenseKeyInput; set => SetProperty(ref _licenseKeyInput, value); }

        private string _activationErrorMessage = string.Empty;
        public string ActivationErrorMessage { get => _activationErrorMessage; set => SetProperty(ref _activationErrorMessage, value); }

        private string _hardwareId = string.Empty;
        public string HardwareId { get => _hardwareId; private set => SetProperty(ref _hardwareId, value); }

        private string _copyHardwareIdStatusMessage = string.Empty;
        public string CopyHardwareIdStatusMessage { get => _copyHardwareIdStatusMessage; set => SetProperty(ref _copyHardwareIdStatusMessage, value); }

        private string _copySupportEmailStatusMessage = string.Empty;
        public string CopySupportEmailStatusMessage { get => _copySupportEmailStatusMessage; set => SetProperty(ref _copySupportEmailStatusMessage, value); }

        public ICommand CopyHardwareIdCommand { get; }
        public ICommand CopySupportEmailCommand { get; }
        public ICommand OpenProjectRepoCommand { get; }
        public ICommand ActivateCommand { get; }

        public ObservableCollection<SearchResultItem> SearchResults { get; } = new ObservableCollection<SearchResultItem>();
        public ObservableCollection<string> FolderOptions { get; } = new ObservableCollection<string> { "All Folders" };
        public ObservableCollection<string> AttachmentOptions { get; } = new ObservableCollection<string> { "All", "With Attachments", "Without Attachments" };
        public ObservableCollection<int> PageSizeOptions { get; } = new ObservableCollection<int> { 25, 50, 100 };

        public ICommand SearchCommand { get; }
        public ICommand ClearFiltersCommand { get; }
        public ICommand OpenResultCommand { get; }
        public ObservableCollection<string> SortOrderOptions { get; } = new ObservableCollection<string> { "Relevance", "Newest First", "Oldest First" };
        private string _selectedSortOption = "Relevance";
        public string SelectedSortOption
        {
            get => _selectedSortOption;
            set
            {
                if (SetProperty(ref _selectedSortOption, value))
                {
                    _ = PerformSearchAsync(resetPage: true);
                }
            }
        }

        public ObservableCollection<string> SearchHistoryList { get; } = new ObservableCollection<string>();

        private bool _isInvalidSyntax;
        public bool IsInvalidSyntax { get => _isInvalidSyntax; set => SetProperty(ref _isInvalidSyntax, value); }

        private string _syntaxErrorMessage = string.Empty;
        public string SyntaxErrorMessage { get => _syntaxErrorMessage; set => SetProperty(ref _syntaxErrorMessage, value); }

        private bool _isSafeMode;
        public bool IsSafeMode { get => _isSafeMode; set => SetProperty(ref _isSafeMode, value); }

        private string _safeModeReason = string.Empty;
        public string SafeModeReason { get => _safeModeReason; set => SetProperty(ref _safeModeReason, value); }

        public ICommand RebuildIndexCommand { get; }
        public ICommand CancelIndexCommand { get; }
        public ICommand TogglePauseResumeCommand { get; }
        public ICommand BackfillActiveStoreCommand { get; }
        public ICommand CopySubjectCommand { get; }
        public ICommand CopySenderCommand { get; }
        public ICommand CopyEmailAddressCommand { get; }
        public ICommand CopyEntryIdCommand { get; }
        public ICommand CopySearchResultCommand { get; }
        public ICommand SelectHistoryQueryCommand { get; }

        public ICommand FirstPageCommand { get; }
        public ICommand PreviousPageCommand { get; }
        public ICommand NextPageCommand { get; }
        public ICommand LastPageCommand { get; }
        public ICommand ChangePageSizeCommand { get; }

        public IIndexHealthService HealthService { get; }
        public IEmailRepository Repository => _repository;
        public IIndexerService IndexerService => _indexerService;
        private readonly DatabaseConnectionFactory? _dbFactory;
        public Action<string, string>? MessageBoxAction { get; set; }

        public SearchViewModel(SearchService searchService, IEmailRepository repository, IIndexerService indexerService, IOutlookAdapter outlookAdapter, ILicensingService? licensingService = null, DatabaseConnectionFactory? dbFactory = null)
        {
            _searchService = searchService ?? throw new ArgumentNullException(nameof(searchService));
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
            _indexerService = indexerService ?? throw new ArgumentNullException(nameof(indexerService));
            _outlookAdapter = outlookAdapter ?? throw new ArgumentNullException(nameof(outlookAdapter));
            _licensingService = licensingService;
            _dbFactory = dbFactory ?? (repository as SqliteEmailRepository)?.ConnectionFactory;
            HealthService = new IndexHealthService(_repository, _indexerService, _outlookAdapter, _dbFactory);

            SearchCommand = new RelayCommand(async _ => await PerformSearchAsync(resetPage: true));
            ClearFiltersCommand = new RelayCommand(_ => ClearFilters());
            OpenResultCommand = new RelayCommand(param => OpenEmail(param as SearchResultItem));
            RebuildIndexCommand = new RelayCommand(async _ => await RebuildIndexAsync());
            CancelIndexCommand = new RelayCommand(async _ => await CancelIndexAsync());
            TogglePauseResumeCommand = new RelayCommand(async _ => await TogglePauseResumeAsync());
            BackfillActiveStoreCommand = new RelayCommand(async _ => await BackfillActiveStoreAsync());

            CopySubjectCommand = new RelayCommand(param => CopyToClipboard((param as SearchResultItem)?.Subject));
            CopySenderCommand = new RelayCommand(param => CopyToClipboard((param as SearchResultItem)?.SenderName));
            CopyEmailAddressCommand = new RelayCommand(param => CopyToClipboard((param as SearchResultItem)?.SenderEmail));
            CopyEntryIdCommand = new RelayCommand(param => CopyToClipboard((param as SearchResultItem)?.EntryId));
            CopySearchResultCommand = new RelayCommand(param => CopySearchResultDetails(param as SearchResultItem));
            SelectHistoryQueryCommand = new RelayCommand(param => { if (param is string q) SearchText = q; });
            ActivateCommand = new RelayCommand(async _ => await ActivateLicenseAsync());
            HardwareId = _licensingService?.GetHardwareId() ?? "Unavailable";
            CopyHardwareIdCommand = new RelayCommand(_ => CopyHardwareIdToClipboard());
            CopySupportEmailCommand = new RelayCommand(_ => CopySupportEmailToClipboard());
            OpenProjectRepoCommand = new RelayCommand(_ => OpenProjectRepository());

            FirstPageCommand = new RelayCommand(async _ => await GoToPageAsync(1), _ => HasPreviousPage);
            PreviousPageCommand = new RelayCommand(async _ => await GoToPageAsync(CurrentPage - 1), _ => HasPreviousPage);
            NextPageCommand = new RelayCommand(async _ => await GoToPageAsync(CurrentPage + 1), _ => HasNextPage);
            LastPageCommand = new RelayCommand(async _ => await GoToPageAsync(TotalPages), _ => HasNextPage);
            ChangePageSizeCommand = new RelayCommand(param =>
            {
                if (param != null && int.TryParse(param.ToString(), out int sz))
                {
                    PageSize = sz;
                }
            });

            _indexerService.ProgressChanged += OnIndexerProgressChanged;
            _ = LoadFoldersAsync();
            _ = LoadSearchHistoryAsync();
            UpdateStatusDisplay();
            UpdateLicensingState();

            _livePollingTimer = new System.Timers.Timer(2000);
            _livePollingTimer.AutoReset = true;
            _livePollingTimer.Elapsed += (s, e) => { _ = PollLiveStatusAsync(); };
            _livePollingTimer.Start();
        }

        private async Task ActivateLicenseAsync()
        {
            if (_licensingService == null || string.IsNullOrWhiteSpace(LicenseKeyInput)) return;
            
            CopyHardwareIdStatusMessage = string.Empty;
            ActivationErrorMessage = "Activating...";
            bool success = await _licensingService.ActivateAsync(LicenseKeyInput.Trim());
            
            if (success)
            {
                ActivationErrorMessage = string.Empty;
                LicenseKeyInput = string.Empty;
                UpdateLicensingState();
                _ = _indexerService.ResumeIndexingAsync();
            }
            else
            {
                ActivationErrorMessage = "Invalid license key or incorrect hardware ID.";
            }
        }

        private void UpdateLicensingState()
        {
            if (_licensingService == null) return;

            var status = _licensingService.GetStatus();
            if (status == LicenseStatus.Expired)
            {
                IsActivationOverlayVisible = true;
                IsTrialWarningVisible = false;
                IsRenewalWarningVisible = false;
            }
            else if (status == LicenseStatus.Trial)
            {
                IsActivationOverlayVisible = false;
                IsTrialWarningVisible = true;
                IsRenewalWarningVisible = false;
                int days = _licensingService.GetTrialDaysRemaining();
                TrialDaysRemainingText = $"Trial Mode: {days} days remaining";
            }
            else if (status == LicenseStatus.Valid)
            {
                IsActivationOverlayVisible = false;
                IsTrialWarningVisible = false;

                int daysUntilExpiry = _licensingService.GetDaysUntilExpiry();
                if (daysUntilExpiry <= 14 && daysUntilExpiry > 0)
                {
                    IsRenewalWarningVisible = true;
                    RenewalWarningText = $"Subscription Renewal: Your annual license expires in {daysUntilExpiry} days. Contact developer.diptesh@gmail.com to renew.";
                }
                else
                {
                    IsRenewalWarningVisible = false;
                }
            }
            else
            {
                IsActivationOverlayVisible = true;
                IsTrialWarningVisible = false;
                IsRenewalWarningVisible = false;
            }
        }

        public string SearchText
        {
            get => _searchText;
            set
            {
                if (SetProperty(ref _searchText, value))
                {
                    CurrentPage = 1;
                    OnPropertyChanged(nameof(ResultsSummaryText));
                }
            }
        }

        public string SelectedSearchHistory
        {
            get => _selectedSearchHistory;
            set
            {
                if (SetProperty(ref _selectedSearchHistory, value))
                {
                    if (!string.IsNullOrEmpty(value))
                    {
                        SearchText = value;
                        _ = PerformSearchAsync(resetPage: true);
                    }
                }
            }
        }
        public string FromFilter { get => _fromFilter; set => SetProperty(ref _fromFilter, value); }
        public string ToFilter { get => _toFilter; set => SetProperty(ref _toFilter, value); }
        public string SubjectFilter { get => _subjectFilter; set => SetProperty(ref _subjectFilter, value); }
        public DateTime? DateFrom { get => _dateFrom; set => SetProperty(ref _dateFrom, value); }
        public DateTime? DateTo { get => _dateTo; set => SetProperty(ref _dateTo, value); }
        public string SelectedFolder { get => _selectedFolder; set => SetProperty(ref _selectedFolder, value); }
        public string SelectedAttachmentOption { get => _selectedAttachmentOption; set => SetProperty(ref _selectedAttachmentOption, value); }
        public bool IsSearching { get => _isSearching; set => SetProperty(ref _isSearching, value); }

        public int CurrentPage
        {
            get => _currentPage;
            set
            {
                if (SetProperty(ref _currentPage, value))
                {
                    NotifyPaginationPropertiesChanged();
                }
            }
        }

        public int PageSize
        {
            get => _pageSize;
            set
            {
                if (SetProperty(ref _pageSize, value))
                {
                    CurrentPage = 1;
                    _ = PerformSearchAsync(resetPage: false);
                    NotifyPaginationPropertiesChanged();
                }
            }
        }

        public int TotalPages => (PageSize > 0 && TotalResultsCount > 0) ? (int)Math.Ceiling((double)TotalResultsCount / PageSize) : 0;
        public int FirstResultNumber => TotalResultsCount > 0 ? (CurrentPage - 1) * PageSize + 1 : 0;
        public int LastResultNumber => Math.Min(CurrentPage * PageSize, TotalResultsCount);
        public bool HasPreviousPage => CurrentPage > 1;
        public bool HasNextPage => CurrentPage < TotalPages && TotalPages > 0;

        public string ShowingResultsSummaryText => TotalResultsCount > 0 ? $"Showing {FirstResultNumber:N0}–{LastResultNumber:N0} of {TotalResultsCount:N0}" : string.Empty;

        public int TotalResultsCount
        {
            get => _totalResultsCount;
            set
            {
                if (SetProperty(ref _totalResultsCount, value))
                {
                    OnPropertyChanged(nameof(ResultsSummaryText));
                    NotifyPaginationPropertiesChanged();
                }
            }
        }
        public string SearchTimeText { get => _searchTimeText; set => SetProperty(ref _searchTimeText, value); }
        public string StatusText { get => _statusText; set => SetProperty(ref _statusText, value); }
        public bool IsIndexing
        {
            get => _isIndexing;
            set
            {
                if (SetProperty(ref _isIndexing, value))
                {
                    OnPropertyChanged(nameof(CanRebuildIndex));
                    OnPropertyChanged(nameof(RebuildButtonText));
                }
            }
        }

        public bool CanRebuildIndex => !IsIndexing;
        public string RebuildButtonText => IsIndexing ? "Indexing..." : "Rebuild Index";

        public int IndexProcessedItems { get => _indexProcessedItems; set => SetProperty(ref _indexProcessedItems, value); }
        public int IndexTotalItems { get => _indexTotalItems; set => SetProperty(ref _indexTotalItems, value); }
        public int IndexFolderProcessedItems { get => _indexFolderProcessedItems; set => SetProperty(ref _indexFolderProcessedItems, value); }
        public int IndexFolderTotalItems { get => _indexFolderTotalItems; set => SetProperty(ref _indexFolderTotalItems, value); }
        public string IndexCurrentFolder { get => _indexCurrentFolder; set => SetProperty(ref _indexCurrentFolder, value); }
        public double IndexProgressPercentage { get => _indexProgressPercentage; set => SetProperty(ref _indexProgressPercentage, value); }

        public string ResultsSummaryText => TotalResultsCount > 0
            ? $"Results: {TotalResultsCount:N0}"
            : (string.IsNullOrWhiteSpace(SearchText) ? "Ready to search" : "No emails found.");

        public void Cleanup()
        {
            try
            {
                _livePollingTimer?.Stop();
                _livePollingTimer?.Dispose();
                _livePollingTimer = null;
                _indexerService.ProgressChanged -= OnIndexerProgressChanged;
                _searchCts?.Cancel();
                _searchCts?.Dispose();
            }
            catch { }
        }

        public async Task GoToPageAsync(int page)
        {
            int targetPage = Math.Max(1, Math.Min(page, TotalPages > 0 ? TotalPages : 1));
            if (CurrentPage != targetPage)
            {
                CurrentPage = targetPage;
                await PerformSearchAsync(resetPage: false);
            }
        }

        private void NotifyPaginationPropertiesChanged()
        {
            OnPropertyChanged(nameof(TotalPages));
            OnPropertyChanged(nameof(FirstResultNumber));
            OnPropertyChanged(nameof(LastResultNumber));
            OnPropertyChanged(nameof(ShowingResultsSummaryText));
            OnPropertyChanged(nameof(HasPreviousPage));
            OnPropertyChanged(nameof(HasNextPage));
            (FirstPageCommand as RelayCommand)?.RaiseCanExecuteChanged();
            (PreviousPageCommand as RelayCommand)?.RaiseCanExecuteChanged();
            (NextPageCommand as RelayCommand)?.RaiseCanExecuteChanged();
            (LastPageCommand as RelayCommand)?.RaiseCanExecuteChanged();
        }

        public async Task PerformSearchAsync(bool resetPage = true)
        {
            if (resetPage)
            {
                _currentPage = 1;
                OnPropertyChanged(nameof(CurrentPage));
            }

            CancellationTokenSource oldCts;
            CancellationTokenSource newCts = new CancellationTokenSource();

            lock (_lock)
            {
                oldCts = _searchCts!;
                _searchCts = newCts;
            }

            if (oldCts != null)
            {
                try { oldCts.Cancel(); oldCts.Dispose(); } catch { }
            }

            var token = newCts.Token;

            IsSearching = true;
            IsInvalidSyntax = false;
            SyntaxErrorMessage = string.Empty;

            try
            {
                var parsedQuery = SearchQueryParser.Parse(SearchText);
                if (parsedQuery.IsInvalidSyntax)
                {
                    IsInvalidSyntax = true;
                    SyntaxErrorMessage = parsedQuery.SyntaxErrorMessage;
                    SearchTimeText = "Invalid search syntax";
                    SearchResults.Clear();
                    TotalResultsCount = 0;
                    NotifyPaginationPropertiesChanged();
                    return;
                }

                parsedQuery.SenderFilter = string.IsNullOrWhiteSpace(parsedQuery.SenderFilter) ? FromFilter : parsedQuery.SenderFilter;
                parsedQuery.RecipientFilter = string.IsNullOrWhiteSpace(parsedQuery.RecipientFilter) ? ToFilter : parsedQuery.RecipientFilter;
                parsedQuery.SubjectFilter = string.IsNullOrWhiteSpace(parsedQuery.SubjectFilter) ? SubjectFilter : parsedQuery.SubjectFilter;
                if (!parsedQuery.DateFrom.HasValue) parsedQuery.DateFrom = DateFrom;
                if (!parsedQuery.DateTo.HasValue) parsedQuery.DateTo = DateTo;
                if (string.IsNullOrWhiteSpace(parsedQuery.FolderFilter) && SelectedFolder != "All Folders") parsedQuery.FolderFilter = SelectedFolder;

                if (SelectedAttachmentOption == "With Attachments") parsedQuery.HasAttachmentFilter = true;
                else if (SelectedAttachmentOption == "Without Attachments") parsedQuery.HasAttachmentFilter = false;

                if (SelectedSortOption == "Oldest First") parsedQuery.SortOrder = SortOption.OldestFirst;
                else if (SelectedSortOption == "Newest First") parsedQuery.SortOrder = SortOption.NewestFirst;
                else parsedQuery.SortOrder = SortOption.Relevance;

                parsedQuery.PageSize = PageSize;
                parsedQuery.Page = CurrentPage;

                var resultContainer = await _searchService.ExecuteSearchAsync(SearchText, parsedQuery, token);

                if (token.IsCancellationRequested) return;

                SearchResults.Clear();
                foreach (var item in resultContainer.Items)
                {
                    SearchResults.Add(item);
                }

                TotalResultsCount = resultContainer.TotalCount;
                NotifyPaginationPropertiesChanged();

                string bgSuffix = IsIndexing ? " • Indexing in background" : "";
                SearchTimeText = $"{TotalResultsCount:N0} results • {resultContainer.SearchTimeMs} ms{bgSuffix}";

                if (!string.IsNullOrWhiteSpace(SearchText))
                {
                    _ = _repository.SaveSearchQueryAsync(SearchText);
                    _ = LoadSearchHistoryAsync();
                }
            }
            catch (OperationCanceledException)
            {
                // Search cancelled cleanly
            }
            catch (Exception ex)
            {
                if (!token.IsCancellationRequested)
                {
                    if (ex is InvalidOperationException && ex.Message.Contains("expired"))
                    {
                        UpdateLicensingState(); // Force UI update if backend caught expiry
                        SearchTimeText = "Trial expired";
                    }
                    else
                    {
                        SearchTimeText = $"Search error: {ex.Message}";
                    }
                }
            }
            finally
            {
                if (!token.IsCancellationRequested)
                {
                    IsSearching = false;
                }
            }
        }

        public static Action<Action>? DispatcherOverride { get; set; }

        private static void DispatchOnUI(Action action)
        {
            if (DispatcherOverride != null)
            {
                DispatcherOverride(action);
                return;
            }

            DispatchViaWpf(action);
        }

        [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]
        private static void DispatchViaWpf(Action action)
        {
            try
            {
                var disp = Application.Current?.Dispatcher;
                if (disp != null && !disp.HasShutdownStarted)
                {
                    if (disp.CheckAccess())
                    {
                        action();
                    }
                    else
                    {
                        disp.BeginInvoke(action);
                    }
                }
                else
                {
                    action();
                }
            }
            catch
            {
                action();
            }
        }

        private async Task LoadSearchHistoryAsync()
        {
            try
            {
                var history = await _repository.GetRecentSearchQueriesAsync(10);
                DispatchOnUI(() =>
                {
                    SearchHistoryList.Clear();
                    foreach (var q in history) SearchHistoryList.Add(q);
                });
            }
            catch { }
        }

        private void CopySearchResultDetails(SearchResultItem? item)
        {
            if (item == null) return;
            string details = $"Subject: {item.Subject}\nFrom: {item.SenderName} <{item.SenderEmail}>\nDate: {item.DisplayDate}\nFolder: {item.FolderName}\nSnippet: {item.BodySnippet}";
            CopyToClipboard(details);
        }

        private void ClearFilters()
        {
            try { _searchCts?.Cancel(); } catch { }
            SearchText = string.Empty;
            FromFilter = string.Empty;
            ToFilter = string.Empty;
            SubjectFilter = string.Empty;
            DateFrom = null;
            DateTo = null;
            SelectedFolder = "All Folders";
            SelectedAttachmentOption = "All";
            _selectedSortOption = "Relevance";
            OnPropertyChanged(nameof(SelectedSortOption));
            _currentPage = 1;
            IsInvalidSyntax = false;
            SyntaxErrorMessage = string.Empty;
            SearchResults.Clear();
            TotalResultsCount = 0;
            SearchTimeText = string.Empty;
            NotifyPaginationPropertiesChanged();
        }

        private void OpenEmail(SearchResultItem? item)
        {
            if (item == null) return;

            try
            {
                FastMailSearch.Infrastructure.Logging.Logger.Info($"[SEARCH] Result double-clicked");
                FastMailSearch.Infrastructure.Logging.Logger.Info($"[SEARCH] Subject: {item.Subject}");
                FastMailSearch.Infrastructure.Logging.Logger.Info($"[SEARCH] StoreId: {item.StoreId}");
                FastMailSearch.Infrastructure.Logging.Logger.Info($"[SEARCH] EntryId: {item.EntryId}");

                if (!_outlookAdapter.IsOutlookAvailable())
                {
                    FastMailSearch.Infrastructure.Logging.Logger.Warn($"[SEARCH] OpenEmail aborted: Outlook is not running.");
                    ShowAlert(
                        "Outlook is not currently running. Please open Outlook to view this email.",
                        "FastMail Search",
                        MessageBoxImage.Information);
                    return;
                }

                FastMailSearch.Infrastructure.Logging.Logger.Info($"[SEARCH] Calling Outlook GetItemFromID");

                bool opened = _outlookAdapter.OpenEmailInOutlook(item.StoreId, item.EntryId);
                if (!opened)
                {
                    FastMailSearch.Infrastructure.Logging.Logger.Warn($"[SEARCH] Unable to open email item in Outlook. StoreId: {item.StoreId}, EntryId: {item.EntryId}");
                    ShowAlert(
                        "Unable to open this email in Outlook. The item may have been moved, deleted, or is temporarily inaccessible.",
                        "FastMail Search",
                        MessageBoxImage.Warning);
                }
                else
                {
                    FastMailSearch.Infrastructure.Logging.Logger.Info($"[SEARCH] Outlook item retrieved");
                    FastMailSearch.Infrastructure.Logging.Logger.Info($"[SEARCH] Display() called successfully");
                }
            }
            catch (Exception ex)
            {
                FastMailSearch.Infrastructure.Logging.Logger.Error($"[SEARCH ERROR] Exception in OpenEmail", ex);
            }
        }

        private void ShowAlert(string message, string title, MessageBoxImage icon = MessageBoxImage.Information)
        {
            if (MessageBoxAction != null)
            {
                MessageBoxAction(message, title);
            }
            else
            {
                MessageBox.Show(message, title, MessageBoxButton.OK, icon);
            }
        }

        private void CopyToClipboard(string? text)
        {
            if (string.IsNullOrEmpty(text)) return;
            try
            {
                Clipboard.SetText(text);
            }
            catch (Exception ex)
            {
                FastMailSearch.Infrastructure.Logging.Logger.Error("Clipboard copy failed", ex);
            }
        }

        private void CopyHardwareIdToClipboard()
        {
            if (string.IsNullOrWhiteSpace(HardwareId) || HardwareId == "Unavailable" || HardwareId == "UNKNOWN-HWID" || HardwareId == "HWID_UNAVAILABLE")
            {
                CopyHardwareIdStatusMessage = "Hardware ID is unavailable to copy.";
                return;
            }

            try
            {
                Clipboard.SetText(HardwareId);
                CopyHardwareIdStatusMessage = "Hardware ID copied to clipboard!";
            }
            catch (Exception ex)
            {
                CopyHardwareIdStatusMessage = "Could not access clipboard. Please copy ID manually.";
                FastMailSearch.Infrastructure.Logging.Logger.Error("Clipboard copy failed for Hardware ID", ex);
            }
        }

        private void CopySupportEmailToClipboard()
        {
            try
            {
                Clipboard.SetText("developer.diptesh@gmail.com");
                CopySupportEmailStatusMessage = "Email copied: developer.diptesh@gmail.com";
            }
            catch (Exception ex)
            {
                CopySupportEmailStatusMessage = "Failed to copy email.";
                FastMailSearch.Infrastructure.Logging.Logger.Warn($"Clipboard copy failed: {ex.Message}");
            }
        }

        private void OpenProjectRepository()
        {
            try
            {
                const string repoUrl = "https://github.com/dipanki111/FastMailSearch";
                var psi = new ProcessStartInfo
                {
                    FileName = repoUrl,
                    UseShellExecute = true
                };
                Process.Start(psi);
            }
            catch (Exception ex)
            {
                FastMailSearch.Infrastructure.Logging.Logger.Warn($"Failed to open project URL: {ex.Message}");
            }
        }

        private async Task LoadFoldersAsync()
        {
            try
            {
                var folders = await _repository.GetFoldersAsync();
                foreach (var f in folders)
                {
                    if (!FolderOptions.Contains(f.Name))
                        FolderOptions.Add(f.Name);
                }
            }
            catch { }
        }

        private async Task RebuildIndexAsync()
        {
            if (_licensingService?.GetStatus() != LicenseStatus.Valid && _licensingService?.GetStatus() != LicenseStatus.Trial)
            {
                MessageBox.Show(
                    "Email indexing is paused because your 30-day trial has expired or a valid license is required.\n\nPlease activate FastMail Search to rebuild the index.",
                    "FastMail Search — License Required",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            if (_indexerService.IsIndexing)
            {
                StatusText = "Index rebuild is already in progress.";
                return;
            }

            var res = MessageBox.Show(
                "Are you sure you want to rebuild the email index?\nThis will re-scan your Outlook mailbox.",
                "Rebuild Index",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (res == MessageBoxResult.Yes)
            {
                StatusText = "Starting index rebuild...";
                IsIndexing = true;
                await _indexerService.RebuildIndexAsync();
            }
        }

        private async Task BackfillActiveStoreAsync()
        {
            if (_licensingService?.GetStatus() != LicenseStatus.Valid && _licensingService?.GetStatus() != LicenseStatus.Trial)
            {
                MessageBox.Show(
                    "Email indexing is paused because your 30-day trial has expired or a valid license is required.\n\nPlease activate FastMail Search to backfill emails.",
                    "FastMail Search — License Required",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            if (_indexerService.IsIndexing) return;
            try
            {
                StatusText = "🔵 Starting Outlook mailbox backfill...";
                IsIndexing = true;
                await _indexerService.StartActiveStoreBackfillAsync();
            }
            catch (Exception ex)
            {
                StatusText = $"🔴 Backfill error: {ex.Message}";
                IsIndexing = false;
            }
        }

        private async Task CancelIndexAsync()
        {
            if (!_indexerService.IsIndexing && !IsIndexing) return;

            var res = MessageBox.Show(
                "Stop the current indexing operation?",
                "Cancel Indexing",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (res == MessageBoxResult.Yes)
            {
                StatusText = "Cancelling indexing...";
                IsIndexing = false;
                await _indexerService.CancelIndexingAsync();
            }
        }

        private async Task TogglePauseResumeAsync()
        {
            if (_indexerService.CurrentStatus == IndexerStatus.Paused)
            {
                await _indexerService.ResumeIndexingAsync();
            }
            else if (_indexerService.CurrentStatus == IndexerStatus.Indexing)
            {
                await _indexerService.PauseIndexingAsync();
            }
        }

        private System.Timers.Timer? _livePollingTimer;
        private int _isPollingStatusInt = 0;
        private IndexerProgressEventArgs? _latestProgressArgs;
        private int _isProgressDispatchScheduled = 0;

        private void OnIndexerProgressChanged(object? sender, IndexerProgressEventArgs e)
        {
            Volatile.Write(ref _latestProgressArgs, e);

            // Immediately dispatch terminal or non-indexing states. Coalesce rapid indexing ticks to keep UI responsive.
            if (e.Status != IndexerStatus.Indexing || Interlocked.CompareExchange(ref _isProgressDispatchScheduled, 1, 0) == 0)
            {
                DispatchOnUI(() =>
                {
                    try
                    {
                        var latest = Volatile.Read(ref _latestProgressArgs) ?? e;
                        ApplyProgressUpdate(latest);
                    }
                    finally
                    {
                        Interlocked.Exchange(ref _isProgressDispatchScheduled, 0);
                    }
                });
            }
        }

        private void ApplyProgressUpdate(IndexerProgressEventArgs e)
        {
            if (e.TotalItems > 0)
            {
                IndexProcessedItems = e.ProcessedItems;
                IndexTotalItems = e.TotalItems;
            }
            else if (IndexTotalItems == 0 && e.ProcessedItems > 0)
            {
                IndexProcessedItems = e.ProcessedItems;
            }

            IndexFolderProcessedItems = e.FolderProcessedItems;
            IndexFolderTotalItems = e.FolderTotalItems;
            IndexCurrentFolder = e.CurrentFolder;
            IndexProgressPercentage = e.Percentage;
            IsIndexing = e.Status == IndexerStatus.Indexing;
            IsSafeMode = e.Status == IndexerStatus.SafeMode;

            if (e.Status == IndexerStatus.SafeMode)
            {
                StatusText = "🟡 Safe Mode Active • Outlook functionality is unaffected.";
            }
            else if (e.Status == IndexerStatus.Indexing)
            {
                if (!string.IsNullOrWhiteSpace(e.CurrentStatusMessage) &&
                    (e.CurrentStatusMessage.Contains("Updating") ||
                     e.CurrentStatusMessage.Contains("Discovering") ||
                     e.CurrentStatusMessage.Contains("Preparing") ||
                     e.CurrentStatusMessage.Contains("remaining")))
                {
                    StatusText = e.CurrentStatusMessage;
                    if (IndexProgressPercentage >= 99.9 && e.FolderTotalItems == 0)
                    {
                        IndexProgressPercentage = 0;
                    }
                }
                else if (e.TotalItems == 0)
                {
                    StatusText = !string.IsNullOrWhiteSpace(e.CurrentStatusMessage)
                        ? e.CurrentStatusMessage
                        : "🔵 Indexing in progress...";
                    IndexProgressPercentage = 0;
                }
                else if (e.RemainingCount > 0)
                {
                    StatusText = $"🔵 Indexing: {e.ProcessedItems:N0} / {e.TotalItems:N0} ({e.Percentage:F0}%) • {e.RemainingCount:N0} remaining" + (!string.IsNullOrEmpty(e.CurrentFolder) ? $" ({e.CurrentFolder})" : "");
                }
                else if (e.FolderTotalItems > 0 && !string.IsNullOrEmpty(e.CurrentFolder))
                {
                    StatusText = $"🔵 Indexing: {e.ProcessedItems:N0} / {e.TotalItems:N0} ({e.Percentage:F0}%) • {e.CurrentFolder}: {e.FolderProcessedItems:N0}/{e.FolderTotalItems:N0}";
                }
                else
                {
                    StatusText = $"🔵 Indexing: {e.ProcessedItems:N0} / {e.TotalItems:N0} ({e.Percentage:F0}%)" + (!string.IsNullOrEmpty(e.CurrentFolder) ? $" • Current folder: {e.CurrentFolder}" : "");
                }
            }
            else if (e.Status == IndexerStatus.Cancelled)
            {
                StatusText = $"🟡 Indexing cancelled ({e.ProcessedItems:N0} emails indexed)";
            }
            else if (e.Status == IndexerStatus.Error)
            {
                StatusText = $"🔴 Index error: {e.LastError}";
            }
            else
            {
                if (e.PendingQueueCount > 0 && e.Status == IndexerStatus.Idle)
                {
                    StatusText = $"⏸ Idle — Queue Paused ({e.PendingQueueCount:N0} pending)";
                }
                else
                {
                    StatusText = $"🟢 Up to date ({e.ProcessedItems:N0} emails indexed)";
                }
            }
        }

        private void UpdateStatusDisplay()
        {
            var status = _indexerService.GetStatus();
            ApplyProgressUpdate(status);

            Task.Run(async () =>
            {
                var summary = await HealthService.GetHealthSummaryAsync();
                int count = summary.IndexedEmailCount;
                DispatchOnUI(() =>
                {
                    if (IndexTotalItems == 0 && count > 0)
                    {
                        IndexTotalItems = count;
                        IndexProcessedItems = count;
                    }
                    if (!IsIndexing)
                    {
                        if (summary.PendingQueueCount > 0 && summary.CurrentlyProcessingCount == 0 && !_indexerService.IsIndexing)
                        {
                            StatusText = $"⏸ Idle — Queue Paused ({summary.PendingQueueCount:N0} pending)";
                        }
                        else if (summary.PendingQueueCount > 0 && (summary.CurrentlyProcessingCount > 0 || _indexerService.IsIndexing))
                        {
                            StatusText = $"🔵 Updating • {summary.PendingQueueCount:N0} email(s) remaining";
                        }
                        else
                        {
                            StatusText = $"🟢 Up to date ({count:N0} emails indexed)";
                        }
                    }
                });
            });
        }

        private async Task PollLiveStatusAsync()
        {
            if (Interlocked.CompareExchange(ref _isPollingStatusInt, 1, 0) != 0) return;
            try
            {
                // If local indexer is actively indexing, its own ProgressChanged event handles UI updates
                if (_indexerService.IsIndexing) return;

                var summary = await HealthService.GetHealthSummaryAsync();
                DispatchOnUI(() =>
                {
                    bool isCrossIndexing = summary.CurrentOperation != null &&
                        (summary.CurrentOperation.StartsWith("Backfilling", StringComparison.OrdinalIgnoreCase) ||
                         summary.CurrentOperation.StartsWith("Preparing", StringComparison.OrdinalIgnoreCase));

                    if (isCrossIndexing)
                    {
                        IsIndexing = true;
                        IndexTotalItems = summary.TotalDiscoveredOutlookItems > 0 ? summary.TotalDiscoveredOutlookItems : summary.IndexedEmailCount;
                        IndexProcessedItems = summary.ScannedCount;
                        IndexProgressPercentage = summary.ProgressPercentage;
                        StatusText = $"🔵 Indexing: {IndexProcessedItems:N0} / {IndexTotalItems:N0} ({IndexProgressPercentage:F0}%) • {summary.RemainingToScan:N0} remaining";
                    }
                    else
                    {
                        if (IsIndexing)
                        {
                            IsIndexing = false;
                        }

                        if (summary.PendingQueueCount > 0)
                        {
                            if (summary.CurrentlyProcessingCount > 0 || _indexerService.IsIndexing)
                            {
                                StatusText = $"🔵 Updating • {summary.PendingQueueCount:N0} email(s) remaining";
                            }
                            else
                            {
                                StatusText = $"⏸ Idle — Queue Paused ({summary.PendingQueueCount:N0} pending)";
                            }
                        }
                        else if (summary.FailedItemCount > 0)
                        {
                            StatusText = $"🟡 {summary.FailedItemCount:N0} indexing issue(s)";
                        }
                        else if (!summary.IntegrityResult.IsHealthy)
                        {
                            StatusText = "🔴 Index requires attention";
                        }
                        else
                        {
                            StatusText = $"🟢 Up to date ({summary.IndexedEmailCount:N0} emails indexed)";
                            if (IndexTotalItems == 0 && summary.IndexedEmailCount > 0)
                            {
                                IndexTotalItems = summary.IndexedEmailCount;
                                IndexProcessedItems = summary.IndexedEmailCount;
                            }
                        }
                    }
                });
            }
            catch (Exception ex)
            {
                Logger.Warn($"Failed to poll live indexing status: {ex.Message}");
            }
            finally
            {
                Interlocked.Exchange(ref _isPollingStatusInt, 0);
            }
        }
    }
}
