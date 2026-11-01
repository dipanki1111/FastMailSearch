using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using FastMailSearch.Core.Interfaces;
using FastMailSearch.Core.Models;
using FastMailSearch.Core.Search;

namespace FastMailSearch.Core.Services
{
    public class SearchService
    {
        private readonly ISearchEngine _searchEngine;
        private readonly IEmailRepository _repository;
        private readonly ILicensingService _licensingService;
        private CancellationTokenSource? _activeCts;
        private readonly object _lock = new object();

        [Obsolete("Use SearchService(ISearchEngine, IEmailRepository, ILicensingService) instead. A licensing service is required for search execution.", false)]
        public SearchService(ISearchEngine searchEngine, IEmailRepository repository)
        {
            _searchEngine = searchEngine ?? throw new ArgumentNullException(nameof(searchEngine));
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
            _licensingService = null!;
        }

        public SearchService(ISearchEngine searchEngine, IEmailRepository repository, ILicensingService licensingService)
        {
            _searchEngine = searchEngine ?? throw new ArgumentNullException(nameof(searchEngine));
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
            _licensingService = licensingService ?? throw new ArgumentNullException(nameof(licensingService));
        }

        public async Task<SearchResultContainer> ExecuteSearchAsync(string rawQuery, SearchQuery? queryFilters = null, CancellationToken cancellationToken = default)
        {
            if (_licensingService == null)
            {
                throw new InvalidOperationException("Licensing service is uninitialized. Your 30-day trial has expired. Please activate FastMail Search to continue.");
            }

            var status = _licensingService.GetStatus();
            if (status != LicenseStatus.Valid && status != LicenseStatus.Trial)
            {
                throw new InvalidOperationException("Your 30-day trial has expired. Please activate FastMail Search to continue.");
            }

            CancellationTokenSource oldCts;
            CancellationTokenSource newCts = new CancellationTokenSource();

            lock (_lock)
            {
                oldCts = _activeCts!;
                _activeCts = newCts;
            }

            if (oldCts != null)
            {
                try
                {
                    oldCts.Cancel();
                    oldCts.Dispose();
                }
                catch { }
            }

            var query = SearchQueryParser.Parse(rawQuery);

            // Merge additional explicit UI filter inputs if provided
            if (queryFilters != null)
            {
                if (!string.IsNullOrWhiteSpace(queryFilters.SenderFilter)) query.SenderFilter = queryFilters.SenderFilter;
                if (!string.IsNullOrWhiteSpace(queryFilters.RecipientFilter)) query.RecipientFilter = queryFilters.RecipientFilter;
                if (!string.IsNullOrWhiteSpace(queryFilters.SubjectFilter)) query.SubjectFilter = queryFilters.SubjectFilter;
                if (!string.IsNullOrWhiteSpace(queryFilters.FolderFilter)) query.FolderFilter = queryFilters.FolderFilter;
                if (!string.IsNullOrWhiteSpace(queryFilters.StoreFilter)) query.StoreFilter = queryFilters.StoreFilter;
                if (queryFilters.HasAttachmentFilter.HasValue) query.HasAttachmentFilter = queryFilters.HasAttachmentFilter;
                if (queryFilters.IsReadFilter.HasValue) query.IsReadFilter = queryFilters.IsReadFilter;
                if (queryFilters.IsFlaggedFilter.HasValue) query.IsFlaggedFilter = queryFilters.IsFlaggedFilter;
                if (queryFilters.DateFrom.HasValue) query.DateFrom = queryFilters.DateFrom;
                if (queryFilters.DateTo.HasValue) query.DateTo = queryFilters.DateTo;
                query.SortOrder = queryFilters.SortOrder;
                if (queryFilters.Limit > 0) query.Limit = queryFilters.Limit;
                if (queryFilters.Offset >= 0) query.Offset = queryFilters.Offset;
            }

            var sw = Stopwatch.StartNew();
            using (var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(newCts.Token, cancellationToken))
            {
                var result = await _searchEngine.SearchAsync(query, linkedCts.Token);
                sw.Stop();

                result.ElapsedSeconds = Math.Round(sw.Elapsed.TotalSeconds, 3);

                // Log query to search history asynchronously
                if (!string.IsNullOrWhiteSpace(rawQuery))
                {
                    _ = Task.Run(async () =>
                    {
                        try { await _repository.SaveSearchQueryAsync(rawQuery); } catch { }
                    });
                }

                return result;
            }
        }

        public Task<List<string>> GetRecentQueriesAsync(int count = 20)
        {
            return _repository.GetRecentSearchQueriesAsync(count);
        }

        public Task ClearSearchHistoryAsync()
        {
            return _repository.ClearSearchHistoryAsync();
        }
    }
}
