using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using FastMailSearch.Core.Models;

namespace FastMailSearch.Core.Interfaces
{
    public interface IIndexHealthService
    {
        Task<IndexHealthSummary> GetHealthSummaryAsync(CancellationToken cancellationToken = default);
        Task<EnvironmentDiagnostics> GetEnvironmentDiagnosticsAsync(CancellationToken cancellationToken = default);
        Task<string> ExportDiagnosticsZipAsync(string targetFolderPath, CancellationToken cancellationToken = default);
        Task<IndexIntegrityResult> VerifyIndexAsync(CancellationToken cancellationToken = default);
        Task RepairIndexAsync(IndexIntegrityResult scanResult, CancellationToken cancellationToken = default);
        Task<StoreCoverageInfo> CheckStoreCoverageAsync(CancellationToken cancellationToken = default);
        Task<BackfillSummary> StartActiveStoreBackfillAsync(CancellationToken cancellationToken = default);
        Task<List<FailedIndexItem>> GetFailedItemsAsync(string? status = null, CancellationToken cancellationToken = default);
        Task<int> GetFailedItemsCountAsync(string? status = null, CancellationToken cancellationToken = default);
        Task RetryFailedItemsBatchAsync(IEnumerable<long> failedItemIds, CancellationToken cancellationToken = default);
        Task<int> RetryAllFailedItemsAsync(CancellationToken cancellationToken = default);
        Task RemoveFailedItemsBatchAsync(IEnumerable<long> failedItemIds, CancellationToken cancellationToken = default);
        Task RemoveAllFailedItemsAsync(CancellationToken cancellationToken = default);
    }
}
