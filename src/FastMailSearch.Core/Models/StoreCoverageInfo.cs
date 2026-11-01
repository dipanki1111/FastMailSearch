using System;

namespace FastMailSearch.Core.Models
{
    public class StoreCoverageInfo
    {
        public int TotalDiscoveredItems { get; set; }
        public int TotalIndexedItems { get; set; }
        public int UnindexedItems { get; set; }
        public bool IsComplete => UnindexedItems == 0 && UnindexedStoreCount == 0;
        public int ActiveStoreCount { get; set; }
        public int UnindexedStoreCount { get; set; }
        public int UnindexedFolderCount { get; set; }
        public string SummaryText { get; set; } = string.Empty;
    }

    public class BackfillSummary
    {
        public int TotalDiscovered { get; set; }
        public int ProcessedCount { get; set; }
        public int NewlyIndexedCount { get; set; }
        public int SkippedCount { get; set; }
        public int FailedCount { get; set; }
        public bool IsCancelled { get; set; }
        public string? ErrorMessage { get; set; }
    }
}
