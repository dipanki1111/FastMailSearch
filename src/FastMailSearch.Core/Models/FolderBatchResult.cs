using System;
using System.Collections.Generic;

namespace FastMailSearch.Core.Models
{
    public class FolderBatchResult
    {
        public List<EmailDocument> Items { get; set; } = new List<EmailDocument>();
        public int StartIndex { get; set; }
        public int LastProcessedIndex { get; set; }
        public int ProcessedCount { get; set; }
        public int IndexedCount { get; set; }
        public int SkippedCount { get; set; }
        public bool WasCancelled { get; set; }
        public bool HasError { get; set; }
        public string? ErrorMessage { get; set; }

        public static FolderBatchResult Empty(int startIndex) => new FolderBatchResult
        {
            StartIndex = startIndex,
            LastProcessedIndex = startIndex - 1,
            ProcessedCount = 0,
            IndexedCount = 0,
            SkippedCount = 0
        };
    }
}
