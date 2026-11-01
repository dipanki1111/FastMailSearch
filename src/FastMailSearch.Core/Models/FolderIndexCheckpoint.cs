using System;

namespace FastMailSearch.Core.Models
{
    public class FolderIndexCheckpoint
    {
        public string StoreId { get; set; } = string.Empty;
        public string FolderId { get; set; } = string.Empty;
        public string Status { get; set; } = string.Empty;
        public int ProcessedOffset { get; set; }
        public int TotalItems { get; set; }
        public string LastSync { get; set; } = string.Empty;
    }
}
