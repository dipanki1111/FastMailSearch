using System;

namespace FastMailSearch.Core.Models
{
    public class PendingChangeItem
    {
        public long Id { get; set; }
        public string StoreId { get; set; } = string.Empty;
        public string EntryId { get; set; } = string.Empty;
        public string Operation { get; set; } = "Add"; // Add, Change, Delete
        public string? FolderEntryId { get; set; }
        public DateTime CreatedAt { get; set; } = DateTime.Now;
        public DateTime UpdatedAt { get; set; } = DateTime.Now;
        public int AttemptCount { get; set; }
        public string Status { get; set; } = "Pending"; // Pending, Processing, Completed, Failed
        public string? LastError { get; set; }
    }
}
