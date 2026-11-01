using System;

namespace FastMailSearch.Core.Models
{
    public class FailedIndexItem
    {
        public long Id { get; set; }
        public string StoreId { get; set; } = string.Empty;
        public string EntryId { get; set; } = string.Empty;
        public string FolderName { get; set; } = string.Empty;
        public string OperationType { get; set; } = string.Empty; // Add, Change, Delete
        public string FailureReason { get; set; } = string.Empty;
        public string ExceptionType { get; set; } = string.Empty;
        public int AttemptCount { get; set; }
        public DateTime FirstFailedAt { get; set; } = DateTime.Now;
        public DateTime LastFailedAt { get; set; } = DateTime.Now;
        public DateTime? LastRetryAt { get; set; }
        public string Status { get; set; } = "Failed"; // Failed, Retrying, Recovered, Ignored
        public string ErrorMessage { get; set; } = string.Empty;
        public bool Resolved { get; set; }

        public string DisplayStatus => Status;
        public string DisplayLastFailed => LastFailedAt.ToString("dd MMM yyyy hh:mm tt");
        public string ShortError => string.IsNullOrWhiteSpace(ErrorMessage) ? FailureReason : ErrorMessage;
    }
}
