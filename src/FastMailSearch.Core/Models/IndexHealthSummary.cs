using System;

namespace FastMailSearch.Core.Models
{
    public class IndexHealthSummary
    {
        public IndexerHealth HealthState { get; set; } = IndexerHealth.Healthy;
        public string StatusState { get; set; } = "🟢 Healthy";
        public int IndexedEmailCount { get; set; }
        public int PendingQueueCount { get; set; }
        public int CurrentlyProcessingCount { get; set; }
        public int FailedItemCount { get; set; }
        public int RecoveredItemCount { get; set; }
        public int IgnoredItemCount { get; set; }
        public int RetryCount { get; set; }
        public DateTime? LastSuccessfulIndexingTime { get; set; }
        public DateTime? LastFailureTime { get; set; }
        public string CurrentStatusMessage { get; set; } = "🟢 Index Healthy";
        public string? CurrentOperation { get; set; } = "Idle";
        public string? LastError { get; set; }
        public string DatabasePath { get; set; } = string.Empty;
        public double DatabaseSizeMb { get; set; }
        public double QueueThroughputPerSec { get; set; }
        public DateTime LastIntegrityScanTime { get; set; } = DateTime.Now;
        public IndexIntegrityResult IntegrityResult { get; set; } = new IndexIntegrityResult();

        public bool IsSafeMode { get; set; }
        public string SafeModeReason { get; set; } = string.Empty;

        // Phase 32 & 37: Mailbox Coverage Disambiguation & Live Progress
        public string MailboxCoverageStatus { get; set; } = "Up to date";
        public int TotalDiscoveredOutlookItems { get; set; }
        public int UnindexedOutlookItems { get; set; }
        public int ProcessedOffset { get; set; }
        public int RemainingToScan { get; set; }
        public int ScannedCount => ProcessedOffset > 0 ? ProcessedOffset : IndexedEmailCount;
        public double ProgressPercentage => TotalDiscoveredOutlookItems > 0 ? Math.Min(100.0, Math.Round((double)ScannedCount / TotalDiscoveredOutlookItems * 100.0, 1)) : 0.0;
        public string DisplayProgressText => TotalDiscoveredOutlookItems > 0 ? $"{ScannedCount:N0} / {TotalDiscoveredOutlookItems:N0} ({ProgressPercentage:F0}%)" : $"{IndexedEmailCount:N0} indexed";
        public bool IsCoverageComplete => (RemainingToScan == 0 || UnindexedOutlookItems == 0) && (TotalDiscoveredOutlookItems == 0 || IndexedEmailCount > 0);

        public int QueueLength => PendingQueueCount;
        public double ProcessingRate => QueueThroughputPerSec;

        public string DisplayLastUpdate => LastSuccessfulIndexingTime.HasValue
            ? LastSuccessfulIndexingTime.Value.ToString("dd MMM yyyy hh:mm tt")
            : "Just now";
        public string DisplayDbSize => $"{DatabaseSizeMb:F2} MB";
        public string DisplayLastError => !string.IsNullOrEmpty(LastError)
            ? LastError
            : (IntegrityResult != null && !IntegrityResult.IsHealthy ? IntegrityResult.SummaryText : "None");
        public string DisplayCoverage
        {
            get
            {
                if (TotalDiscoveredOutlookItems > 0)
                {
                    return $"{IndexedEmailCount:N0} indexed / {TotalDiscoveredOutlookItems:N0} in Outlook ({MailboxCoverageStatus})";
                }

                if (!string.IsNullOrEmpty(MailboxCoverageStatus) &&
                    MailboxCoverageStatus != "Up to date" &&
                    MailboxCoverageStatus != "No active store discovered")
                {
                    return $"{IndexedEmailCount:N0} indexed ({MailboxCoverageStatus})";
                }

                return $"{IndexedEmailCount:N0} indexed (No active store discovered)";
            }
        }
    }
}
