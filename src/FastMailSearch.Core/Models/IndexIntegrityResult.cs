using System;

namespace FastMailSearch.Core.Models
{
    public class IndexIntegrityResult
    {
        public bool IsHealthy => MissingFtsCount == 0 && OrphanFtsCount == 0 && DuplicateCount == 0 && InvalidRecordCount == 0;
        public int EmailCount { get; set; }
        public int FtsCount { get; set; }
        public int MissingFtsCount { get; set; }
        public int OrphanFtsCount { get; set; }
        public int DuplicateCount { get; set; }
        public int InvalidRecordCount { get; set; }
        public bool RepairRequired => !IsHealthy;
        public DateTime ScannedAt { get; set; } = DateTime.Now;
        public string SummaryText => IsHealthy
            ? $"🟢 Index Healthy ({EmailCount:N0} emails, {FtsCount:N0} FTS records)"
            : $"🔴 Integrity Problem: {MissingFtsCount} missing FTS, {OrphanFtsCount} orphan FTS, {DuplicateCount} duplicates";
    }
}
