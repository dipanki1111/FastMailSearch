using System;

namespace FastMailSearch.Core.Models
{
    public class EnvironmentDiagnostics
    {
        public string OsVersion { get; set; } = string.Empty;
        public string OsArchitecture { get; set; } = string.Empty;
        public string OutlookVersion { get; set; } = string.Empty;
        public string OutlookBitness { get; set; } = string.Empty;
        public string OfficeVersion { get; set; } = string.Empty;
        public bool HasExchangeAccount { get; set; }
        public bool HasM365Account { get; set; }
        public bool HasImapAccount { get; set; }
        public bool HasPopAccount { get; set; }
        public int StoreCount { get; set; }
        public string DefaultStoreType { get; set; } = string.Empty;
        public string AddInInstallationPath { get; set; } = string.Empty;
        public string FastMailSearchVersion { get; set; } = string.Empty;
        public string DatabasePath { get; set; } = string.Empty;
        public double DatabaseSizeMb { get; set; }
        public string SqliteVersion { get; set; } = string.Empty;
        public bool IsFts5Available { get; set; }
        public DateTime Timestamp { get; set; } = DateTime.Now;
    }
}
