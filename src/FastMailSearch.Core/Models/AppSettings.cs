using System;
using System.IO;

namespace FastMailSearch.Core.Models
{
    public class AppSettings
    {
        public const string DatabasePathEnvironmentVariable = "FASTMAILSEARCH_DB_PATH";
        public const string RequireOverrideEnvironmentVariable = "FASTMAILSEARCH_REQUIRE_OVERRIDE";
        public const string TestModeEnvironmentVariable = TestModeStoreFilter.TestModeEnvironmentVariable;
        public const string ApprovedStoresEnvironmentVariable = TestModeStoreFilter.ApprovedStoresEnvironmentVariable;

        public string DatabasePath { get; set; }
        public int MaxResults { get; set; } = 100;
        public int MaxBodySizeBytes { get; set; } = 1024 * 1024; // 1 MB limit
        public bool IndexOnStartup { get; set; } = true;
        public bool BackgroundIndexing { get; set; } = true;
        public bool SearchHistoryEnabled { get; set; } = true;
        public int MaxSearchHistory { get; set; } = 20;
        public bool IsBetaModeEnabled { get; set; } = false;

        public AppSettings()
        {
            DatabasePath = GetDatabasePath();
        }

        public static string NormalizePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return string.Empty;
            try
            {
                string fullPath = Path.GetFullPath(path.Trim());
                return fullPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            }
            catch
            {
                return path.Trim().TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            }
        }

        public static bool IsProductionDatabasePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return false;
            try
            {
                string normResolved = NormalizePath(path);
                string normDefault = NormalizePath(GetDefaultDatabasePath());
                return string.Equals(normResolved, normDefault, StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        public static void ValidateDatabasePath(string path)
        {
            string? requireOverride = Environment.GetEnvironmentVariable(RequireOverrideEnvironmentVariable);
            bool isRequireOverrideActive = string.Equals(requireOverride?.Trim(), "1", StringComparison.OrdinalIgnoreCase) ||
                                          string.Equals(requireOverride?.Trim(), "true", StringComparison.OrdinalIgnoreCase);

            if (isRequireOverrideActive && IsProductionDatabasePath(path))
            {
                throw new InvalidOperationException($"[SAFETY HALT] Database override is strictly required ({RequireOverrideEnvironmentVariable}=1), but resolved path matches canonical production database: '{path}'.");
            }
        }

        public static string GetDatabasePath()
        {
            string? overridePath = Environment.GetEnvironmentVariable(DatabasePathEnvironmentVariable);
            string resolvedPath;
            if (!string.IsNullOrWhiteSpace(overridePath))
            {
                resolvedPath = overridePath!.Trim();
            }
            else
            {
                resolvedPath = GetDefaultDatabasePath();
            }

            ValidateDatabasePath(resolvedPath);
            return resolvedPath;
        }

        public static string GetDefaultDatabasePath()
        {
            string docsFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string appFolder = Path.Combine(docsFolder, "FastMailSearch");
            return Path.Combine(appFolder, "FastMailSearch.db");
        }
    }

    public class IndexMetadata
    {
        public int TotalEmailsIndexed { get; set; }
        public DateTime? LastIndexedTime { get; set; }
        public string Status { get; set; } = "Up to date"; // Up to date, Indexing..., Paused, Error
        public int ItemsInQueue { get; set; }
        public int CurrentlyIndexedCount { get; set; }
        public int TotalItemsToIndex { get; set; }
        public string CurrentStatusDetail { get; set; } = string.Empty;
    }

    public class DiagnosticsInfo
    {
        public string OutlookVersion { get; set; } = "Unknown";
        public string OfficeBitness { get; set; } = "Unknown";
        public string WindowsVersion { get; set; } = "Unknown";
        public string DatabasePath { get; set; } = string.Empty;
        public long DatabaseSizeBytes { get; set; }
        public int IndexedEmailCount { get; set; }
        public int MailboxCount { get; set; }
        public int FolderCount { get; set; }
        public DateTime? LastIndexingTime { get; set; }
        public string IndexStatus { get; set; } = "Up to date";
    }
}
