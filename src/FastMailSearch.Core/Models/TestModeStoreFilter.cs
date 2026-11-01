using System;
using System.Collections.Generic;
using System.IO;

namespace FastMailSearch.Core.Models
{
    /// <summary>
    /// Enforces fail-closed store protection during test runs.
    /// In normal production mode (IsTestModeActive == false), all stores are processed normally.
    /// In test mode (IsTestModeActive == true):
    /// 1. The canonical personal PST is permanently and unconditionally rejected.
    /// 2. All OST files are permanently rejected.
    /// 3. All unapproved PST files are rejected.
    /// 4. An explicitly approved synthetic PST is permitted by exact path or StoreId.
    /// 5. If approved stores list is empty, all stores are rejected (fail closed).
    /// </summary>
    public static class TestModeStoreFilter
    {
        public const string TestModeEnvironmentVariable = "FASTMAILSEARCH_TEST_MODE";
        public const string ApprovedStoresEnvironmentVariable = "FASTMAILSEARCH_APPROVED_STORES";

        public static readonly string HardcodedPersonalPstPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Outlook Files", "personal-test@example.com - test.pst");
        public static readonly string CanonicalPersonalPstPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Outlook Files", "personal-test@example.com - test.pst");
        public const string CanonicalPersonalDisplayName = "personal-test@example.com";
        public const string CanonicalPersonalStoreId = "0000000038A1BB1005E5101AA1BB08002B2A56C200006D73707374";

        private static readonly object _lock = new object();
        private static bool? _programmaticTestMode;
        private static HashSet<string>? _programmaticApprovedStores;
        private static readonly HashSet<string> _dynamicallyApprovedStoreIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private static readonly HashSet<string> _blockedStoreIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Indicates whether test mode is active.
        /// Checks programmatic setting first, then the FASTMAILSEARCH_TEST_MODE environment variable.
        /// </summary>
        public static bool IsTestModeActive
        {
            get
            {
                lock (_lock)
                {
                    if (_programmaticTestMode.HasValue)
                        return _programmaticTestMode.Value;
                }

                string? env = Environment.GetEnvironmentVariable(TestModeEnvironmentVariable);
                return string.Equals(env?.Trim(), "1", StringComparison.OrdinalIgnoreCase) ||
                       string.Equals(env?.Trim(), "true", StringComparison.OrdinalIgnoreCase);
            }
        }

        /// <summary>
        /// Explicitly configure test mode state and approved store identifiers for testing.
        /// </summary>
        public static void SetTestMode(bool enabled, IEnumerable<string>? approvedStoreIdentifiers = null)
        {
            lock (_lock)
            {
                _programmaticTestMode = enabled;
                _dynamicallyApprovedStoreIds.Clear();
                _blockedStoreIds.Clear();
                if (approvedStoreIdentifiers != null)
                {
                    _programmaticApprovedStores = new HashSet<string>(approvedStoreIdentifiers, StringComparer.OrdinalIgnoreCase);
                }
                else
                {
                    _programmaticApprovedStores = null;
                }
            }
        }

        /// <summary>
        /// Resets programmatic test mode settings back to environment defaults.
        /// </summary>
        public static void ResetTestMode()
        {
            lock (_lock)
            {
                _programmaticTestMode = null;
                _programmaticApprovedStores = null;
                _dynamicallyApprovedStoreIds.Clear();
                _blockedStoreIds.Clear();
            }
        }

        /// <summary>
        /// Safely normalizes and resolves full file paths for comparison, handling casing,
        /// relative traversals, separators, and trailing slashes.
        /// </summary>
        public static string NormalizePath(string? path)
        {
            if (string.IsNullOrWhiteSpace(path)) return string.Empty;
            try
            {
                string trimmed = path!.Trim();
                string fullPath = Path.GetFullPath(trimmed);
                return fullPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            }
            catch
            {
                return path!.Trim().TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            }
        }

        /// <summary>
        /// Evaluates whether a file path points to the canonical personal PST.
        /// Handles case variations, relative path traversals, mixed slashes, and filename checks.
        /// </summary>
        public static bool IsPersonalPstPath(string? path)
        {
            if (string.IsNullOrWhiteSpace(path)) return false;
            string norm = NormalizePath(path);
            if (string.IsNullOrEmpty(norm)) return false;

            if (string.Equals(norm, NormalizePath(HardcodedPersonalPstPath), StringComparison.OrdinalIgnoreCase))
                return true;

            try
            {
                if (string.Equals(norm, NormalizePath(CanonicalPersonalPstPath), StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            catch { }

            try
            {
                string fileName = Path.GetFileName(norm);
                if (string.Equals(fileName, "personal-test@example.com - test.pst", StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            catch { }

            if (norm.EndsWith("personal.pst", StringComparison.OrdinalIgnoreCase))
                return true;

            return false;
        }

        /// <summary>
        /// Determines whether a mailbox corresponds to the user's personal PST.
        /// </summary>
        public static bool IsPersonalPst(MailboxInfo? mailbox)
        {
            if (mailbox == null) return false;

            if (IsPersonalPstPath(mailbox.FilePath)) return true;

            if (!string.IsNullOrWhiteSpace(mailbox.DisplayName))
            {
                string disp = mailbox.DisplayName.Trim();
                if (string.Equals(disp, CanonicalPersonalDisplayName, StringComparison.OrdinalIgnoreCase) ||
                    disp.IndexOf("personal-test@example.com", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    string.Equals(disp, "personal@gmail.com", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            if (!string.IsNullOrWhiteSpace(mailbox.StoreId) &&
                string.Equals(mailbox.StoreId.Trim(), CanonicalPersonalStoreId, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Determines whether a mailbox corresponds to an OST file.
        /// </summary>
        public static bool IsOstStore(MailboxInfo? mailbox)
        {
            if (mailbox == null) return false;
            return IsOstPath(mailbox.FilePath);
        }

        /// <summary>
        /// Determines whether a file path has an OST extension.
        /// </summary>
        public static bool IsOstPath(string? filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath)) return false;
            string norm = NormalizePath(filePath);
            return norm.EndsWith(".ost", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Determines whether a mailbox corresponds to a PST file.
        /// </summary>
        public static bool IsPstStore(MailboxInfo? mailbox)
        {
            if (mailbox == null) return false;
            if (mailbox.IsPst) return true;
            if (!string.IsNullOrWhiteSpace(mailbox.FilePath))
            {
                string norm = NormalizePath(mailbox.FilePath);
                if (norm.EndsWith(".pst", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Retrieves the set of approved store identifiers (by exact FilePath, StoreId, or DisplayName).
        /// </summary>
        public static HashSet<string> GetApprovedStores()
        {
            lock (_lock)
            {
                if (_programmaticApprovedStores != null)
                {
                    return new HashSet<string>(_programmaticApprovedStores, StringComparer.OrdinalIgnoreCase);
                }
            }

            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            string? env = Environment.GetEnvironmentVariable(ApprovedStoresEnvironmentVariable);
            if (!string.IsNullOrWhiteSpace(env))
            {
                var tokens = env!.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var t in tokens)
                {
                    string trimmed = t.Trim();
                    if (!string.IsNullOrEmpty(trimmed))
                    {
                        set.Add(trimmed);
                    }
                }
            }
            return set;
        }

        /// <summary>
        /// Evaluates whether a mailbox is approved for enumeration or indexing.
        /// If test mode is inactive, returns true (preserving normal production behavior).
        /// If test mode is active:
        /// 1. The canonical personal PST is permanently and unconditionally rejected.
        /// 2. All OST files are permanently rejected.
        /// 3. All unapproved PST files are rejected.
        /// 4. An explicitly approved synthetic PST is permitted by exact path or StoreId.
        /// 5. If approved stores list is empty, all stores are rejected (fail closed).
        /// </summary>
        public static bool IsStoreApproved(MailboxInfo? mailbox)
        {
            if (mailbox == null) return false;

            // Rule 1: Preserve existing production behavior when test mode is disabled
            if (!IsTestModeActive) return true;

            // Rule 2: Permanently reject the canonical personal PST (regardless of whitelist)
            if (IsPersonalPst(mailbox))
            {
                RecordBlockedStoreId(mailbox.StoreId);
                return false;
            }

            // Rule 3: Reject every OST file
            if (IsOstStore(mailbox))
            {
                RecordBlockedStoreId(mailbox.StoreId);
                return false;
            }

            // Rule 4: If test mode is enabled and the approved-store list is empty, reject all stores
            var approved = GetApprovedStores();
            if (approved.Count == 0) return false;

            // Rule 5: Check whitelist match
            bool matchesPath = false;
            if (!string.IsNullOrWhiteSpace(mailbox.FilePath))
            {
                string normMailboxPath = NormalizePath(mailbox.FilePath);
                if (!string.IsNullOrEmpty(normMailboxPath))
                {
                    foreach (var entry in approved)
                    {
                        string normApproved = NormalizePath(entry);
                        if (!string.IsNullOrEmpty(normApproved) &&
                            string.Equals(normMailboxPath, normApproved, StringComparison.OrdinalIgnoreCase))
                        {
                            matchesPath = true;
                            break;
                        }
                    }
                }
            }

            bool matchesId = !string.IsNullOrWhiteSpace(mailbox.StoreId) && approved.Contains(mailbox.StoreId!.Trim());

            bool isPst = IsPstStore(mailbox);

            if (isPst)
            {
                // For synthetic PST stores: Must match by exact path or StoreId
                if (matchesPath || matchesId)
                {
                    RecordApprovedStoreId(mailbox.StoreId);
                    return true;
                }

                // Reject every unapproved PST
                return false;
            }
            else
            {
                // For non-PST synthetic/mock stores: Match by StoreId, DisplayName, or Path
                bool matchesName = !string.IsNullOrWhiteSpace(mailbox.DisplayName) && approved.Contains(mailbox.DisplayName!.Trim());
                if (matchesId || matchesName || matchesPath)
                {
                    RecordApprovedStoreId(mailbox.StoreId);
                    return true;
                }

                return false;
            }
        }

        /// <summary>
        /// Evaluates whether a given storeId/metadata is approved.
        /// </summary>
        public static bool IsStoreIdApproved(string storeId, string? displayName = null, string? filePath = null, bool isPst = false)
        {
            if (!IsTestModeActive) return true;

            // Check canonical personal PST
            if (IsPersonalPstPath(filePath)) return false;
            if (!string.IsNullOrWhiteSpace(displayName))
            {
                string disp = displayName.Trim();
                if (string.Equals(disp, CanonicalPersonalDisplayName, StringComparison.OrdinalIgnoreCase) ||
                    disp.IndexOf("personal-test@example.com", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    string.Equals(disp, "personal@gmail.com", StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }
            }

            if (!string.IsNullOrWhiteSpace(storeId))
            {
                if (string.Equals(storeId.Trim(), CanonicalPersonalStoreId, StringComparison.OrdinalIgnoreCase))
                    return false;

                lock (_lock)
                {
                    if (_blockedStoreIds.Contains(storeId.Trim()))
                        return false;
                }
            }

            // Reject every OST file
            if (IsOstPath(filePath)) return false;

            var approved = GetApprovedStores();
            if (approved.Count == 0) return false;

            // If previously validated and approved dynamically, accept it
            if (!string.IsNullOrWhiteSpace(storeId))
            {
                lock (_lock)
                {
                    if (_dynamicallyApprovedStoreIds.Contains(storeId.Trim()))
                        return true;
                }
            }

            bool matchesId = !string.IsNullOrWhiteSpace(storeId) && approved.Contains(storeId.Trim());

            bool matchesPath = false;
            if (!string.IsNullOrWhiteSpace(filePath))
            {
                string normFilePath = NormalizePath(filePath);
                if (!string.IsNullOrEmpty(normFilePath))
                {
                    foreach (var entry in approved)
                    {
                        string normApproved = NormalizePath(entry);
                        if (!string.IsNullOrEmpty(normApproved) &&
                            string.Equals(normFilePath, normApproved, StringComparison.OrdinalIgnoreCase))
                        {
                            matchesPath = true;
                            break;
                        }
                    }
                }
            }

            bool isPstStore = isPst || (!string.IsNullOrWhiteSpace(filePath) && filePath.EndsWith(".pst", StringComparison.OrdinalIgnoreCase));

            if (isPstStore)
            {
                return matchesPath || matchesId;
            }
            else
            {
                bool matchesName = !string.IsNullOrWhiteSpace(displayName) && approved.Contains(displayName.Trim());
                return matchesId || matchesName || matchesPath;
            }
        }

        private static void RecordBlockedStoreId(string? storeId)
        {
            if (!string.IsNullOrWhiteSpace(storeId))
            {
                lock (_lock)
                {
                    _blockedStoreIds.Add(storeId!.Trim());
                }
            }
        }

        private static void RecordApprovedStoreId(string? storeId)
        {
            if (!string.IsNullOrWhiteSpace(storeId))
            {
                lock (_lock)
                {
                    _dynamicallyApprovedStoreIds.Add(storeId!.Trim());
                }
            }
        }

        /// <summary>
        /// Filters a collection of mailboxes according to the test mode safety guard.
        /// Returns all mailboxes if test mode is inactive.
        /// </summary>
        public static List<MailboxInfo> FilterMailboxes(IEnumerable<MailboxInfo>? mailboxes)
        {
            if (mailboxes == null) return new List<MailboxInfo>();
            if (!IsTestModeActive) return new List<MailboxInfo>(mailboxes);

            var list = new List<MailboxInfo>();
            foreach (var mb in mailboxes)
            {
                if (IsStoreApproved(mb))
                {
                    list.Add(mb);
                }
            }
            return list;
        }
    }
}
