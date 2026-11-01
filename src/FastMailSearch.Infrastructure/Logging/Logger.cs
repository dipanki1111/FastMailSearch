using System;
using System.IO;
using System.Text.RegularExpressions;

namespace FastMailSearch.Infrastructure.Logging
{
    public static class Logger
    {
        public const string LogDirectoryEnvironmentVariable = "FASTMAILSEARCH_LOG_DIR";

        private static readonly object _logLock = new object();
        private static string? _customLogDirectory;
        private static bool _isBetaModeEnabled = false;

        public static bool IsBetaModeEnabled
        {
            get => _isBetaModeEnabled;
            set => _isBetaModeEnabled = value;
        }

        public static string GetDefaultLogDirectory()
        {
            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            return Path.Combine(localAppData, "FastMailSearch", "Logs");
        }

        public static string GetLogDirectory()
        {
            if (!string.IsNullOrWhiteSpace(_customLogDirectory))
            {
                return _customLogDirectory!;
            }

            string? overrideDir = Environment.GetEnvironmentVariable(LogDirectoryEnvironmentVariable);
            if (!string.IsNullOrWhiteSpace(overrideDir))
            {
                return overrideDir!.Trim();
            }

            return GetDefaultLogDirectory();
        }

        public static void ResetLogDirectory()
        {
            _customLogDirectory = null;
        }

        public static void Initialize(string? customLogDir = null)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(customLogDir))
                {
                    _customLogDirectory = customLogDir!.Trim();
                }

                string dir = GetLogDirectory();
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                CleanupOldLogs();
            }
            catch { }
        }

        public static string LogDirectory => GetLogDirectory();

        public static void Info(string message) => Log("INFO", message);
        public static void Warn(string message) => Log("WARN", message);
        public static void Error(string message, Exception? ex = null) => Log("ERROR", ex != null ? $"{message} | Exception: {ex.Message}\n{ex.StackTrace}" : message);
        public static void Critical(string message, Exception? ex = null) => Log("CRITICAL", ex != null ? $"{message} | Exception: {ex.Message}\n{ex.StackTrace}" : message);

        public static void LogStage(int stageNumber, string description, Exception? ex = null)
        {
            try
            {
                Initialize();
                string logFile = Path.Combine(LogDirectory, "UILaunch.log");
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                string message = $"[{timestamp}] [{stageNumber}] {description}";

                if (ex != null)
                {
                    message += $"{Environment.NewLine}  Exception Type: {ex.GetType().FullName}";
                    message += $"{Environment.NewLine}  Message: {ex.Message}";
                    message += $"{Environment.NewLine}  StackTrace: {ex.StackTrace}";
                    if (ex.InnerException != null)
                    {
                        message += $"{Environment.NewLine}  InnerException: {ex.InnerException.GetType().FullName}: {ex.InnerException.Message}";
                    }
                }

                lock (_logLock)
                {
                    File.AppendAllText(logFile, message + Environment.NewLine);
                }

                // Also forward to main logger
                Log("STAGE", $"[{stageNumber}] {description}" + (ex != null ? $" | Ex: {ex.Message}" : ""));
            }
            catch { }
        }

        public static void BetaLog(string category, string message)
        {
            if (_isBetaModeEnabled)
            {
                Log("BETA", $"[{category}] {message}");
            }
        }

        public static string ScrubPii(string input)
        {
            if (string.IsNullOrEmpty(input)) return string.Empty;

            // Scrub email addresses: user@domain.com -> [REDACTED_EMAIL]
            string scrubbed = Regex.Replace(input, @"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", "[REDACTED_EMAIL]");

            // Scrub authentication tokens or passwords if present in log strings
            scrubbed = Regex.Replace(scrubbed, @"(?i)(password|pwd|token|secret|auth)\s*[:=]\s*[^\s]+", "$1=[REDACTED_SECRET]");

            return scrubbed;
        }

        private static void Log(string level, string message)
        {
            try
            {
                Initialize();

                string safeMessage = ScrubPii(message);
                string dateSuffix = DateTime.Now.ToString("yyyyMMdd");
                string dailyFileName = $"FastMailSearch_{dateSuffix}.log";
                string dailyPath = Path.Combine(LogDirectory, dailyFileName);
                string masterPath = Path.Combine(LogDirectory, "FastMailSearch.log");

                string entry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [{level}] {safeMessage}{Environment.NewLine}";

                lock (_logLock)
                {
                    // Truncate daily log file if it exceeds 10MB
                    CheckAndRotateFile(dailyPath, 10 * 1024 * 1024);
                    File.AppendAllText(dailyPath, entry);

                    // Dual log to master FastMailSearch.log
                    CheckAndRotateFile(masterPath, 10 * 1024 * 1024);
                    File.AppendAllText(masterPath, entry);
                }
            }
            catch
            {
                // Fallback: zero exception propagation from logging system
            }
        }

        private static void CheckAndRotateFile(string filePath, long maxSizeBytes)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    var fi = new FileInfo(filePath);
                    if (fi.Length > maxSizeBytes)
                    {
                        string backupPath = $"{filePath}.old";
                        if (File.Exists(backupPath)) File.Delete(backupPath);
                        File.Move(filePath, backupPath);
                    }
                }
            }
            catch { }
        }

        public static void CleanupOldLogs(int maxRetentionDays = 14, string? targetDirectory = null)
        {
            try
            {
                string dir = !string.IsNullOrWhiteSpace(targetDirectory) ? targetDirectory : LogDirectory;
                if (!Directory.Exists(dir)) return;

                var files = Directory.GetFiles(dir, "FastMailSearch_*.log");
                DateTime threshold = DateTime.Now.AddDays(-maxRetentionDays);

                foreach (var file in files)
                {
                    var fi = new FileInfo(file);
                    bool shouldDelete = fi.LastWriteTime < threshold || fi.CreationTime < threshold;

                    if (!shouldDelete)
                    {
                        string fileName = Path.GetFileNameWithoutExtension(file);
                        if (fileName.StartsWith("FastMailSearch_", StringComparison.OrdinalIgnoreCase) && fileName.Length >= 23)
                        {
                            string datePart = fileName.Substring(15, 8);
                            if (DateTime.TryParseExact(datePart, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out DateTime fileDate))
                            {
                                if (fileDate < threshold)
                                {
                                    shouldDelete = true;
                                }
                            }
                        }
                    }

                    if (shouldDelete)
                    {
                        try { fi.Delete(); } catch { }
                    }
                }
            }
            catch { }
        }
    }
}
