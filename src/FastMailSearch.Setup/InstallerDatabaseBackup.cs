using System;
using System.Collections.Generic;
using System.IO;

namespace FastMailSearch.Setup
{
    public enum ReparseCheckResult
    {
        DoesNotExist,
        NotReparsePoint,
        IsReparsePoint,
        QueryFailed
    }

    public class BackupResult
    {
        public bool Success { get; set; }
        public string? ErrorMessage { get; set; }
        public List<string> BackedUpFiles { get; } = new List<string>();
        public bool DatabaseFound { get; set; }
    }

    public static class InstallerDatabaseBackup
    {
        public const string CanonicalDatabaseFileName = "FastMailSearch.db";
        public const string LegacyDatabaseFileName = "fastmail_search.db";
        public const string DatabaseDirName = "FastMailSearch";

        public static string GetCanonicalDataDirectory()
        {
            string docsFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            return Path.Combine(docsFolder, DatabaseDirName);
        }

        public static string GetCanonicalDatabasePath()
        {
            return Path.Combine(GetCanonicalDataDirectory(), CanonicalDatabaseFileName);
        }

        public static string GetLegacyDataDirectory()
        {
            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            return Path.Combine(localAppData, DatabaseDirName);
        }

        public static string GetLegacyDatabasePath()
        {
            return Path.Combine(GetLegacyDataDirectory(), LegacyDatabaseFileName);
        }

        public static ReparseCheckResult GetReparseStatus(string path, out string? errorDetail)
        {
            errorDetail = null;
            try
            {
                FileAttributes attrs = File.GetAttributes(path);
                if ((attrs & FileAttributes.ReparsePoint) != 0)
                {
                    return ReparseCheckResult.IsReparsePoint;
                }

                return ReparseCheckResult.NotReparsePoint;
            }
            catch (FileNotFoundException)
            {
                return ReparseCheckResult.DoesNotExist;
            }
            catch (DirectoryNotFoundException)
            {
                return ReparseCheckResult.DoesNotExist;
            }
            catch (Exception ex)
            {
                errorDetail = ex.Message;
                return ReparseCheckResult.QueryFailed;
            }
        }

        /// <summary>
        /// Backs up a single database file and its associated WAL/SHM files to the destination directory.
        /// </summary>
        public static BackupResult BackupDatabase(
            string dbFilePath,
            string? destinationDir = null,
            Func<string, (ReparseCheckResult Status, string? Error)>? reparseChecker = null,
            Action<string>? logger = null)
        {
            var result = new BackupResult();
            var reparseCheck = reparseChecker ?? (p =>
            {
                var status = GetReparseStatus(p, out string? err);
                return (status, err);
            });

            try
            {
                var dbStatus = reparseCheck(dbFilePath);
                if (dbStatus.Status == ReparseCheckResult.DoesNotExist)
                {
                    result.DatabaseFound = false;
                    result.Success = true;
                    return result;
                }

                if (dbStatus.Status == ReparseCheckResult.IsReparsePoint)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Database file '{dbFilePath}' is a junction or reparse point. Aborting setup to prevent arbitrary file redirection.";
                    logger?.Invoke($"[SECURITY ERROR] {result.ErrorMessage}");
                    return result;
                }

                if (dbStatus.Status == ReparseCheckResult.QueryFailed)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Could not verify attributes of database file '{dbFilePath}': {dbStatus.Error}. Aborting setup fail-closed.";
                    logger?.Invoke($"[SECURITY ERROR] {result.ErrorMessage}");
                    return result;
                }

                result.DatabaseFound = true;
                string targetDir = destinationDir ?? Path.GetDirectoryName(dbFilePath) ?? string.Empty;
                if (string.IsNullOrWhiteSpace(targetDir))
                {
                    result.Success = false;
                    result.ErrorMessage = $"Invalid target directory for database backup: '{dbFilePath}'.";
                    logger?.Invoke($"[ERROR] {result.ErrorMessage}");
                    return result;
                }

                // Verify destination directory reparse status
                var dirStatus = reparseCheck(targetDir);
                if (dirStatus.Status == ReparseCheckResult.IsReparsePoint)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Backup destination directory '{targetDir}' is a junction or reparse point. Aborting setup.";
                    logger?.Invoke($"[SECURITY ERROR] {result.ErrorMessage}");
                    return result;
                }
                if (dirStatus.Status == ReparseCheckResult.QueryFailed)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Could not verify attributes of backup directory '{targetDir}': {dirStatus.Error}. Aborting setup fail-closed.";
                    logger?.Invoke($"[SECURITY ERROR] {result.ErrorMessage}");
                    return result;
                }

                if (dirStatus.Status == ReparseCheckResult.DoesNotExist)
                {
                    Directory.CreateDirectory(targetDir);
                }

                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string baseFileName = Path.GetFileNameWithoutExtension(dbFilePath);
                string backupDbName = $"{baseFileName}_Backup_PreUpgrade_{timestamp}.db";
                string backupDbPath = Path.Combine(targetDir, backupDbName);

                // Verify backup path
                var backupStatus = reparseCheck(backupDbPath);
                if (backupStatus.Status == ReparseCheckResult.IsReparsePoint)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Backup destination file '{backupDbPath}' is a junction or reparse point. Aborting setup.";
                    logger?.Invoke($"[SECURITY ERROR] {result.ErrorMessage}");
                    return result;
                }
                if (backupStatus.Status == ReparseCheckResult.QueryFailed)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Could not verify attributes of backup destination '{backupDbPath}': {backupStatus.Error}. Aborting setup fail-closed.";
                    logger?.Invoke($"[SECURITY ERROR] {result.ErrorMessage}");
                    return result;
                }

                logger?.Invoke($"Creating pre-upgrade database backup: {backupDbName}...");
                File.Copy(dbFilePath, backupDbPath, overwrite: true);

                // Verify DB copy
                var sourceFi = new FileInfo(dbFilePath);
                var backupFi = new FileInfo(backupDbPath);
                if (!backupFi.Exists || backupFi.Length != sourceFi.Length)
                {
                    throw new IOException($"Database backup verification failed: Expected length {sourceFi.Length}, found {backupFi.Length}.");
                }
                result.BackedUpFiles.Add(backupDbPath);

                // Check and back up WAL file
                string walPath = dbFilePath + "-wal";
                var walStatus = reparseCheck(walPath);
                if (walStatus.Status == ReparseCheckResult.IsReparsePoint)
                {
                    throw new InvalidOperationException($"WAL file '{walPath}' is a junction or reparse point. Aborting setup.");
                }
                if (walStatus.Status == ReparseCheckResult.QueryFailed)
                {
                    throw new InvalidOperationException($"Could not verify attributes of WAL file '{walPath}': {walStatus.Error}. Aborting setup fail-closed.");
                }

                if (walStatus.Status == ReparseCheckResult.NotReparsePoint)
                {
                    string backupWalPath = backupDbPath + "-wal";
                    var backupWalStatus = reparseCheck(backupWalPath);
                    if (backupWalStatus.Status == ReparseCheckResult.IsReparsePoint)
                    {
                        throw new InvalidOperationException($"Backup destination WAL '{backupWalPath}' is a junction or reparse point.");
                    }
                    if (backupWalStatus.Status == ReparseCheckResult.QueryFailed)
                    {
                        throw new InvalidOperationException($"Could not verify attributes of backup destination WAL '{backupWalPath}': {backupWalStatus.Error}.");
                    }

                    logger?.Invoke($"Backing up SQLite WAL file: {Path.GetFileName(backupWalPath)}...");
                    File.Copy(walPath, backupWalPath, overwrite: true);

                    var sourceWalFi = new FileInfo(walPath);
                    var backupWalFi = new FileInfo(backupWalPath);
                    if (!backupWalFi.Exists || backupWalFi.Length != sourceWalFi.Length)
                    {
                        throw new IOException($"WAL backup verification failed: Expected length {sourceWalFi.Length}, found {backupWalFi.Length}.");
                    }
                    result.BackedUpFiles.Add(backupWalPath);
                }

                // Check and back up SHM file
                string shmPath = dbFilePath + "-shm";
                var shmStatus = reparseCheck(shmPath);
                if (shmStatus.Status == ReparseCheckResult.IsReparsePoint)
                {
                    throw new InvalidOperationException($"SHM file '{shmPath}' is a junction or reparse point. Aborting setup.");
                }
                if (shmStatus.Status == ReparseCheckResult.QueryFailed)
                {
                    throw new InvalidOperationException($"Could not verify attributes of SHM file '{shmPath}': {shmStatus.Error}. Aborting setup fail-closed.");
                }

                if (shmStatus.Status == ReparseCheckResult.NotReparsePoint)
                {
                    string backupShmPath = backupDbPath + "-shm";
                    var backupShmStatus = reparseCheck(backupShmPath);
                    if (backupShmStatus.Status == ReparseCheckResult.IsReparsePoint)
                    {
                        throw new InvalidOperationException($"Backup destination SHM '{backupShmPath}' is a junction or reparse point.");
                    }
                    if (backupShmStatus.Status == ReparseCheckResult.QueryFailed)
                    {
                        throw new InvalidOperationException($"Could not verify attributes of backup destination SHM '{backupShmPath}': {backupShmStatus.Error}.");
                    }

                    logger?.Invoke($"Backing up SQLite SHM file: {Path.GetFileName(backupShmPath)}...");
                    File.Copy(shmPath, backupShmPath, overwrite: true);

                    var sourceShmFi = new FileInfo(shmPath);
                    var backupShmFi = new FileInfo(backupShmPath);
                    if (!backupShmFi.Exists || backupShmFi.Length != sourceShmFi.Length)
                    {
                        throw new IOException($"SHM backup verification failed: Expected length {sourceShmFi.Length}, found {backupShmFi.Length}.");
                    }
                    result.BackedUpFiles.Add(backupShmPath);
                }

                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                // Roll back any partially created backup files to avoid leaving corrupted artifacts
                foreach (var file in result.BackedUpFiles)
                {
                    try
                    {
                        if (File.Exists(file))
                        {
                            File.Delete(file);
                        }
                    }
                    catch { }
                }
                result.BackedUpFiles.Clear();
                result.Success = false;
                result.ErrorMessage = $"Failed to create pre-upgrade database backup for '{dbFilePath}': {ex.Message}";
                logger?.Invoke($"[ERROR] {result.ErrorMessage}");
                return result;
            }
        }

        /// <summary>
        /// Scans canonical and legacy locations and backs up all existing databases safely.
        /// </summary>
        public static BackupResult BackupAllExistingDatabases(
            Action<string>? logger = null,
            string? overrideCanonicalPath = null,
            string? overrideLegacyPath = null,
            Func<string, (ReparseCheckResult Status, string? Error)>? reparseChecker = null)
        {
            var aggregatedResult = new BackupResult();
            var candidatePaths = new List<string>();

            // 1. Canonical DB location (Documents\FastMailSearch\FastMailSearch.db)
            string canonicalPath = overrideCanonicalPath ?? GetCanonicalDatabasePath();
            candidatePaths.Add(canonicalPath);

            // 2. Canonical dir with legacy name (Documents\FastMailSearch\fastmail_search.db)
            string canonicalDir = Path.GetDirectoryName(canonicalPath) ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(canonicalDir))
            {
                candidatePaths.Add(Path.Combine(canonicalDir, LegacyDatabaseFileName));
            }

            // 3. Legacy AppData location (%LOCALAPPDATA%\FastMailSearch\fastmail_search.db)
            string legacyPath = overrideLegacyPath ?? GetLegacyDatabasePath();
            candidatePaths.Add(legacyPath);

            // 4. Legacy AppData with canonical name (%LOCALAPPDATA%\FastMailSearch\FastMailSearch.db)
            string legacyDir = Path.GetDirectoryName(legacyPath) ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(legacyDir))
            {
                candidatePaths.Add(Path.Combine(legacyDir, CanonicalDatabaseFileName));
            }

            // Deduplicate candidate paths
            var processedPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var candidate in candidatePaths)
            {
                string norm;
                try
                {
                    norm = Path.GetFullPath(candidate);
                }
                catch
                {
                    continue;
                }

                if (!processedPaths.Add(norm))
                {
                    continue;
                }

                var singleResult = BackupDatabase(norm, null, reparseChecker, logger);
                if (!singleResult.Success)
                {
                    // Fail closed immediately on any backup error
                    aggregatedResult.Success = false;
                    aggregatedResult.ErrorMessage = singleResult.ErrorMessage;
                    return aggregatedResult;
                }

                if (singleResult.DatabaseFound)
                {
                    aggregatedResult.DatabaseFound = true;
                    aggregatedResult.BackedUpFiles.AddRange(singleResult.BackedUpFiles);
                }
            }

            if (!aggregatedResult.DatabaseFound)
            {
                logger?.Invoke("No existing database found to back up (fresh installation).");
            }

            aggregatedResult.Success = true;
            return aggregatedResult;
        }
    }
}
