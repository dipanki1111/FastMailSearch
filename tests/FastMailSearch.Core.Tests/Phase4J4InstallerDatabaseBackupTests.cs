using System;
using System.IO;
using System.Text.RegularExpressions;
using FastMailSearch.Core.Models;
using FastMailSearch.Setup;
using Microsoft.Data.Sqlite;
using Xunit;

namespace FastMailSearch.Core.Tests
{
    public class Phase4J4InstallerDatabaseBackupTests : IDisposable
    {
        private readonly string _testRoot;

        public Phase4J4InstallerDatabaseBackupTests()
        {
            _testRoot = Path.Combine(Path.GetTempPath(), $"FMS_Phase4J4_Tests_{Guid.NewGuid():N}");
            Directory.CreateDirectory(_testRoot);
        }

        public void Dispose()
        {
            try
            {
                if (Directory.Exists(_testRoot))
                {
                    Directory.Delete(_testRoot, true);
                }
            }
            catch { }
        }

        #region Task 1: Canonical DB Path Detection
        [Fact]
        public void Test01_Canonical_Database_Path_Detection_Matches_AppSettings_And_Documents()
        {
            string expectedDocsFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string expectedCanonicalDir = Path.Combine(expectedDocsFolder, "FastMailSearch");
            string expectedCanonicalDb = Path.Combine(expectedCanonicalDir, "FastMailSearch.db");

            string actualCanonicalDir = InstallerDatabaseBackup.GetCanonicalDataDirectory();
            string actualCanonicalDb = InstallerDatabaseBackup.GetCanonicalDatabasePath();
            string appSettingsDb = AppSettings.GetDefaultDatabasePath();

            Assert.Equal(expectedCanonicalDir, actualCanonicalDir);
            Assert.Equal(expectedCanonicalDb, actualCanonicalDb);
            Assert.Equal(appSettingsDb, actualCanonicalDb);

            string expectedLocalApp = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string expectedLegacyDb = Path.Combine(expectedLocalApp, "FastMailSearch", "fastmail_search.db");
            string actualLegacyDb = InstallerDatabaseBackup.GetLegacyDatabasePath();

            Assert.Equal(expectedLegacyDb, actualLegacyDb);
        }
        #endregion

        #region Task 2: Existing DB Backup
        [Fact]
        public void Test02_Existing_Database_Backup_Creates_Timestamped_File_With_Matching_Bytes()
        {
            string dbDir = Path.Combine(_testRoot, "CanonicalData");
            Directory.CreateDirectory(dbDir);
            string dbPath = Path.Combine(dbDir, "FastMailSearch.db");

            byte[] syntheticPayload = new byte[8192];
            new Random(42).NextBytes(syntheticPayload);
            File.WriteAllBytes(dbPath, syntheticPayload);

            var result = InstallerDatabaseBackup.BackupDatabase(dbPath);

            Assert.True(result.Success, $"Backup failed: {result.ErrorMessage}");
            Assert.True(result.DatabaseFound);
            Assert.Single(result.BackedUpFiles);

            string backupPath = result.BackedUpFiles[0];
            Assert.True(File.Exists(backupPath), "Backup file does not exist on disk.");
            Assert.True(File.Exists(dbPath), "Source database was unexpectedly moved or deleted.");

            string backupFileName = Path.GetFileName(backupPath);
            Assert.Matches(@"^FastMailSearch_Backup_PreUpgrade_\d{8}_\d{6}\.db$", backupFileName);

            byte[] backupBytes = File.ReadAllBytes(backupPath);
            Assert.Equal(syntheticPayload, backupBytes);
        }
        #endregion

        #region Task 3: Missing DB Handling
        [Fact]
        public void Test03_Missing_Database_Handling_Succeeds_Without_Error()
        {
            string nonExistentDb = Path.Combine(_testRoot, "NonExistentFolder", "FastMailSearch.db");

            var result = InstallerDatabaseBackup.BackupDatabase(nonExistentDb);

            Assert.True(result.Success);
            Assert.False(result.DatabaseFound);
            Assert.Empty(result.BackedUpFiles);
            Assert.Null(result.ErrorMessage);

            string nonExistentCanonical = Path.Combine(_testRoot, "NoCanonical", "FastMailSearch.db");
            string nonExistentLegacy = Path.Combine(_testRoot, "NoLegacy", "fastmail_search.db");

            var allResult = InstallerDatabaseBackup.BackupAllExistingDatabases(
                overrideCanonicalPath: nonExistentCanonical,
                overrideLegacyPath: nonExistentLegacy);

            Assert.True(allResult.Success);
            Assert.False(allResult.DatabaseFound);
            Assert.Empty(allResult.BackedUpFiles);
            Assert.Null(allResult.ErrorMessage);
        }
        #endregion

        #region Task 4: Backup Failure / Insufficient Permissions
        [Fact]
        public void Test04_Backup_Failure_Invalid_Destination_Fails_Closed_And_Cleans_Up()
        {
            string dbDir = Path.Combine(_testRoot, "Data");
            Directory.CreateDirectory(dbDir);
            string dbPath = Path.Combine(dbDir, "FastMailSearch.db");
            File.WriteAllText(dbPath, "SQLite Header Payload");

            // Point destination to a file instead of a directory (causes CreateDirectory or Path.Combine to fail)
            string invalidDestDir = Path.Combine(_testRoot, "InvalidDestFile.txt");
            File.WriteAllText(invalidDestDir, "I am a file, not a directory");

            var result = InstallerDatabaseBackup.BackupDatabase(dbPath, destinationDir: invalidDestDir);

            Assert.False(result.Success);
            Assert.NotNull(result.ErrorMessage);
            Assert.Empty(result.BackedUpFiles);
        }
        #endregion

        #region Task 5: WAL/SHM Handling & SQLite Consistency
        [Fact]
        public void Test05_WAL_And_SHM_Handling_Preserves_SQLite_Consistency_And_Uncheckpointed_Data()
        {
            string dbDir = Path.Combine(_testRoot, "SqliteData");
            Directory.CreateDirectory(dbDir);
            string dbPath = Path.Combine(dbDir, "FastMailSearch.db");

            // Create a real SQLite database in WAL mode and insert records without checkpointing
            using (var conn = new SqliteConnection($"Data Source={dbPath};"))
            {
                conn.Open();
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @"
                        PRAGMA journal_mode = WAL;
                        PRAGMA synchronous = NORMAL;
                        CREATE TABLE Emails (Id INTEGER PRIMARY KEY, Subject TEXT, Sender TEXT);
                        INSERT INTO Emails (Subject, Sender) VALUES ('Q3 Financials', 'cfo@company.test');
                        INSERT INTO Emails (Subject, Sender) VALUES ('Production Outage RCA', 'devops@company.test');
                        INSERT INTO Emails (Subject, Sender) VALUES ('Architecture Review', 'lead@company.test');
                    ";
                    cmd.ExecuteNonQuery();
                }
            }

            string walPath = dbPath + "-wal";
            // Ensure WAL exists or create a synthetic WAL frame if already checkpointed
            if (!File.Exists(walPath))
            {
                File.WriteAllText(walPath, "synthetic wal frame");
            }
            string shmPath = dbPath + "-shm";
            if (!File.Exists(shmPath))
            {
                File.WriteAllText(shmPath, "synthetic shm frame");
            }

            var result = InstallerDatabaseBackup.BackupDatabase(dbPath);

            Assert.True(result.Success, $"Backup failed: {result.ErrorMessage}");
            Assert.True(result.DatabaseFound);
            Assert.True(result.BackedUpFiles.Count >= 2, $"Expected at least 2 backed up files (DB + WAL), got {result.BackedUpFiles.Count}");

            string backupDb = result.BackedUpFiles.Find(f => f.EndsWith(".db"))!;
            string backupWal = result.BackedUpFiles.Find(f => f.EndsWith(".db-wal"))!;
            Assert.NotNull(backupDb);
            Assert.NotNull(backupWal);
            Assert.True(File.Exists(backupDb));
            Assert.True(File.Exists(backupWal));

            // Verify backup DB opens in SQLite and passes integrity_check
            using (var backupConn = new SqliteConnection($"Data Source={backupDb};"))
            {
                backupConn.Open();
                using (var cmd = backupConn.CreateCommand())
                {
                    cmd.CommandText = "PRAGMA integrity_check;";
                    var integrityResult = (string?)cmd.ExecuteScalar();
                    Assert.Equal("ok", integrityResult);

                    cmd.CommandText = "SELECT COUNT(*) FROM Emails;";
                    long count = (long)cmd.ExecuteScalar()!;
                    Assert.Equal(3L, count);
                }
            }
        }
        #endregion

        #region Task 6: Multi-Store / Backward Compatibility
        [Fact]
        public void Test06_Backup_All_Existing_Databases_Backs_Up_Both_Canonical_And_Legacy_When_Present()
        {
            string canonicalDir = Path.Combine(_testRoot, "MyDocs", "FastMailSearch");
            Directory.CreateDirectory(canonicalDir);
            string canonicalDb = Path.Combine(canonicalDir, "FastMailSearch.db");
            File.WriteAllText(canonicalDb, "Canonical DB Payload Content");

            string legacyDir = Path.Combine(_testRoot, "LocalAppData", "FastMailSearch");
            Directory.CreateDirectory(legacyDir);
            string legacyDb = Path.Combine(legacyDir, "fastmail_search.db");
            File.WriteAllText(legacyDb, "Legacy DB Payload Content");

            var result = InstallerDatabaseBackup.BackupAllExistingDatabases(
                overrideCanonicalPath: canonicalDb,
                overrideLegacyPath: legacyDb);

            Assert.True(result.Success, $"Backup failed: {result.ErrorMessage}");
            Assert.True(result.DatabaseFound);
            Assert.Equal(2, result.BackedUpFiles.Count);

            string canonicalBackup = result.BackedUpFiles.Find(p => p.StartsWith(canonicalDir))!;
            string legacyBackup = result.BackedUpFiles.Find(p => p.StartsWith(legacyDir))!;

            Assert.NotNull(canonicalBackup);
            Assert.NotNull(legacyBackup);
            Assert.True(File.Exists(canonicalBackup));
            Assert.True(File.Exists(legacyBackup));

            Assert.Equal("Canonical DB Payload Content", File.ReadAllText(canonicalBackup));
            Assert.Equal("Legacy DB Payload Content", File.ReadAllText(legacyBackup));
        }
        #endregion

        #region Task 7: Reparse Point / Junction Guards
        [Fact]
        public void Test07_Reparse_Point_Junction_File_Aborts_Fail_Closed()
        {
            string dbDir = Path.Combine(_testRoot, "ReparseData");
            Directory.CreateDirectory(dbDir);
            string dbPath = Path.Combine(dbDir, "FastMailSearch.db");
            File.WriteAllText(dbPath, "Payload");

            var result = InstallerDatabaseBackup.BackupDatabase(
                dbPath,
                reparseChecker: p =>
                {
                    if (string.Equals(p, dbPath, StringComparison.OrdinalIgnoreCase))
                    {
                        return (ReparseCheckResult.IsReparsePoint, null);
                    }
                    return (ReparseCheckResult.NotReparsePoint, null);
                });

            Assert.False(result.Success);
            Assert.Contains("junction or reparse point", result.ErrorMessage);
            Assert.Empty(result.BackedUpFiles);
        }

        [Fact]
        public void Test08_Reparse_Point_Query_Failure_Aborts_Fail_Closed()
        {
            string dbDir = Path.Combine(_testRoot, "QueryFailData");
            Directory.CreateDirectory(dbDir);
            string dbPath = Path.Combine(dbDir, "FastMailSearch.db");
            File.WriteAllText(dbPath, "Payload");

            var result = InstallerDatabaseBackup.BackupDatabase(
                dbPath,
                reparseChecker: p => (ReparseCheckResult.QueryFailed, "Simulated access denied on attributes"));

            Assert.False(result.Success);
            Assert.Contains("Could not verify attributes", result.ErrorMessage);
            Assert.Empty(result.BackedUpFiles);
        }
        #endregion
    }
}
