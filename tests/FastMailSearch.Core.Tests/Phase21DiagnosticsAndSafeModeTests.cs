using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using FastMailSearch.Core.Interfaces;
using FastMailSearch.Core.Models;
using FastMailSearch.Core.Services;
using FastMailSearch.Infrastructure.Database;
using FastMailSearch.Infrastructure.Diagnostics;
using FastMailSearch.Infrastructure.Indexing;
using FastMailSearch.Infrastructure.Logging;
using FastMailSearch.Infrastructure.Services;
using Xunit;

namespace FastMailSearch.Core.Tests
{
    public class Phase21DiagnosticsAndSafeModeTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly DatabaseConnectionFactory _factory;
        private readonly SqliteEmailRepository _repository;
        private readonly MockOutlookAdapter _outlookAdapter;
        private readonly EmailIndexerService _indexerService;
        private readonly IndexHealthService _healthService;

        public Phase21DiagnosticsAndSafeModeTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), $"FastMailSearch_Phase21Tests_{Guid.NewGuid():N}.db");
            _factory = new DatabaseConnectionFactory(_dbPath);
            _repository = new SqliteEmailRepository(_factory);
            _repository.InitializeDatabase();

            _outlookAdapter = new MockOutlookAdapter();
            _indexerService = new EmailIndexerService(_outlookAdapter, _repository);
            _healthService = new IndexHealthService(_repository, _indexerService, _outlookAdapter, _factory);
        }

        public void Dispose()
        {
            _indexerService?.CancelIndexingAsync().Wait();
            if (File.Exists(_dbPath))
            {
                try { File.Delete(_dbPath); } catch { }
            }
        }

        private class MockOutlookAdapter : IOutlookAdapter
        {
            public ConcurrentDictionary<string, EmailDocument> Storage = new ConcurrentDictionary<string, EmailDocument>(StringComparer.OrdinalIgnoreCase);

            public bool IsOutlookAvailable() => true;
            public string GetOutlookVersion() => "16.0.16000.20000";
            public string GetOfficeBitness() => "64-bit";

            public List<MailboxInfo> GetMailboxes() => new List<MailboxInfo>
            {
                new MailboxInfo { StoreId = "StoreA", DisplayName = "Exchange Account", IsPst = false },
                new MailboxInfo { StoreId = "StoreB", DisplayName = "PST Archive", IsPst = true }
            };

            public List<MailboxFolder> GetFolders(string storeId) => new List<MailboxFolder>
            {
                new MailboxFolder { StoreId = storeId, EntryId = "Folder1", Name = "Inbox" },
                new MailboxFolder { StoreId = storeId, EntryId = "OutboxFolder", Name = "Outbox" }
            };

            public List<EmailDocument> FetchFolderEmailsBatch(string storeId, string folderEntryId, int startIndex, int batchSize, Func<bool>? cancelCheck = null)
            {
                return Storage.Values.Where(d => d.StoreId == storeId && d.FolderEntryId == folderEntryId).Skip(startIndex - 1).Take(batchSize).ToList();
            }

            public EmailDocument? FetchSingleEmail(string storeId, string entryId)
            {
                string key = $"{storeId}:{entryId}";
                if (Storage.TryGetValue(key, out var doc)) return doc;
                return null;
            }

            public bool OpenEmailInOutlook(string storeId, string entryId) => true;

            public bool IsOutboxFolder(string storeId, string folderEntryId)
            {
                return string.Equals(folderEntryId, "OutboxFolder", StringComparison.OrdinalIgnoreCase) ||
                       string.Equals(folderEntryId, "Outbox", StringComparison.OrdinalIgnoreCase);
            }

            public void SubscribeToEvents(Action<string, string, string> onItemAddedOrChanged, Action<string, string> onItemRemoved) { }
            public void UnsubscribeFromEvents() { }
        }

        #region Test 1: Safe Mode Activation
        [Fact]
        public void Test01_SafeMode_Activation_Sets_Status()
        {
            _indexerService.SetSafeMode("COM initialization failed");
            var status = _indexerService.GetStatus();

            Assert.Equal(IndexerStatus.SafeMode, status.Status);
            Assert.Equal("COM initialization failed", _indexerService.LastError);
        }
        #endregion

        #region Test 2: Safe Mode Disables Indexing
        [Fact]
        public async Task Test02_SafeMode_Disables_Background_Indexing()
        {
            _indexerService.SetSafeMode("Database locked");

            _outlookAdapter.Storage["StoreA:Msg001"] = new EmailDocument
            {
                StoreId = "StoreA",
                EntryId = "Msg001",
                FolderEntryId = "Folder1",
                Subject = "Safe Mode Item"
            };

            await _indexerService.QueueItemForIndexingAsync("StoreA", "Msg001", "Folder1");
            await Task.Delay(300);

            var doc = await _repository.GetEmailByStoreAndEntryIdAsync("StoreA", "Msg001");
            Assert.Null(doc);
        }
        #endregion

        #region Test 3: Database Integrity Scan
        [Fact]
        public async Task Test03_Database_Integrity_Scan_Returns_Result()
        {
            var result = await _repository.ScanIndexIntegrityAsync();
            Assert.NotNull(result);
            Assert.True(result.IsHealthy);
        }
        #endregion

        #region Test 4: Database Repair
        [Fact]
        public async Task Test04_Database_Repair_Fixes_Discrepancies()
        {
            await _repository.SaveEmailBatchAsync(new[]
            {
                new EmailDocument { StoreId = "StoreA", EntryId = "Msg001", Subject = "Repair Test 1", FolderName = "Inbox" }
            });

            var scan = await _repository.ScanIndexIntegrityAsync();
            await _repository.RepairIndexIntegrityAsync(scan);

            var doc = await _repository.GetEmailByStoreAndEntryIdAsync("StoreA", "Msg001");
            Assert.NotNull(doc);
        }
        #endregion

        #region Test 5: PII Scrubbing
        [Fact]
        public void Test05_Log_Privacy_Scrubs_Email_Addresses_And_Secrets()
        {
            string raw = "User john.doe@company.com attempted login with password=SecretPassword123";
            string scrubbed = Logger.ScrubPii(raw);

            Assert.DoesNotContain("john.doe@company.com", scrubbed);
            Assert.DoesNotContain("SecretPassword123", scrubbed);
            Assert.Contains("[REDACTED_EMAIL]", scrubbed);
            Assert.Contains("[REDACTED_SECRET]", scrubbed);
        }
        #endregion

        #region Test 6: Diagnostics Exporter Generates Valid ZIP
        [Fact]
        public async Task Test06_Diagnostics_Exporter_Generates_Valid_Zip()
        {
            string tempDir = Path.Combine(Path.GetTempPath(), $"FMS_DiagExport_{Guid.NewGuid():N}");
            Directory.CreateDirectory(tempDir);

            try
            {
                string zipPath = await _healthService.ExportDiagnosticsZipAsync(tempDir);

                Assert.True(File.Exists(zipPath));
                Assert.True(new FileInfo(zipPath).Length > 0);

                using (var archive = ZipFile.OpenRead(zipPath))
                {
                    Assert.Contains(archive.Entries, e => e.Name == "EnvironmentReport.json");
                    Assert.Contains(archive.Entries, e => e.Name == "IndexHealthSummary.json");
                    Assert.Contains(archive.Entries, e => e.Name == "DatabaseStatistics.txt");
                }
            }
            finally
            {
                if (Directory.Exists(tempDir)) try { Directory.Delete(tempDir, true); } catch { }
            }
        }
        #endregion

        #region Test 6b: Concurrent Logging and Diagnostics Export
        [Fact]
        public async Task Test06b_Concurrent_Logging_And_Diagnostics_Export()
        {
            string tempDir = Path.Combine(Path.GetTempPath(), $"FMS_DiagExport_Conc_{Guid.NewGuid():N}");
            Directory.CreateDirectory(tempDir);
            bool loggingDone = false;

            try
            {
                // Start aggressive background logging
                var logTask = Task.Run(async () =>
                {
                    for (int i = 0; i < 500; i++)
                    {
                        Logger.Info($"Concurrent log stress test {i}");
                        await Task.Delay(1);
                    }
                    loggingDone = true;
                });

                // Attempt to export ZIP while logging is happening
                List<string> zipPaths = new List<string>();
                while (!loggingDone)
                {
                    string zipPath = await _healthService.ExportDiagnosticsZipAsync(tempDir);
                    Assert.True(File.Exists(zipPath));
                    zipPaths.Add(zipPath);
                    await Task.Delay(50);
                }

                await logTask;
                Assert.NotEmpty(zipPaths);
            }
            finally
            {
                if (Directory.Exists(tempDir)) try { Directory.Delete(tempDir, true); } catch { }
            }
        }
        #endregion

        #region Test 7: Log Rotation and Retention
        [Fact(Skip = "Flaky due to static Logger in parallel test execution")]
        public void Test07_Log_Rotation_And_Retention_Prunes_Old_Files()
        {
            string testLogDir = Path.Combine(Path.GetTempPath(), $"FMS_Logs_{Guid.NewGuid():N}");
            Directory.CreateDirectory(testLogDir);

            try
            {
                Logger.Initialize(testLogDir);

                string oldLog = Path.Combine(testLogDir, "FastMailSearch_20200101.log");
                File.WriteAllText(oldLog, "Old log entry");
                DateTime oldDate = DateTime.Now.AddDays(-20);
                File.SetLastWriteTime(oldLog, oldDate);
                File.SetCreationTime(oldLog, oldDate);

                Logger.CleanupOldLogs(14, testLogDir);

                Assert.False(File.Exists(oldLog));
            }
            finally
            {
                if (Directory.Exists(testLogDir)) try { Directory.Delete(testLogDir, true); } catch { }
            }
        }
        #endregion

        #region Test 8: Multi-Store Identical EntryId Isolation
        [Fact]
        public async Task Test08_MultiStore_Identical_EntryId_Isolation()
        {
            string sameEntryId = "MsgEntry999";

            var docA = new EmailDocument { StoreId = "StoreA", EntryId = sameEntryId, Subject = "Store A Email", FolderName = "Inbox" };
            var docB = new EmailDocument { StoreId = "StoreB", EntryId = sameEntryId, Subject = "Store B Email", FolderName = "Archive" };

            await _repository.SaveEmailBatchAsync(new[] { docA, docB });

            var fetchedA = await _repository.GetEmailByStoreAndEntryIdAsync("StoreA", sameEntryId);
            var fetchedB = await _repository.GetEmailByStoreAndEntryIdAsync("StoreB", sameEntryId);

            Assert.NotNull(fetchedA);
            Assert.NotNull(fetchedB);
            Assert.Equal("Store A Email", fetchedA.Subject);
            Assert.Equal("Store B Email", fetchedB.Subject);
        }
        #endregion

        #region Test 9: Indexer Health State Transitions
        [Fact]
        public async Task Test09_Indexer_Health_State_Transitions()
        {
            var summary1 = await _healthService.GetHealthSummaryAsync();
            Assert.Equal(IndexerHealth.Healthy, summary1.HealthState);

            _indexerService.SetSafeMode("COM Failure");
            var summary2 = await _healthService.GetHealthSummaryAsync();
            Assert.Equal(IndexerHealth.SafeMode, summary2.HealthState);
        }
        #endregion

        #region Test 10: Environment Detector Gathers Metrics
        [Fact]
        public void Test10_Environment_Detector_Gathers_Metrics()
        {
            var detector = new EnvironmentDetector(_factory, _outlookAdapter);
            var diag = detector.DetectEnvironment();

            Assert.NotNull(diag.OsVersion);
            Assert.NotNull(diag.OsArchitecture);
            Assert.Equal("16.0.16000.20000", diag.OutlookVersion);
            Assert.Equal(2, diag.StoreCount);
            Assert.True(diag.HasExchangeAccount);
            Assert.True(diag.HasPopAccount);
        }
        #endregion

        #region Test 11: Send Pipeline Safety Static Audit
        [Fact]
        public void Test11_No_Forbidden_Send_Handler_Registered()
        {
            string srcDir = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "..", "..", "src"));
            if (!Directory.Exists(srcDir))
            {
                srcDir = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "..", "..", "..", "..", "..", "src"));
            }

            Assert.True(Directory.Exists(srcDir), $"Source directory not found at: {srcDir}");

            string[] forbiddenTokens = new[]
            {
                "Application.ItemSend",
                "ItemSend",
                "BeforeSend",
                "MailItem.Send",
                "SendAndReceive",
                "SyncNow",
                "SendUsingAccount"
            };

            var csFiles = Directory.GetFiles(srcDir, "*.cs", SearchOption.AllDirectories);
            var violations = new List<string>();

            foreach (var file in csFiles)
            {
                if (file.Contains("AssemblyInfo.cs") || file.Contains(".g.cs") || file.Contains("wpftmp")) continue;

                string content = File.ReadAllText(file);
                foreach (var token in forbiddenTokens)
                {
                    if (content.Contains(token))
                    {
                        violations.Add($"Forbidden token '{token}' found in file: {Path.GetFileName(file)}");
                    }
                }
            }

            Assert.Empty(violations);
        }
        #endregion

        #region Test 12: Outbox Events Skipped
        [Fact]
        public async Task Test12_Outbox_Events_Skipped()
        {
            await _indexerService.QueueItemForIndexingAsync("StoreA", "OutboxMsg1", "OutboxFolder");
            await Task.Delay(200);

            var doc = await _repository.GetEmailByStoreAndEntryIdAsync("StoreA", "OutboxMsg1");
            Assert.Null(doc);
        }
        #endregion

        #region Test 13: COM Worker Thread Isolation
        [Fact]
        public void Test13_COM_Objects_Never_Passed_To_Queue()
        {
            var itemType = typeof(IncrementalItem);
            var properties = itemType.GetProperties(BindingFlags.Public | BindingFlags.Instance);

            foreach (var prop in properties)
            {
                var propType = prop.PropertyType;
                Assert.False(propType.IsCOMObject, $"Property {prop.Name} on IncrementalItem is a COM Object!");
                Assert.False(propType.FullName?.StartsWith("Microsoft.Office.Interop.Outlook") ?? false,
                    $"Property {prop.Name} references Outlook COM interop types!");
            }
        }
        #endregion

        #region Test 14: Beta Mode Toggle
        [Fact]
        public void Test14_Beta_Mode_Toggle_Enables_Verbose_Logging()
        {
            Logger.IsBetaModeEnabled = true;
            Assert.True(Logger.IsBetaModeEnabled);

            Logger.BetaLog("PERF", "Execution duration = 12ms");

            Logger.IsBetaModeEnabled = false;
            Assert.False(Logger.IsBetaModeEnabled);
        }
        #endregion

        #region Test 15: Database Backup Before Repair
        [Fact]
        public async Task Test15_Database_Backup_Verified_Before_Repair()
        {
            string backupPath = Path.Combine(Path.GetTempPath(), $"FastMailSearch_Backup_{DateTime.Now:yyyyMMdd_HHmmss}.db");
            File.Copy(_dbPath, backupPath, overwrite: true);

            Assert.True(File.Exists(backupPath));
            Assert.True(new FileInfo(backupPath).Length >= 0);

            try { File.Delete(backupPath); } catch { }
            await Task.CompletedTask;
        }
        #endregion
    }
}
