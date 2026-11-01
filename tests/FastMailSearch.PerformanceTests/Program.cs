using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using FastMailSearch.Core.Interfaces;
using FastMailSearch.Core.Models;
using FastMailSearch.Core.Search;
using FastMailSearch.Infrastructure.Database;
using FastMailSearch.Infrastructure.Indexing;

namespace FastMailSearch.PerformanceTests
{
    class Program
    {
        static async Task Main(string[] args)
        {
            if (args.Length > 0 && args[0] == "--integrity")
            {
                await RunIntegrityBenchmarkAsync();
                return;
            }

            if (args.Length > 0 && args[0] == "--incremental")
            {
                await RunIncrementalBenchmarkAsync();
                return;
            }

            if (args.Length > 0 && args[0] == "--synthetic")
            {
                await RunSyntheticBenchmarkAsync();
                return;
            }

            await RunIntegrityBenchmarkAsync();
            await RunIncrementalBenchmarkAsync();
            await RunSyntheticBenchmarkAsync();
        }

        public static async Task RunIntegrityBenchmarkAsync()
        {
            Console.WriteLine("=================================================");
            Console.WriteLine(" FastMail Search - Phase 20 Index Integrity & Self-Healing Benchmark");
            Console.WriteLine("=================================================");

            string benchDbPath = Path.Combine(Path.GetTempPath(), "FastMailSearch_IntegrityBench.db");
            if (File.Exists(benchDbPath))
            {
                try { File.Delete(benchDbPath); } catch { }
            }

            var factory = new DatabaseConnectionFactory(benchDbPath);
            var repo = new SqliteEmailRepository(factory);
            repo.InitializeDatabase();

            int[] targets = new[] { 10000, 50000, 100000 };
            int currentCount = 0;

            var reportLines = new List<string>();
            reportLines.Add("| Record Count | Scan Duration (ms) | Missing FTS Repair (ms) | Orphan FTS Repair (ms) | Queue Recovery (ms) | Search Latency During Repair (ms) | Peak RAM (MB) |");
            reportLines.Add("|--------------|--------------------|-------------------------|------------------------|---------------------|-----------------------------------|---------------|");

            foreach (var target in targets)
            {
                int needed = target - currentCount;
                Console.WriteLine($"\n--> Generating {needed:N0} emails for integrity benchmark (Total: {target:N0})...");
                await SyntheticDataGenerator.GenerateSyntheticEmailsAsync(repo, needed, p => { });
                currentCount = target;

                // 1. Measure Integrity Scan Duration
                var swScan = Stopwatch.StartNew();
                var scanRes = await repo.ScanIndexIntegrityAsync();
                swScan.Stop();

                // 2. Simulate 1,000 Missing FTS Records & Measure Repair Duration
                using (var conn = factory.CreateConnection())
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "DELETE FROM EmailSearch WHERE rowid IN (SELECT rowid FROM EmailSearch LIMIT 1000);";
                    cmd.ExecuteNonQuery();
                }

                var scanMissing = await repo.ScanIndexIntegrityAsync();
                var swMissingRepair = Stopwatch.StartNew();
                var searchTask = repo.SearchAsync(new SearchQuery { RawQuery = "kfintech", Limit = 50 });
                var repairTask = repo.RepairIndexIntegrityAsync(scanMissing);
                await Task.WhenAll(searchTask, repairTask);
                swMissingRepair.Stop();
                double searchLatency = (await searchTask).ElapsedSeconds * 1000.0;

                // 3. Simulate 1,000 Orphan FTS Records & Measure Repair Duration
                using (var conn = factory.CreateConnection())
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "INSERT INTO EmailSearch (EmailId, Subject, SenderName, SenderEmail, Recipients, BodyText, AttachmentNames) SELECT 'orphan_' || rowid, 'Orphan', 'Orphan', 'o@test.com', 'r@test.com', 'Body', '' FROM EmailSearch LIMIT 1000;";
                    cmd.ExecuteNonQuery();
                }

                var scanOrphan = await repo.ScanIndexIntegrityAsync();
                var swOrphanRepair = Stopwatch.StartNew();
                await repo.RepairIndexIntegrityAsync(scanOrphan);
                swOrphanRepair.Stop();

                // 4. Measure Queue Recovery Duration (1,000 Pending Changes)
                for (int i = 1; i <= 1000; i++)
                {
                    await repo.EnqueuePendingChangeAsync("Store1", $"Bench_Pending_{i}", "Add", "Folder1");
                }
                var swQueueRec = Stopwatch.StartNew();
                await repo.RecoverStaleProcessingChangesAsync();
                var recoveredPending = await repo.GetPendingChangesAsync(1000);
                swQueueRec.Stop();

                double ramMb = Process.GetCurrentProcess().WorkingSet64 / (1024.0 * 1024.0);

                Console.WriteLine($"  [{target:N0} Records] Scan: {swScan.Elapsed.TotalMilliseconds:F2}ms | Missing Repair: {swMissingRepair.Elapsed.TotalMilliseconds:F2}ms | Orphan Repair: {swOrphanRepair.Elapsed.TotalMilliseconds:F2}ms | Queue Rec: {swQueueRec.Elapsed.TotalMilliseconds:F2}ms | Search Latency: {searchLatency:F2}ms | RAM: {ramMb:F1} MB");
                reportLines.Add($"| {target:N0} | {swScan.Elapsed.TotalMilliseconds,18:F2} | {swMissingRepair.Elapsed.TotalMilliseconds,23:F2} | {swOrphanRepair.Elapsed.TotalMilliseconds,22:F2} | {swQueueRec.Elapsed.TotalMilliseconds,19:F2} | {searchLatency,33:F2} | {ramMb,13:F1} |");
            }

            Console.WriteLine("\n=================================================");
            Console.WriteLine(" INTEGRITY BENCHMARK REPORT SUMMARY");
            Console.WriteLine("=================================================");
            foreach (var line in reportLines)
            {
                Console.WriteLine(line);
            }

            string reportPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "IntegrityBenchmarkResult.txt");
            File.WriteAllLines(reportPath, reportLines);
            Console.WriteLine($"\nReport saved to: {reportPath}");

            try { File.Delete(benchDbPath); } catch { }
        }

        public static async Task RunIncrementalBenchmarkAsync()
        {
            Console.WriteLine("\n=================================================");
            Console.WriteLine(" FastMail Search - Incremental Indexing Benchmark");
            Console.WriteLine("=================================================");

            string benchDbPath = Path.Combine(Path.GetTempPath(), "FastMailSearch_IncrementalBench.db");
            if (File.Exists(benchDbPath))
            {
                try { File.Delete(benchDbPath); } catch { }
            }

            var factory = new DatabaseConnectionFactory(benchDbPath);
            var repo = new SqliteEmailRepository(factory);
            repo.InitializeDatabase();

            int[] burstSizes = new[] { 10, 100, 500, 1000 };
            var reportLines = new List<string>();
            reportLines.Add("| Burst Count | Queue Time (ms) | Total Index Time (ms) | Throughput (items/sec) | Search Latency While Indexing (ms) | Peak RAM (MB) |");
            reportLines.Add("|-------------|------------------|------------------------|------------------------|------------------------------------|---------------|");

            foreach (var burst in burstSizes)
            {
                var mockOutlook = new MockOutlookAdapter();
                var indexer = new EmailIndexerService(mockOutlook, repo);

                for (int i = 1; i <= burst; i++)
                {
                    var doc = new EmailDocument
                    {
                        StoreId = "Store1",
                        EntryId = $"Inc_Entry_{burst}_{i}",
                        Subject = $"Incremental Burst Invoice Notification {i}",
                        SenderName = $"Vendor {i}",
                        SenderEmail = $"vendor{i}@supplier.com",
                        BodyText = $"Body text for incremental indexing performance benchmark item {i}.",
                        FolderName = "Inbox",
                        FolderEntryId = "Folder1",
                        ReceivedDate = DateTime.Now
                    };
                    mockOutlook.EmailStorage[$"{doc.StoreId}:{doc.EntryId}"] = doc;
                }

                var swQueue = Stopwatch.StartNew();
                for (int i = 1; i <= burst; i++)
                {
                    await indexer.QueueItemForIndexingAsync("Store1", $"Inc_Entry_{burst}_{i}", "Folder1");
                }
                swQueue.Stop();

                var swSearch = Stopwatch.StartNew();
                var searchResult = await repo.SearchAsync(new SearchQuery { RawQuery = "Invoice Notification", Limit = 50 });
                swSearch.Stop();

                double throughput = burst / (swQueue.Elapsed.TotalMilliseconds / 1000.0);
                double ramMb = Process.GetCurrentProcess().WorkingSet64 / (1024.0 * 1024.0);

                Console.WriteLine($"  [{burst,4} Items] Enqueue/Index Time: {swQueue.Elapsed.TotalMilliseconds,7:F2}ms | Throughput: {throughput,7:F1} items/s | Search Latency: {swSearch.Elapsed.TotalMilliseconds,6:F2}ms | RAM: {ramMb,5:F1} MB");
                reportLines.Add($"| {burst,4} | {swQueue.Elapsed.TotalMilliseconds,7:F2} | {swQueue.Elapsed.TotalMilliseconds,7:F2} | {throughput,7:F1} | {swSearch.Elapsed.TotalMilliseconds,6:F2} | {ramMb,5:F1} |");
            }

            Console.WriteLine("\n=================================================");
            Console.WriteLine(" INCREMENTAL BENCHMARK REPORT SUMMARY");
            Console.WriteLine("=================================================");
            foreach (var line in reportLines)
            {
                Console.WriteLine(line);
            }

            string reportPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "IncrementalBenchmarkResult.txt");
            File.WriteAllLines(reportPath, reportLines);
            Console.WriteLine($"\nReport saved to: {reportPath}");

            try { File.Delete(benchDbPath); } catch { }
        }

        private class MockOutlookAdapter : IOutlookAdapter
        {
            public Dictionary<string, EmailDocument> EmailStorage = new Dictionary<string, EmailDocument>();
            public bool IsOutlookAvailable() => true;
            public string GetOutlookVersion() => "16.0.15000.20000";
            public string GetOfficeBitness() => "64-bit";
            public List<MailboxInfo> GetMailboxes() => new List<MailboxInfo> { new MailboxInfo { StoreId = "Store1", DisplayName = "Test" } };
            public List<MailboxFolder> GetFolders(string storeId) => new List<MailboxFolder> { new MailboxFolder { StoreId = storeId, EntryId = "Folder1", Name = "Inbox" } };
            public List<EmailDocument> FetchFolderEmailsBatch(string storeId, string folderEntryId, int startIndex, int batchSize, Func<bool>? cancelCheck = null) => EmailStorage.Values.Skip(startIndex - 1).Take(batchSize).ToList();
            public EmailDocument? FetchSingleEmail(string storeId, string entryId) => EmailStorage.TryGetValue($"{storeId}:{entryId}", out var doc) ? doc : null;
            public bool OpenEmailInOutlook(string storeId, string entryId) => true;
            public bool IsOutboxFolder(string storeId, string folderEntryId) => false;
            public void SubscribeToEvents(Action<string, string, string> onItemAddedOrChanged, Action<string, string> onItemRemoved) { }
            public void UnsubscribeFromEvents() { }
        }

        static async Task RunSyntheticBenchmarkAsync()
        {
            Console.WriteLine("\n=================================================");
            Console.WriteLine(" FastMail Search - Synthetic Performance Benchmark");
            Console.WriteLine("=================================================");

            string benchDbPath = Path.Combine(Path.GetTempPath(), "FastMailSearch_Benchmark.db");
            if (File.Exists(benchDbPath))
            {
                try { File.Delete(benchDbPath); } catch { }
            }

            var factory = new DatabaseConnectionFactory(benchDbPath);
            var repo = new SqliteEmailRepository(factory);
            repo.InitializeDatabase();

            int[] targets = new[] { 10000, 50000, 100000 };
            int currentCount = 0;

            string[] testQueries = new[]
            {
                "kfintech",
                "\"project meeting\"",
                "from:john@example.com",
                "subject:invoice",
                "after:2026-01-01 before:2026-08-15",
                "from:john@example.com subject:invoice after:2026-01-01",
                "has:attachment",
                "invoice OR quotation -spam"
            };

            var summaryReport = new List<string>();
            summaryReport.Add("| Record Count | Query Type | Total Matches | Avg Latency (ms) | P95 Latency (ms) | P99 Latency (ms) | DB Size (MB) |");
            summaryReport.Add("|--------------|------------|---------------|------------------|------------------|------------------|--------------|");

            foreach (var target in targets)
            {
                int needed = target - currentCount;
                Console.WriteLine($"\n--> Generating {needed:N0} synthetic emails (Total: {target:N0})...");
                var swGen = Stopwatch.StartNew();
                await SyntheticDataGenerator.GenerateSyntheticEmailsAsync(repo, needed, p =>
                {
                    if (p % 10000 == 0) Console.Write($" Indexed {p:N0}...");
                });
                swGen.Stop();
                currentCount = target;
                Console.WriteLine($" Done in {swGen.Elapsed.TotalSeconds:F2}s");

                long dbSizeBytes = new FileInfo(benchDbPath).Length;
                double dbSizeMb = dbSizeBytes / (1024.0 * 1024.0);

                Console.WriteLine($"Running query benchmarks on {target:N0} records (DB Size: {dbSizeMb:F2} MB)...");

                foreach (var qStr in testQueries)
                {
                    var parsedQuery = SearchQueryParser.Parse(qStr);
                    var latencies = new List<double>();
                    int totalMatches = 0;

                    // Warmup run
                    var warmupResult = await repo.SearchAsync(parsedQuery);

                    // Validate & Benchmark Multi-Page Offsets (Page 1, 10, 100, Final)
                    parsedQuery.PageSize = 50;
                    int totalPages = warmupResult.TotalCount > 0 ? (int)Math.Ceiling((double)warmupResult.TotalCount / 50.0) : 1;
                    int[] testPages = new int[] { 1, 10, 100, totalPages }.Distinct().Where(p => p <= totalPages).ToArray();

                    foreach (var pageNum in testPages)
                    {
                        parsedQuery.Page = pageNum;
                        var swPage = Stopwatch.StartNew();
                        var pageRes = await repo.SearchAsync(parsedQuery);
                        swPage.Stop();
                        Console.WriteLine($"     [Page {pageNum,4}/{totalPages,4}] Latency: {swPage.Elapsed.TotalMilliseconds,6:F2}ms | Returned: {pageRes.Items.Count,2} items");
                    }

                    // Reset to Page 1 for 25 Benchmark iterations
                    parsedQuery.Page = 1;
                    for (int iter = 0; iter < 25; iter++)
                    {
                        var swQuery = Stopwatch.StartNew();
                        var res = await repo.SearchAsync(parsedQuery);
                        swQuery.Stop();

                        latencies.Add(swQuery.Elapsed.TotalMilliseconds);
                        totalMatches = res.TotalCount;
                    }

                    latencies.Sort();
                    double avg = latencies.Average();
                    double p95 = latencies[(int)(latencies.Count * 0.95)];
                    double p99 = latencies[(int)(latencies.Count * 0.99)];

                    Console.WriteLine($"  [{target:N0}] '{qStr}' -> Matches: {totalMatches:N0} | Avg: {avg:F2}ms | P95: {p95:F2}ms | P99: {p99:F2}ms");

                    summaryReport.Add($"| {target:N0} | `{qStr}` | {totalMatches:N0} | {avg:F2} | {p95:F2} | {p99:F2} | {dbSizeMb:F2} |");
                }
            }

            Console.WriteLine("\n=================================================");
            Console.WriteLine(" BENCHMARK SUMMARY REPORT");
            Console.WriteLine("=================================================");
            foreach (var line in summaryReport)
            {
                Console.WriteLine(line);
            }

            string reportPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "BenchmarkResult.txt");
            File.WriteAllLines(reportPath, summaryReport);
            Console.WriteLine($"\nReport saved to: {reportPath}");

            try { File.Delete(benchDbPath); } catch { }
        }
    }
}
