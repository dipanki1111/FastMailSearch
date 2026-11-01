using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using FastMailSearch.Core.Models;
using FastMailSearch.Core.Search;
using FastMailSearch.Infrastructure.Database;
using FastMailSearch.Infrastructure.Indexing;
using FastMailSearch.Infrastructure.Logging;
using FastMailSearch.Infrastructure.Outlook;

namespace FastMailSearch.PerformanceTests
{
    public static class RealOutlookTestRunner
    {
        public static async Task RunAsync()
        {
            Console.WriteLine("=================================================");
            Console.WriteLine(" FastMail Search - REAL CLASSIC OUTLOOK INTEGRATION TEST");
            Console.WriteLine("=================================================");

            Logger.Initialize();
            var outlookAdapter = new OutlookAdapter();

            bool available = outlookAdapter.IsOutlookAvailable();
            Console.WriteLine($"Is Outlook Classic Available: {available}");
            if (!available)
            {
                Console.WriteLine("Classic Outlook is NOT available on this system.");
                return;
            }

            Console.WriteLine($"Outlook Version: {outlookAdapter.GetOutlookVersion()} ({outlookAdapter.GetOfficeBitness()})");

            string dbPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "FastMailSearch", "FastMailSearch.db");
            Console.WriteLine($"Database Path: {dbPath}");

            var dbFactory = new DatabaseConnectionFactory(dbPath);
            var repo = new SqliteEmailRepository(dbFactory);
            repo.InitializeDatabase();

            var indexer = new EmailIndexerService(outlookAdapter, repo);

            Console.WriteLine("\n[Step 1] Fetching Real Mailboxes & Folders...");
            var mailboxes = outlookAdapter.GetMailboxes();
            Console.WriteLine($"Found {mailboxes.Count} Mailbox(es):");
            foreach (var mb in mailboxes)
            {
                Console.WriteLine($" - {mb.DisplayName} (PST: {mb.IsPst})");
                var folders = outlookAdapter.GetFolders(mb.StoreId);
                Console.WriteLine($"   Folders ({folders.Count}):");
                foreach (var f in folders)
                {
                    Console.WriteLine($"     * {f.Name} ({f.ItemCount} items)");
                }
            }

            indexer.ProgressChanged += (sender, progress) =>
            {
                Console.WriteLine($" [Indexer] {progress.CurrentStatusMessage} (Indexed: {progress.ProcessedItems:N0}, Skipped: {progress.SkippedItems:N0})");
            };

            Console.WriteLine("\n[Step 2] Indexing Real Mailbox Emails...");
            var swIndex = Stopwatch.StartNew();
            await indexer.StartInitialIndexingAsync();
            swIndex.Stop();

            var meta = await repo.GetIndexMetadataAsync();
            Console.WriteLine($"\nIndexing Finished in {swIndex.Elapsed.TotalSeconds:F2} seconds.");
            Console.WriteLine($"Total Real Emails Indexed: {meta.TotalEmailsIndexed:N0}");

            Console.WriteLine("\n[Step 3] Executing Required Searches on Real Indexed Emails...");

            string[] testQueries = new[]
            {
                "invoice",
                "from:abc.com",
                "subject:invoice",
                "hasattachment:true"
            };

            foreach (var qStr in testQueries)
            {
                var query = SearchQueryParser.Parse(qStr);
                var swSearch = Stopwatch.StartNew();
                var result = await repo.SearchAsync(query);
                swSearch.Stop();

                Console.WriteLine($" -> Query: '{qStr}'");
                Console.WriteLine($"    Matches: {result.TotalCount:N0}");
                Console.WriteLine($"    Search Time: {swSearch.Elapsed.TotalMilliseconds:F3} ms ({result.ElapsedSeconds:F3} sec)");

                if (result.Items.Count > 0)
                {
                    var top = result.Items[0];
                    Console.WriteLine($"    Top Result: Subject='{top.Subject}', Sender='{top.DisplaySender}', Folder='{top.FolderName}', Date='{top.DisplayDate}'");
                }
                else
                {
                    Console.WriteLine("    Top Result: (No matching items in real mailbox)");
                }
            }

            Console.WriteLine("\n[Step 4] Testing COM OpenEmailInOutlook for Top Indexed Item...");
            if (meta.TotalEmailsIndexed > 0)
            {
                var sampleResult = await repo.SearchAsync(new SearchQuery { Limit = 1 });
                if (sampleResult.Items.Count > 0)
                {
                    var item = sampleResult.Items[0];
                    Console.WriteLine($"Resolving Outlook COM Item: StoreId='{item.StoreId}', EntryId='{item.EntryId}'...");
                    bool opened = outlookAdapter.OpenEmailInOutlook(item.StoreId, item.EntryId);
                    Console.WriteLine($"Outlook COM Display Result: {opened}");
                }
            }

            Console.WriteLine("\n=================================================");
            Console.WriteLine(" Real Outlook Integration Run Complete!");
            Console.WriteLine("=================================================");
        }
    }
}
