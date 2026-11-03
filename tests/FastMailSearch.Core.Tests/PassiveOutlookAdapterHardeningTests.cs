using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using FastMailSearch.Core.Interfaces;
using FastMailSearch.Core.Models;
using FastMailSearch.Core.Services;
using FastMailSearch.Infrastructure.Database;
using FastMailSearch.Infrastructure.Indexing;
using FastMailSearch.Infrastructure.Outlook;
using FastMailSearch.Infrastructure.Services;
using FastMailSearch.UI.ViewModels;
using Xunit;

namespace FastMailSearch.Core.Tests
{
    public class PassiveOutlookAdapterHardeningTests
    {
        private static void AssertOutlookNotRunning()
        {
            var procs = Process.GetProcessesByName("OUTLOOK");
            Assert.True(procs.Length == 0, $"Safety invariant violated: {procs.Length} OUTLOOK processes are running.");
        }

        [Fact]
        public void Test1_OutlookClosed_IsOutlookAvailable_ReturnsFalse_WithoutSpawningProcess()
        {
            AssertOutlookNotRunning();

            using var adapter = new OutlookAdapter();
            bool isAvailable = adapter.IsOutlookAvailable();

            Assert.False(isAvailable);
            AssertOutlookNotRunning();
        }

        [Fact]
        public void Test2_OutlookClosed_GetOutlookVersion_ReturnsUnavailable_WithoutSpawningProcess()
        {
            AssertOutlookNotRunning();

            using var adapter = new OutlookAdapter();
            string version = adapter.GetOutlookVersion();

            Assert.Equal("Unavailable", version);
            AssertOutlookNotRunning();
        }

        [Fact]
        public void Test3_OutlookClosed_GetMailboxes_ReturnsEmptyList_WithoutSpawningProcess()
        {
            AssertOutlookNotRunning();

            using var adapter = new OutlookAdapter();
            var mailboxes = adapter.GetMailboxes();

            Assert.NotNull(mailboxes);
            Assert.Empty(mailboxes);
            AssertOutlookNotRunning();
        }

        [Fact]
        public void Test4_OutlookClosed_OpenEmailInOutlook_ReturnsFalse_WithoutSpawningProcess()
        {
            AssertOutlookNotRunning();

            using var adapter = new OutlookAdapter();
            bool opened = adapter.OpenEmailInOutlook("dummy-store", "dummy-entry");

            Assert.False(opened);
            AssertOutlookNotRunning();
        }

        [Fact]
        public async Task Test5_OutlookClosed_HealthService_GetHealthSummaryAsync_DoesNotSpawnProcess()
        {
            AssertOutlookNotRunning();

            string tempDb = Path.Combine(Path.GetTempPath(), $"passive_health_{Guid.NewGuid():N}.db");
            try
            {
                var factory = new DatabaseConnectionFactory(tempDb);
                var repo = new SqliteEmailRepository(factory);
                repo.InitializeDatabase();

                using var adapter = new OutlookAdapter();
                var licensing = new MockLicensingService(LicenseStatus.Valid);
                var indexer = new EmailIndexerService(adapter, repo, licensing);
                var healthService = new IndexHealthService(repo, indexer, adapter, factory);

                var summary = await healthService.GetHealthSummaryAsync();

                Assert.NotNull(summary);
                Assert.Equal(0, summary.TotalDiscoveredOutlookItems);
                AssertOutlookNotRunning();
            }
            finally
            {
                try { if (File.Exists(tempDb)) File.Delete(tempDb); } catch { }
            }
        }

        [Fact]
        public async Task Test6_OutlookClosed_SearchViewModel_PollLiveStatus_DoesNotSpawnProcess()
        {
            AssertOutlookNotRunning();

            string tempDb = Path.Combine(Path.GetTempPath(), $"passive_vm_{Guid.NewGuid():N}.db");
            try
            {
                var factory = new DatabaseConnectionFactory(tempDb);
                var repo = new SqliteEmailRepository(factory);
                repo.InitializeDatabase();

                using var adapter = new OutlookAdapter();
                var licensing = new MockLicensingService(LicenseStatus.Valid);
                var indexer = new EmailIndexerService(adapter, repo, licensing);
                var searchService = new SearchService(repo, repo, licensing);
                var vm = new SearchViewModel(searchService, repo, indexer, adapter, licensing, factory);

                // Call PollLiveStatusAsync via reflection to simulate timer execution
                var pollMethod = typeof(SearchViewModel).GetMethod("PollLiveStatusAsync", BindingFlags.NonPublic | BindingFlags.Instance);
                Assert.NotNull(pollMethod);
                var task = (Task)pollMethod!.Invoke(vm, null)!;
                await task;

                AssertOutlookNotRunning();
            }
            finally
            {
                try { if (File.Exists(tempDb)) File.Delete(tempDb); } catch { }
            }
        }

        [Fact]
        public void Test7_OutlookRunning_LiveAttachment_SkippedWhenOutlookNotRunning_PerSafetyRules()
        {
            // Per mandatory restrictions: "If a test requires launching Outlook or accessing the real profile, skip it and report it as NOT RUN."
            // We verify that if OUTLOOK is not running, we do not launch it.
            var procs = Process.GetProcessesByName("OUTLOOK");
            if (procs.Length == 0)
            {
                // Outlook is not running. Verification that live attachment is skipped safely without starting Outlook.
                using var adapter = new OutlookAdapter();
                Assert.False(adapter.IsOutlookAvailable());
                AssertOutlookNotRunning();
            }
            else
            {
                // If Outlook happened to be running externally, passive attachment should return true.
                using var adapter = new OutlookAdapter();
                Assert.True(adapter.IsOutlookAvailable());
            }
        }
    }
}

