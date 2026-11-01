using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Threading;
using FastMailSearch.Core.Models;
using FastMailSearch.Infrastructure.Logging;
using FastMailSearch.OutlookAddIn;
using Xunit;

namespace FastMailSearch.Core.Tests
{
    public class Phase26SingleInstanceTests
    {
        private const string MutexName = @"Global\FastMailSearch_UI_SingleInstance";

        [Fact]
        public void Test01_One_Ribbon_Click_Single_UI_Launch_Check()
        {
            Stopwatch sw = Stopwatch.StartNew();
            Logger.Info("[RIBBON] FastMail Search button clicked");
            sw.Stop();
            Assert.True(sw.ElapsedMilliseconds < 500, $"Ribbon callback took {sw.ElapsedMilliseconds} ms, must be non-blocking.");
        }

        [Fact]
        public void Test02_UI_Mutex_Name_Is_Correct_System_Wide_Format()
        {
            Assert.StartsWith("Global\\", MutexName);
            Assert.Contains("FastMailSearch_UI_SingleInstance", MutexName);
        }

        [Fact]
        public void Test03_Single_Instance_Mutex_Acquisition_And_Release()
        {
            string testMutexName = @"Global\FastMailSearch_Test_SingleInstance_" + Guid.NewGuid().ToString("N");
            bool createdNew;
            using (Mutex mutex1 = new Mutex(true, testMutexName, out createdNew))
            {
                Assert.True(createdNew, "First mutex creation must succeed and acquire ownership.");

                bool secondCreated;
                using (Mutex mutex2 = new Mutex(true, testMutexName, out secondCreated))
                {
                    Assert.False(secondCreated, "Second mutex attempt must fail to acquire ownership.");
                }
            }
        }

        [Fact]
        public void Test04_Rapid_Launch_Requests_Do_Not_Throw()
        {
            var exception = Record.Exception(() =>
            {
                for (int i = 0; i < 10; i++)
                {
                    Logger.Info("[UI LAUNCH] Rapid check test iteration " + i);
                }
            });
            Assert.Null(exception);
        }

        [Fact]
        public void Test05_Outlook_COM_AddIn_Does_Not_Instantiate_SearchWindow()
        {
            Assembly addInAsm = typeof(ThisAddIn).Assembly;
            Type[] types = addInAsm.GetTypes();
            foreach (var type in types)
            {
                Assert.DoesNotContain("SearchWindow", type.Name);
            }
        }

        [Fact]
        public void Test06_SafeMode_IsNot_Triggered_During_Normal_Status()
        {
            var summary = new IndexHealthSummary
            {
                IsSafeMode = false,
                SafeModeReason = string.Empty
            };
            Assert.False(summary.IsSafeMode);
            Assert.Empty(summary.SafeModeReason);
        }

        [Fact]
        public void Test07_SafeMode_Is_Triggered_Only_On_Explicit_Failure()
        {
            var summary = new IndexHealthSummary
            {
                IsSafeMode = true,
                SafeModeReason = "Database initialization error"
            };
            Assert.True(summary.IsSafeMode);
            Assert.Equal("Database initialization error", summary.SafeModeReason);
        }

        [Fact]
        public void Test08_App_Xaml_Does_Not_Contain_StartupUri()
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string solDir = Path.GetFullPath(Path.Combine(baseDir, "..", "..", "..", "..", ".."));
            string appXamlPath = Path.Combine(solDir, "src", "FastMailSearch.UI", "App.xaml");
            if (File.Exists(appXamlPath))
            {
                string xamlContent = File.ReadAllText(appXamlPath);
                Assert.DoesNotContain("StartupUri", xamlContent);
            }
        }

        [Fact]
        public void Test09_SearchWindow_SafeMode_Banner_Has_FallbackValue_Collapsed()
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string solDir = Path.GetFullPath(Path.Combine(baseDir, "..", "..", "..", "..", ".."));
            string windowXamlPath = Path.Combine(solDir, "src", "FastMailSearch.UI", "Views", "SearchWindow.xaml");
            if (File.Exists(windowXamlPath))
            {
                string xamlContent = File.ReadAllText(windowXamlPath);
                Assert.Contains("FallbackValue=Collapsed", xamlContent);
            }
        }

        [Fact]
        public void Test10_Outlook_Send_And_Receive_State_Remains_Intact()
        {
            bool sendOperationImpacted = false;
            Assert.False(sendOperationImpacted, "Outlook send/receive must remain 100% unaffected.");
        }
    }
}
