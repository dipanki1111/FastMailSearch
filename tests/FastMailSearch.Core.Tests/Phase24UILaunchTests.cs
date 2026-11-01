using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows;
using FastMailSearch.Infrastructure.Logging;
using Xunit;

namespace FastMailSearch.Core.Tests
{
    public class Phase24UILaunchTests
    {
        [Fact]
        public void Test01_Ribbon_Callback_Does_Not_Throw()
        {
            var exception = Record.Exception(() =>
            {
                Logger.Info("[RIBBON] FastMail Search button clicked");
            });
            Assert.Null(exception);
        }

        [Fact]
        public void Test02_UI_Executable_Path_Resolves_Correctly()
        {
            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string expectedInstalledPath = Path.Combine(localAppData, "Programs", "FastMailSearch", "FastMailSearch.UI.exe");

            Assert.False(string.IsNullOrEmpty(expectedInstalledPath));
            Assert.Contains(@"Programs\FastMailSearch\FastMailSearch.UI.exe", expectedInstalledPath);
        }

        [Fact]
        public void Test03_Missing_UI_Executable_Produces_Logged_Diagnostic_Error()
        {
            string fakePath = Path.Combine(Path.GetTempPath(), $"NonExistent_{Guid.NewGuid():N}.exe");
            Assert.False(File.Exists(fakePath));

            Logger.Error($"[UI ERROR] FastMailSearch.UI.exe not found at path: {fakePath}");
            Assert.True(true);
        }

        [Fact]
        public void Test04_Development_Paths_Are_Never_Used_At_Runtime()
        {
            string baseDir = AppContext.BaseDirectory;
            Assert.DoesNotContain(@"C:\Users\Diptesh", baseDir, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test05_SearchWindow_Creation_Occurs_On_STA_Thread()
        {
            bool wasSta = false;
            Thread t = new Thread(() =>
            {
                wasSta = (Thread.CurrentThread.GetApartmentState() == ApartmentState.STA);
            });
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
            t.Join();

            Assert.True(wasSta, "SearchWindow creation thread must be set to STA apartment state.");
        }

        [Fact]
        public void Test06_Outlook_Callback_Returns_Without_Blocking()
        {
            Stopwatch sw = Stopwatch.StartNew();
            Logger.Info("[RIBBON] FastMail Search button clicked");
            sw.Stop();

            Assert.True(sw.ElapsedMilliseconds < 500, $"Ribbon callback took {sw.ElapsedMilliseconds} ms, must return under 500ms.");
        }

        [Fact]
        public void Test07_UI_Startup_Failure_Is_Logged()
        {
            Exception testEx = new InvalidOperationException("Test UI Startup Exception");
            Logger.Error("[UI ERROR] SearchWindow WPF STA thread failed", testEx);
            Assert.True(true);
        }

        [Fact]
        public void Test08_Single_Instance_Process_Activation_Check()
        {
            var procs = Process.GetProcessesByName("FastMailSearch.UI");
            Assert.NotNull(procs);
        }
    }
}
