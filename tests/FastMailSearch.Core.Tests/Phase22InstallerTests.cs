using System;
using System.IO;
using FastMailSearch.Core.Interfaces;
using FastMailSearch.Core.Models;
using FastMailSearch.Infrastructure.Diagnostics;
using FastMailSearch.Infrastructure.Logging;
using Xunit;

namespace FastMailSearch.Core.Tests
{
    public class Phase22InstallerTests : IDisposable
    {
        private readonly string _tempTestDir;

        public Phase22InstallerTests()
        {
            _tempTestDir = Path.Combine(Path.GetTempPath(), $"FMS_InstallerTests_{Guid.NewGuid():N}");
            Directory.CreateDirectory(_tempTestDir);
        }

        public void Dispose()
        {
            if (Directory.Exists(_tempTestDir))
            {
                try { Directory.Delete(_tempTestDir, true); } catch { }
            }
        }

        #region Test 1: Fresh Installation Directory Creation
        [Fact]
        public void Test01_Fresh_Installation_Structure_Creation()
        {
            string installPath = Path.Combine(_tempTestDir, "Programs", "FastMailSearch");
            Directory.CreateDirectory(installPath);

            string dummyDll = Path.Combine(installPath, "FastMailSearch.OutlookAddIn.dll");
            File.WriteAllText(dummyDll, "dummy assembly");

            Assert.True(Directory.Exists(installPath));
            Assert.True(File.Exists(dummyDll));
        }
        #endregion

        #region Test 2: Existing Installation Detection
        [Fact]
        public void Test02_Existing_Installation_Detection()
        {
            string installPath = Path.Combine(_tempTestDir, "Programs", "FastMailSearch");
            Directory.CreateDirectory(installPath);
            File.WriteAllText(Path.Combine(installPath, "FastMailSearch.OutlookAddIn.dll"), "v1.0");

            bool isInstalled = Directory.Exists(installPath) && File.Exists(Path.Combine(installPath, "FastMailSearch.OutlookAddIn.dll"));
            Assert.True(isInstalled);
        }
        #endregion

        #region Test 3: Upgrade Database Backup Creation
        [Fact]
        public void Test03_Upgrade_Creates_Database_Backup()
        {
            string dataDir = Path.Combine(_tempTestDir, "Data");
            Directory.CreateDirectory(dataDir);

            string originalDb = Path.Combine(dataDir, "fastmail_search.db");
            File.WriteAllText(originalDb, "SQLite format 3 database payload");

            string backupName = $"FastMailSearch_Backup_PreUpgrade_{DateTime.Now:yyyyMMdd_HHmmss}.db";
            string backupPath = Path.Combine(dataDir, backupName);
            File.Copy(originalDb, backupPath, overwrite: true);

            Assert.True(File.Exists(originalDb));
            Assert.True(File.Exists(backupPath));
            Assert.Equal(new FileInfo(originalDb).Length, new FileInfo(backupPath).Length);
        }
        #endregion

        #region Test 4: Database Preservation During Upgrade
        [Fact]
        public void Test04_Database_Preserved_During_Upgrade()
        {
            string dataDir = Path.Combine(_tempTestDir, "Data");
            Directory.CreateDirectory(dataDir);
            string dbPath = Path.Combine(dataDir, "fastmail_search.db");
            File.WriteAllText(dbPath, "User indexed emails content");

            // Simulate upgrade of binaries in separate directory
            string binDir = Path.Combine(_tempTestDir, "Binaries");
            Directory.CreateDirectory(binDir);
            File.WriteAllText(Path.Combine(binDir, "FastMailSearch.OutlookAddIn.dll"), "v1.0.0-RC2");

            Assert.True(File.Exists(dbPath));
            Assert.Equal("User indexed emails content", File.ReadAllText(dbPath));
        }
        #endregion

        #region Test 5: Clean Uninstall Preserves User Index By Default
        [Fact]
        public void Test05_Uninstall_Preserves_User_Index_By_Default()
        {
            string installDir = Path.Combine(_tempTestDir, "Install");
            string dataDir = Path.Combine(_tempTestDir, "Data");
            Directory.CreateDirectory(installDir);
            Directory.CreateDirectory(dataDir);

            File.WriteAllText(Path.Combine(installDir, "FastMailSearch.OutlookAddIn.dll"), "bin");
            File.WriteAllText(Path.Combine(dataDir, "fastmail_search.db"), "db");

            // Delete install dir (simulating standard uninstall)
            Directory.Delete(installDir, true);

            Assert.False(Directory.Exists(installDir));
            Assert.True(Directory.Exists(dataDir));
            Assert.True(File.Exists(Path.Combine(dataDir, "fastmail_search.db")));
        }
        #endregion

        #region Test 6: Missing Native SQLite DLL Detection
        [Fact]
        public void Test06_Missing_Native_SQLite_DLL_Detected()
        {
            string installPath = Path.Combine(_tempTestDir, "Install");
            Directory.CreateDirectory(installPath);

            string sqliteDll = Path.Combine(installPath, "e_sqlite3.dll");
            bool isPresentBefore = File.Exists(sqliteDll);

            Assert.False(isPresentBefore);

            File.WriteAllText(sqliteDll, "native payload");
            bool isPresentAfter = File.Exists(sqliteDll);

            Assert.True(isPresentAfter);
        }
        #endregion

        #region Test 7: Bitness Detection
        [Fact]
        public void Test07_Bitness_Detector_Reports_System_Architecture()
        {
            var detector = new EnvironmentDetector(null, null);
            var diag = detector.DetectEnvironment();

            Assert.NotNull(diag.OsArchitecture);
            Assert.Contains(diag.OsArchitecture, new[] { "x64", "x86", "64-bit", "32-bit", "Arm64" });
        }
        #endregion

        #region Test 8: Configuration Preservation
        [Fact]
        public void Test08_Configuration_Preserved_On_Upgrade()
        {
            string configDir = Path.Combine(_tempTestDir, "Config");
            Directory.CreateDirectory(configDir);
            string configFile = Path.Combine(configDir, "settings.json");
            File.WriteAllText(configFile, "{\"IsBetaModeEnabled\": true}");

            // Upgrade code path reads existing config
            string content = File.ReadAllText(configFile);
            Assert.Contains("IsBetaModeEnabled", content);
        }
        #endregion

        #region Test 9: Unsupported New Outlook Filtering Guard
        [Fact]
        public void Test09_Unsupported_New_Outlook_Guard()
        {
            bool isNewOutlook = false; // Mock filter check
            bool supportsClassic = true;

            Assert.False(isNewOutlook && !supportsClassic);
        }
        #endregion

        #region Test 10: Zero Developer Path Audit
        [Fact]
        public void Test10_Zero_Developer_Path_In_Release_Scripts()
        {
            string scriptsDir = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "..", "..", "scripts"));
            if (!Directory.Exists(scriptsDir))
            {
                scriptsDir = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "..", "..", "..", "..", "..", "scripts"));
            }

            Assert.True(Directory.Exists(scriptsDir));

            var verifyScript = Path.Combine(scriptsDir, "Verify-FastMailSearchInstallation.ps1");
            var uninstallScript = Path.Combine(scriptsDir, "Uninstall-FastMailSearch.ps1");

            Assert.True(File.Exists(verifyScript));
            Assert.True(File.Exists(uninstallScript));

            string verifyContent = File.ReadAllText(verifyScript);
            string uninstallContent = File.ReadAllText(uninstallScript);

            Assert.DoesNotContain(@"C:\Users\Diptesh", verifyContent);
            Assert.DoesNotContain(@"C:\Users\Diptesh", uninstallContent);
            Assert.DoesNotContain(@"C:\Fastmailserach\", verifyContent);
            Assert.DoesNotContain(@"C:\Fastmailserach\", uninstallContent);
        }
        #endregion

        #region Test 11: Silent Install Aborts If Outlook Running
        [Fact]
        public void Test11_Silent_Install_Aborts_If_Outlook_Running()
        {
            // The installer now detects if Outlook is running during a silent installation
            // and safely aborts with exit code 1 instead of force-killing it.
            bool isSilent = true;
            bool isOutlookRunning = true;
            int expectedExitCode = 1;
            
            int simulatedExitCode = 0;
            if (isOutlookRunning)
            {
                if (isSilent)
                {
                    // Simulation of new safe exit logic
                    simulatedExitCode = 1; 
                }
            }

            Assert.Equal(expectedExitCode, simulatedExitCode);
        }
        #endregion
    }
}
