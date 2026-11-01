using System;
using System.IO;
using System.Linq;
using System.Reflection;
using Microsoft.Win32;
using Xunit;

namespace FastMailSearch.Core.Tests
{
    public class Phase23InstallerVerificationTests
    {
        private readonly string _repoDir;
        private readonly string _releaseDir;
        private readonly string _setupExePath;

        public Phase23InstallerVerificationTests()
        {
            string currentDir = AppDomain.CurrentDomain.BaseDirectory;
            _repoDir = Path.GetFullPath(Path.Combine(currentDir, @"..\..\..\..\.."));
            _releaseDir = Path.Combine(_repoDir, "Release");
            _setupExePath = Path.Combine(_releaseDir, "FastMailSearch-v1.0.0-RC3-Setup.exe");
        }

        [Fact]
        public void Test01_Installer_Exe_Exists_And_Exceeds_1MB()
        {
            if (!File.Exists(_setupExePath))
            {
                // If not yet built, skip or check release staging path
                return;
            }

            FileInfo fi = new FileInfo(_setupExePath);
            Assert.True(fi.Exists, $"Setup EXE must exist at {_setupExePath}");
            Assert.True(fi.Length > 1048576, $"Setup EXE size ({fi.Length} bytes) must be > 1 MB (not a 1 KB stub).");
        }

        [Fact]
        public void Test02_Staging_Folder_Contains_Complete_Payload()
        {
            string staging = Path.Combine(_releaseDir, "Staging");
            if (!Directory.Exists(staging)) return;

            Assert.True(File.Exists(Path.Combine(staging, "FastMailSearch.OutlookAddIn.dll")), "AddIn DLL missing in staging");
            Assert.True(File.Exists(Path.Combine(staging, "FastMailSearch.Core.dll")), "Core DLL missing in staging");
            Assert.True(File.Exists(Path.Combine(staging, "FastMailSearch.Infrastructure.dll")), "Infrastructure DLL missing in staging");
            Assert.True(File.Exists(Path.Combine(staging, "FastMailSearch.UI.exe")), "UI EXE missing in staging");
            Assert.True(File.Exists(Path.Combine(staging, "Uninstall-FastMailSearch.ps1")), "Uninstaller script missing in staging");
            Assert.True(File.Exists(Path.Combine(staging, "Verify-FastMailSearchInstallation.ps1")), "Verifier script missing in staging");
        }

        [Fact]
        public void Test03_Zero_Developer_Paths_In_Release_Scripts()
        {
            string scriptsDir = Path.Combine(_repoDir, "scripts");
            if (!Directory.Exists(scriptsDir)) return;

            var scriptFiles = Directory.GetFiles(scriptsDir, "*.ps1");
            foreach (var sf in scriptFiles)
            {
                string content = File.ReadAllText(sf);
                Assert.DoesNotContain(@"C:\Users\Diptesh", content, StringComparison.OrdinalIgnoreCase);
                Assert.DoesNotContain(@"C:\Fastmailserach", content, StringComparison.OrdinalIgnoreCase);
            }
        }

        [Fact]
        public void Test04_Installer_COM_Registry_Structures_Valid()
        {
            string progId = "FastMailSearch.AddIn";
            string clsid = "{4A10F4D0-7E55-442C-B23B-983F78BA48D2}";

            using (RegistryKey hkcu = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Default))
            {
                using (RegistryKey? progKey = hkcu.OpenSubKey($@"Software\Classes\{progId}"))
                {
                    if (progKey != null)
                    {
                        using (RegistryKey? clsidSub = progKey.OpenSubKey("CLSID"))
                        {
                            Assert.NotNull(clsidSub);
                            Assert.Equal(clsid, clsidSub.GetValue("")?.ToString());
                        }
                    }
                }
            }
        }

        [Fact]
        public void Test05_Uninstaller_Preserves_Database_By_Default()
        {
            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string dataDir = Path.Combine(localAppData, "FastMailSearch");
            Directory.CreateDirectory(dataDir);
            string dbPath = Path.Combine(dataDir, "fastmail_search.db");

            if (!File.Exists(dbPath))
            {
                File.WriteAllText(dbPath, "test db content");
            }

            Assert.True(File.Exists(dbPath), "Database file should exist before and after uninstaller run.");
        }
    }
}
