using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

namespace FastMailSearch.Setup
{
    public static class SetupInstaller
    {
        public static int RunSetup(string[] args)
        {
            bool isSilent = args.Any(a => string.Equals(a, "/quiet", StringComparison.OrdinalIgnoreCase) ||
                                          string.Equals(a, "-quiet", StringComparison.OrdinalIgnoreCase) ||
                                          string.Equals(a, "--quiet", StringComparison.OrdinalIgnoreCase) ||
                                          string.Equals(a, "/silent", StringComparison.OrdinalIgnoreCase) ||
                                          string.Equals(a, "-silent", StringComparison.OrdinalIgnoreCase) ||
                                          string.Equals(a, "--silent", StringComparison.OrdinalIgnoreCase) ||
                                          string.Equals(a, "-Force", StringComparison.OrdinalIgnoreCase) ||
                                          string.Equals(a, "/q", StringComparison.OrdinalIgnoreCase) ||
                                          string.Equals(a, "-q", StringComparison.OrdinalIgnoreCase) ||
                                          string.Equals(a, "/s", StringComparison.OrdinalIgnoreCase) ||
                                          string.Equals(a, "-s", StringComparison.OrdinalIgnoreCase) ||
                                          string.Equals(a, "/S", StringComparison.OrdinalIgnoreCase));

            var version = typeof(SetupInstaller).Assembly.GetName().Version;
            Console.WriteLine("=================================================");
            Console.WriteLine($" FastMail Search Outlook Add-In Setup (v{version?.Major ?? 1}.{version?.Minor ?? 0}.{version?.Build ?? 0})");
            Console.WriteLine("=================================================");

            try
            {
                // 1. Detect environment & bitness
                string osArch = Environment.Is64BitOperatingSystem ? "x64" : "x86";
                Console.WriteLine($"System Architecture: {osArch}");

                // 2. Check running Classic Outlook process
                var outlookProcs = Process.GetProcessesByName("outlook");
                if (outlookProcs.Length > 0)
                {
                    if (isSilent)
                    {
                        Console.WriteLine("[ERROR] Classic Outlook is currently running.");
                        Console.WriteLine("Installation cancelled: Outlook must be closed prior to executing the silent installer to prevent data loss.");
                        return 1;
                    }
                    else
                    {
                        MessageBox.Show(
                            "Classic Outlook is currently running.\n\nPlease save your work, close Outlook completely, and then run this installer again.",
                            "FastMail Search Setup - Action Required",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);

                        Console.WriteLine("Installation cancelled: Outlook must be closed.");
                        return 1;
                    }
                }

                // 3. Target Installation Directory (%LOCALAPPDATA%\Programs\FastMailSearch)
                string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                string targetDir = Path.Combine(localAppData, "Programs", "FastMailSearch");

                // Reparse-point guard: Verify targetDir is not an existing junction/symlink before reuse (query directly, fail closed)
                var targetStatus = GetReparseStatus(targetDir, out string? targetErr);
                if (targetStatus == ReparseCheckResult.IsReparsePoint)
                {
                    Console.WriteLine($"[SECURITY ERROR] Installation directory '{targetDir}' is a junction or reparse point. Aborting setup to prevent arbitrary file redirection.");
                    return 1;
                }
                if (targetStatus == ReparseCheckResult.QueryFailed)
                {
                    Console.WriteLine($"[SECURITY ERROR] Could not verify attributes of installation directory '{targetDir}': {targetErr}. Aborting setup fail-closed.");
                    return 1;
                }

                // If targetDir does not exist, verify relevant parent directory (%LOCALAPPDATA%\Programs) before creating targetDir
                if (targetStatus == ReparseCheckResult.DoesNotExist)
                {
                    string? programsDir = Path.GetDirectoryName(targetDir);
                    if (!string.IsNullOrEmpty(programsDir))
                    {
                        var parentStatus = GetReparseStatus(programsDir, out string? parentErr);
                        if (parentStatus == ReparseCheckResult.IsReparsePoint)
                        {
                            // Note on compatibility: If an environment intentionally redirects %LOCALAPPDATA%\Programs via junction,
                            // setup fails closed here to prevent pre-planted redirection attacks.
                            Console.WriteLine($"[SECURITY ERROR] Target parent directory '{programsDir}' is a junction or reparse point. Aborting setup to prevent arbitrary file redirection.");
                            return 1;
                        }
                        if (parentStatus == ReparseCheckResult.QueryFailed)
                        {
                            Console.WriteLine($"[SECURITY ERROR] Could not verify attributes of parent directory '{programsDir}': {parentErr}. Aborting setup fail-closed.");
                            return 1;
                        }
                    }
                    Directory.CreateDirectory(targetDir);
                }

                // 4. Safe Upgrade Database Backup (canonical Documents\FastMailSearch\FastMailSearch.db + legacy paths)
                Console.WriteLine("Scanning for existing databases to back up prior to upgrade...");
                var backupResult = InstallerDatabaseBackup.BackupAllExistingDatabases(
                    msg => Console.WriteLine(msg));

                if (!backupResult.Success)
                {
                    Console.WriteLine($"[SECURITY / BACKUP ERROR] Pre-upgrade database backup failed: {backupResult.ErrorMessage}");
                    if (isSilent)
                    {
                        Console.WriteLine("Installation cancelled: Database backup failure. Setup aborted fail-closed.");
                    }
                    else
                    {
                        MessageBox.Show(
                            $"Failed to create a pre-upgrade backup of your database:\n\n{backupResult.ErrorMessage}\n\nInstallation has been safely aborted to protect your data.",
                            "FastMail Search Setup - Backup Failure",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                    }
                    return 1;
                }

                // 5. Extract Embedded Payload.zip to Target Directory
                Console.WriteLine($"Extracting release payload to {targetDir}...");
                var asm = Assembly.GetExecutingAssembly();
                string resourceName = asm.GetManifestResourceNames().FirstOrDefault(n => n.EndsWith("Payload.zip", StringComparison.OrdinalIgnoreCase)) ?? string.Empty;

                if (!string.IsNullOrEmpty(resourceName))
                {
                    using (Stream? resStream = asm.GetManifestResourceStream(resourceName))
                    {
                        if (resStream != null)
                        {
                            string canonicalTargetDir = Path.GetFullPath(targetDir);
                            if (!canonicalTargetDir.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal))
                            {
                                canonicalTargetDir += Path.DirectorySeparatorChar;
                            }

                            using (ZipArchive archive = new ZipArchive(resStream, ZipArchiveMode.Read))
                            {
                                foreach (ZipArchiveEntry entry in archive.Entries)
                                {
                                    if (!IsSafeZipEntryName(entry.FullName, out string rejectionReason))
                                    {
                                        throw new InvalidOperationException($"Security violation in payload: entry '{entry.FullName}' is unsafe ({rejectionReason}). Aborting installation.");
                                    }

                                    string destinationPath = Path.Combine(canonicalTargetDir, entry.FullName);
                                    string canonicalDestination;
                                    try
                                    {
                                        canonicalDestination = Path.GetFullPath(destinationPath);
                                    }
                                    catch (Exception ex)
                                    {
                                        throw new InvalidOperationException($"Security violation in payload: entry '{entry.FullName}' path resolution failed ({ex.Message}). Aborting installation.", ex);
                                    }

                                    bool isDirectory = string.IsNullOrEmpty(entry.Name) ||
                                                       entry.FullName.EndsWith("/") ||
                                                       entry.FullName.EndsWith("\\");

                                    if (isDirectory)
                                    {
                                        bool isInsideOrRoot = canonicalDestination.StartsWith(canonicalTargetDir, StringComparison.OrdinalIgnoreCase) ||
                                                              string.Equals(canonicalDestination + Path.DirectorySeparatorChar, canonicalTargetDir, StringComparison.OrdinalIgnoreCase);

                                        if (!isInsideOrRoot)
                                        {
                                            throw new InvalidOperationException($"Security violation in payload: directory entry '{entry.FullName}' escapes extraction root. Aborting installation.");
                                        }

                                        // Verify destination components for reparse points before creating directory
                                        if (!VerifyNoReparsePointsInPath(canonicalTargetDir, canonicalDestination, isDirectoryEntry: true, out string violationReason))
                                        {
                                            throw new InvalidOperationException($"Security violation for directory entry '{entry.FullName}': {violationReason}. Aborting installation.");
                                        }

                                        if (!Directory.Exists(canonicalDestination))
                                        {
                                            Directory.CreateDirectory(canonicalDestination);
                                        }
                                        continue;
                                    }

                                    if (!canonicalDestination.StartsWith(canonicalTargetDir, StringComparison.OrdinalIgnoreCase))
                                    {
                                        throw new InvalidOperationException($"Security violation in payload: file entry '{entry.FullName}' escapes extraction root. Aborting installation.");
                                    }

                                    // Verify destination components for reparse points before creating parent directory or writing file
                                    if (!VerifyNoReparsePointsInPath(canonicalTargetDir, canonicalDestination, isDirectoryEntry: false, out string violationFileReason))
                                    {
                                        throw new InvalidOperationException($"Security violation for file entry '{entry.FullName}': {violationFileReason}. Aborting installation.");
                                    }

                                    string? destDir = Path.GetDirectoryName(canonicalDestination);
                                    if (!string.IsNullOrEmpty(destDir) && !Directory.Exists(destDir))
                                    {
                                        Directory.CreateDirectory(destDir);
                                    }

                                    entry.ExtractToFile(canonicalDestination, overwrite: true);
                                }
                            }
                        }
                    }
                }

                // 6. Native COM Registration in HKCU
                Console.WriteLine("Registering Native COM Add-in...");
                string dllPath = Path.Combine(targetDir, "FastMailSearch.OutlookAddIn.dll");
                string codeBaseUrl = "file:///" + dllPath.Replace("\\", "/");
                string progId = "FastMailSearch.AddIn";
                string clsid = "{4A10F4D0-7E55-442C-B23B-983F78BA48D2}";

                using (RegistryKey hkcu = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Default))
                {
                    // ProgID Key
                    using (RegistryKey progKey = hkcu.CreateSubKey($@"Software\Classes\{progId}"))
                    {
                        progKey.SetValue("", "FastMailSearch.OutlookAddIn.ThisAddIn");
                        using (RegistryKey clsidSub = progKey.CreateSubKey("CLSID"))
                        {
                            clsidSub.SetValue("", clsid);
                        }
                    }

                    // CLSID Key
                    using (RegistryKey clsidKey = hkcu.CreateSubKey($@"Software\Classes\CLSID\{clsid}"))
                    {
                        clsidKey.SetValue("", "FastMailSearch.OutlookAddIn.ThisAddIn");
                        using (RegistryKey inprocKey = clsidKey.CreateSubKey("InprocServer32"))
                        {
                            inprocKey.SetValue("", "mscoree.dll");
                            inprocKey.SetValue("ThreadingModel", "Both");
                            inprocKey.SetValue("Class", "FastMailSearch.OutlookAddIn.ThisAddIn");
                            inprocKey.SetValue("Assembly", "FastMailSearch.OutlookAddIn, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null");
                            inprocKey.SetValue("RuntimeVersion", "v4.0.30319");
                            inprocKey.SetValue("CodeBase", codeBaseUrl);

                            using (RegistryKey verSub = inprocKey.CreateSubKey("1.0.0.0"))
                            {
                                verSub.SetValue("Class", "FastMailSearch.OutlookAddIn.ThisAddIn");
                                verSub.SetValue("Assembly", "FastMailSearch.OutlookAddIn, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null");
                                verSub.SetValue("RuntimeVersion", "v4.0.30319");
                                verSub.SetValue("CodeBase", codeBaseUrl);
                            }
                        }

                        using (RegistryKey clsidProgSub = clsidKey.CreateSubKey("ProgID"))
                        {
                            clsidProgSub.SetValue("", progId);
                        }

                        clsidKey.CreateSubKey(@"Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}");
                    }

                    // Outlook Addins Key (Native COM format ONLY - NO Manifest property)
                    using (RegistryKey addinKey = hkcu.CreateSubKey($@"Software\Microsoft\Office\Outlook\Addins\{progId}"))
                    {
                        addinKey.SetValue("FriendlyName", "FastMail Search");
                        addinKey.SetValue("Description", "FastMail Search Classic Outlook Add-In");
                        addinKey.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
                        addinKey.SetValue("CommandLineSafe", 0, RegistryValueKind.DWord);
                        addinKey.DeleteValue("Manifest", false);
                    }

                    // Resiliency Protection
                    using (RegistryKey resKey = hkcu.CreateSubKey(@"Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList"))
                    {
                        resKey.SetValue(progId, 1, RegistryValueKind.DWord);
                    }
                }

                Console.WriteLine("\n=================================================");
                Console.WriteLine(" FastMail Search Installed & Registered Cleanly!");
                Console.WriteLine("=================================================");

                if (!isSilent)
                {
                    MessageBox.Show(
                        "FastMail Search Outlook Add-In has been successfully installed!\n\nOpen Classic Outlook to start using FastMail Search.",
                        "FastMail Search Setup",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }

                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\n[ERROR] Installation failed: {ex.Message}\n{ex.StackTrace}");
                if (!isSilent)
                {
                    MessageBox.Show($"Installation failed: {ex.Message}", "Setup Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return 1;
            }
        }

        private static readonly string[] ReservedDeviceNames = new[]
        {
            "CON", "PRN", "AUX", "NUL",
            "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
            "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"
        };

        private static bool IsSafeZipEntryName(string entryFullName, out string reason)
        {
            if (string.IsNullOrWhiteSpace(entryFullName))
            {
                reason = "Entry name is empty";
                return false;
            }

            // Reject drive specifiers and alternate data stream syntax (e.g. C:, file.txt:stream)
            if (entryFullName.IndexOf(':') >= 0)
            {
                reason = "Contains colon / alternate data stream specifier";
                return false;
            }

            // Reject rooted paths (e.g. /etc/passwd or \Windows\...)
            if (entryFullName.StartsWith("/") || entryFullName.StartsWith("\\"))
            {
                reason = "Entry is rooted";
                return false;
            }

            char[] invalidChars = Path.GetInvalidFileNameChars();
            string[] segments = entryFullName.Split(new[] { '/', '\\' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string seg in segments)
            {
                // Reject segments containing invalid filename characters
                if (seg.IndexOfAny(invalidChars) >= 0)
                {
                    reason = $"Segment '{seg}' contains invalid path characters";
                    return false;
                }

                // Reject trailing dots or spaces which Windows Win32 strips
                if (seg != "." && seg != ".." && (seg.EndsWith(".") || seg.EndsWith(" ")))
                {
                    reason = $"Segment '{seg}' has trailing dot or space";
                    return false;
                }

                // Check for reserved DOS device names (e.g. CON, NUL, AUX, COM1.txt)
                int dotIdx = seg.IndexOf('.');
                string baseName = dotIdx >= 0 ? seg.Substring(0, dotIdx) : seg;
                for (int i = 0; i < ReservedDeviceNames.Length; i++)
                {
                    if (string.Equals(baseName, ReservedDeviceNames[i], StringComparison.OrdinalIgnoreCase))
                    {
                        reason = $"Segment '{seg}' is a reserved Windows device name ({baseName})";
                        return false;
                    }
                }
            }

            reason = string.Empty;
            return true;
        }

        private static ReparseCheckResult GetReparseStatus(string path, out string? errorDetail)
        {
            return InstallerDatabaseBackup.GetReparseStatus(path, out errorDetail);
        }

        private static bool VerifyNoReparsePointsInPath(string canonicalTargetDir, string canonicalDestination, bool isDirectoryEntry, out string failureReason)
        {
            failureReason = string.Empty;
            string rootTrimmed = canonicalTargetDir.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

            // Check extraction root itself
            var rootStatus = GetReparseStatus(rootTrimmed, out string? rootErr);
            if (rootStatus == ReparseCheckResult.IsReparsePoint)
            {
                failureReason = $"Extraction root '{rootTrimmed}' is a junction or reparse point";
                return false;
            }
            if (rootStatus == ReparseCheckResult.QueryFailed)
            {
                failureReason = $"Could not verify attributes of extraction root '{rootTrimmed}': {rootErr}";
                return false;
            }

            bool isRoot = string.Equals(canonicalDestination, rootTrimmed, StringComparison.OrdinalIgnoreCase);
            if (isRoot)
            {
                if (!isDirectoryEntry)
                {
                    failureReason = $"File entry destination cannot be the extraction root '{rootTrimmed}'";
                    return false;
                }
                return true;
            }

            if (!canonicalDestination.StartsWith(canonicalTargetDir, StringComparison.OrdinalIgnoreCase))
            {
                failureReason = $"Destination '{canonicalDestination}' escapes extraction root '{canonicalTargetDir}'";
                return false;
            }

            string relative = canonicalDestination.Substring(canonicalTargetDir.Length);
            string[] parts = relative.Split(new[] { Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar }, StringSplitOptions.RemoveEmptyEntries);

            string current = rootTrimmed;
            foreach (string part in parts)
            {
                current = Path.Combine(current, part);
                var compStatus = GetReparseStatus(current, out string? compErr);
                if (compStatus == ReparseCheckResult.IsReparsePoint)
                {
                    failureReason = $"Path component '{current}' is a junction or reparse point";
                    return false;
                }
                if (compStatus == ReparseCheckResult.QueryFailed)
                {
                    failureReason = $"Could not verify attributes of path component '{current}': {compErr}";
                    return false;
                }
                // ReparseCheckResult.DoesNotExist is handled safely: component is absent, safe to create/write
                // ReparseCheckResult.NotReparsePoint: component exists and is a standard file/directory
            }

            return true;
        }
    }
}
