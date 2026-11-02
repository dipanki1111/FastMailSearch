using System;
using System.IO;
using System.Management;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using FastMailSearch.Core.Interfaces;
using Microsoft.Win32;
using Newtonsoft.Json;

[assembly: InternalsVisibleTo("FastMailSearch.Core.Tests")]

namespace FastMailSearch.Infrastructure.Services
{
    public class LicensePayload
    {
        public string Email { get; set; } = string.Empty;
        public string HardwareId { get; set; } = string.Empty;
        public DateTime? IssueDate { get; set; }
        public DateTime? ExpiryDate { get; set; }
    }

    public class LicensingService : ILicensingService
    {
        private const int TrialDays = 30;
        private const string RegKeyPath = @"Software\FastMailSearch";
        private const string SecondaryRegKeyPath = @"Software\Classes\CLSID\{A3289035-7798-4665-A226-5F792CD0F8BC}";
        private const string InstallDateValueName = "InstallDate";
        private const string SecondaryValueName = "Data";
        private const string SecondaryLastRunValueName = "LastTime";
        private const string LicenseKeyValueName = "LicenseKey";
        private const string LastRunTimeValueName = "LastRunTime";
        private static readonly byte[] FileAnchorMagic = new byte[] { 0x46, 0x4D, 0x53, 0x31 }; // "FMS1"

        // This is the public key used for offline validation
        // WARNING: Replacing this key invalidates all previously generated licenses!
        public const string DefaultRsaPublicKeyXml = @"<RSAKeyValue><Modulus>tSq7b7pHYUKVmD3E8hhkiDtb1MGkRqxa9NwBe3p+KfsRPJtu2e8RTkOWDvIPmTSVl1JPtcvqpPpz78hndWdMxZLxCYy+hzhOnp1dlbGwpouqtPKImEbiB4T4w+U184t7xtguDlcrw4MDp5EOoEyNX4rNB1j+AXQ39pNpYoI8Z0Yrsw3yIiSrLhhGsrCEVjyPx6gm31kuOdkvCn2w9DmauNu87J85sdAaTz4Aq+gkWSTQEDCqwLGQDFxi7ZP9WaF3clQMS1idr6xwCNQEK/GMd2/bfzCGVTVv6tI8plt/l1VGeXtl1EXQc0XzWPKZIQU1WyAFGk0LFNYoUGPJExWYcQ==</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>";

        private readonly string _rsaPublicKeyXml;
        private readonly string _fileAnchorPath;

        private static string DefaultFileAnchorPath =>
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "FastMailSearch", ".fms_identity");

        public LicensingService() : this(DefaultRsaPublicKeyXml, null)
        {
        }

        internal LicensingService(string publicKeyXml, string? fileAnchorPath = null)
        {
            _rsaPublicKeyXml = publicKeyXml;
            _fileAnchorPath = fileAnchorPath ?? DefaultFileAnchorPath;
            EnsureTrialInitialized();
        }

        public string GetHardwareId()
        {
            try
            {
                string rawId = "";
                try
                {
#pragma warning disable CA1416 // Validate platform compatibility
                    using (ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT SerialNumber FROM Win32_BaseBoard"))
                    {
                        foreach (ManagementObject obj in searcher.Get())
                        {
                            var serial = obj["SerialNumber"]?.ToString()?.Trim();
                            if (!string.IsNullOrWhiteSpace(serial) &&
                                !serial.Equals("None", StringComparison.OrdinalIgnoreCase) &&
                                !serial.Equals("Default string", StringComparison.OrdinalIgnoreCase) &&
                                !serial.Equals("To be filled by O.E.M.", StringComparison.OrdinalIgnoreCase) &&
                                !serial.Equals("Base Board Serial Number", StringComparison.OrdinalIgnoreCase))
                            {
                                rawId += serial;
                            }
                        }
                    }
#pragma warning restore CA1416 // Validate platform compatibility
                }
                catch
                {
                    // WMI service stopped, restricted permissions, or unavailable; fall back to MachineName
                }
                
                if (string.IsNullOrWhiteSpace(rawId))
                {
                    try
                    {
                        rawId = Environment.MachineName?.Trim() ?? "";
                    }
                    catch { }
                }

                if (string.IsNullOrWhiteSpace(rawId))
                {
                    return "HWID_UNAVAILABLE";
                }

                using (var sha = SHA256.Create())
                {
                    var hash = sha.ComputeHash(Encoding.UTF8.GetBytes(rawId));
                    return Convert.ToBase64String(hash).Substring(0, 16).Replace("/", "").Replace("+", "");
                }
            }
            catch
            {
                return "HWID_UNAVAILABLE";
            }
        }

        public LicenseStatus GetStatus()
        {
            try
            {
                // Check if activated
                var licenseKeyStr = GetStoredLicenseKey();
                if (!string.IsNullOrWhiteSpace(licenseKeyStr) && ValidateLicenseKey(licenseKeyStr))
                {
                    return LicenseStatus.Valid;
                }

                // Check trial
                var installDate = GetInstallDate();
                if (installDate == null) return LicenseStatus.Expired; // Tampered or error

                var now = DateTime.UtcNow;
                var lastRunTime = GetLastRunTime();

                if (lastRunTime.HasValue && now < lastRunTime.Value)
                {
                    return LicenseStatus.Expired; // Clock rollback detected
                }

                if (now > installDate.Value.AddDays(TrialDays))
                {
                    return LicenseStatus.Expired;
                }

                UpdateLastRunTime(now);

                return LicenseStatus.Trial;
            }
            catch
            {
                return LicenseStatus.Expired;
            }
        }

        public int GetTrialDaysRemaining()
        {
            if (GetStatus() != LicenseStatus.Trial) return 0;
            
            var installDate = GetInstallDate();
            if (installDate == null) return 0;

            var remaining = (installDate.Value.AddDays(TrialDays) - DateTime.UtcNow).TotalDays;
            return remaining > 0 ? (int)Math.Ceiling(remaining) : 0;
        }

        public int GetDaysUntilExpiry()
        {
            if (GetStatus() != LicenseStatus.Valid) return 0;
            var expiry = GetLicenseExpiryDate();
            if (!expiry.HasValue) return 0;
            var remaining = (expiry.Value - DateTime.UtcNow).TotalDays;
            return remaining > 0 ? (int)Math.Ceiling(remaining) : 0;
        }

        public DateTime? GetLicenseExpiryDate()
        {
            var key = GetStoredLicenseKey();
            if (string.IsNullOrWhiteSpace(key)) return null;
            var payload = ParsePayload(key);
            return payload?.ExpiryDate;
        }

        public Task<bool> ActivateAsync(string licenseKey)
        {
            if (ValidateLicenseKey(licenseKey))
            {
                var newPayload = ParsePayload(licenseKey);
                if (newPayload == null || !newPayload.ExpiryDate.HasValue) return Task.FromResult(false);

                // Prevent unintended expiry shortening if an active license is currently installed
                var existingKey = GetStoredLicenseKey();
                if (!string.IsNullOrWhiteSpace(existingKey) && ValidateLicenseKey(existingKey))
                {
                    var existingPayload = ParsePayload(existingKey);
                    if (existingPayload?.ExpiryDate.HasValue == true && existingPayload.ExpiryDate.Value > DateTime.UtcNow)
                    {
                        if (newPayload.ExpiryDate.Value < existingPayload.ExpiryDate.Value)
                        {
                            return Task.FromResult(false);
                        }
                    }
                }

#pragma warning disable CA1416
                var bytes = Encoding.UTF8.GetBytes(licenseKey);
                var encrypted = ProtectedData.Protect(bytes, null, DataProtectionScope.CurrentUser);
                using (var key = Registry.CurrentUser.CreateSubKey(RegKeyPath))
                {
                    key.SetValue(LicenseKeyValueName, encrypted, RegistryValueKind.Binary);
                }
#pragma warning restore CA1416
                return Task.FromResult(true);
            }
            return Task.FromResult(false);
        }

        private bool ValidateLicenseKey(string licenseKey)
        {
            // Format: Base64(json).Base64(signature)
            try
            {
                var parts = licenseKey.Split('.');
                if (parts.Length != 2) return false;

                var payloadBytes = Convert.FromBase64String(parts[0]);
                var signatureBytes = Convert.FromBase64String(parts[1]);

                // 1. Verify Signature
                using (var rsa = new RSACryptoServiceProvider())
                {
                    rsa.FromXmlString(_rsaPublicKeyXml);
                    
                    using (var sha = SHA256.Create())
                    {
                        bool isSignatureValid = rsa.VerifyData(payloadBytes, sha, signatureBytes);
                        if (!isSignatureValid) return false;
                    }
                }

                // 2. Verify Payload
                var json = Encoding.UTF8.GetString(payloadBytes);
                var payload = JsonConvert.DeserializeObject<LicensePayload>(json);
                if (payload == null) return false;

                // 3. Verify Hardware ID (fail closed on unavailable, unknown, or wildcard)
                var currentHwid = GetHardwareId();
                if (string.IsNullOrWhiteSpace(currentHwid) ||
                    currentHwid.Equals("HWID_UNAVAILABLE", StringComparison.OrdinalIgnoreCase) ||
                    currentHwid.Equals("UNKNOWN-HWID", StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }

                if (string.IsNullOrWhiteSpace(payload.HardwareId) ||
                    payload.HardwareId.Equals("HWID_UNAVAILABLE", StringComparison.OrdinalIgnoreCase) ||
                    payload.HardwareId.Equals("UNKNOWN-HWID", StringComparison.OrdinalIgnoreCase) ||
                    payload.HardwareId == "*")
                {
                    return false;
                }

                if (!string.Equals(payload.HardwareId, currentHwid, StringComparison.Ordinal))
                {
                    return false;
                }

                // 4. Verify Dates
                var now = DateTime.UtcNow;
                if (!payload.IssueDate.HasValue)
                {
                    return false;
                }

                if (payload.IssueDate.Value > now.AddDays(1)) // 1 day tolerance for timezones/clock skew
                {
                    return false;
                }
                
                // Paid licenses must possess an ExpiryDate (lifetime licenses are not offered in this release)
                if (!payload.ExpiryDate.HasValue)
                {
                    return false;
                }

                if (payload.ExpiryDate.Value < now)
                {
                    return false;
                }

                // Verify license validity does not exceed 1 year + 35 days early renewal window from IssueDate
                if (payload.ExpiryDate.Value > payload.IssueDate.Value.AddYears(1).AddDays(35))
                {
                    return false;
                }

                return true;
            }
            catch { }
            return false;
        }

        private string? GetStoredLicenseKey()
        {
            try
            {
#pragma warning disable CA1416
                using (var key = Registry.CurrentUser.OpenSubKey(RegKeyPath))
                {
                    if (key != null)
                    {
                        var licenseBlob = key.GetValue(LicenseKeyValueName) as byte[];
                        if (licenseBlob != null)
                        {
                            var unencrypted = ProtectedData.Unprotect(licenseBlob, null, DataProtectionScope.CurrentUser);
                            return Encoding.UTF8.GetString(unencrypted);
                        }
                    }
                }
#pragma warning restore CA1416
            }
            catch { }
            return null;
        }

        private static LicensePayload? ParsePayload(string licenseKey)
        {
            try
            {
                var parts = licenseKey.Split('.');
                if (parts.Length != 2) return null;
                var payloadBytes = Convert.FromBase64String(parts[0]);
                var json = Encoding.UTF8.GetString(payloadBytes);
                return JsonConvert.DeserializeObject<LicensePayload>(json);
            }
            catch
            {
                return null;
            }
        }

        private (DateTime? InstallDate, DateTime? LastRunTime) ReadFileAnchor()
        {
            try
            {
                if (string.IsNullOrWhiteSpace(_fileAnchorPath) || !File.Exists(_fileAnchorPath))
                {
                    return (null, null);
                }

#pragma warning disable CA1416
                var encrypted = File.ReadAllBytes(_fileAnchorPath);
                var raw = ProtectedData.Unprotect(encrypted, null, DataProtectionScope.CurrentUser);
#pragma warning restore CA1416
                if (raw.Length >= 20 &&
                    raw[0] == FileAnchorMagic[0] &&
                    raw[1] == FileAnchorMagic[1] &&
                    raw[2] == FileAnchorMagic[2] &&
                    raw[3] == FileAnchorMagic[3])
                {
                    long installTicks = BitConverter.ToInt64(raw, 4);
                    long lastRunTicks = BitConverter.ToInt64(raw, 12);

                    DateTime? inst = installTicks > 0 ? new DateTime(installTicks, DateTimeKind.Utc) : null;
                    DateTime? last = lastRunTicks > 0 ? new DateTime(lastRunTicks, DateTimeKind.Utc) : null;
                    return (inst, last);
                }
            }
            catch { }
            return (null, null);
        }

        private void WriteFileAnchor(DateTime? installDate, DateTime? lastRunTime)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(_fileAnchorPath)) return;
                var dir = Path.GetDirectoryName(_fileAnchorPath);
                if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                using (var ms = new MemoryStream())
                using (var bw = new BinaryWriter(ms))
                {
                    bw.Write(FileAnchorMagic);
                    bw.Write(installDate.HasValue ? installDate.Value.Ticks : 0L);
                    bw.Write(lastRunTime.HasValue ? lastRunTime.Value.Ticks : 0L);
                    var raw = ms.ToArray();
#pragma warning disable CA1416
                    var encrypted = ProtectedData.Protect(raw, null, DataProtectionScope.CurrentUser);
#pragma warning restore CA1416
                    File.WriteAllBytes(_fileAnchorPath, encrypted);
                }
            }
            catch { }
        }

        private void EnsureTrialInitialized()
        {
            var primaryInstall = ReadEncryptedDate(RegKeyPath, InstallDateValueName);
            var secondaryInstall = ReadEncryptedDate(SecondaryRegKeyPath, SecondaryValueName);
            var (fileInstall, fileLastRun) = ReadFileAnchor();

            var primaryLastRun = ReadEncryptedDate(RegKeyPath, LastRunTimeValueName);
            var secondaryLastRun = ReadEncryptedDate(SecondaryRegKeyPath, SecondaryLastRunValueName);

            bool hasAnyInstallDate = primaryInstall.HasValue || secondaryInstall.HasValue || fileInstall.HasValue;
            bool hasAnyLastRunTime = primaryLastRun.HasValue || secondaryLastRun.HasValue || fileLastRun.HasValue;

            if (!hasAnyInstallDate)
            {
                if (hasAnyLastRunTime)
                {
                    // Tampering detected: InstallDate missing across all anchors, but LastRunTime exists.
                    // Hard-fault the trial by setting InstallDate to MinValue across all anchors.
                    var expiredDate = DateTime.MinValue;
                    var preservedLastRun = primaryLastRun ?? secondaryLastRun ?? fileLastRun ?? DateTime.UtcNow;
                    WriteEncryptedDate(RegKeyPath, InstallDateValueName, expiredDate);
                    WriteEncryptedDate(SecondaryRegKeyPath, SecondaryValueName, expiredDate);
                    WriteFileAnchor(expiredDate, preservedLastRun);
                    return;
                }

                // First run: brand new legitimate installation
                var now = DateTime.UtcNow;
                WriteEncryptedDate(RegKeyPath, InstallDateValueName, now);
                WriteEncryptedDate(SecondaryRegKeyPath, SecondaryValueName, now);
                WriteEncryptedDate(RegKeyPath, LastRunTimeValueName, now);
                WriteEncryptedDate(SecondaryRegKeyPath, SecondaryLastRunValueName, now);
                WriteFileAnchor(now, now);
                return;
            }

            // At least one valid install date exists: determine oldest install date
            DateTime? oldestInstall = null;
            if (primaryInstall.HasValue) oldestInstall = primaryInstall.Value;
            if (secondaryInstall.HasValue) oldestInstall = oldestInstall == null || secondaryInstall.Value < oldestInstall.Value ? secondaryInstall.Value : oldestInstall;
            if (fileInstall.HasValue) oldestInstall = oldestInstall == null || fileInstall.Value < oldestInstall.Value ? fileInstall.Value : oldestInstall;

            // Determine latest last run time
            DateTime? latestLastRun = null;
            if (primaryLastRun.HasValue) latestLastRun = primaryLastRun.Value;
            if (secondaryLastRun.HasValue) latestLastRun = latestLastRun == null || secondaryLastRun.Value > latestLastRun.Value ? secondaryLastRun.Value : latestLastRun;
            if (fileLastRun.HasValue) latestLastRun = latestLastRun == null || fileLastRun.Value > latestLastRun.Value ? fileLastRun.Value : latestLastRun;

            // Synchronize oldest install date to all anchors
            if (oldestInstall.HasValue)
            {
                if (primaryInstall == null || primaryInstall.Value != oldestInstall.Value)
                    WriteEncryptedDate(RegKeyPath, InstallDateValueName, oldestInstall.Value);

                if (secondaryInstall == null || secondaryInstall.Value != oldestInstall.Value)
                    WriteEncryptedDate(SecondaryRegKeyPath, SecondaryValueName, oldestInstall.Value);

                if (fileInstall == null || fileInstall.Value != oldestInstall.Value)
                    WriteFileAnchor(oldestInstall.Value, latestLastRun);
            }

            // Synchronize latest last run time to all anchors
            if (latestLastRun.HasValue)
            {
                if (primaryLastRun == null || primaryLastRun.Value != latestLastRun.Value)
                    WriteEncryptedDate(RegKeyPath, LastRunTimeValueName, latestLastRun.Value);

                if (secondaryLastRun == null || secondaryLastRun.Value != latestLastRun.Value)
                    WriteEncryptedDate(SecondaryRegKeyPath, SecondaryLastRunValueName, latestLastRun.Value);

                if (fileLastRun == null || fileLastRun.Value != latestLastRun.Value)
                    WriteFileAnchor(oldestInstall ?? DateTime.UtcNow, latestLastRun.Value);
            }
        }

        private DateTime? GetInstallDate()
        {
            var primary = ReadEncryptedDate(RegKeyPath, InstallDateValueName);
            var secondary = ReadEncryptedDate(SecondaryRegKeyPath, SecondaryValueName);
            var (fileInstall, _) = ReadFileAnchor();

            DateTime? oldest = null;
            if (primary.HasValue) oldest = primary.Value;
            if (secondary.HasValue) oldest = oldest == null || secondary.Value < oldest.Value ? secondary.Value : oldest;
            if (fileInstall.HasValue) oldest = oldest == null || fileInstall.Value < oldest.Value ? fileInstall.Value : oldest;

            return oldest;
        }

        private DateTime? ReadEncryptedDate(string keyPath, string valueName)
        {
            try
            {
#pragma warning disable CA1416
                using (var key = Registry.CurrentUser.OpenSubKey(keyPath))
                {
                    if (key != null)
                    {
                        var existing = key.GetValue(valueName) as byte[];
                        if (existing != null)
                        {
                            var decrypted = ProtectedData.Unprotect(existing, null, DataProtectionScope.CurrentUser);
                            var ticks = BitConverter.ToInt64(decrypted, 0);
                            return new DateTime(ticks, DateTimeKind.Utc);
                        }
                    }
                }
#pragma warning restore CA1416
            }
            catch { }
            return null;
        }

        private void WriteEncryptedDate(string keyPath, string valueName, DateTime date)
        {
            try
            {
#pragma warning disable CA1416
                using (var key = Registry.CurrentUser.CreateSubKey(keyPath))
                {
                    var dateBytes = BitConverter.GetBytes(date.Ticks);
                    var encrypted = ProtectedData.Protect(dateBytes, null, DataProtectionScope.CurrentUser);
                    key.SetValue(valueName, encrypted, RegistryValueKind.Binary);
                }
#pragma warning restore CA1416
            }
            catch { }
        }

        private DateTime? GetLastRunTime()
        {
            var primary = ReadEncryptedDate(RegKeyPath, LastRunTimeValueName);
            var secondary = ReadEncryptedDate(SecondaryRegKeyPath, SecondaryLastRunValueName);
            var (_, fileLastRun) = ReadFileAnchor();

            DateTime? latest = null;
            if (primary.HasValue) latest = primary.Value;
            if (secondary.HasValue) latest = latest == null || secondary.Value > latest.Value ? secondary.Value : latest;
            if (fileLastRun.HasValue) latest = latest == null || fileLastRun.Value > latest.Value ? fileLastRun.Value : latest;

            return latest;
        }

        private void UpdateLastRunTime(DateTime now)
        {
            try
            {
                WriteEncryptedDate(RegKeyPath, LastRunTimeValueName, now);
                WriteEncryptedDate(SecondaryRegKeyPath, SecondaryLastRunValueName, now);
                var installDate = GetInstallDate();
                WriteFileAnchor(installDate, now);
            }
            catch { }
        }
    }
}
