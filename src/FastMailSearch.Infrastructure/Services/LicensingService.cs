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
        private const string LicenseKeyValueName = "LicenseKey";
        private const string LastRunTimeValueName = "LastRunTime";
        // This is the public key used for offline validation
        // WARNING: Replacing this key invalidates all previously generated licenses!
        public const string DefaultRsaPublicKeyXml = @"<RSAKeyValue><Modulus>tSq7b7pHYUKVmD3E8hhkiDtb1MGkRqxa9NwBe3p+KfsRPJtu2e8RTkOWDvIPmTSVl1JPtcvqpPpz78hndWdMxZLxCYy+hzhOnp1dlbGwpouqtPKImEbiB4T4w+U184t7xtguDlcrw4MDp5EOoEyNX4rNB1j+AXQ39pNpYoI8Z0Yrsw3yIiSrLhhGsrCEVjyPx6gm31kuOdkvCn2w9DmauNu87J85sdAaTz4Aq+gkWSTQEDCqwLGQDFxi7ZP9WaF3clQMS1idr6xwCNQEK/GMd2/bfzCGVTVv6tI8plt/l1VGeXtl1EXQc0XzWPKZIQU1WyAFGk0LFNYoUGPJExWYcQ==</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>";

        private readonly string _rsaPublicKeyXml;

        public LicensingService() : this(DefaultRsaPublicKeyXml)
        {
        }

        internal LicensingService(string publicKeyXml)
        {
            _rsaPublicKeyXml = publicKeyXml;
            EnsureTrialInitialized();
        }

        public string GetHardwareId()
        {
            try
            {
                string id = "";
#pragma warning disable CA1416 // Validate platform compatibility
                using (ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT SerialNumber FROM Win32_BaseBoard"))
                {
                    foreach (ManagementObject obj in searcher.Get())
                    {
                        id += obj["SerialNumber"]?.ToString() ?? "";
                    }
                }
#pragma warning restore CA1416 // Validate platform compatibility
                
                if (string.IsNullOrWhiteSpace(id))
                    id = Environment.MachineName;

                using (var sha = SHA256.Create())
                {
                    var hash = sha.ComputeHash(Encoding.UTF8.GetBytes(id));
                    return Convert.ToBase64String(hash).Substring(0, 16).Replace("/", "").Replace("+", "");
                }
            }
            catch
            {
                return "UNKNOWN-HWID";
            }
        }

        public LicenseStatus GetStatus()
        {
            try
            {
                // Check if activated
#pragma warning disable CA1416
                using (var key = Registry.CurrentUser.OpenSubKey(RegKeyPath))
                {
                    if (key != null)
                    {
                        var licenseBlob = key.GetValue(LicenseKeyValueName) as byte[];
                        if (licenseBlob != null)
                        {
                            try
                            {
                                var unencrypted = ProtectedData.Unprotect(licenseBlob, null, DataProtectionScope.CurrentUser);
                                var licenseKeyStr = Encoding.UTF8.GetString(unencrypted);
                                if (ValidateLicenseKey(licenseKeyStr))
                                {
                                    return LicenseStatus.Valid;
                                }
                            }
                            catch { }
                        }
                    }
                }
#pragma warning restore CA1416

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

        public Task<bool> ActivateAsync(string licenseKey)
        {
            if (ValidateLicenseKey(licenseKey))
            {
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

                // 3. Verify Hardware ID
                if (payload.HardwareId != GetHardwareId())
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

                // Verify license validity does not exceed 1 year from IssueDate (plus 2 days leap year / clock skew tolerance)
                if (payload.ExpiryDate.Value > payload.IssueDate.Value.AddYears(1).AddDays(2))
                {
                    return false;
                }

                return true;
            }
            catch { }
            return false;
        }

        private void EnsureTrialInitialized()
        {
            var primary = ReadEncryptedDate(RegKeyPath, InstallDateValueName);
            var secondary = ReadEncryptedDate(SecondaryRegKeyPath, SecondaryValueName);

            if (primary == null && secondary == null)
            {
                var lastRunTime = GetLastRunTime();
                if (lastRunTime.HasValue)
                {
                    // Tampering detected: Both InstallDate anchors missing, but LastRunTime exists.
                    // Hard-fault the trial by conservatively setting InstallDate to MinValue.
                    var expiredDate = DateTime.MinValue;
                    WriteEncryptedDate(RegKeyPath, InstallDateValueName, expiredDate);
                    WriteEncryptedDate(SecondaryRegKeyPath, SecondaryValueName, expiredDate);
                    return;
                }

                // First run
                var now = DateTime.UtcNow;
                WriteEncryptedDate(RegKeyPath, InstallDateValueName, now);
                WriteEncryptedDate(SecondaryRegKeyPath, SecondaryValueName, now);
                return;
            }

            // Sync missing or mismatched anchors
            var oldest = primary;
            if (secondary != null)
            {
                if (oldest == null || secondary.Value < oldest.Value)
                    oldest = secondary;
            }

            if (oldest.HasValue)
            {
                if (primary == null || primary.Value != oldest.Value)
                    WriteEncryptedDate(RegKeyPath, InstallDateValueName, oldest.Value);
                
                if (secondary == null || secondary.Value != oldest.Value)
                    WriteEncryptedDate(SecondaryRegKeyPath, SecondaryValueName, oldest.Value);
            }
        }

        private DateTime? GetInstallDate()
        {
            var primary = ReadEncryptedDate(RegKeyPath, InstallDateValueName);
            var secondary = ReadEncryptedDate(SecondaryRegKeyPath, SecondaryValueName);

            if (primary != null && secondary != null)
            {
                return primary.Value < secondary.Value ? primary : secondary;
            }

            if (primary != null) return primary;
            if (secondary != null) return secondary;

            return null;
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
            try
            {
#pragma warning disable CA1416
                using (var key = Registry.CurrentUser.OpenSubKey(RegKeyPath))
                {
                    if (key != null)
                    {
                        var existing = key.GetValue(LastRunTimeValueName) as byte[];
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

        private void UpdateLastRunTime(DateTime now)
        {
            try
            {
#pragma warning disable CA1416
                using (var key = Registry.CurrentUser.CreateSubKey(RegKeyPath))
                {
                    var dateBytes = BitConverter.GetBytes(now.Ticks);
                    var encrypted = ProtectedData.Protect(dateBytes, null, DataProtectionScope.CurrentUser);
                    key.SetValue(LastRunTimeValueName, encrypted, RegistryValueKind.Binary);
                }
#pragma warning restore CA1416
            }
            catch { }
        }
    }
}
