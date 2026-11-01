using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using FastMailSearch.Core.Interfaces;
using FastMailSearch.Infrastructure.Services;
using Microsoft.Win32;
using Newtonsoft.Json;
using Xunit;

namespace FastMailSearch.Core.Tests
{
    public class LicensingServiceTests : IDisposable
    {
        private const string RegKeyPath = @"Software\FastMailSearch";
        private const string SecondaryRegKeyPath = @"Software\Classes\CLSID\{A3289035-7798-4665-A226-5F792CD0F8BC}";
        private const string SecondaryValueName = "Data";
        private const string InstallDateValueName = "InstallDate";
        private const string LicenseKeyValueName = "LicenseKey";
        private const string TestPrivateKeyXml = @"<RSAKeyValue><Modulus>8wNM8raekxgmB/88a1CyL6c5gGme7tiW18WcQssw/jGTiy7DxE0q3ScEXuIK78Gz8sT4uwOSMu8CeeKBvVdN4rnBXonDcenAQWA/5rKPKcj6bCIMxoDRPYXngxYZ/Wbz0IhbO3M6ct4HQ7+K9faJGA+R48RIZMBeOTXRTnv7wWRbRoigvZNMzL3poGmFs3FioGf5EZrAgPaLBrfruVHLrrfU7gM+rJhVCq2OWQ3fRrDVGwU4FhQGuYqetvmpsauwuwlOukDKTleySKuC3rxXCMQdpXSoJF14fmcM1XFWzHsrwKkYR0YFa40fGpK0IBWO7cp1WMmfyXhn/Uxvm9GM9Q==</Modulus><Exponent>AQAB</Exponent><P>+KGGsdZMIDkEcjUcpEjB91dIN1XnDppT2/mqV3k0tqXwORvCakaY963OY1kmKIOQZoQNyzOsZ6Xq9lIRFcotqe3SBsfexPx8BoFPRNU1HEJwZkGw3HQDHcHfmKdrMVlAigI99jWurWyJDyrEghGMYfqUdJ7dQDfoRRseAacizYM=</P><Q>+jcly58qlWwoxmvZXv19m7GaqmkciPiQzg5qJ7TYRA2AFRM70u/ARiHfjgkTvw4yyECI8O2aj1SpAwj8tYDXl93EKALRM/5IGaslK0uTqYwuiCGcC6oA5nG9ZfexzAsixICZexuL0KsYL1lBE4fAlthZNmsiIuBLizvT8PIkaic=</Q><DP>BfirKX5CGcwExVFtPKzaq7wW6e1QPVOLjA70/mDU2UYm/Ngj1TWY2un7uPsdJX/zaeLmCelVzK95cy0JaAKREjG6B5upSAXgXK+iXb7ZJbKazHnra4AFyyMq4Zmzw12nZSkkn9MuPNAMgqa+TXlqa7f+9MYT/kTZeT9rexKujU0=</DP><DQ>RdtJsQyx/d+hjqwjlvKcs/Ak5zjLX65N90ls98L0siM2xCuHgz+uZaYlXOlJ2whiwyfJU8JyaVvQh+llbGMZTev7TNRkheiTRp9SlEDYa/iYMSd5AoSnFa9hmmtJ91aQkV35+7XYFhz+dpGrqoE3EqY6WggVErmePbQpWtavtfc=</DQ><InverseQ>P9hGoOkjfX9pk/hBIDNemX3TSdRJDAAqrHOOZkpdPaD2eSnDnJfXwaWammbwEM0fpfWO9L+dWCWaHV59po0Sv+aAs3a2vK+v943Ret+w2tx3UOhitxEc8hKRXDz/qeDfBoTZUaedvtTWzVU7bGFZHDXi1ddp3GrzUuCWvPa8QMk=</InverseQ><D>iHZ4BWvrjGppIOCBSzobdYsQTT4Ck9loi25am1+C+ZtZYTU3lcpHobzBRY7l2VEK/e8CRfJxbMyutUcu1+ZoxELPhYkTCarb3VpxR9lJb4t+pK5SQj4gQc4p/YCxw4zGqXmI0c+ZKNvpdG/hVtHElpBXunGEvZ8QykugZ9RTwXrDbb3KyjnT39HusuRof20TR+DjAewEffvoPtYCpApW2GSIFEh5U6+GGglkWW8akK1LVQnh7cwhZXkXzpCbaUVY0nnxmxPVmHsr91At5CEckmxDJ4Eu4eHl3gGtDXE8a6nLAWXzvgMnUUcJ7UD9vW5ybE9EBA6pGlqMVkxGGuEXmQ==</D></RSAKeyValue>";
        private const string TestPublicKeyXml = @"<RSAKeyValue><Modulus>8wNM8raekxgmB/88a1CyL6c5gGme7tiW18WcQssw/jGTiy7DxE0q3ScEXuIK78Gz8sT4uwOSMu8CeeKBvVdN4rnBXonDcenAQWA/5rKPKcj6bCIMxoDRPYXngxYZ/Wbz0IhbO3M6ct4HQ7+K9faJGA+R48RIZMBeOTXRTnv7wWRbRoigvZNMzL3poGmFs3FioGf5EZrAgPaLBrfruVHLrrfU7gM+rJhVCq2OWQ3fRrDVGwU4FhQGuYqetvmpsauwuwlOukDKTleySKuC3rxXCMQdpXSoJF14fmcM1XFWzHsrwKkYR0YFa40fGpK0IBWO7cp1WMmfyXhn/Uxvm9GM9Q==</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>";

        public LicensingServiceTests()
        {
            CleanupRegistry();
        }

        public void Dispose()
        {
            CleanupRegistry();
        }

        private void CleanupRegistry()
        {
#pragma warning disable CA1416
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(RegKeyPath, true))
                {
                    if (key != null)
                    {
                        key.DeleteValue(InstallDateValueName, false);
                        key.DeleteValue(LicenseKeyValueName, false);
                        key.DeleteValue("LastRunTime", false);
                    }
                }
                using (var key = Registry.CurrentUser.OpenSubKey(SecondaryRegKeyPath, true))
                {
                    if (key != null)
                    {
                        key.DeleteValue(SecondaryValueName, false);
                    }
                }
            }
            catch { }
#pragma warning restore CA1416
        }

        private void SetRegistryDate(string valueName, DateTime date, string keyPath = RegKeyPath)
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

        private void SetCorruptedRegistryData(string valueName, string keyPath = RegKeyPath)
        {
#pragma warning disable CA1416
            using (var key = Registry.CurrentUser.CreateSubKey(keyPath))
            {
                var randomBytes = new byte[64];
                new Random().NextBytes(randomBytes);
                // Write raw garbage WITHOUT DPAPI protection to simulate corruption
                key.SetValue(valueName, randomBytes, RegistryValueKind.Binary);
            }
#pragma warning restore CA1416
        }

        private string GenerateTestLicense(string hardwareId, DateTime? issueDate = null, DateTime? expiryDate = null, bool omitExpiry = false)
        {
            var resolvedIssue = issueDate ?? DateTime.UtcNow;
            var payload = new LicensePayload
            {
                Email = "test@fastmail.com",
                HardwareId = hardwareId,
                IssueDate = resolvedIssue,
                ExpiryDate = omitExpiry ? null : (expiryDate ?? resolvedIssue.AddYears(1))
            };
            var json = JsonConvert.SerializeObject(payload);
            var payloadBytes = Encoding.UTF8.GetBytes(json);

            using (var rsa = new RSACryptoServiceProvider())
            {
                rsa.FromXmlString(TestPrivateKeyXml);
                using (var sha = SHA256.Create())
                {
                    var signatureBytes = rsa.SignData(payloadBytes, sha);
                    return $"{Convert.ToBase64String(payloadBytes)}.{Convert.ToBase64String(signatureBytes)}";
                }
            }
        }

        [Fact]
        public void LicensingService_InitializesTrial_OnFirstRun_CreatesBothAnchors()
        {
            var service = new LicensingService(TestPublicKeyXml);
            Assert.Equal(LicenseStatus.Trial, service.GetStatus());
            Assert.True(service.GetTrialDaysRemaining() > 29); // Starts at 30
        }

        [Fact]
        public void LicensingService_MissingPrimaryAnchor_RestoresFromSecondary()
        {
            // Set only the secondary anchor
            SetRegistryDate(SecondaryValueName, DateTime.UtcNow.AddDays(-10), SecondaryRegKeyPath);
            
            var service = new LicensingService(TestPublicKeyXml);
            Assert.Equal(LicenseStatus.Trial, service.GetStatus());
            Assert.Equal(20, service.GetTrialDaysRemaining());
        }

        [Fact]
        public void LicensingService_MissingSecondaryAnchor_RestoresFromPrimary()
        {
            // Set only the primary anchor
            SetRegistryDate(InstallDateValueName, DateTime.UtcNow.AddDays(-10));
            
            var service = new LicensingService(TestPublicKeyXml);
            Assert.Equal(LicenseStatus.Trial, service.GetStatus());
            Assert.Equal(20, service.GetTrialDaysRemaining());
        }

        [Fact]
        public void LicensingService_MismatchedAnchors_UsesOlderDate()
        {
            // Primary says 10 days ago, Secondary says 20 days ago. It should pick 20 days ago.
            SetRegistryDate(InstallDateValueName, DateTime.UtcNow.AddDays(-10));
            SetRegistryDate(SecondaryValueName, DateTime.UtcNow.AddDays(-20), SecondaryRegKeyPath);
            
            var service = new LicensingService(TestPublicKeyXml);
            Assert.Equal(LicenseStatus.Trial, service.GetStatus());
            Assert.Equal(10, service.GetTrialDaysRemaining());
        }

        [Fact]
        public void LicensingService_CorruptedSecondaryAnchor_RestoresFromPrimary()
        {
            SetRegistryDate(InstallDateValueName, DateTime.UtcNow.AddDays(-10));
            SetCorruptedRegistryData(SecondaryValueName, SecondaryRegKeyPath);
            
            var service = new LicensingService(TestPublicKeyXml);
            Assert.Equal(LicenseStatus.Trial, service.GetStatus());
            Assert.Equal(20, service.GetTrialDaysRemaining());
        }

        [Fact]
        public void LicensingService_TrialExpires_After30Days()
        {
            SetRegistryDate(InstallDateValueName, DateTime.UtcNow.AddDays(-31));
            var service = new LicensingService(TestPublicKeyXml);
            Assert.Equal(LicenseStatus.Expired, service.GetStatus());
        }

        [Fact]
        public void LicensingService_ClockTampering_Backward_FailsTrial()
        {
            // Simulate that the app was last run tomorrow
            SetRegistryDate(InstallDateValueName, DateTime.UtcNow.AddDays(-1));
            SetRegistryDate("LastRunTime", DateTime.UtcNow.AddDays(1));
            
            var service = new LicensingService(TestPublicKeyXml);
            Assert.Equal(LicenseStatus.Expired, service.GetStatus());
        }

        [Fact]
        public void LicensingService_ClockTampering_MissingLastRunTime_Succeeds()
        {
            // Simulate upgrading from an older version (no LastRunTime recorded yet)
            SetRegistryDate(InstallDateValueName, DateTime.UtcNow.AddDays(-1));
            
            var service = new LicensingService(TestPublicKeyXml);
            Assert.Equal(LicenseStatus.Trial, service.GetStatus());
        }

        [Fact]
        public void LicensingService_ClockTampering_CorruptedLastRunTime_Succeeds()
        {
            // If LastRunTime is unreadable, it should default to ignoring it and re-saving
            SetRegistryDate(InstallDateValueName, DateTime.UtcNow.AddDays(-1));
            SetCorruptedRegistryData("LastRunTime");
            
            var service = new LicensingService(TestPublicKeyXml);
            Assert.Equal(LicenseStatus.Trial, service.GetStatus());
        }

        [Fact]
        public void LicensingService_OrphanedLastRunTime_ExpiresTrial()
        {
            // Both anchors missing, but LastRunTime is present (indicates trial was reset)
            SetRegistryDate("LastRunTime", DateTime.UtcNow.AddDays(-5));
            
            var service = new LicensingService(TestPublicKeyXml);
            Assert.Equal(LicenseStatus.Expired, service.GetStatus());
        }

        [Fact]
        public void LicensingService_OrphanedLastRunTime_RepeatedGetStatus_RemainsExpired()
        {
            // Both anchors missing, but LastRunTime is present
            SetRegistryDate("LastRunTime", DateTime.UtcNow.AddDays(-5));
            
            var service = new LicensingService(TestPublicKeyXml);
            Assert.Equal(LicenseStatus.Expired, service.GetStatus());
            Assert.Equal(LicenseStatus.Expired, service.GetStatus());
            Assert.Equal(LicenseStatus.Expired, service.GetStatus());
        }

        [Fact]
        public void LicensingService_OrphanedLastRunTime_WithCorruptedLastRunTime_BehavesAsFirstRun()
        {
            // If LastRunTime is corrupted and both anchors are missing, it can't be read.
            // It safely defaults to First Run.
            SetCorruptedRegistryData("LastRunTime");
            
            var service = new LicensingService(TestPublicKeyXml);
            Assert.Equal(LicenseStatus.Trial, service.GetStatus());
            Assert.True(service.GetTrialDaysRemaining() > 29);
        }

        [Fact]
        public async Task LicensingService_MissingAnchors_DoesNotInterfereWithValidLicense()
        {
            var service = new LicensingService(TestPublicKeyXml);
            var hwid = service.GetHardwareId();
            var license = GenerateTestLicense(hwid);

            var activated = await service.ActivateAsync(license);
            Assert.True(activated);

            // Delete anchors to simulate wipe
#pragma warning disable CA1416
            using (var key = Registry.CurrentUser.OpenSubKey(RegKeyPath, true))
            {
                key?.DeleteValue(InstallDateValueName, false);
                key?.DeleteValue("LastRunTime", false);
            }
            using (var key = Registry.CurrentUser.OpenSubKey(SecondaryRegKeyPath, true))
            {
                key?.DeleteValue(SecondaryValueName, false);
            }
#pragma warning restore CA1416
            
            // Re-instantiate to test startup behavior
            var newService = new LicensingService(TestPublicKeyXml);
            Assert.Equal(LicenseStatus.Valid, newService.GetStatus());
        }

        [Fact]
        public async Task LicensingService_ClockTampering_DoesNotInterfereWithValidLicense()
        {
            // Even if the clock is rolled back, a valid paid license should remain valid
            SetRegistryDate("LastRunTime", DateTime.UtcNow.AddDays(1));

            var service = new LicensingService(TestPublicKeyXml);
            var hwid = service.GetHardwareId();
            var license = GenerateTestLicense(hwid);

            var activated = await service.ActivateAsync(license);
            Assert.True(activated);
            Assert.Equal(LicenseStatus.Valid, service.GetStatus());
        }

        [Fact]
        public async Task LicensingService_Activate_WithValidKey_SetsStatusToValid()
        {
            var service = new LicensingService(TestPublicKeyXml);
            var hwid = service.GetHardwareId();
            var license = GenerateTestLicense(hwid);

            var activated = await service.ActivateAsync(license);
            
            Assert.True(activated);
            Assert.Equal(LicenseStatus.Valid, service.GetStatus());
        }

        [Fact]
        public async Task LicensingService_Activate_WithInvalidSignature_Fails()
        {
            var service = new LicensingService(TestPublicKeyXml);
            var hwid = service.GetHardwareId();
            var license = GenerateTestLicense(hwid);

            // Tamper with the payload part
            var parts = license.Split('.');
            var fakePayload = Convert.ToBase64String(Encoding.UTF8.GetBytes("{\"Email\":\"hacked@hack.com\"}"));
            var tamperedLicense = $"{fakePayload}.{parts[1]}";

            var activated = await service.ActivateAsync(tamperedLicense);
            
            Assert.False(activated);
            Assert.Equal(LicenseStatus.Trial, service.GetStatus()); // Remains trial since activation failed
        }

        [Fact]
        public async Task LicensingService_Activate_WithWrongHardwareId_Fails()
        {
            var service = new LicensingService(TestPublicKeyXml);
            var wrongHwid = "SOME-OTHER-PC";
            var license = GenerateTestLicense(wrongHwid);

            var activated = await service.ActivateAsync(license);
            
            Assert.False(activated);
        }

        [Fact]
        public async Task LicensingService_Activate_WithWildcardHardwareId_Fails()
        {
            var service = new LicensingService(TestPublicKeyXml);
            var license = GenerateTestLicense("*"); // Wildcard

            var activated = await service.ActivateAsync(license);
            
            Assert.False(activated);
        }

        [Fact]
        public async Task LicensingService_Activate_WithExpiredLicense_Fails()
        {
            var service = new LicensingService(TestPublicKeyXml);
            var hwid = service.GetHardwareId();
            var license = GenerateTestLicense(hwid, expiryDate: DateTime.UtcNow.AddDays(-1));

            var activated = await service.ActivateAsync(license);
            
            Assert.False(activated);
        }

        [Fact]
        public async Task LicensingService_Activate_WithFutureIssueDate_Fails()
        {
            var service = new LicensingService(TestPublicKeyXml);
            var hwid = service.GetHardwareId();
            var license = GenerateTestLicense(hwid, issueDate: DateTime.UtcNow.AddDays(2)); // Beyond 1 day tolerance

            var activated = await service.ActivateAsync(license);
            
            Assert.False(activated);
        }

        [Fact]
        public async Task LicensingService_Activate_WithoutExpiryDate_FailsClosed_InAnnualModel()
        {
            var service = new LicensingService(TestPublicKeyXml);
            var hwid = service.GetHardwareId();
            var licenseWithoutExpiry = GenerateTestLicense(hwid, omitExpiry: true);

            var activated = await service.ActivateAsync(licenseWithoutExpiry);
            Assert.False(activated);
        }

        [Fact]
        public async Task LicensingService_Activate_ExceedingOneYearDuration_FailsClosed()
        {
            var service = new LicensingService(TestPublicKeyXml);
            var hwid = service.GetHardwareId();
            var issue = DateTime.UtcNow;
            var invalidMultiYearExpiry = issue.AddDays(400); // Exceeds 1-year annual limit
            var license = GenerateTestLicense(hwid, issueDate: issue, expiryDate: invalidMultiYearExpiry);

            var activated = await service.ActivateAsync(license);
            Assert.False(activated);
        }

        [Fact]
        public async Task LicensingService_Activate_WithExactOneYearLicense_Succeeds()
        {
            var service = new LicensingService(TestPublicKeyXml);
            var hwid = service.GetHardwareId();
            var issue = DateTime.UtcNow;
            var exactOneYearExpiry = issue.AddYears(1);
            var license = GenerateTestLicense(hwid, issueDate: issue, expiryDate: exactOneYearExpiry);

            var activated = await service.ActivateAsync(license);
            Assert.True(activated);
            Assert.Equal(LicenseStatus.Valid, service.GetStatus());
        }

        [Fact]
        public async Task LicensingService_Activate_MalformedInput_Fails()
        {
            var service = new LicensingService(TestPublicKeyXml);
            Assert.False(await service.ActivateAsync("not.a.base64.string"));
            Assert.False(await service.ActivateAsync("onlyonestring"));
            Assert.False(await service.ActivateAsync(""));
        }
    }
}
