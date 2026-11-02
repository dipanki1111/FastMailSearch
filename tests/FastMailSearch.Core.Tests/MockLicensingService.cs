using System;
using System.Threading.Tasks;
using FastMailSearch.Core.Interfaces;

namespace FastMailSearch.Core.Tests
{
    public class MockLicensingService : ILicensingService
    {
        public LicenseStatus Status { get; set; } = LicenseStatus.Valid;
        public Func<LicenseStatus>? StatusFunc { get; set; }
        public Exception? ExceptionToThrow { get; set; }
        public int TrialDaysRemaining { get; set; } = 30;
        public bool ActivateResult { get; set; } = true;
        public string HardwareId { get; set; } = "MOCK-HWID-12345";
        public int GetStatusCallCount { get; private set; }

        public MockLicensingService(LicenseStatus status = LicenseStatus.Valid)
        {
            Status = status;
        }

        public LicenseStatus GetStatus()
        {
            GetStatusCallCount++;
            if (ExceptionToThrow != null) throw ExceptionToThrow;
            if (StatusFunc != null) return StatusFunc();
            return Status;
        }

        public int DaysUntilExpiry { get; set; } = 365;
        public DateTime? LicenseExpiryDate { get; set; } = DateTime.UtcNow.AddYears(1);

        public int GetTrialDaysRemaining() => TrialDaysRemaining;
        public int GetDaysUntilExpiry() => DaysUntilExpiry;
        public DateTime? GetLicenseExpiryDate() => LicenseExpiryDate;

        public Task<bool> ActivateAsync(string licenseKey) => Task.FromResult(ActivateResult);

        public string GetHardwareId() => HardwareId;
    }
}
