using System;
using System.Threading.Tasks;

namespace FastMailSearch.Core.Interfaces
{
    public enum LicenseStatus
    {
        Valid,
        Trial,
        Expired,
        InvalidKey
    }

    public interface ILicensingService
    {
        LicenseStatus GetStatus();
        int GetTrialDaysRemaining();
        int GetDaysUntilExpiry();
        DateTime? GetLicenseExpiryDate();
        Task<bool> ActivateAsync(string licenseKey);
        string GetHardwareId();
    }
}
