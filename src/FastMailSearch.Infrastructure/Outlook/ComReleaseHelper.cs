using System;
using System.Runtime.InteropServices;
using FastMailSearch.Infrastructure.Logging;

namespace FastMailSearch.Infrastructure.Outlook
{
    public static class ComReleaseHelper
    {
        public static void Release<T>(ref T? comObject) where T : class
        {
            if (comObject != null && Marshal.IsComObject(comObject))
            {
                try
                {
                    Marshal.ReleaseComObject(comObject);
                }
                catch (Exception ex)
                {
                    Logger.Warn($"Exception releasing COM object: {ex.Message}");
                }
                finally
                {
                    comObject = null;
                }
            }
            else
            {
                comObject = null;
            }
        }

        public static void FinalRelease<T>(ref T? comObject) where T : class
        {
            if (comObject != null && Marshal.IsComObject(comObject))
            {
                try
                {
                    Marshal.FinalReleaseComObject(comObject);
                }
                catch (Exception ex)
                {
                    Logger.Warn($"Exception final releasing COM object: {ex.Message}");
                }
                finally
                {
                    comObject = null;
                }
            }
            else
            {
                comObject = null;
            }
        }
    }
}
