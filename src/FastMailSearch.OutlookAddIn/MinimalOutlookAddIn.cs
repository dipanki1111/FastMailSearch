using System;
using System.IO;
using System.Runtime.InteropServices;
using Extensibility;

namespace FastMailSearch.OutlookAddIn
{
    [ComVisible(true)]
    [Guid("B7C4D6A1-3D5E-4F91-8A20-7C6B9E123456")]
    [ProgId("FastMailSearch.MinimalAddIn")]
    public class MinimalOutlookAddIn : IDTExtensibility2
    {
        public MinimalOutlookAddIn()
        {
            try
            {
                Log("MinimalOutlookAddIn constructor invoked via COM.");
            }
            catch { }
        }

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                Log($"MinimalOutlookAddIn OnConnection reached successfully! (ConnectMode={ConnectMode})");
            }
            catch (Exception ex)
            {
                Log($"MinimalOutlookAddIn OnConnection exception: {ex.Message}");
            }
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            Log("MinimalOutlookAddIn OnDisconnection reached.");
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        private static void Log(string message)
        {
            try
            {
                string logDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "FastMailSearch", "Logs");
                if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                string entry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [MINIMAL_ADDIN] {message}{Environment.NewLine}";
                File.AppendAllText(Path.Combine(logDir, "MinimalOutlookAddIn.log"), entry);
                File.AppendAllText(Path.Combine(logDir, "FastMailSearch.log"), entry);
            }
            catch { }
        }
    }
}
