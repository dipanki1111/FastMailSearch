using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using FastMailSearch.Infrastructure.Logging;
using OfficeCore = Microsoft.Office.Core;

namespace FastMailSearch.OutlookAddIn.Ribbon
{
    [ComVisible(true)]
    [Guid("E3A2F9B8-6C7D-4F1E-9B5A-3A8B1C2D4E5F")]
    public class FastMailRibbon : OfficeCore.IRibbonExtensibility
    {
        private OfficeCore.IRibbonUI? _ribbon;

        public string GetCustomUI(string ribbonID)
        {
            Logger.Info($"[RIBBON] FastMailRibbon.GetCustomUI requested for RibbonID='{ribbonID}'");
            if (ribbonID == "Microsoft.Outlook.Explorer")
            {
                return GetResourceText("FastMailSearch.OutlookAddIn.Ribbon.FastMailRibbon.xml");
            }
            return string.Empty;
        }

        [ComVisible(true)]
        public void Ribbon_Load(OfficeCore.IRibbonUI ribbonUI)
        {
            Logger.Info("[RIBBON] Ribbon_Load callback invoked by Outlook on FastMailRibbon.");
            _ribbon = ribbonUI;
        }

        [ComVisible(true)]
        public void OnOpenSearchWindowClicked(OfficeCore.IRibbonControl control)
        {
            Logger.Info("[RIBBON] FastMail Search button clicked");
            ThisAddIn.LaunchOrActivateUIProcess();
        }

        [ComVisible(true)]
        public void OnBackfillActiveStoreClicked(OfficeCore.IRibbonControl control)
        {
            Logger.Info("[RIBBON] Backfill POP Mailbox button clicked");
            ThisAddIn.Instance?.StartInProcessBackfill();
        }

        private static string GetResourceText(string resourceName)
        {
            var asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return string.Empty;
        }
    }
}
