using System;
using System.IO;
using System.Reflection;
using System.Threading;
using Xunit;

namespace FastMailSearch.Core.Tests
{
    public class RibbonCallbackTests
    {
        private static string GetBinFolder()
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string addInPath = Path.Combine(baseDir, "FastMailSearch.OutlookAddIn.dll");
            if (File.Exists(addInPath)) return baseDir;

            string solDir = Path.GetFullPath(Path.Combine(baseDir, "..", "..", "..", "..", ".."));
#if NETFRAMEWORK
            string releasePath = Path.Combine(solDir, "src", "FastMailSearch.OutlookAddIn", "bin", "Release", "net48");
#else
            string releasePath = Path.Combine(solDir, "src", "FastMailSearch.UI", "bin", "Release", "net8.0-windows");
            if (!Directory.Exists(releasePath)) releasePath = Path.Combine(solDir, "src", "FastMailSearch.OutlookAddIn", "bin", "Release", "net48");
#endif
            if (Directory.Exists(releasePath)) return releasePath;

            return baseDir;
        }

        [Fact]
        public void ThisAddIn_Exposes_Ribbon_Callback_Methods()
        {
            string binFolder = GetBinFolder();
            string dllPath = Path.Combine(binFolder, "FastMailSearch.OutlookAddIn.dll");

            Assembly addInAsm;
            if (File.Exists(dllPath))
            {
                addInAsm = Assembly.LoadFrom(dllPath);
            }
            else
            {
                addInAsm = Assembly.Load("FastMailSearch.OutlookAddIn");
            }

            var type = addInAsm.GetType("FastMailSearch.OutlookAddIn.ThisAddIn");
            Assert.NotNull(type);

            var clickMethod = type.GetMethod("OnOpenSearchWindowClicked");
            Assert.NotNull(clickMethod);
            Assert.True(clickMethod.IsPublic);

            var loadMethod = type.GetMethod("Ribbon_Load");
            Assert.NotNull(loadMethod);
            Assert.True(loadMethod.IsPublic);

            var customUiMethod = type.GetMethod("GetCustomUI");
            Assert.NotNull(customUiMethod);
            Assert.True(customUiMethod.IsPublic);
        }

    }
}
