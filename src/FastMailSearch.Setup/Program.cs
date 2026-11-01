using System;
using System.Windows.Forms;

namespace FastMailSearch.Setup
{
    public static class Program
    {
        [STAThread]
        public static int Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            return SetupInstaller.RunSetup(args);
        }
    }
}
