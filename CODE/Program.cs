using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace PrintSelector
{
    internal static class Program
    {
        [DllImport("user32.dll")]
        private static extern bool SetProcessDPIAware();

        [STAThread]
        private static void Main()
        {
            try { SetProcessDPIAware(); } catch { }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            Application.ThreadException += (s, e) =>
            {
                CrashLogger.Log("ui_thread", e.Exception);
            };
            AppDomain.CurrentDomain.UnhandledException += (s, e) =>
            {
                CrashLogger.Log("unhandled", e.ExceptionObject as Exception);
            };

            try
            {
                Application.Run(new MainForm());
            }
            catch (Exception ex)
            {
                CrashLogger.Log("application_startup", ex);
                throw;
            }
        }
    }
}
