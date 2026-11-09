using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace PrintSelector
{
    internal static class CrashLogger
    {
        private static string GetLogsDir()
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string logsDir = Path.Combine(baseDir, "LOGS");

            try
            {
                Directory.CreateDirectory(logsDir);
                // Verify write access by touching a temp file path check
                string probe = Path.Combine(logsDir, ".write_test");
                File.WriteAllText(probe, "");
                File.Delete(probe);
                return logsDir;
            }
            catch
            {
                // Fall back to AppData if running from a protected folder like Program Files
                string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                if (string.IsNullOrEmpty(appData))
                    appData = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);

                logsDir = Path.Combine(appData, "Print Selector", "LOGS");
                Directory.CreateDirectory(logsDir);
                return logsDir;
            }
        }

        public static string Log(
            string operation,
            Exception ex,
            string filePath = null,
            string printerName = null,
            Dictionary<string, string> extraContext = null)
        {
            try
            {
                DateTime now = DateTime.Now;
                string fileStamp = now.ToString("yyyyMMdd_HHmmss_fff");
                string humanStamp = now.ToString("yyyy-MM-dd HH:mm:ss.fff");

                string fileName = "N/A";
                string fileExt = "N/A";
                if (!string.IsNullOrEmpty(filePath))
                {
                    fileName = Path.GetFileName(filePath);
                    fileExt = Path.GetExtension(filePath).ToLowerInvariant();
                    if (string.IsNullOrEmpty(fileExt)) fileExt = "(no extension)";
                }

                Exception root = ex != null ? (ex.GetBaseException() ?? ex) : null;
                string rootCause = root != null
                    ? string.Format("{0}: {1}", root.GetType().FullName, root.Message)
                    : "Unknown error";

                var sb = new StringBuilder();
                sb.AppendLine(new string('=', 70));
                sb.AppendLine("  PRINT SELECTOR - CRASH REPORT");
                sb.AppendLine(new string('=', 70));
                sb.AppendLine();
                sb.AppendLine("  Timestamp      : " + humanStamp);
                sb.AppendLine("  Operation      : " + operation);
                sb.AppendLine("  File Name      : " + fileName);
                sb.AppendLine("  File Type      : " + fileExt);
                sb.AppendLine("  Printer        : " + (string.IsNullOrEmpty(printerName) ? "N/A" : printerName));
                sb.AppendLine();
                sb.AppendLine(new string('-', 70));
                sb.AppendLine("  ROOT CAUSE");
                sb.AppendLine(new string('-', 70));
                sb.AppendLine("  " + rootCause);
                sb.AppendLine();
                sb.AppendLine(new string('-', 70));
                sb.AppendLine("  FULL TRACEBACK");
                sb.AppendLine(new string('-', 70));
                sb.AppendLine(ex != null ? ex.ToString() : "N/A");

                if (extraContext != null && extraContext.Count > 0)
                {
                    sb.AppendLine(new string('-', 70));
                    sb.AppendLine("  ADDITIONAL CONTEXT");
                    sb.AppendLine(new string('-', 70));
                    foreach (var kv in extraContext)
                        sb.AppendLine(string.Format("  {0}: {1}", kv.Key, kv.Value));
                    sb.AppendLine();
                }

                sb.AppendLine(new string('-', 70));
                sb.AppendLine("  SYSTEM INFO");
                sb.AppendLine(new string('-', 70));
                sb.AppendLine("  CLR Version    : " + Environment.Version);
                sb.AppendLine("  OS Version     : " + Environment.OSVersion);
                sb.AppendLine("  64-Bit Process : " + Environment.Is64BitProcess);
                sb.AppendLine();
                sb.AppendLine(new string('=', 70));
                sb.AppendLine("  END OF CRASH REPORT");
                sb.AppendLine(new string('=', 70));

                string logPath = Path.Combine(GetLogsDir(), string.Format("crash_{0}_{1}.log", fileStamp, operation));
                File.WriteAllText(logPath, sb.ToString(), Encoding.UTF8);
                return logPath;
            }
            catch
            {
                return "";
            }
        }
    }
}
