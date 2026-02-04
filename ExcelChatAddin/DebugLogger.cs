using System;
using System.IO;
using System.Text;

namespace ExcelChatAddin
{
    internal static class DebugLogger
    {
        private static readonly object _lock = new object();
        private static readonly string _logPath;

        static DebugLogger()
        {
            try
            {
                var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelChatAddin");
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                _logPath = Path.Combine(dir, "log.txt");
            }
            catch
            {
                _logPath = Path.Combine(Path.GetTempPath(), "ExcelChatAddin_log.txt");
            }
        }

        public static void LogInfo(string msg)
        {
            WriteLine("INFO", msg);
        }

        public static void LogError(string msg)
        {
            WriteLine("ERROR", msg);
        }

        public static void LogException(Exception ex, string context = null)
        {
            try
            {
                var sb = new StringBuilder();
                if (!string.IsNullOrEmpty(context)) sb.AppendLine(context);
                sb.AppendLine(ex.ToString());
                WriteLine("EX", sb.ToString());
            }
            catch { }
        }

        private static void WriteLine(string level, string msg)
        {
            try
            {
                lock (_lock)
                {
                    File.AppendAllText(_logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {level}: {msg}{Environment.NewLine}");
                }
            }
            catch { }
        }
    }
}
