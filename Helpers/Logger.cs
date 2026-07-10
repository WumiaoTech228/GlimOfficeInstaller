using System;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace GOI.Helpers
{
    public static class Logger
    {
        private static readonly object _lock = new object();
        private static string _logFile;

        private static string LogFile
        {
            get
            {
                if (_logFile == null)
                    _logFile = Path.Combine(AppConfig.LogPath,
                        $"install_{DateTime.Now:yyyyMMdd_HHmmss}.log");
                return _logFile;
            }
        }
        public static string LogFilePath => LogFile;
        public static void Info(string msg) => Log("INFO", msg);
        public static void Warn(string msg) => Log("WARN", msg);
        public static void Error(string msg, Exception ex = null)
        {
            var full = ex == null ? msg : $"{msg} | {ex.GetType().Name}: {ex.Message}\n{ex.StackTrace}";
            Log("ERROR", full);
        }

        private static void Log(string level, string msg)
        {
            var line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [{level}] {msg}";
            Debug.WriteLine(line);

            lock (_lock)
            {
                try
                {
                    if (Directory.Exists(AppConfig.LogPath))
                        File.AppendAllText(LogFile, line + Environment.NewLine, Encoding.UTF8);
                }
                catch { /* 写日志失败不能影响主流程 */ }
            }
        }

        public static string ReadAll()
        {
            try
            {
                return File.Exists(LogFile) ? File.ReadAllText(LogFile, Encoding.UTF8) : "暂无日志。";
            }
            catch (Exception ex)
            {
                return "读取日志失败: " + ex.Message;
            }
        }
    }
}
