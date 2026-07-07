using System;
using System.IO;

namespace GOI.Helpers
{
    /// <summary>全局路径配置</summary>
    public static class AppConfig
    {
        public static readonly string RootPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "GOI_Data");
        public static readonly string LogPath = Path.Combine(RootPath, "logs");
        public static readonly string ToolsPath = Path.Combine(RootPath, "tools");
        public static readonly string SetupPath = Path.Combine(RootPath, "setup.exe");
        public static readonly string XmlConfigPath = Path.Combine(RootPath, "configuration.xml");

        public static void Initialize()
        {
            EnsureDir(RootPath);
            EnsureDir(LogPath);
            EnsureDir(ToolsPath);
        }

        private static void EnsureDir(string path)
        {
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
        }
    }
}
