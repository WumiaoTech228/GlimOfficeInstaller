using System;
using System.IO;
using System.Reflection;

namespace GOI.Helpers
{
    public static class ResourceHelper
    {
        private static readonly (string, string)[] Scripts =
        {
            ("GOI.Resources.Ohook_Activation_AIO.cmd", "Ohook_Activation_AIO.cmd"),
        };

        public static void ExtractAllScripts()
        {
            if (!Directory.Exists(AppConfig.ToolsPath))
                Directory.CreateDirectory(AppConfig.ToolsPath);

            var asm = Assembly.GetExecutingAssembly();
            foreach (var (resName, fileName) in Scripts)
            {
                var dest = Path.Combine(AppConfig.ToolsPath, fileName);
                try
                {
                    using var stream = asm.GetManifestResourceStream(resName);
                    if (stream == null) { Logger.Warn($"资源未找到: {resName}"); continue; }
                    using var fs = new FileStream(dest, FileMode.Create, FileAccess.Write);
                    stream.CopyTo(fs);
                }
                catch (Exception ex) { Logger.Error($"提取 {fileName} 失败", ex); }
            }
            Logger.Info("MAS 脚本已提取完毕。");
        }

        public static string GetScriptPath(string fileName)
        {
            var p = Path.Combine(AppConfig.ToolsPath, fileName);
            return File.Exists(p) ? p : null;
        }
    }
}
