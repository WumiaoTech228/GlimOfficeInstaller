using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Threading.Tasks;
using GOI.Helpers;

namespace GOI.Services
{
    public class DownloadService
    {
        private const string ODT_URL = "https://c2rsetup.officeapps.live.com/c2r/officeDeploymentTool/officedeploymenttool.exe";

        /// <summary>下载 ODT 并解压出 setup.exe，返回成功/失败</summary>
        public async Task<bool> DownloadODTAsync(IProgress<int> progressPercent = null)
        {
            var odtPath = Path.Combine(AppConfig.RootPath, "officedeploymenttool.exe");
            var setupPath = AppConfig.SetupPath;

            // 如果 setup.exe 已存在，跳过
            if (File.Exists(setupPath))
                return true;

            Logger.Info("开始下载 ODT: " + ODT_URL);

            using (var client = new WebClient())
            {
                client.DownloadProgressChanged += (s, e) =>
                    progressPercent?.Report(e.ProgressPercentage);

                await client.DownloadFileTaskAsync(new Uri(ODT_URL), odtPath);
            }

            // 解压 ODT
            Logger.Info("ODT 下载完成，开始解压...");
            var psi = new ProcessStartInfo(odtPath, $"/quiet /extract:\"{AppConfig.RootPath.TrimEnd('\\')}\"")
            {
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using (var proc = Process.Start(psi))
            {
                if (proc == null) return false;
                await Task.Run(() => proc.WaitForExit());
                Logger.Info($"ODT 解压完成，退出码: {proc.ExitCode}");
            }

            // 清理 odt 安装包
            try { if (File.Exists(odtPath)) File.Delete(odtPath); } catch { }

            return File.Exists(setupPath);
        }
    }
}
