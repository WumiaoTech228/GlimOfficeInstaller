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
        private const string ODT_URL = "https://officecdn.microsoft.com/pr/wsus/setup.exe";

        /// <summary>直接从 Office CDN 下载 setup.exe，返回成功/失败</summary>
        public async Task<bool> DownloadODTAsync(IProgress<int> progressPercent = null)
        {
            var setupPath = AppConfig.SetupPath;

            // 如果 setup.exe 已存在，跳过
            if (File.Exists(setupPath))
                return true;

            Logger.Info("开始下载 Setup.exe: " + ODT_URL);

            try
            {
                using (var client = new WebClient())
                {
                    client.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36");
                    client.DownloadProgressChanged += (s, e) =>
                        progressPercent?.Report(e.ProgressPercentage);

                    await client.DownloadFileTaskAsync(new Uri(ODT_URL), setupPath);
                }

                Logger.Info("Setup.exe 下载完成。");
            }
            catch (Exception ex)
            {
                Logger.Error("下载 Setup.exe 失败", ex);
                try { if (File.Exists(setupPath)) File.Delete(setupPath); } catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in DownloadService.cs at UnknownMethod", ex_captured); }
                return false;
            }

            return File.Exists(setupPath);
        }
    }
}
