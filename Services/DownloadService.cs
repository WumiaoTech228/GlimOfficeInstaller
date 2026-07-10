using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using GOI.Helpers;

namespace GOI.Services
{
    public class DownloadService
    {
        private const string ODT_URL = "https://officecdn.microsoft.com/pr/wsus/setup.exe";
        private static readonly HttpClient _httpClient;

        static DownloadService()
        {
            var handler = new HttpClientHandler();
            try
            {
                handler.UseProxy = true;
            }
            catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in DownloadService.cs at static constructor", ex_captured); }
            _httpClient = new HttpClient(handler);
            _httpClient.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36");
            _httpClient.Timeout = TimeSpan.FromSeconds(60);
        }

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
                using (var response = await _httpClient.GetAsync(ODT_URL, HttpCompletionOption.ResponseHeadersRead))
                {
                    response.EnsureSuccessStatusCode();
                    var totalBytes = response.Content.Headers.ContentLength ?? -1L;

                    using (var contentStream = await response.Content.ReadAsStreamAsync())
                    using (var fileStream = new FileStream(setupPath, FileMode.Create, FileAccess.Write, FileShare.None, 8192, true))
                    {
                        var buffer = new byte[8192];
                        long totalRead = 0L;
                        int bytesRead;
                        while ((bytesRead = await contentStream.ReadAsync(buffer, 0, buffer.Length)) > 0)
                        {
                            await fileStream.WriteAsync(buffer, 0, bytesRead);
                            totalRead += bytesRead;
                            if (totalBytes != -1L)
                            {
                                int percentage = (int)((double)totalRead / totalBytes * 100);
                                progressPercent?.Report(percentage);
                            }
                        }
                    }
                }

                Logger.Info("Setup.exe 下载完成。");
            }
            catch (Exception ex)
            {
                Logger.Error("下载 Setup.exe 失败", ex);
                try { if (File.Exists(setupPath)) File.Delete(setupPath); } catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in DownloadService.cs at DownloadODTAsync", ex_captured); }
                return false;
            }

            return File.Exists(setupPath);
        }
    }
}
