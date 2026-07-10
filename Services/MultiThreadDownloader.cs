using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using GOI.Helpers;
using SharpCompress.Readers;
using SharpCompress.Common;

namespace GOI.Services
{
    /// <summary>
    /// 多线程分段下载器。
    /// 优先使用 aria2c 进行高速下载，若不可用则回退至内置的多线程分段/单线程下载引擎。
    /// </summary>
    public class MultiThreadDownloader
    {
        private const int DefaultThreadCount = 8;
        private const int BufferSize = 81920; // 80KB
        private const string Aria2ZipUrl = "https://ghproxy.net/https://github.com/aria2/aria2/releases/download/release-1.37.0/aria2-1.37.0-win-64bit-build1.zip";
        private static readonly HttpClient _httpClient;

        static MultiThreadDownloader()
        {
            var handler = new HttpClientHandler();
            try
            {
                handler.UseProxy = true;
            }
            catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in MultiThreadDownloader.cs static constructor", ex_captured); }
            _httpClient = new HttpClient(handler);
            _httpClient.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36");
            _httpClient.Timeout = TimeSpan.FromSeconds(60);
        }

        /// <summary>
        /// 下载文件到指定路径。
        /// </summary>
        public async Task DownloadAsync(
            string url, string savePath,
            IProgress<int> progress = null,
            int threadCount = DefaultThreadCount,
            CancellationToken ct = default)
        {
            // 尝试使用 aria2c 高速下载
            try
            {
                string aria2Path = await GetOrDownloadAria2Async(progress, ct);
                if (!string.IsNullOrEmpty(aria2Path) && File.Exists(aria2Path))
                {
                    Logger.Info("检测到 aria2c，开始通过 aria2c 执行极速下载...");
                    bool success = await DownloadWithAria2Async(aria2Path, url, savePath, progress, ct);
                    if (success)
                    {
                        Logger.Info("通过 aria2c 极速下载完成！");
                        return;
                    }
                    Logger.Warn("aria2c 下载异常退出，开始回退至内置下载引擎...");
                }
            }
            catch (Exception ex)
            {
                Logger.Warn("启动 aria2c 下载失败，准备回退至内置下载引擎: " + ex.Message);
            }

            // ── 内置下载引擎回退 ──
            // 1. HEAD 请求获取文件大小和 Range 支持
            long totalSize = -1;
            bool supportsRange = false;

            try
            {
                using (var request = new HttpRequestMessage(HttpMethod.Head, url))
                using (var response = await _httpClient.SendAsync(request, ct))
                {
                    if (response.IsSuccessStatusCode)
                    {
                        totalSize = response.Content.Headers.ContentLength ?? -1;
                        var acceptRanges = response.Headers.AcceptRanges;
                        supportsRange = acceptRanges != null && acceptRanges.Contains("bytes");
                        if (totalSize > 0 && !supportsRange)
                        {
                            supportsRange = true;
                        }
                    }
                }
            }
            catch
            {
                Logger.Info("HEAD 请求探测失败，尝试使用 GET Range 探测...");
                try
                {
                    using (var request = new HttpRequestMessage(HttpMethod.Get, url))
                    {
                        request.Headers.Range = new RangeHeaderValue(0, 0);
                        using (var response = await _httpClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, ct))
                        {
                            if (response.StatusCode == HttpStatusCode.PartialContent)
                            {
                                supportsRange = true;
                                var contentRange = response.Content.Headers.ContentRange;
                                if (contentRange != null && contentRange.Length.HasValue)
                                {
                                    totalSize = contentRange.Length.Value;
                                    Logger.Info($"GET Range 探测成功！文件大小: {totalSize / 1024 / 1024}MB");
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.Warn("GET Range 探测失败，回退单线程: " + ex.Message);
                    supportsRange = false;
                }
            }

            if (!supportsRange || totalSize <= 0 || totalSize < 2 * 1024 * 1024)
            {
                await SingleThreadDownloadAsync(url, savePath, totalSize, progress, ct);
                return;
            }

            Logger.Info($"内置多线程下载: {threadCount} 线程, 大小: {totalSize / 1024 / 1024}MB");

            long segmentSize = totalSize / threadCount;
            var segments = new SegmentInfo[threadCount];
            for (int i = 0; i < threadCount; i++)
            {
                segments[i] = new SegmentInfo
                {
                    Index = i,
                    Start = i * segmentSize,
                    End = (i == threadCount - 1) ? totalSize - 1 : (i + 1) * segmentSize - 1,
                    Downloaded = 0
                };
            }

            // 1. 预先设置目标文件长度
            using (var fsPre = new FileStream(savePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
            {
                fsPre.SetLength(totalSize);
            }

            long totalDownloaded = 0;
            var lockObj = new object();

            var tasks = new Task[threadCount];
            for (int i = 0; i < threadCount; i++)
            {
                var seg = segments[i];
                tasks[i] = Task.Run(async () =>
                {
                    int retries = 0;
                    const int maxRetries = 3;

                    while (retries < maxRetries)
                    {
                        try
                        {
                            using (var request = new HttpRequestMessage(HttpMethod.Get, url))
                            {
                                request.Headers.Range = new RangeHeaderValue(seg.Start + seg.Downloaded, seg.End);
                                using (var response = await _httpClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, ct))
                                {
                                    response.EnsureSuccessStatusCode();
                                    using (var stream = await response.Content.ReadAsStreamAsync())
                                    // 2. 每个分段下载线程打开独立的 FileStream 句柄，通过 FileShare.ReadWrite 并行无锁写入各自的分段区间
                                    using (var fs = new FileStream(savePath, FileMode.Open, FileAccess.Write, FileShare.ReadWrite, 8192, true))
                                    {
                                        fs.Position = seg.Start + seg.Downloaded;
                                        var buffer = new byte[BufferSize];
                                        int bytesRead;

                                        while ((bytesRead = await stream.ReadAsync(buffer, 0, buffer.Length, ct)) > 0)
                                        {
                                            ct.ThrowIfCancellationRequested();
                                            await fs.WriteAsync(buffer, 0, bytesRead, ct);
                                            seg.Downloaded += bytesRead;

                                            lock (lockObj)
                                            {
                                                totalDownloaded += bytesRead;
                                                int pct = (int)(totalDownloaded * 100 / totalSize);
                                                progress?.Report(Math.Min(pct, 100));
                                            }
                                        }
                                    }
                                }
                            }
                            break;
                        }
                        catch (OperationCanceledException) { throw; }
                        catch (Exception ex)
                        {
                            retries++;
                            if (retries >= maxRetries)
                            {
                                Logger.Error($"分段 {seg.Index} 下载失败", ex);
                                throw;
                            }
                            Logger.Info($"分段 {seg.Index} 下载失败，重试 {retries}...");
                            await Task.Delay(1000 * retries, ct);
                        }
                    }
                }, ct);
            }

            await Task.WhenAll(tasks);
            progress?.Report(100);
            Logger.Info("内置多线程下载完成");
        }

        /// <summary>获取或下载 aria2c.exe</summary>
        private async Task<string> GetOrDownloadAria2Async(IProgress<int> progress, CancellationToken ct)
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string exePath = Path.Combine(baseDir, "aria2c.exe");
            if (File.Exists(exePath)) return exePath;

            string configPath = Path.Combine(AppConfig.RootPath, "aria2c.exe");
            if (File.Exists(configPath)) return configPath;

            Logger.Info("本地未检测到 aria2c，开始通过代理拉取官方稳定版...");
            progress?.Report(0);
            string zipPath = Path.Combine(AppConfig.RootPath, "aria2c.zip");
            try
            {
                using (var response = await _httpClient.GetAsync(Aria2ZipUrl, ct))
                {
                    response.EnsureSuccessStatusCode();
                    using (var fs = File.Create(zipPath))
                    {
                        await response.Content.CopyToAsync(fs);
                    }
                }

                Logger.Info("aria2c.zip 下载完成，正在解压提取...");
                using (var fs = File.OpenRead(zipPath))
                using (var reader = ReaderFactory.OpenReader(fs))
                {
                    while (reader.MoveToNextEntry())
                    {
                        if (!reader.Entry.IsDirectory && reader.Entry.Key.EndsWith("aria2c.exe", StringComparison.OrdinalIgnoreCase))
                        {
                            reader.WriteEntryToDirectory(AppConfig.RootPath, new ExtractionOptions
                            {
                                ExtractFullPath = false,
                                Overwrite = true
                            });
                            break;
                        }
                    }
                }
                Logger.Info("aria2c.exe 提取成功！");
                return configPath;
            }
            catch (Exception ex)
            {
                Logger.Warn("下载或解压 aria2c 失败，回退至传统下载: " + ex.Message);
                return null;
            }
            finally
            {
                try { if (File.Exists(zipPath)) File.Delete(zipPath); } catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in MultiThreadDownloader.cs at GetOrDownloadAria2Async finally", ex_captured); }
            }
        }

        /// <summary>调用 aria2c 下载</summary>
        private async Task<bool> DownloadWithAria2Async(
            string aria2Path, string url, string savePath, IProgress<int> progress, CancellationToken ct)
        {
            string dir = Path.GetDirectoryName(savePath);
            string file = Path.GetFileName(savePath);

            string proxyArg = "";
            try
            {
                var systemProxy = WebRequest.GetSystemWebProxy();
                var destinationUri = new Uri(url);
                var proxyUri = systemProxy.GetProxy(destinationUri);
                if (proxyUri != null && proxyUri != destinationUri)
                {
                    proxyArg = $" --all-proxy=\"{proxyUri}\"";
                    Logger.Info($"Aria2c 已自动侦测并挂载系统代理: {proxyUri}");
                }
            }
            catch (Exception ex)
            {
                Logger.Warn("获取系统代理失败: " + ex.Message);
            }

            var psi = new ProcessStartInfo(aria2Path, $"--max-connection-per-server=16 --split=16 --min-split-size=1M --file-allocation=none --disk-cache=32M --disable-ipv6=true --summary-interval=1 --console-log-level=info{proxyArg} --user-agent=\"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36\" -d \"{dir}\" -o \"{file}\" \"{url}\"")
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };

            using (var proc = new Process { StartInfo = psi })
            {
                proc.Start();

                using (ct.Register(() => { try { proc.Kill(); } catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in MultiThreadDownloader.cs at DownloadWithAria2Async cancel registration", ex_captured); } }))
                {
                    var reader = proc.StandardOutput;
                    while (!reader.EndOfStream)
                    {
                        string line = await reader.ReadLineAsync();
                        if (line == null) break;

                        var match = Regex.Match(line, @"\((\d+)%\)");
                        if (match.Success)
                        {
                            if (int.TryParse(match.Groups[1].Value, out int pct))
                            {
                                progress?.Report(pct);
                            }
                        }
                    }

                    await Task.Run(() => proc.WaitForExit());
                }

                return proc.ExitCode == 0;
            }
        }

        private async Task SingleThreadDownloadAsync(
            string url, string savePath, long totalSize,
            IProgress<int> progress, CancellationToken ct)
        {
            Logger.Info("回退单线程下载");
            using (var response = await _httpClient.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, ct))
            {
                response.EnsureSuccessStatusCode();
                if (totalSize <= 0)
                {
                    totalSize = response.Content.Headers.ContentLength ?? -1;
                }

                using (var stream = await response.Content.ReadAsStreamAsync())
                using (var fs = new FileStream(savePath, FileMode.Create, FileAccess.Write, FileShare.None, 8192, true))
                {
                    var buffer = new byte[BufferSize];
                    long downloaded = 0;
                    int bytesRead;

                    while ((bytesRead = await stream.ReadAsync(buffer, 0, buffer.Length, ct)) > 0)
                    {
                        ct.ThrowIfCancellationRequested();
                        await fs.WriteAsync(buffer, 0, bytesRead, ct);
                        downloaded += bytesRead;

                        if (totalSize > 0)
                        {
                            int pct = (int)(downloaded * 100 / totalSize);
                            progress?.Report(Math.Min(pct, 100));
                        }
                    }
                }
            }
            progress?.Report(100);
        }

        private class SegmentInfo
        {
            public int Index;
            public long Start;
            public long End;
            public long Downloaded;
        }
    }
}
