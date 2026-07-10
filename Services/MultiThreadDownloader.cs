using System;
using System.Diagnostics;
using System.IO;
using System.Net;
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

            var headReq = (HttpWebRequest)WebRequest.Create(url);
            headReq.Method = "HEAD";
            headReq.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36";
            headReq.Timeout = 15000;

            try
            {
                using (var headResp = (HttpWebResponse)await Task.Factory.FromAsync(
                    headReq.BeginGetResponse, headReq.EndGetResponse, null))
                {
                    totalSize = headResp.ContentLength;
                    var acceptRanges = headResp.Headers["Accept-Ranges"];
                    supportsRange = !string.IsNullOrEmpty(acceptRanges) &&
                                    acceptRanges.IndexOf("bytes", StringComparison.OrdinalIgnoreCase) >= 0;

                    if (totalSize > 0 && !supportsRange)
                        supportsRange = true;
                }
            }
            catch
            {
                Logger.Info("HEAD 请求探测失败，尝试使用 GET Range 探测...");
                try
                {
                    var getRangeReq = (HttpWebRequest)WebRequest.Create(url);
                    getRangeReq.Method = "GET";
                    getRangeReq.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36";
                    getRangeReq.AddRange(0, 0);
                    getRangeReq.Timeout = 15000;

                    using (var getRangeResp = (HttpWebResponse)await Task.Factory.FromAsync(
                        getRangeReq.BeginGetResponse, getRangeReq.EndGetResponse, null))
                    {
                        if (getRangeResp.StatusCode == HttpStatusCode.PartialContent)
                        {
                            supportsRange = true;
                            var contentRange = getRangeResp.Headers["Content-Range"];
                            if (!string.IsNullOrEmpty(contentRange))
                            {
                                int slashIdx = contentRange.LastIndexOf('/');
                                if (slashIdx >= 0 && long.TryParse(contentRange.Substring(slashIdx + 1), out long parsedSize))
                                {
                                    totalSize = parsedSize;
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

            ServicePointManager.DefaultConnectionLimit = 512;
            ServicePointManager.Expect100Continue = false;

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

            using (var fs = new FileStream(savePath, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                fs.SetLength(totalSize);
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
                                var req = (HttpWebRequest)WebRequest.Create(url);
                                req.Method = "GET";
                                req.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36";
                                req.AddRange(seg.Start + seg.Downloaded, seg.End);
                                req.Timeout = 30000;
                                req.ReadWriteTimeout = 30000;

                                using (var resp = (HttpWebResponse)await Task.Factory.FromAsync(
                                    req.BeginGetResponse, req.EndGetResponse, null))
                                using (var stream = resp.GetResponseStream())
                                {
                                    var buffer = new byte[BufferSize];
                                    int bytesRead;

                                    while ((bytesRead = await stream.ReadAsync(buffer, 0, buffer.Length, ct)) > 0)
                                    {
                                        ct.ThrowIfCancellationRequested();
                                        
                                        lock (lockObj)
                                        {
                                            fs.Seek(seg.Start + seg.Downloaded, SeekOrigin.Begin);
                                            fs.Write(buffer, 0, bytesRead);
                                        }
                                        
                                        seg.Downloaded += bytesRead;

                                        lock (lockObj)
                                        {
                                            totalDownloaded += bytesRead;
                                            int pct = (int)(totalDownloaded * 100 / totalSize);
                                            progress?.Report(Math.Min(pct, 100));
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
            }
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
                using (var client = new WebClient())
                {
                    ct.Register(() => client.CancelAsync());
                    await client.DownloadFileTaskAsync(new Uri(Aria2ZipUrl), zipPath);
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
                try { if (File.Exists(zipPath)) File.Delete(zipPath); } catch {}
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

                using (ct.Register(() => { try { proc.Kill(); } catch {} }))
                {
                    var reader = proc.StandardOutput;
                    while (!reader.EndOfStream)
                    {
                        string line = await reader.ReadLineAsync();
                        if (line == null) break;

                        // 匹配类似 (2%), (99%) 的进度表示
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
            var req = (HttpWebRequest)WebRequest.Create(url);
            req.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36";
            req.Timeout = 30000;

            using (var resp = (HttpWebResponse)await Task.Factory.FromAsync(
                req.BeginGetResponse, req.EndGetResponse, null))
            {
                if (totalSize <= 0)
                {
                    totalSize = resp.ContentLength;
                }

                using (var stream = resp.GetResponseStream())
                using (var fs = new FileStream(savePath, FileMode.Create, FileAccess.Write))
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
