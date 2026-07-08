using System;
using System.IO;
using System.Net;
using System.Threading;
using System.Threading.Tasks;
using GOI.Helpers;

namespace GOI.Services
{
    /// <summary>
    /// 多线程分段下载器（类似 IDM 的实现原理）。
    /// 通过 HTTP Range 请求将文件拆分为多个分段并发下载，最终合并。
    /// </summary>
    public class MultiThreadDownloader
    {
        private const int DefaultThreadCount = 8;
        private const int BufferSize = 81920; // 80KB

        /// <summary>
        /// 多线程下载文件到指定路径。
        /// </summary>
        /// <param name="url">下载 URL</param>
        /// <param name="savePath">保存路径</param>
        /// <param name="progress">进度回调 (0-100)</param>
        /// <param name="threadCount">并发线程数</param>
        /// <param name="ct">取消令牌</param>
        public async Task DownloadAsync(
            string url, string savePath,
            IProgress<int> progress = null,
            int threadCount = DefaultThreadCount,
            CancellationToken ct = default)
        {
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

                    // 某些 CDN 即使不返回 Accept-Ranges 头也支持 Range
                    // 如果文件大小已知，我们尝试 Range 请求
                    if (totalSize > 0 && !supportsRange)
                        supportsRange = true;
                }
            }
            catch
            {
                // HEAD 失败，回退单线程
                supportsRange = false;
            }

            // 如果不支持 Range 或文件太小（< 2MB），走单线程
            if (!supportsRange || totalSize <= 0 || totalSize < 2 * 1024 * 1024)
            {
                await SingleThreadDownloadAsync(url, savePath, totalSize, progress, ct);
                return;
            }

            // 2. 提升 .NET HttpWebRequest 默认的最大并发连接限制 (默认为 2 极其缓慢)
            ServicePointManager.DefaultConnectionLimit = 512;
            ServicePointManager.Expect100Continue = false;

            Logger.Info($"多线程下载: {threadCount} 线程, 文件大小: {totalSize / 1024 / 1024}MB");

            // 3. 计算每个分段的范围
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

            // 4. 创建目标文件（预分配大小）
            using (var fs = new FileStream(savePath, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                fs.SetLength(totalSize);
            }

            // 5. 并发下载所有分段
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
                            using (var fs = new FileStream(savePath, FileMode.Open, FileAccess.Write, FileShare.ReadWrite))
                            {
                                var buffer = new byte[BufferSize];
                                int bytesRead;

                                while ((bytesRead = await stream.ReadAsync(buffer, 0, buffer.Length, ct)) > 0)
                                {
                                    ct.ThrowIfCancellationRequested();
                                    
                                    // 线程安全定位写入
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

                            // 分段下载完成
                            break;
                        }
                        catch (OperationCanceledException) { throw; }
                        catch (Exception ex)
                        {
                            retries++;
                            if (retries >= maxRetries)
                            {
                                Logger.Error($"分段 {seg.Index} 下载失败（已重试 {maxRetries} 次）", ex);
                                throw;
                            }
                            Logger.Info($"分段 {seg.Index} 下载失败，第 {retries} 次重试...");
                            await Task.Delay(1000 * retries, ct);
                        }
                    }
                }, ct);
            }

            await Task.WhenAll(tasks);
            progress?.Report(100);
            Logger.Info("多线程下载完成");
        }

        /// <summary>单线程回退下载</summary>
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
