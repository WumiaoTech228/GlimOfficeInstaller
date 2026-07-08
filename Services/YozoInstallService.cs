using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading;
using System.Threading.Tasks;
using GOI.Helpers;
using GOI.Models;
using SharpCompress.Archives;
using SharpCompress.Common;
using SharpCompress.Readers;

namespace GOI.Services
{
    public class YozoInstallService
    {
        private const string YozoUrl = "https://dl.yozosoft.com/yozo/project/file/20251224_131531_158622/9.0.6589.101ZH.S1.rar";
        private const string YozoFileName = "YozoOffice_Setup.rar";

        /// <summary>下载并静默安装永中Office，以伪进度条报告进度</summary>
        public async Task<bool> InstallAsync(
            YozoVersion version,
            IProgress<string> phaseText,
            IProgress<int> progressPercent,
            CancellationToken ct = default)
        {
            string localPath = Path.Combine(AppConfig.RootPath, YozoFileName);
            string extractPath = Path.Combine(AppConfig.RootPath, "YozoExtract");

            // ── 阶段 1：下载 ──
            phaseText.Report("正在下载 永中Office 官方压缩包...");
            try
            {
                using (var client = new WebClient())
                {
                    client.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36");
                    client.DownloadProgressChanged += (s, e) =>
                        progressPercent?.Report(e.ProgressPercentage / 2); // 下载占 0-50%
                    await client.DownloadFileTaskAsync(new Uri(YozoUrl), localPath);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("下载 永中Office 失败", ex);
                phaseText.Report("下载失败，请检查网络连接。");
                try { if (File.Exists(localPath)) File.Delete(localPath); } catch { }
                return false;
            }

            // ── 阶段 2：解压 RAR ──
            phaseText.Report("正在解压 永中Office 安装包...");
            progressPercent?.Report(52);
            try
            {
                if (Directory.Exists(extractPath)) Directory.Delete(extractPath, true);
                Directory.CreateDirectory(extractPath);

                await Task.Run(() =>
                {
                    using (var fs = File.OpenRead(localPath))
                    using (var reader = ReaderFactory.OpenReader(fs))
                    {
                        while (reader.MoveToNextEntry())
                        {
                            if (!reader.Entry.IsDirectory)
                            {
                                reader.WriteEntryToDirectory(extractPath, new ExtractionOptions
                                {
                                    ExtractFullPath = true,
                                    Overwrite = true
                                });
                            }
                        }
                    }
                }, ct);
            }
            catch (Exception ex)
            {
                Logger.Error("解压 永中Office 失败", ex);
                phaseText.Report("安装包解压失败: " + ex.Message);
                CleanTempFiles(localPath, extractPath);
                return false;
            }

            // ── 阶段 3：定位并静默安装（/S 参数）──
            phaseText.Report("正在静默安装 永中Office 个人版...");
            progressPercent?.Report(60);

            try
            {
                // 在解压出来的文件夹中递归查找可执行安装程序
                string installExe = null;
                if (Directory.Exists(extractPath))
                {
                    var files = Directory.GetFiles(extractPath, "*.exe", SearchOption.AllDirectories);
                    foreach (var file in files)
                    {
                        string nameLower = Path.GetFileName(file).ToLower();
                        if (nameLower.Contains("setup") || nameLower.Contains("install"))
                        {
                            installExe = file;
                            break;
                        }
                    }
                    if (installExe == null && files.Length > 0)
                    {
                        installExe = files[0];
                    }
                }

                if (string.IsNullOrEmpty(installExe) || !File.Exists(installExe))
                {
                    phaseText.Report("未能在压缩包内找到安装执行文件。");
                    CleanTempFiles(localPath, extractPath);
                    return false;
                }

                Logger.Info($"找到永中安装程序: {installExe}，开始执行静默安装...");
                var psi = new ProcessStartInfo(installExe, "/S")
                {
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                // 启动伪进度推进（60 → 95）
                var fakeTask = FakeProgressAsync(progressPercent, 60, 95, durationMs: 45000, ct: ct);

                using (var proc = Process.Start(psi))
                {
                    if (proc == null)
                    {
                        phaseText.Report("无法启动安装程序。");
                        CleanTempFiles(localPath, extractPath);
                        return false;
                    }
                    await Task.Run(() => proc.WaitForExit(), ct);
                    Logger.Info($"永中Office 安装退出码: {proc.ExitCode}");
                }

                progressPercent?.Report(100);
                phaseText.Report("永中Office 安装完成！");

                CleanTempFiles(localPath, extractPath);
                return true;
            }
            catch (Exception ex)
            {
                Logger.Error("永中Office 安装失败", ex);
                phaseText.Report("安装失败: " + ex.Message);
                CleanTempFiles(localPath, extractPath);
                return false;
            }
        }

        private static void CleanTempFiles(string rarPath, string extractDir)
        {
            try { if (File.Exists(rarPath)) File.Delete(rarPath); } catch { }
            try { if (Directory.Exists(extractDir)) Directory.Delete(extractDir, true); } catch { }
        }

        private static async Task FakeProgressAsync(
            IProgress<int> progress, int from, int to, int durationMs, CancellationToken ct)
        {
            int steps = to - from;
            int intervalMs = steps > 0 ? durationMs / steps : durationMs;
            for (int i = from; i <= to && !ct.IsCancellationRequested; i++)
            {
                progress?.Report(i);
                await Task.Delay(intervalMs, ct).ContinueWith(_ => { });
            }
        }
    }
}
