using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Threading;
using System.Threading.Tasks;
using GOI.Helpers;
using GOI.Models;

namespace GOI.Services
{
    public class LibreOfficeInstallService
    {
        private const string LibreOfficeUrl = "https://download.documentfoundation.org/libreoffice/stable/26.2.4/win/x86_64/LibreOffice_26.2.4_Win_x86-64.msi";
        private const string LibreOfficeFileName = "LibreOffice_Setup.msi";

        /// <summary>下载并静默安装 LibreOffice，以伪进度条报告进度</summary>
        public async Task<bool> InstallAsync(
            LibreOfficeVersion version,
            IProgress<string> phaseText,
            IProgress<int> progressPercent,
            CancellationToken ct = default)
        {
            string localPath = Path.Combine(AppConfig.RootPath, LibreOfficeFileName);

            // ── 阶段 1：下载 ──
            phaseText.Report("正在下载 LibreOffice 稳定版...");
            try
            {
                using (var client = new WebClient())
                {
                    client.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36");
                    client.DownloadProgressChanged += (s, e) =>
                        progressPercent?.Report(e.ProgressPercentage / 2); // 下载占 0-50%
                    await client.DownloadFileTaskAsync(new Uri(LibreOfficeUrl), localPath);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("下载 LibreOffice 失败", ex);
                phaseText.Report("下载失败，请检查网络连接。");
                try { if (File.Exists(localPath)) File.Delete(localPath); } catch { }
                return false;
            }

            // ── 阶段 2：静默安装（msiexec /i /qn /norestart）──
            phaseText.Report("正在静默安装 LibreOffice...");
            progressPercent?.Report(55);

            try
            {
                var psi = new ProcessStartInfo("msiexec.exe", $"/i \"{localPath}\" /qn /norestart")
                {
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                // 启动伪进度推进（55 → 95）
                var fakeTask = FakeProgressAsync(progressPercent, 55, 95, durationMs: 90000, ct: ct);

                using (var proc = Process.Start(psi))
                {
                    if (proc == null)
                    {
                        phaseText.Report("无法启动 MSI 安装引擎。");
                        return false;
                    }
                    await Task.Run(() => proc.WaitForExit(), ct);
                    Logger.Info($"LibreOffice 安装退出码: {proc.ExitCode}");
                }

                progressPercent?.Report(100);
                phaseText.Report("LibreOffice 安装完成！");

                // 清理安装包
                try { if (File.Exists(localPath)) File.Delete(localPath); } catch { }
                return true;
            }
            catch (Exception ex)
            {
                Logger.Error("LibreOffice 安装失败", ex);
                phaseText.Report("安装失败: " + ex.Message);
                return false;
            }
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
