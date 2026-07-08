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
    public class OnlyOfficeInstallService
    {
        private const string OnlyOfficeUrl = "https://download.onlyoffice.com/install/desktop/editors/windows/distrib/onlyoffice/DesktopEditors_x64.exe";
        private const string OnlyOfficeFileName = "OnlyOffice_Setup.exe";

        /// <summary>下载并静默安装 OnlyOffice，以伪进度条报告进度</summary>
        public async Task<bool> InstallAsync(
            OnlyOfficeVersion version,
            IProgress<string> phaseText,
            IProgress<int> progressPercent,
            CancellationToken ct = default)
        {
            string localPath = Path.Combine(AppConfig.RootPath, OnlyOfficeFileName);

            // ── 阶段 1：下载 ──
            phaseText.Report("正在下载 OnlyOffice Desktop Editors...");
            try
            {
                var downloader = new MultiThreadDownloader();
                var downloadProgress = new Progress<int>(pct =>
                {
                    progressPercent?.Report(pct / 2); // 下载占 0-50%
                });
                await downloader.DownloadAsync(OnlyOfficeUrl, localPath, downloadProgress, 8, ct);
            }
            catch (Exception ex)
            {
                Logger.Error("下载 OnlyOffice 失败", ex);
                phaseText.Report("下载失败，请检查网络连接。");
                try { if (File.Exists(localPath)) File.Delete(localPath); } catch { }
                return false;
            }

            // ── 阶段 2：静默安装（/VERYSILENT /NORESTART 参数）──
            phaseText.Report("正在静默安装 OnlyOffice...");
            progressPercent?.Report(55);

            try
            {
                var psi = new ProcessStartInfo(localPath, "/VERYSILENT /NORESTART")
                {
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                // 启动伪进度推进（55 → 95）
                var fakeTask = FakeProgressAsync(progressPercent, 55, 95, durationMs: 60000, ct: ct);

                using (var proc = Process.Start(psi))
                {
                    if (proc == null)
                    {
                        phaseText.Report("无法启动安装程序。");
                        return false;
                    }
                    await Task.Run(() => proc.WaitForExit(), ct);
                    Logger.Info($"OnlyOffice 安装退出码: {proc.ExitCode}");
                }

                progressPercent?.Report(100);
                phaseText.Report("OnlyOffice 安装完成！");

                // 清理安装包
                try { if (File.Exists(localPath)) File.Delete(localPath); } catch { }
                return true;
            }
            catch (Exception ex)
            {
                Logger.Error("OnlyOffice 安装失败", ex);
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
