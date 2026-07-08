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
    public class WpsInstallService
    {
        /// <summary>WPS Office 官方个人版 CDN 直链（wps.cn 官网实际下发的安装包）</summary>
        private const string WpsOfficialUrl = "https://official-package.wpscdn.cn/wps/download/WPS_Setup.exe";
        private const string WpsFileName = "WPS_Setup.exe";

        /// <summary>下载并静默安装 WPS 官方个人版</summary>
        public async Task<bool> InstallAsync(
            WpsVersion version,
            IProgress<string> phaseText,
            IProgress<int> progressPercent,
            CancellationToken ct = default)
        {
            string localPath = Path.Combine(AppConfig.RootPath, WpsFileName);

            // ── 阶段 1：下载 ──
            phaseText.Report("正在从 WPS 官方服务器下载最新版...");
            try
            {
                using (var client = new WebClient())
                {
                    client.Headers.Add("User-Agent",
                        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36");
                    client.DownloadProgressChanged += (s, e) =>
                        progressPercent?.Report(e.ProgressPercentage / 2); // 下载占 0-50%
                    await client.DownloadFileTaskAsync(new Uri(WpsOfficialUrl), localPath);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("下载 WPS 失败", ex);
                phaseText.Report("WPS 下载失败，请检查网络连接。");
                try { if (File.Exists(localPath)) File.Delete(localPath); } catch { }
                return false;
            }

            // ── 阶段 2：静默安装 ──
            phaseText.Report("正在静默安装 WPS Office...");
            progressPercent?.Report(55);

            try
            {
                // 官方个人版安装包支持 /S 静默参数
                var psi = new ProcessStartInfo(localPath, "/S")
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
                        phaseText.Report("无法启动 WPS 安装程序。");
                        return false;
                    }
                    await Task.Run(() => proc.WaitForExit(), ct);
                    Logger.Info($"WPS 安装退出码: {proc.ExitCode}");
                }

                progressPercent?.Report(100);
                phaseText.Report("WPS Office 安装完成！");

                // 清理安装包
                try { if (File.Exists(localPath)) File.Delete(localPath); } catch { }
                return true;
            }
            catch (Exception ex)
            {
                Logger.Error("WPS 安装失败", ex);
                phaseText.Report("WPS 安装失败: " + ex.Message);
                return false;
            }
        }

        /// <summary>在指定时间内将进度从 from 匀速推进到 to，可被取消</summary>
        private static async Task FakeProgressAsync(
            IProgress<int> progress, int from, int to, int durationMs, CancellationToken ct)
        {
            int steps = to - from;
            int intervalMs = steps > 0 ? durationMs / steps : durationMs;
            for (int i = from; i <= to && !ct.IsCancellationRequested; i++)
            {
                progress?.Report(i);
                await Task.Delay(intervalMs, ct).ContinueWith(_ => { }); // 忽略取消异常
            }
        }
    }
}

