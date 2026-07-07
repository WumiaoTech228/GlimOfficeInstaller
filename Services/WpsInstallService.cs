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
        private static readonly string[] WpsUrls =
        {
            /* 2013 */ "https://share.osbox.top/d/Office/WPS/WPSOffice_Professional_2013_9.1.0.5026(2024.08.15).exe?sign=RN7RcHfiroz81R-gZzoUljzXjGzveyRTIiRDmpJSw48=:1783540085",
            /* 2016 */ "https://share.osbox.top/d/Office/WPS/WPS%20Office%202016%20%E4%B8%93%E4%B8%9A%E5%A2%9E%E5%BC%BA%E7%89%88_10.8.2.7164_mefcl_Setup.exe?sign=79sBvillY6_oqXPyUl4CKdDA_ooAG2duWXJ7D7tSsPg=:1783540085",
            /* 2019 */ "https://share.osbox.top/d/Office/WPS/WPSOffice2019ProPlus_11.8.2.12330_mefcl.exe?sign=qL5uRhU1DuyQj5M7NvKPXllmsNP6sRLw9UdYx8ITyTk=:1783540085",
            /* 2023 */ "https://share.osbox.top/d/Office/WPS/WPSOfficePro_12.1.0.26884_mefcl_x64_20260703.exe?sign=01zoUKonLwKJzR6Pj0ZFs8T6tbVYW739zJIJuTtp_vU=:1783540085"
        };

        private static readonly string[] WpsFileNames =
        {
            "WPS2013.exe", "WPS2016.exe", "WPS2019.exe", "WPS2023.exe"
        };

        /// <summary>下载并静默安装 WPS，以伪进度条报告进度</summary>
        public async Task<bool> InstallAsync(
            WpsVersion version,
            IProgress<string> phaseText,
            IProgress<int> progressPercent,
            CancellationToken ct = default)
        {
            int idx = (int)version;
            string url = WpsUrls[idx];
            string fileName = WpsFileNames[idx];
            string localPath = Path.Combine(AppConfig.RootPath, fileName);

            // ── 阶段 1：下载 ──
            phaseText.Report($"正在下载 WPS {GetVersionYear(version)}...");
            try
            {
                using (var client = new WebClient())
                {
                    client.Headers.Add("User-Agent",
                        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36");
                    client.DownloadProgressChanged += (s, e) =>
                        progressPercent?.Report(e.ProgressPercentage / 2); // 下载占 0-50%
                    await client.DownloadFileTaskAsync(new Uri(url), localPath);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("下载 WPS 失败", ex);
                phaseText.Report("WPS 下载失败，请检查网络连接。");
                try { if (File.Exists(localPath)) File.Delete(localPath); } catch { }
                return false;
            }

            // ── 阶段 2：静默安装（/S 参数）──
            phaseText.Report($"正在静默安装 WPS {GetVersionYear(version)}...");
            progressPercent?.Report(55);

            try
            {
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
                phaseText.Report($"WPS {GetVersionYear(version)} 安装完成！");

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

        private static string GetVersionYear(WpsVersion v) => v switch
        {
            WpsVersion.Wps2013 => "2013",
            WpsVersion.Wps2016 => "2016",
            WpsVersion.Wps2019 => "2019",
            WpsVersion.Wps2023 => "2023",
            _ => ""
        };
    }
}
