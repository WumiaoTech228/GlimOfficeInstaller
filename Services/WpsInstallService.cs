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
            /* 2013 */ "https://share.osbox.top/d/CloudService/WPS%20Pro/WPSPRO2013.exe",
            /* 2016 */ "https://share.osbox.top/d/CloudService/WPS%20Pro/WPSPRO2016.exe",
            /* 2019 */ "https://share.osbox.top/d/CloudService/WPS%20Pro/WPSPRO2019.exe",
            /* 2023 */ "https://share.osbox.top/d/CloudService/WPS%20Pro/WPSPRO2023.exe",
            /* 最新版(官方) */ "https://official-package.wpscdn.cn/wps/download/WPS_Setup_26899.exe"
        };

        private static readonly string[] WpsFileNames =
        {
            "WPS2013.exe", "WPS2016.exe", "WPS2019.exe", "WPS2023.exe", "WPS_Setup_26899.exe"
        };

        /// <summary>下载并静默安装 WPS，以伪进度条报告进度</summary>
        public async Task<bool> InstallAsync(
            WpsVersion version,
            IProgress<string> phaseText,
            IProgress<int> progressPercent,
            IProgress<InstallPhase> phaseProgress,
            CancellationToken ct = default)
        {
            int idx = (int)version;
            string url = WpsUrls[idx];
            string fileName = WpsFileNames[idx];
            string localPath = Path.Combine(AppConfig.RootPath, fileName);

            // ── 阶段 1：下载 ──
            phaseProgress?.Report(InstallPhase.Downloading);
            phaseText.Report(LocalizationStrings.Instance.StatusDownloadingProduct($"WPS {GetVersionLabel(version)}"));
            try
            {
                var downloader = new MultiThreadDownloader();
                var downloadProgress = new Progress<int>(pct =>
                {
                    progressPercent?.Report(pct / 2); // 下载占 0-50%
                });
                await downloader.DownloadAsync(url, localPath, downloadProgress, 8, ct);
            }
            catch (Exception ex)
            {
                Logger.Error("下载 WPS 失败", ex);
                phaseText.Report(LocalizationStrings.Instance.ErrDownloadFailedWithMsg);
                try { if (File.Exists(localPath)) File.Delete(localPath); } catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in WpsInstallService.cs", ex_captured); }
                return false;
            }

            // ── 阶段 2：静默安装 ──
            phaseProgress?.Report(InstallPhase.Installing);
            phaseText.Report(LocalizationStrings.Instance.StatusInstallingProduct($"WPS {GetVersionLabel(version)}"));
            progressPercent?.Report(55);

            try
            {
                string args = "/S";
                if (version == WpsVersion.Wps2019)
                {
                    args = "/NoCloud";
                }

                var psi = new ProcessStartInfo(localPath, args)
                {
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                // 启动并关联局部 CancellationTokenSource 以便在安装结束时立即取消伪进度，防止进度条后台继续跳动
                using (var ctsFake = new CancellationTokenSource())
                {
                    var fakeTask = FakeProgressAsync(progressPercent, 55, 95, durationMs: 90000, ct: ctsFake.Token);

                    using (var proc = Process.Start(psi))
                    {
                        if (proc == null)
                        {
                            phaseText.Report(LocalizationStrings.Instance.ErrCannotStartInstallerWithMsg);
                            return false;
                        }
                        await Task.Run(() => proc.WaitForExit(), ct);
                        Logger.Info($"WPS 安装退出码: {proc.ExitCode}");
 
                        ctsFake.Cancel();
                        
                        if (proc.ExitCode != 0)
                        {
                            string wpsVer = RegistryHelper.GetInstalledProductVersion(ProductType.Wps);
                            if (!string.IsNullOrEmpty(wpsVer))
                            {
                                Logger.Info($"WPS 安装退出码虽为 {proc.ExitCode}，但检测到系统注册表中已成功注册 WPS ({wpsVer})，判定为安装成功！");
                            }
                            else
                            {
                                phaseText.Report(LocalizationStrings.Instance.ErrInstallerAbortedWithCode("WPS", proc.ExitCode));
                                try { if (File.Exists(localPath)) File.Delete(localPath); } catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in WpsInstallService.cs", ex_captured); }
                                return false;
                            }
                        }
                    }
                }
 
                progressPercent?.Report(100);
                phaseText.Report(LocalizationStrings.Instance.StatusProductInstalled($"WPS {GetVersionLabel(version)}"));
 
                // 清理安装包
                try { if (File.Exists(localPath)) File.Delete(localPath); } catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in WpsInstallService.cs", ex_captured); }
                return true;
            }
            catch (Exception ex)
            {
                Logger.Error("WPS 安装失败", ex);
                phaseText.Report(LocalizationStrings.Instance.StatusProductInstallFailed("WPS", ex.Message));
                try { if (File.Exists(localPath)) File.Delete(localPath); } catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in WpsInstallService.cs", ex_captured); }
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

        private static string GetVersionLabel(WpsVersion v) => v switch
        {
            WpsVersion.Wps2013 => "2013",
            WpsVersion.Wps2016 => "2016",
            WpsVersion.Wps2019 => "2019",
            WpsVersion.Wps2023 => "2023",
            WpsVersion.WpsLatest => LocalizationStrings.Instance.WpsVersionLatestLabel,
            _ => ""
        };
    }
}
