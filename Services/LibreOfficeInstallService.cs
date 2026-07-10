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
            IProgress<string> phaseText,
            IProgress<int> progressPercent,
            IProgress<InstallPhase> phaseProgress,
            CancellationToken ct = default)
        {
            string localPath = Path.Combine(AppConfig.RootPath, LibreOfficeFileName);

            // ── 阶段 1：下载 ──
            phaseProgress?.Report(InstallPhase.Downloading);
            phaseText.Report(LocalizationStrings.Instance.StatusDownloadingProduct("LibreOffice"));
            try
            {
                var downloader = new MultiThreadDownloader();
                var downloadProgress = new Progress<int>(pct =>
                {
                    progressPercent?.Report(pct / 2); // 下载占 0-50%
                });
                await downloader.DownloadAsync(LibreOfficeUrl, localPath, downloadProgress, 8, ct);
            }
            catch (Exception ex)
            {
                Logger.Error("下载 LibreOffice 失败", ex);
                phaseText.Report(LocalizationStrings.Instance.ErrDownloadFailedWithMsg);
                try { if (File.Exists(localPath)) File.Delete(localPath); } catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in LibreOfficeInstallService.cs at UnknownMethod", ex_captured); }
                return false;
            }

            // ── 阶段 2：静默安装（msiexec /i /qn /norestart）──
            phaseProgress?.Report(InstallPhase.Installing);
            phaseText.Report(LocalizationStrings.Instance.StatusInstallingProduct("LibreOffice"));
            progressPercent?.Report(55);

            try
            {
                var psi = new ProcessStartInfo("msiexec.exe", $"/i \"{localPath}\" /qn /norestart")
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
                            phaseText.Report(LocalizationStrings.Instance.ErrCannotStartMsiWithMsg);
                            return false;
                        }
                        await Task.Run(() => proc.WaitForExit(), ct);
                        Logger.Info($"LibreOffice 安装退出码: {proc.ExitCode}");

                        ctsFake.Cancel();

                        if (proc.ExitCode != 0 && proc.ExitCode != 3010)
                        {
                            string libreOfficeVer = RegistryHelper.GetInstalledProductVersion(ProductType.LibreOffice);
                            if (!string.IsNullOrEmpty(libreOfficeVer))
                            {
                                Logger.Info($"LibreOffice 安装退出码虽为 {proc.ExitCode}，但检测到系统注册表中已成功注册 LibreOffice ({libreOfficeVer})，判定为安装成功！");
                            }
                            else
                            {
                                phaseText.Report(LocalizationStrings.Instance.ErrInstallerAbortedWithCode("LibreOffice", proc.ExitCode));
                                try { if (File.Exists(localPath)) File.Delete(localPath); } catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in LibreOfficeInstallService.cs at UnknownMethod", ex_captured); }
                                return false;
                            }
                        }
                    }
                }

                progressPercent?.Report(100);
                phaseText.Report(LocalizationStrings.Instance.StatusProductInstalled("LibreOffice"));

                // 清理安装包
                try { if (File.Exists(localPath)) File.Delete(localPath); } catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in LibreOfficeInstallService.cs at UnknownMethod", ex_captured); }
                return true;
            }
            catch (Exception ex)
            {
                Logger.Error("LibreOffice 安装失败", ex);
                phaseText.Report(LocalizationStrings.Instance.StatusProductInstallFailed("LibreOffice", ex.Message));
                try { if (File.Exists(localPath)) File.Delete(localPath); } catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in LibreOfficeInstallService.cs at UnknownMethod", ex_captured); }
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
