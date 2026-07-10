using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading;
using System.Threading.Tasks;
using GOI.Helpers;
using GOI.Models;
using System.Windows;
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
            IProgress<string> phaseText,
            IProgress<int> progressPercent,
            IProgress<InstallPhase> phaseProgress,
            CancellationToken ct = default)
        {
            string localPath = Path.Combine(AppConfig.RootPath, YozoFileName);
            string extractPath = Path.Combine(AppConfig.RootPath, "YozoExtract");

            // ── 阶段 1：下载 ──
            phaseProgress?.Report(InstallPhase.Downloading);
            phaseText.Report(LocalizationStrings.Instance.StatusDownloadYozoRar);
            try
            {
                var downloader = new MultiThreadDownloader();
                var downloadProgress = new Progress<int>(pct =>
                {
                    progressPercent?.Report(pct / 2); // 下载占 0-50%
                });
                await downloader.DownloadAsync(YozoUrl, localPath, downloadProgress, 8, ct);
            }
            catch (Exception ex)
            {
                Logger.Error("下载 永中Office 失败", ex);
                phaseText.Report(LocalizationStrings.Instance.ErrDownloadFailedWithMsg);
                try { if (File.Exists(localPath)) File.Delete(localPath); } catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in YozoInstallService.cs at UnknownMethod", ex_captured); }
                return false;
            }

            // ── 阶段 2：解压 RAR ──
            phaseProgress?.Report(InstallPhase.Installing);
            phaseText.Report(LocalizationStrings.Instance.StatusExtractingProduct(LocalizationStrings.Instance.YozoTitle));
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
                phaseText.Report(LocalizationStrings.Instance.ErrExtractFailed(ex.Message));
                CleanTempFiles(localPath, extractPath);
                return false;
            }

            // ── 阶段 3：引导安装 ──
            phaseText.Report(LocalizationStrings.Instance.StatusInstallingProductGuide(LocalizationStrings.Instance.YozoTitle));
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
                    phaseText.Report(LocalizationStrings.Instance.ErrYozoExeNotFound);
                    CleanTempFiles(localPath, extractPath);
                    return false;
                }

                // 弹出提示框告知用户官方暂不支持静默安装，需要人工交互安装
                MessageBox.Show(
                    LocalizationStrings.Instance.DlgConfirmYozoMsg,
                    LocalizationStrings.Instance.DlgConfirmYozoTitle,
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);

                Logger.Info($"找到永中安装程序: {installExe}，开始引导交互式安装...");
                var psi = new ProcessStartInfo(installExe)
                {
                    UseShellExecute = true
                };

                // 启动并关联局部 CancellationTokenSource 以便在安装结束时立即取消伪进度，防止进度条后台继续跳动
                using (var ctsFake = new CancellationTokenSource())
                {
                    var fakeTask = FakeProgressAsync(progressPercent, 60, 95, durationMs: 45000, ct: ctsFake.Token);

                    using (var proc = Process.Start(psi))
                    {
                        if (proc == null)
                        {
                            phaseText.Report(LocalizationStrings.Instance.ErrCannotStartInstallerWithMsg);
                            CleanTempFiles(localPath, extractPath);
                            return false;
                        }
                        await Task.Run(() => proc.WaitForExit(), ct);
                        Logger.Info($"永中Office 安装退出码: {proc.ExitCode}");

                        ctsFake.Cancel();

                        if (proc.ExitCode != 0)
                        {
                            phaseText.Report(LocalizationStrings.Instance.ErrInstallerAbortedWithCode(LocalizationStrings.Instance.YozoTitle, proc.ExitCode));
                            CleanTempFiles(localPath, extractPath);
                            return false;
                        }
                    }
                }

                progressPercent?.Report(100);
                phaseText.Report(LocalizationStrings.Instance.StatusProductInstalled(LocalizationStrings.Instance.YozoTitle));
 
                CleanTempFiles(localPath, extractPath);
                return true;
            }
            catch (Exception ex)
            {
                Logger.Error("永中Office 安装失败", ex);
                phaseText.Report(LocalizationStrings.Instance.StatusProductInstallFailed(LocalizationStrings.Instance.YozoTitle, ex.Message));
                CleanTempFiles(localPath, extractPath);
                return false;
            }
        }

        private static void CleanTempFiles(string rarPath, string extractDir)
        {
            try { if (File.Exists(rarPath)) File.Delete(rarPath); } catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in YozoInstallService.cs at UnknownMethod", ex_captured); }
            try { if (Directory.Exists(extractDir)) Directory.Delete(extractDir, true); } catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in YozoInstallService.cs at UnknownMethod", ex_captured); }
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
