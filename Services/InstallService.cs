using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using GOI.Activation;
using GOI.Helpers;
using GOI.Models;

namespace GOI.Services
{
    public class InstallService
    {
        private readonly CleanupService _cleanup = new CleanupService();
        private readonly DownloadService _download = new DownloadService();

        /// <summary>全流程安装：清理 → 下载 → 生成配置 → 安装（伪进度）→ 激活</summary>
        public async Task<bool> RunAsync(
            OfficeVersion version,
            string bitness,
            string channel,
            string lang,
            System.Collections.Generic.HashSet<OfficeComponent> selected,
            bool autoActivate,
            IProgress<string> phaseText,
            IProgress<int> downloadProgress,
            IProgress<InstallPhase> phaseProgress)
        {
            var loc = new LocalizationStrings();
            try
            {
                // 阶段 1: 清理
                phaseProgress?.Report(InstallPhase.Cleaning);
                phaseText.Report(loc.StatusClean);
                downloadProgress.Report(0);
                await _cleanup.CleanAsync(ProductType.MsOffice, phaseText);

                // 阶段 2: 下载 setup.exe（进度 0-5%）
                phaseProgress?.Report(InstallPhase.Downloading);
                phaseText.Report(loc.StatusDownloading);
                var dlProgress = new Progress<int>(p => downloadProgress.Report(p / 20)); // 0-100% → 0-5%
                var ok = await _download.DownloadODTAsync(dlProgress);
                if (!ok)
                {
                    phaseText.Report(loc.ErrDownloadFailed);
                    return false;
                }
                downloadProgress.Report(5);

                // 阶段 3: 生成配置
                phaseText.Report(loc.StatusConfiguringXml);
                var xml = XmlConfigHelper.Generate(version, bitness, channel, lang, selected);
                File.WriteAllText(AppConfig.XmlConfigPath, xml, System.Text.Encoding.UTF8);
                Logger.Info("配置文件已生成:\n" + xml);
                downloadProgress.Report(6);

                // 阶段 4: 运行安装 + 伪进度条（6-95%）
                // 启动 Microsoft ODT 安装程序，使用 UseShellExecute=true 确保微软官方安装动画向导能完美弹出和交互
                phaseProgress?.Report(InstallPhase.Installing);
                phaseText.Report(loc.StatusInstallingWizard);
                var psi = new ProcessStartInfo(AppConfig.SetupPath,
                    $"/configure \"{AppConfig.XmlConfigPath}\"")
                {
                    UseShellExecute = true,
                    CreateNoWindow = false,
                    WorkingDirectory = AppConfig.RootPath
                };

                using (var proc = Process.Start(psi))
                {
                    if (proc == null)
                    {
                        phaseText.Report(loc.ErrCannotStartInstaller);
                        return false;
                    }

                    // 等待 setup.exe 结束
                    await Task.Run(() => proc.WaitForExit());
                    Logger.Info($"ODT 安装完成，退出码: {proc.ExitCode}");

                    if (proc.ExitCode != 0)
                    {
                        phaseText.Report(loc.ErrInstallerExitCode(proc.ExitCode));
                        return false;
                    }
                }

                downloadProgress.Report(96);

                // 阶段 5: 激活 (Ohook)
                if (autoActivate)
                {
                    phaseProgress?.Report(InstallPhase.Activating);
                    phaseText.Report(loc.StatusActivating);
                    try
                    {
                        await ActivateOhookAsync();
                    }
                    catch (Exception ex)
                    {
                        Logger.Error("自动激活过程中断", ex);
                    }
                }
                downloadProgress.Report(100);
                return true;
            }
            catch (Exception ex)
            {
                Logger.Error("安装流程出错", ex);
                phaseText.Report(loc.ErrInstallFailed(ex.Message));
                return false;
            }
            finally
            {
                // 彻底删除临时缓存和下载的安装包
                try
                {
                    if (Directory.Exists(AppConfig.RootPath))
                    {
                        Directory.Delete(AppConfig.RootPath, true);
                        Logger.Info("已成功清理部署临时缓存目录：" + AppConfig.RootPath);
                    }
                }
                catch (Exception ex)
                {
                    Logger.Warn("清理部署临时目录失败：" + ex.Message);
                }
            }
        }



        /// <summary>Office Ohook 激活（C# 原生实现，不再依赖外部脚本）</summary>
        public async Task<bool> ActivateOhookAsync()
        {
            var result = await OhookActivator.ActivateAsync(
                new Progress<string>(msg => Logger.Info($"[Ohook] {msg}")));
            return result.Success;
        }

        /// <summary>删除 Office Ohook 激活（C# 原生实现）</summary>
        public async Task<bool> RemoveOfficeActivationAsync()
        {
            var result = await OhookActivator.DeactivateAsync(
                new Progress<string>(msg => Logger.Info($"[Ohook] {msg}")));
            return result.Success;
        }
    }
}
