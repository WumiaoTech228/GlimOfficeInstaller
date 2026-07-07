using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using GOI.Helpers;
using GOI.Models;

namespace GOI.Services
{
    public class InstallService
    {
        private readonly CleanupService _cleanup = new CleanupService();
        private readonly DownloadService _download = new DownloadService();

        /// <summary>全流程安装：清理 → 下载 → 生成配置 → 安装 → 激活</summary>
        public async Task<bool> RunAsync(
            OfficeVersion version,
            Architecture arch,
            System.Collections.Generic.HashSet<OfficeComponent> selected,
            IProgress<string> phaseText,
            IProgress<int> downloadProgress)
        {
            try
            {
                // 阶段 1: 清理
                phaseText.Report("正在清理旧版本 Office 残留...");
                await _cleanup.CleanAsync(phaseText);

                // 阶段 2: 下载 ODT
                phaseText.Report("正在下载安装组件...");
                var ok = await _download.DownloadODTAsync(downloadProgress);
                if (!ok)
                {
                    phaseText.Report("下载失败，请检查网络连接。");
                    return false;
                }

                // 阶段 3: 生成配置
                phaseText.Report("正在生成安装配置...");
                var xml = XmlConfigHelper.Generate(version, arch, selected);
                File.WriteAllText(AppConfig.XmlConfigPath, xml, System.Text.Encoding.UTF8);
                Logger.Info("配置文件已生成:\n" + xml);

                // 阶段 4: 运行安装
                // GOI.exe 已通过 manifest requireAdministrator 提权，setup.exe 继承管理员权限无需再次 runas
                phaseText.Report("正在安装 Office …");
                var psi = new ProcessStartInfo(AppConfig.SetupPath,
                    $"/configure \"{AppConfig.XmlConfigPath}\"")
                {
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WorkingDirectory = AppConfig.RootPath
                };

                using (var proc = Process.Start(psi))
                {
                    if (proc == null)
                    {
                        phaseText.Report("无法启动安装程序。");
                        return false;
                    }
                    await Task.Run(() => proc.WaitForExit());
                    Logger.Info($"ODT 安装完成，退出码: {proc.ExitCode}");
                }

                // 阶段 5: 激活 (Ohook)
                phaseText.Report("正在激活 Office...");
                return await ActivateOhookAsync();
            }
            catch (Exception ex)
            {
                Logger.Error("安装流程出错", ex);
                phaseText.Report("安装失败: " + ex.Message);
                return false;
            }
        }

        /// <summary>Office Ohook 激活</summary>
        public async Task<bool> ActivateOhookAsync()
        {
            var script = ResourceHelper.GetScriptPath("Ohook_Activation_AIO.cmd");
            if (script == null) return false;
            return await RunScriptAsync(script, "/Ohook");
        }

        /// <summary>删除 Office Ohook 激活</summary>
        public async Task<bool> RemoveOfficeActivationAsync()
        {
            var script = ResourceHelper.GetScriptPath("Ohook_Activation_AIO.cmd");
            if (script == null) return false;
            return await RunScriptAsync(script, "/Ohook-Uninstall");
        }

        /// <summary>低层：以管理员权限静默运行脚本并等待退出，返回是否成功</summary>
        private async Task<bool> RunScriptAsync(string scriptPath, string args)
        {
            try
            {
                // 必须以 runas 提权运行，Ohook 需要管理员权限写入注册表
                var psi = new ProcessStartInfo("cmd.exe",
                    $"/c \"\"{scriptPath}\" {args}\"")
                {
                    UseShellExecute = true,
                    Verb = "runas",
                    WindowStyle = ProcessWindowStyle.Hidden,
                    WorkingDirectory = Path.GetDirectoryName(scriptPath)
                };

                using (var proc = Process.Start(psi))
                {
                    if (proc == null) return false;
                    await Task.Run(() => proc.WaitForExit());
                    Logger.Info($"MAS [{Path.GetFileName(scriptPath)} {args}] 退出码: {proc.ExitCode}");
                    return proc.ExitCode == 0;
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"运行激活脚本失败 [{args}]", ex);
                return false;
            }
        }
    }
}
