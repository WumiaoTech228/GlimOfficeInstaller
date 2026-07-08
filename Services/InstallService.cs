using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
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
            Architecture arch,
            System.Collections.Generic.HashSet<OfficeComponent> selected,
            IProgress<string> phaseText,
            IProgress<int> downloadProgress)
        {
            try
            {
                // 阶段 1: 清理
                phaseText.Report("正在清理旧版本 Office 残留...");
                downloadProgress.Report(0);
                await _cleanup.CleanAsync(ProductType.MsOffice, phaseText);

                // 阶段 2: 下载 setup.exe（进度 0-5%）
                phaseText.Report("正在下载安装组件...");
                var dlProgress = new Progress<int>(p => downloadProgress.Report(p / 20)); // 0-100% → 0-5%
                var ok = await _download.DownloadODTAsync(dlProgress);
                if (!ok)
                {
                    phaseText.Report("下载失败，请检查网络连接。");
                    return false;
                }
                downloadProgress.Report(5);

                // 阶段 3: 生成配置
                phaseText.Report("正在生成安装配置...");
                var xml = XmlConfigHelper.Generate(version, arch, selected);
                File.WriteAllText(AppConfig.XmlConfigPath, xml, System.Text.Encoding.UTF8);
                Logger.Info("配置文件已生成:\n" + xml);
                downloadProgress.Report(6);

                // 阶段 4: 运行安装 + 伪进度条（6-95%）
                // GOI.exe 已通过 manifest requireAdministrator 提权，setup.exe 继承管理员权限
                phaseText.Report("正在安装 Office，请耐心等待（可能需要 10-30 分钟）…");
                var psi = new ProcessStartInfo(AppConfig.SetupPath,
                    $"/configure \"{AppConfig.XmlConfigPath}\"")
                {
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WorkingDirectory = AppConfig.RootPath
                };

                using (var cts = new CancellationTokenSource())
                using (var proc = Process.Start(psi))
                {
                    if (proc == null)
                    {
                        phaseText.Report("无法启动安装程序。");
                        return false;
                    }

                    // 启动伪进度任务：监控 OfficeClickToRun.exe 进程，缓慢推进进度
                    var fakeProgressTask = RunFakeInstallProgressAsync(
                        downloadProgress, phaseText, cts.Token);

                    // 等待 setup.exe 结束
                    await Task.Run(() => proc.WaitForExit());
                    Logger.Info($"ODT 安装完成，退出码: {proc.ExitCode}");

                    // 通知伪进度停止并推到 95%
                    cts.Cancel();
                    await fakeProgressTask;
                }

                downloadProgress.Report(96);

                // 阶段 5: 激活 (Ohook)
                phaseText.Report("正在激活 Office...");
                bool activated = await ActivateOhookAsync();
                downloadProgress.Report(100);
                return activated;
            }
            catch (Exception ex)
            {
                Logger.Error("安装流程出错", ex);
                phaseText.Report("安装失败: " + ex.Message);
                return false;
            }
        }

        /// <summary>
        /// 伪进度条：每隔 3 秒轮询 OfficeClickToRun.exe / officec2rclient 进程是否存在。
        /// 进程存在 → 缓慢推进（每次 +1%）。
        /// 进程消失或收到取消信号 → 快速推到 95%。
        /// </summary>
        private static async Task RunFakeInstallProgressAsync(
            IProgress<int> progress,
            IProgress<string> phaseText,
            CancellationToken ct)
        {
            int current = 6;
            bool c2rSeen = false;

            while (!ct.IsCancellationRequested && current < 95)
            {
                try { await Task.Delay(3000, ct); }
                catch (TaskCanceledException) { break; }

                bool c2rRunning = Process.GetProcessesByName("OfficeClickToRun").Length > 0
                               || Process.GetProcessesByName("officec2rclient").Length > 0;

                if (c2rRunning) c2rSeen = true;

                // Office C2R 曾出现过但现在消失 → 安装完成，跳出
                if (c2rSeen && !c2rRunning) break;

                // 缓慢推进
                current = Math.Min(current + 1, 94);
                progress?.Report(current);

                // 更新文字提示
                string dots = new string('.', ((current - 6) / 5 % 4) + 1);
                if (current < 30)
                    phaseText.Report($"正在下载 Office 文件{dots}");
                else if (current < 70)
                    phaseText.Report($"正在安装 Office 组件{dots}");
                else
                    phaseText.Report($"即将完成，请稍候{dots}");
            }

            // 快速推进到 95%
            for (int i = current; i <= 95; i++)
            {
                progress?.Report(i);
                await Task.Delay(30);
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
