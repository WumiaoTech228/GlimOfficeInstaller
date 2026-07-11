using System;
using System.Linq;
using System.Threading.Tasks;
using GOI.Helpers;
using GOI.Models;
using Microsoft.Win32;

namespace GOI.Services
{
    public class CleanupService
    {
        /// <summary>深度清理特定 Office 残留</summary>
        public async Task<int> CleanAsync(ProductType product, IProgress<string> progress = null)
        {
            int count = 0;
            Logger.Info($">>> 开始卸载与清理 {product} <<<");

            // 1. 对于 MS Office，首先尝试运行官方 Click-to-Run 卸载程序 (拷贝自 OTP 卸载流程，优先安全卸载)
            if (product == ProductType.MsOffice)
            {
                progress?.Report(LocalizationStrings.Instance.StatusCleanStartC2R);
                string c2rPath = @"C:\Program Files\Common Files\microsoft shared\ClickToRun\OfficeClickToRun.exe";
                if (!System.IO.File.Exists(c2rPath))
                    c2rPath = @"C:\Program Files (x86)\Common Files\microsoft shared\ClickToRun\OfficeClickToRun.exe";

                if (System.IO.File.Exists(c2rPath))
                {
                    try
                    {
                        var psi = new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = c2rPath,
                            Arguments = "scenario=install action=uninstall",
                            UseShellExecute = true,
                            Verb = "runas"
                        };
                        using (var p = System.Diagnostics.Process.Start(psi))
                        {
                            if (p != null)
                            {
                                await Task.Run(() => p.WaitForExit());
                                count++;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Error("启动官方 C2R 卸载失败", ex);
                    }
                }
            }

            progress?.Report(LocalizationStrings.Instance.StatusCleanKillProcesses);

            // 2. 终止进程
            await Task.Run(() => RegistryHelper.KillOfficeProcesses(product));
            count++;

            // 3. 停止/删除服务 (仅限 MS Office)
            if (product == ProductType.MsOffice)
            {
                progress?.Report(LocalizationStrings.Instance.StatusCleanC2RService);
                await Task.Run(() => RegistryHelper.RemoveClickToRunService());
                count++;
            }

            // 4. 清理注册表
            progress?.Report(LocalizationStrings.Instance.StatusCleanRegistry);
            await Task.Run(() => RegistryHelper.CleanRegistryKeys(product));
            count++;

            // 5. 清理卸载项
            progress?.Report(LocalizationStrings.Instance.StatusCleanUninstallEntries);
            await Task.Run(() => RegistryHelper.CleanUninstallEntries(product));
            count++;

            // 6. 删除残留文件
            progress?.Report(LocalizationStrings.Instance.StatusCleanResidualFiles);
            await Task.Run(() => RegistryHelper.CleanResidualFolders(product));
            count++;

            // 7. 清理快捷方式和文件关联并恢复已有的 Microsoft Office 关联
            progress?.Report(LocalizationStrings.Instance.StatusCleanAssociations);
            await Task.Run(() =>
            {
                RegistryHelper.CleanShortcuts(product);
                RegistryHelper.CleanFileAssociations(product);
                
                // 如果卸载的不是 MS Office，我们尝试恢复存在的 MS Office 关联并刷新图标缓存，防止文件关联变成空白或混乱
                if (product != ProductType.MsOffice)
                {
                    RegistryHelper.RestoreInstalledProductAssociations(progress);
                    RegistryHelper.RefreshIconCache();
                }
            });
            count++;

            Logger.Info($"<<< {product} 深度清理完成 >>>");
            return count;
        }
    }
}
