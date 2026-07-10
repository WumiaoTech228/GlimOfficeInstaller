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

            // 3. 清理注册表
            progress?.Report(LocalizationStrings.Instance.StatusCleanRegistry);
            await Task.Run(() =>
            {
                string[] paths;
                switch (product)
                {
                    case ProductType.MsOffice:
                        paths = new[]
                        {
                            @"SOFTWARE\Microsoft\Office",
                            @"SOFTWARE\Microsoft\Office\ClickToRun",
                            @"SOFTWARE\Microsoft\AppVisv",
                            @"SOFTWARE\WOW6432Node\Microsoft\Office",
                            @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe",
                            @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Excel.exe",
                            @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Powerpnt.exe",
                            @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Outlook.exe",
                            @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Onenote.exe",
                            @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Visio.exe",
                            @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winproj.exe",
                            @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Msaccess.exe",
                            @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Mspub.exe",
                            @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Lync.exe",
                            @"SOFTWARE\Microsoft\Office\16.0\Registration",
                            @"SOFTWARE\Microsoft\Office\15.0\Registration",
                            @"SOFTWARE\Microsoft\Office\14.0\Registration",
                            @"SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\00006109C80000000000000000F01FEC",
                            @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Office16.PROPLUS",
                            @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Office16.VISIOPRO",
                            @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Office16.PROJECTPRO",
                            @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Office16.OUTLOOK"
                        };
                        break;
                    case ProductType.Wps:
                        paths = new[]
                        {
                            @"SOFTWARE\Kingsoft",
                            @"SOFTWARE\WOW6432Node\Kingsoft"
                        };
                        break;
                    case ProductType.Yozo:
                        paths = new[]
                        {
                            @"SOFTWARE\Yozosoft",
                            @"SOFTWARE\WOW6432Node\Yozosoft"
                        };
                        break;
                    case ProductType.OnlyOffice:
                        paths = new[]
                        {
                            @"SOFTWARE\ONLYOFFICE",
                            @"SOFTWARE\WOW6432Node\ONLYOFFICE"
                        };
                        break;
                    case ProductType.LibreOffice:
                        paths = new[]
                        {
                            @"SOFTWARE\The Document Foundation",
                            @"SOFTWARE\LibreOffice",
                            @"SOFTWARE\WOW6432Node\The Document Foundation",
                            @"SOFTWARE\WOW6432Node\LibreOffice"
                        };
                        break;
                    default:
                        paths = new string[0];
                        break;
                }
                foreach (var p in paths) RegistryHelper.DeleteKey(p);
            });
            count++;

            // 4. 清理卸载项
            progress?.Report(LocalizationStrings.Instance.StatusCleanUninstallEntries);
            await Task.Run(() => RegistryHelper.CleanUninstallEntries(product));
            count++;

            // 5. 删除残留文件
            progress?.Report(LocalizationStrings.Instance.StatusCleanResidualFiles);
            await Task.Run(() => RegistryHelper.CleanResidualFolders(product));
            count++;

            // 6. 清理快捷方式和文件关联并恢复已有的 Microsoft Office 关联
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
