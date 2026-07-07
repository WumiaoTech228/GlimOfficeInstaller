using System;
using System.Linq;
using System.Threading.Tasks;
using GOI.Helpers;
using Microsoft.Win32;

namespace GOI.Services
{
    public class CleanupService
    {
        /// <summary>深度清理 Office 残留，返回清理项数量</summary>
        public async Task<int> CleanAsync(IProgress<string> progress = null)
        {
            int count = 0;
            progress?.Report("正在终止 Office 进程...");
            Logger.Info(">>> 开始深度清理 Office <<<");

            // 1. 终止进程
            await Task.Run(() => RegistryHelper.KillOfficeProcesses());
            count++;

            // 2. 停止/删除服务
            progress?.Report("正在清理 ClickToRun 服务...");
            await Task.Run(() => RegistryHelper.RemoveClickToRunService());
            count++;

            // 3. 清理注册表
            progress?.Report("正在清理注册表...");
            await Task.Run(() =>
            {
                var paths = new[]
                {
                    // MS Office
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
                    @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Office16.OUTLOOK",
                    
                    // WPS
                    @"SOFTWARE\Kingsoft",
                    @"SOFTWARE\WOW6432Node\Kingsoft",

                    // Yozo
                    @"SOFTWARE\Yozosoft",
                    @"SOFTWARE\WOW6432Node\Yozosoft",

                    // OnlyOffice
                    @"SOFTWARE\ONLYOFFICE",
                    @"SOFTWARE\WOW6432Node\ONLYOFFICE",

                    // LibreOffice
                    @"SOFTWARE\The Document Foundation",
                    @"SOFTWARE\LibreOffice",
                    @"SOFTWARE\WOW6432Node\The Document Foundation",
                    @"SOFTWARE\WOW6432Node\LibreOffice"
                };
                foreach (var p in paths) RegistryHelper.DeleteKey(p);
            });
            count++;

            // 4. 清理卸载项
            progress?.Report("正在清理卸载记录...");
            await Task.Run(() => RegistryHelper.CleanUninstallEntries());
            count++;

            // 5. 删除残留文件
            progress?.Report("正在清理残留文件...");
            await Task.Run(() => RegistryHelper.CleanResidualFolders());
            count++;

            Logger.Info("<<< 深度清理完成 >>>");
            return count;
        }
    }
}
