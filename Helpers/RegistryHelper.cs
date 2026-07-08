using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Microsoft.Win32;
using GOI.Models;

namespace GOI.Helpers
{
    public static class RegistryHelper
    {
        /// <summary>删除指定路径下的注册表键，HKCU 和 HKLM 均尝试</summary>
        public static void DeleteKey(string subKeyPath)
        {
            DeleteKey(Registry.CurrentUser, subKeyPath);
            DeleteKey(Registry.LocalMachine, subKeyPath);
        }

        private static void DeleteKey(RegistryKey root, string path)
        {
            try { root.DeleteSubKeyTree(path, false); }
            catch { /* 键不存在或无权限，忽略 */ }
        }

        /// <summary>终止特定 Office 相关进程</summary>
        public static void KillOfficeProcesses(ProductType product)
        {
            string[] names;
            switch (product)
            {
                case ProductType.MsOffice:
                    names = new[] {
                        "winword","excel","powerpnt","outlook","onenote","publisher",
                        "infopath","visio","winproj","msaccess","lync","groove",
                        "teams","officeclicktorun","officeintegration","setuphost",
                        "msoev","msosync","msoia","setup"
                    };
                    break;
                case ProductType.Wps:
                    names = new[] { "wps","wpp","et","wpscloudsv","wpscenter","wpscloudsvr" };
                    break;
                case ProductType.Yozo:
                    names = new[] { "yozo_office","yozo","yozooffice","yozoword","yozosheet","yozopresent","yozopresentation","yozo_binder" };
                    break;
                case ProductType.OnlyOffice:
                    names = new[] { "DesktopEditors","ONLYOFFICE" };
                    break;
                case ProductType.LibreOffice:
                    names = new[] { "soffice.bin","soffice.exe" };
                    break;
                default:
                    return;
            }

            foreach (var name in names)
            {
                try
                {
                    foreach (var proc in Process.GetProcessesByName(name))
                    {
                        proc.Kill();
                        proc.WaitForExit(2000);
                        Logger.Info("已终止进程: " + proc.ProcessName);
                    }
                }
                catch { }
            }
        }

        /// <summary>获取当前已安装的特定产品版本</summary>
        public static string GetInstalledProductVersion(ProductType product)
        {
            string[] keywords;
            switch (product)
            {
                case ProductType.MsOffice:
                    keywords = new[] { "Microsoft Office", "Microsoft 365", "Office 16", "Office 15" };
                    break;
                case ProductType.Wps:
                    keywords = new[] { "WPS Office" };
                    break;
                case ProductType.Yozo:
                    keywords = new[] { "永中Office", "Yozo Office", "Yozosoft" };
                    break;
                case ProductType.OnlyOffice:
                    keywords = new[] { "ONLYOFFICE" };
                    break;
                case ProductType.LibreOffice:
                    keywords = new[] { "LibreOffice" };
                    break;
                default:
                    return null;
            }

            var basePaths = new[] {
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
            };
            var hives = new[] { Registry.LocalMachine, Registry.CurrentUser };

            foreach (var hive in hives)
            {
                foreach (var bp in basePaths)
                {
                    try
                    {
                        using (var root = hive.OpenSubKey(bp))
                        {
                            if (root == null) continue;
                            foreach (var sub in root.GetSubKeyNames())
                            {
                                try
                                {
                                    using (var key = root.OpenSubKey(sub))
                                    {
                                        if (key == null) continue;
                                        var dn = key.GetValue("DisplayName") as string;
                                        if (string.IsNullOrEmpty(dn)) continue;
                                        
                                        if (product == ProductType.MsOffice && (dn.Contains("Access Runtime") || dn.Contains("Language Pack")))
                                            continue;

                                        foreach (var kw in keywords)
                                        {
                                            if (dn.ToLower().Contains(kw.ToLower()))
                                            {
                                                var dv = key.GetValue("DisplayVersion") as string;
                                                return string.IsNullOrEmpty(dv) ? dn : $"{dn} ({dv})";
                                            }
                                        }
                                    }
                                }
                                catch { }
                            }
                        }
                    }
                    catch { }
                }
            }
            return null;
        }

        /// <summary>停止并删除 ClickToRun 服务</summary>
        public static void RemoveClickToRunService()
        {
            RunSc("stop ClickToRunSvc");
            RunSc("delete ClickToRunSvc");
        }

        private static void RunSc(string args)
        {
            try
            {
                var psi = new ProcessStartInfo("sc.exe", args)
                {
                    CreateNoWindow = true,
                    UseShellExecute = false
                };
                Process.Start(psi)?.WaitForExit(5000);
            }
            catch { }
        }

        /// <summary>清理并深度卸载系统内已注册的特定 Office 卸载项</summary>
        public static void CleanUninstallEntries(ProductType product)
        {
            var basePaths = new[] {
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
            };

            var hives = new[] { Registry.LocalMachine, Registry.CurrentUser };

            foreach (var hive in hives)
            {
                foreach (var bp in basePaths)
                {
                    try
                    {
                        using (var root = hive.OpenSubKey(bp, writable: true))
                        {
                            if (root == null) continue;
                            foreach (var sub in root.GetSubKeyNames())
                            {
                                try
                                {
                                    using (var key = root.OpenSubKey(sub))
                                    {
                                        if (key == null) continue;
                                        var dn = (key.GetValue("DisplayName") as string) ?? "";
                                        var pub = (key.GetValue("Publisher") as string) ?? "";
                                        var us = (key.GetValue("UninstallString") as string) ?? "";

                                        bool isMsOffice = dn.Contains("Microsoft Office") ||
                                                          dn.Contains("Microsoft 365") ||
                                                          dn.Contains("Office 16") ||
                                                          dn.Contains("Office 15") ||
                                                          sub.StartsWith("Office1") ||
                                                          us.Contains("OfficeClickToRun") ||
                                                          (pub.Contains("Microsoft Corporation") && (dn.Contains("Office") || dn.Contains("365")));

                                        bool isWps = dn.Contains("WPS Office") || sub.Contains("WPS Office");
                                        bool isYozo = dn.Contains("永中") || dn.Contains("Yozo");
                                        bool isOnlyOffice = dn.Contains("ONLYOFFICE") || sub.Contains("ONLYOFFICE");
                                        bool isLibreOffice = dn.Contains("LibreOffice") || sub.Contains("LibreOffice");

                                        bool shouldClean = false;
                                        switch (product)
                                        {
                                            case ProductType.MsOffice: shouldClean = isMsOffice; break;
                                            case ProductType.Wps: shouldClean = isWps; break;
                                            case ProductType.Yozo: shouldClean = isYozo; break;
                                            case ProductType.OnlyOffice: shouldClean = isOnlyOffice; break;
                                            case ProductType.LibreOffice: shouldClean = isLibreOffice; break;
                                        }

                                        if (shouldClean)
                                        {
                                            Logger.Info($"发现卸载项: {dn}，准备静默调用卸载器...");

                                            // 如果有卸载字符串，静默调用它
                                            if (!string.IsNullOrEmpty(us))
                                            {
                                                RunUninstaller(us);
                                            }

                                            root.DeleteSubKeyTree(sub, throwOnMissingSubKey: false);
                                            Logger.Info("已清理注册表卸载项: " + dn);
                                        }
                                    }
                                }
                                catch { }
                            }
                        }
                    }
                    catch { }
                }
            }
        }

        private static void RunUninstaller(string uninstallString)
        {
            try
            {
                string cmd = uninstallString.Trim();
                string args = "";

                // 如果是 msi 卸载，把 /I 替换为 /X
                if (cmd.ToLower().Contains("msiexec"))
                {
                    cmd = "msiexec.exe";
                    // 提取 GUID
                    var parts = uninstallString.Split(' ');
                    var guid = parts.FirstOrDefault(p => p.Contains("{") && p.Contains("}"));
                    if (!string.IsNullOrEmpty(guid))
                    {
                        args = $"/x {guid} /qn /norestart";
                    }
                    else
                    {
                        args = uninstallString.Replace("/I", "/X").Replace("/i", "/x") + " /qn /norestart";
                        args = args.Replace("msiexec.exe", "").Replace("msiexec", "").Trim();
                    }
                }
                else
                {
                    // WPS, Yozo, OnlyOffice 均为 exe，通常加上 /S 或 /verysilent
                    if (cmd.StartsWith("\""))
                    {
                        int nextQuote = cmd.IndexOf("\"", 1);
                        if (nextQuote > 0)
                        {
                            args = cmd.Substring(nextQuote + 1).Trim();
                            cmd = cmd.Substring(1, nextQuote - 1);
                        }
                    }
                    else
                    {
                        int space = cmd.IndexOf(' ');
                        if (space > 0)
                        {
                            args = cmd.Substring(space + 1);
                            cmd = cmd.Substring(0, space);
                        }
                    }

                    // 附加静默参数
                    if (args.Contains("/S") || args.Contains("/s") || args.Contains("/silent") || args.Contains("/verysilent"))
                    {
                        // 已经有静默参数，保持原样
                    }
                    else
                    {
                        if (uninstallString.ToLower().Contains("wps") || uninstallString.ToLower().Contains("yozo"))
                            args += " /S";
                        else
                            args += " /VERYSILENT /NORESTART";
                    }
                }

                Logger.Info($"执行静默卸载命令: {cmd} {args}");
                var psi = new ProcessStartInfo(cmd, args)
                {
                    CreateNoWindow = true,
                    UseShellExecute = false
                };
                var proc = Process.Start(psi);
                proc?.WaitForExit(30000); // 最多等待 30 秒
            }
            catch (Exception ex)
            {
                Logger.Warn($"调用卸载命令失败: {uninstallString}, 错误: {ex.Message}");
            }
        }

        /// <summary>删除特定 Office 品牌的残留文件目录</summary>
        public static void CleanResidualFolders(ProductType product)
        {
            var folders = new List<string>();
            switch (product)
            {
                case ProductType.MsOffice:
                    folders.AddRange(new[] {
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Microsoft Office"),
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Microsoft Office"),
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "Microsoft\\Office"),
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft\\Office"),
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft\\Office"),
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Common Files\\microsoft shared\\OFFICE16"),
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Common Files\\microsoft shared\\OFFICE16"),
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Common Files\\Microsoft Shared\\ClickToRun"),
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "Microsoft\\ClickToRun")
                    });
                    break;
                case ProductType.Wps:
                    folders.AddRange(new[] {
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Kingsoft"),
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Kingsoft"),
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Kingsoft"),
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Kingsoft")
                    });
                    break;
                case ProductType.Yozo:
                    folders.AddRange(new[] {
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Yozosoft"),
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Yozosoft"),
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Yozosoft"),
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Yozosoft")
                    });
                    break;
                case ProductType.OnlyOffice:
                    folders.AddRange(new[] {
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "ONLYOFFICE"),
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "ONLYOFFICE"),
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ONLYOFFICE"),
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ONLYOFFICE")
                    });
                    break;
                case ProductType.LibreOffice:
                    folders.AddRange(new[] {
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "LibreOffice"),
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "LibreOffice"),
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "LibreOffice"),
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "LibreOffice")
                    });
                    break;
            }

            foreach (var f in folders)
            {
                try
                {
                    if (Directory.Exists(f))
                    {
                        Directory.Delete(f, recursive: true);
                        Logger.Info("已删除残留目录: " + f);
                    }
                }
                catch (Exception ex)
                {
                    Logger.Warn($"删除残留目录失败: {f}, 错误: {ex.Message}");
                }
            }
        }
    }
}
