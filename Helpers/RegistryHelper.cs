using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Microsoft.Win32;

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

        /// <summary>终止所有 Office 相关进程</summary>
        public static void KillOfficeProcesses()
        {
            var names = new[] {
                "winword","excel","powerpnt","outlook","onenote","publisher",
                "infopath","visio","winproj","msaccess","lync","groove",
                "teams","officeclicktorun","officeintegration","setuphost",
                "msoev","msosync","msoia"
            };

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

        /// <summary>清理卸载注册表中 Office 相关条目</summary>
        public static void CleanUninstallEntries()
        {
            var basePaths = new[] {
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
            };

            foreach (var bp in basePaths)
            {
                try
                {
                    using (var root = Registry.LocalMachine.OpenSubKey(bp, writable: true))
                    {
                        if (root == null) continue;
                        foreach (var sub in root.GetSubKeyNames())
                        {
                            try
                            {
                                using (var key = root.OpenSubKey(sub))
                                {
                                    var dn = (key?.GetValue("DisplayName") as string) ?? "";
                                    var pub = (key?.GetValue("Publisher") as string) ?? "";
                                    var us = (key?.GetValue("UninstallString") as string) ?? "";

                                    var isOffice = dn.Contains("Microsoft Office") ||
                                                   dn.Contains("Microsoft 365") ||
                                                   dn.Contains("Office 16") ||
                                                   dn.Contains("Office 15") ||
                                                   sub.StartsWith("Office1") ||
                                                   us.Contains("OfficeClickToRun") ||
                                                   (pub.Contains("Microsoft Corporation") &&
                                                    (dn.Contains("Office") || dn.Contains("365") || dn.Contains("ClickToRun")));

                                    if (isOffice)
                                    {
                                        root.DeleteSubKeyTree(sub, throwOnMissingSubKey: false);
                                        Logger.Info("已清理卸载项: " + dn);
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

        /// <summary>删除 Office 残留文件目录</summary>
        public static void CleanResidualFolders()
        {
            var folders = new List<string>
            {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Microsoft Office"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Microsoft Office"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "Microsoft\\Office"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft\\Office"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft\\Office"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Common Files\\microsoft shared\\OFFICE16"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Common Files\\microsoft shared\\OFFICE16"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Common Files\\Microsoft Shared\\ClickToRun"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "Microsoft\\ClickToRun")
            };

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
                catch (Exception ex) { Logger.Warn($"删除目录失败 [{f}]: {ex.Message}"); }
            }
        }
    }
}
