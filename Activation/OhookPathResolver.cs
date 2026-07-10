using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Win32;
using GOI.Helpers;

namespace GOI.Activation
{
    /// <summary>
    /// 探测本地已安装的 Microsoft Office 版本及相关路径信息。
    /// 等效于 MAS 脚本中的 :oh_getpath 和 :oh_ppcpath 子程序。
    /// 支持 C2R（即点即用）和 MSI（传统安装）两种安装方式。
    /// </summary>
    public static class OhookPathResolver
    {
        /// <summary>返回所有已安装的 Office 信息</summary>
        public static List<OfficeInstallation> FindAllInstallations()
        {
            var results = new List<OfficeInstallation>();

            // C2R (Click-to-Run) 安装
            var o16c2r = FindC2R(null);           // 16.0 C2R (Office 2016/2019/2021/2024/M365)
            var o15c2r = FindC2R("15.0");          // 15.0 C2R (Office 2013)

            if (o16c2r != null) results.Add(o16c2r);
            if (o15c2r != null) results.Add(o15c2r);

            // MSI 安装
            var o16msi = FindMsi("16.0");          // Office 2016/2019 MSI
            var o15msi = FindMsi("15.0");          // Office 2013 MSI
            var o14msi = FindMsi("14.0");          // Office 2010 MSI

            if (o16msi != null) results.Add(o16msi);
            if (o15msi != null) results.Add(o15msi);
            if (o14msi != null) results.Add(o14msi);

            // 去重（同一路径可能被多个分支命中）
            var seen = new HashSet<string>();
            var distinct = new List<OfficeInstallation>();
            foreach (var inst in results)
            {
                if (seen.Add(inst.RootPath ?? inst.VfsPath ?? Guid.NewGuid().ToString()))
                    distinct.Add(inst);
            }
            return distinct;
        }

        /// <summary>探测 C2R 安装</summary>
        private static OfficeInstallation FindC2R(string version)
        {
            string subKeyPath = version != null
                ? $@"SOFTWARE\Microsoft\Office\{version}\ClickToRun"
                : @"SOFTWARE\Microsoft\Office\ClickToRun";

            // 先查 64 位注册表，再查 32 位
            foreach (var (hive, view) in new[] {
                (RegistryHive.LocalMachine, RegistryView.Registry64),
                (RegistryHive.LocalMachine, RegistryView.Registry32) })
            {
                try
                {
                    using var baseKey = RegistryKey.OpenBaseKey(hive, view);
                    using var key = baseKey.OpenSubKey(subKeyPath);
                    if (key == null) continue;

                    var installPath = key.GetValue("InstallPath") as string;
                    if (string.IsNullOrEmpty(installPath) || !Directory.Exists(installPath))
                        continue;

                    var root = Path.Combine(installPath, "root");

                    // 验证 Licenses 目录存在（C2R 的标志）
                    var license16 = Path.Combine(root, "Licenses16");
                    var license15 = Path.Combine(root, "Licenses");
                    if (!Directory.Exists(license16) && !Directory.Exists(license15))
                        continue;

                    // 读取架构和版本
                    var arch = "x64";
                    var v = "16.0";
                    using (var configKey = key.OpenSubKey("Configuration"))
                    {
                        if (configKey != null)
                        {
                            var platform = configKey.GetValue("Platform") as string;
                            if (!string.IsNullOrEmpty(platform))
                                arch = platform;

                            var ver = configKey.GetValue("VersionToReport") as string;
                            if (!string.IsNullOrEmpty(ver))
                                v = ver;
                        }
                    }

                    // 兼容旧版 propertyBag
                    if (arch == "x64")
                    {
                        using (var propBagKey = key.OpenSubKey("propertyBag"))
                        {
                            if (propBagKey != null)
                            {
                                var platform = propBagKey.GetValue("Platform") as string;
                                if (!string.IsNullOrEmpty(platform))
                                    arch = platform;
                            }
                        }
                    }

                    var is64 = arch.Equals("x64", StringComparison.OrdinalIgnoreCase);
                    var vfsSystem = is64 ? "System" : "SystemX86";

                    return new OfficeInstallation
                    {
                        Type = OfficeInstallType.C2R,
                        Version = v,
                        Architecture = arch,
                        RootPath = root,
                        VfsPath = Path.Combine(root, "vfs", vfsSystem),
                        LicensePath = Directory.Exists(license16) ? license16 : license15,
                        Is64BitOffice = is64,
                        RegistryKeyPath = subKeyPath
                    };
                }
                catch { }
            }

            return null;
        }

        /// <summary>探测 MSI 安装</summary>
        private static OfficeInstallation FindMsi(string version)
        {
            foreach (var (hive, view) in new[] {
                (RegistryHive.LocalMachine, RegistryView.Registry64),
                (RegistryHive.LocalMachine, RegistryView.Registry32) })
            {
                try
                {
                    var subKeyPath = $@"SOFTWARE\Microsoft\Office\{version}\Common\InstallRoot";
                    using var baseKey = RegistryKey.OpenBaseKey(hive, view);
                    using var key = baseKey.OpenSubKey(subKeyPath);
                    if (key == null) continue;

                    var path = key.GetValue("Path") as string;
                    if (string.IsNullOrEmpty(path) || !Directory.Exists(path))
                        continue;

                    // 验证 *Picker.dll 存在（MSI 的标志）
                    var pickerFiles = Directory.GetFiles(path, "*Picker.dll");
                    if (pickerFiles.Length == 0)
                        continue;

                    return new OfficeInstallation
                    {
                        Type = OfficeInstallType.MSI,
                        Version = version,
                        RootPath = path,
                        VfsPath = null, // MSI 没有 vfs，走 OSPP 路径
                        LicensePath = null,
                        Is64BitOffice = view == RegistryView.Registry64,
                        RegistryKeyPath = subKeyPath
                    };
                }
                catch { }
            }

            return null;
        }

        /// <summary>获取系统 sppc.dll 路径</summary>
        public static string GetSystemSppcPath(bool is64BitOffice)
        {
            var sysDir = is64BitOffice
                ? Environment.GetFolderPath(Environment.SpecialFolder.System)
                : Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                    "SysWOW64");

            return Path.Combine(sysDir, "sppc.dll");
        }

        /// <summary>获取 OSPP（Office Software Protection Platform）路径</summary>
        public static string GetOsppPath()
        {
            foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
            {
                try
                {
                    using var baseKey = RegistryKey.OpenBaseKey(
                        RegistryHive.LocalMachine, view);
                    using var key = baseKey.OpenSubKey(
                        @"SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform");
                    if (key == null) continue;

                    var path = key.GetValue("Path") as string;
                    if (!string.IsNullOrEmpty(path) && Directory.Exists(path))
                        return path.TrimEnd('\\');
                }
                catch { }
            }
            return null;
        }

        /// <summary>判断是否为 OSPP 模式（传统 Office 2010 及部分 2013）</summary>
        public static bool IsOsppMode(OfficeInstallation install)
        {
            if (install.Type != OfficeInstallType.MSI)
                return false;

            // Office 14.0 一定是 OSPP
            if (install.Version?.StartsWith("14") == true)
                return true;

            // Win7/8 环境且 Office 版本 ≤ 15 时走 OSPP
            if (Environment.OSVersion.Version.Major < 10)
                return true;

            return false;
        }

        /// <summary>获取 Common Files\Microsoft Shared 路径</summary>
        public static string GetCommonOfficeSharedPath()
        {
            var commonFiles = Environment.GetFolderPath(
                Environment.SpecialFolder.CommonProgramFiles);
            return Path.Combine(commonFiles, "Microsoft Shared", "OfficeSoftwareProtectionPlatform");
        }

        /// <summary>获取当前已安装的 Office 产品 ID 列表</summary>
        public static List<string> GetInstalledProductIds()
        {
            var productIds = new List<string>();
            try
            {
                // 先在 64 位下查，再查 32 位
                foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
                {
                    using var baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view);
                    using var key = baseKey.OpenSubKey(@"SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs");
                    if (key == null) continue;

                    foreach (var subKeyName in key.GetSubKeyNames())
                    {
                        using var subKey = key.OpenSubKey(subKeyName);
                        if (subKey == null) continue;

                        foreach (var prodName in subKey.GetSubKeyNames())
                        {
                            // 排除 "culture" 等非产品子键
                            if (prodName.Equals("culture", StringComparison.OrdinalIgnoreCase))
                                continue;

                            string cleanName = prodName;
                            if (cleanName.EndsWith(".16", StringComparison.OrdinalIgnoreCase))
                            {
                                cleanName = cleanName.Substring(0, cleanName.Length - 3);
                            }
                            if (!productIds.Contains(cleanName))
                            {
                                productIds.Add(cleanName);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("[Ohook] 获取安装产品列表失败", ex);
            }
            return productIds;
        }
    }

    /// <summary>Office 安装信息</summary>
    public class OfficeInstallation
    {
        /// <summary>C2R 或 MSI</summary>
        public OfficeInstallType Type { get; set; }

        /// <summary>"16.0" / "15.0" / "14.0"</summary>
        public string Version { get; set; }

        /// <summary>"x64" 或 "x86"</summary>
        public string Architecture { get; set; }

        /// <summary>Office 安装根目录（含 root 子目录或 MSI 根目录）</summary>
        public string RootPath { get; set; }

        /// <summary>vfs\System 或 vfs\SystemX86（仅 C2R，MSI 为 null）</summary>
        public string VfsPath { get; set; }

        /// <summary>Licenses16 或 Licenses 目录（仅 C2R）</summary>
        public string LicensePath { get; set; }

        public bool Is64BitOffice { get; set; }

        public string RegistryKeyPath { get; set; }
    }

    /// <summary>Office 安装类型</summary>
    public enum OfficeInstallType
    {
        C2R,    // Click-to-Run（Office 2013+ 即点即用）
        MSI     // 传统 MSI 安装（Office 2010/2013 及部分 2016+）
    }
}
