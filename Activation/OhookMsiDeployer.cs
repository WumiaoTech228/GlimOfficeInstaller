using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

namespace GOI.Activation
{
    /// <summary>
    /// MSI（传统安装）版本的 Ohook 部署器。
    /// 在 Common Files\Microsoft Shared\OfficeSoftwareProtectionPlatform 中操作 OSPPC.DLL。
    /// 逻辑：备份原版 OSPPC.DLL → 写入自定义 OSPPC.DLL（实际为 hook DLL）
    /// 等效于 MAS 脚本中的 :oh_hookinstall_ospp 子程序。
    /// </summary>
    public static class OhookMsiDeployer
    {
        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool CreateSymbolicLink(
            string symlinkFileName, string targetFileName, uint flags);

        /// <summary>在 MSI 安装上部署 Ohook</summary>
        public static DeployResult Deploy(OfficeInstallation install, byte[] dllBytes)
        {
            var result = new DeployResult { Phase = "MSI Deploy" };
            var osppPath = OhookPathResolver.GetOsppPath();

            // 如果没有 OSPP 路径，使用默认 Common Files 路径
            if (string.IsNullOrEmpty(osppPath))
            {
                osppPath = OhookPathResolver.GetCommonOfficeSharedPath();
            }

            if (string.IsNullOrEmpty(osppPath) || !Directory.Exists(osppPath))
            {
                result.Success = false;
                result.Error = $"OSPP 目录不存在: {osppPath}";
                return result;
            }

            try
            {
                // 1. 删除旧的 hook DLL（大小 < 100KB 的 OSPPC.DLL/sppcs.dll）
                foreach (var file in new[] { "OSPPC.DLL", "sppcs.dll" })
                {
                    var fullPath = Path.Combine(osppPath, file);
                    try
                    {
                        if (File.Exists(fullPath))
                        {
                            var info = new FileInfo(fullPath);
                            if (info.Length > 0 && info.Length < 100_000)
                            {
                                // 小于 100KB = hook DLL，删除
                                File.Delete(fullPath);
                                result.Steps.Add($"已删除旧 hook: {file}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        result.Warnings.Add($"清理 {file} 失败: {ex.Message}");
                    }
                }

                // 2. 如果 sppcs.dll 存在（> 100KB = 原版备份），恢复为 OSPPC.DLL
                var sppcsBackupPath = Path.Combine(osppPath, "sppcs.dll");
                var osppcPath = Path.Combine(osppPath, "OSPPC.DLL");
                if (File.Exists(sppcsBackupPath) && !File.Exists(osppcPath))
                {
                    var info = new FileInfo(sppcsBackupPath);
                    if (info.Length >= 100_000)
                    {
                        File.Move(sppcsBackupPath, osppcPath);
                        result.Steps.Add("已恢复原版 OSPPC.DLL（从 sppcs.dll 重命名）");
                    }
                }

                // 3. 备份原版 OSPPC.DLL → sppcs.dll
                if (File.Exists(osppcPath))
                {
                    if (File.Exists(sppcsBackupPath))
                    {
                        File.Delete(sppcsBackupPath);
                    }
                    File.Move(osppcPath, sppcsBackupPath);
                    result.Steps.Add("已备份原版 OSPPC.DLL → sppcs.dll");
                }

                // 4. 写入自定义 hook DLL → OSPPC.DLL
                File.WriteAllBytes(osppcPath, dllBytes);
                result.Steps.Add($"已写入 hook DLL: {osppcPath}");

                // 5. 在 Office 根目录下创建 sppcs.dll 符号链接指向 OSPP 下的 sppcs.dll 备份（原版 OSPPC.DLL）。
                // 这是为了使 Office 应用程序在加载 hook OSPPC.DLL 后，能正确将原版函数调用重定向到真实备份（等效于 MAS CMD 原理）。
                if (!string.IsNullOrEmpty(install.VfsPath) && Directory.Exists(install.VfsPath))
                {
                    var symlinkPath = Path.Combine(install.VfsPath, "sppcs.dll");
                    if (!File.Exists(symlinkPath))
                    {
                        var ok = CreateSymbolicLink(
                            symlinkPath, sppcsBackupPath, 0);
                        if (ok)
                            result.Steps.Add($"符号链接创建: sppcs.dll → {sppcsBackupPath}");
                        else
                            result.Warnings.Add($"符号链接创建失败 (error {Marshal.GetLastWin32Error()})");
                    }
                }

                result.Success = true;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Error = $"MSI 部署异常: {ex.Message}";
            }

            return result;
        }

        /// <summary>卸载 MSI Ohook</summary>
        public static DeployResult Undeploy(OfficeInstallation install)
        {
            var result = new DeployResult { Phase = "MSI Undeploy" };
            var osppPath = OhookPathResolver.GetOsppPath()
                ?? OhookPathResolver.GetCommonOfficeSharedPath();

            if (string.IsNullOrEmpty(osppPath) || !Directory.Exists(osppPath))
            {
                result.Success = true;
                result.Steps.Add("OSPP 目录不存在，无需清理");
                return result;
            }

            try
            {
                var osppcPath = Path.Combine(osppPath, "OSPPC.DLL");
                var sppcsBackupPath = Path.Combine(osppPath, "sppcs.dll");

                // 删除 hook OSPPC.DLL（小文件）
                if (File.Exists(osppcPath))
                {
                    var info = new FileInfo(osppcPath);
                    if (info.Length < 100_000)
                    {
                        File.Delete(osppcPath);
                        result.Steps.Add("已删除 hook OSPPC.DLL");
                    }
                }

                // 恢复原版：sppcs.dll → OSPPC.DLL
                if (File.Exists(sppcsBackupPath) && !File.Exists(osppcPath))
                {
                    var info = new FileInfo(sppcsBackupPath);
                    if (info.Length >= 100_000)
                    {
                        File.Move(sppcsBackupPath, osppcPath);
                        result.Steps.Add("已恢复原版 OSPPC.DLL");
                    }
                }

                // 清理 vfs 中的符号链接
                if (!string.IsNullOrEmpty(install.VfsPath))
                {
                    var symlink = Path.Combine(install.VfsPath, "sppcs.dll");
                    try
                    {
                        if (File.Exists(symlink))
                        {
                            var attr = File.GetAttributes(symlink);
                            if (attr.HasFlag(FileAttributes.ReparsePoint))
                            {
                                File.Delete(symlink);
                                result.Steps.Add("已删除 vfs 符号链接 sppcs.dll");
                            }
                        }
                    }
                    catch { }
                }

                result.Success = true;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Error = $"MSI 卸载异常: {ex.Message}";
            }

            return result;
        }
    }
}
