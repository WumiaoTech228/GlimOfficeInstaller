using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

namespace GOI.Activation
{
    /// <summary>
    /// C2R（Click-to-Run）版本的 Ohook 部署器。
    /// 在 vfs\System（或 SystemX86）中创建符号链接并写入自定义 sppc.dll。
    /// 等效于 MAS 脚本中的 :oh_hookinstall 子程序。
    /// </summary>
    public static class OhookC2RDeployer
    {
        // 创建符号链接（要求 SeCreateSymbolicLinkPrivilege 权限）
        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool CreateSymbolicLink(
            string symlinkFileName, string targetFileName, SymbolicLinkFlags flags);

        [Flags]
        private enum SymbolicLinkFlags : uint
        {
            File = 0,
            Directory = 1,
            AllowUnprivilegedCreate = 0x2
        }

        /// <summary>
        /// 在 C2R 安装上部署 Ohook。
        /// </summary>
        /// <param name="install">Office 安装信息</param>
        /// <param name="dllBytes">打好 PE 补丁的 sppc64.dll 或 sppc32.dll 字节</param>
        public static DeployResult Deploy(OfficeInstallation install, byte[] dllBytes)
        {
            var result = new DeployResult { Phase = "C2R Deploy" };
            var vfsPath = install.VfsPath;

            if (string.IsNullOrEmpty(vfsPath) || !Directory.Exists(vfsPath))
            {
                result.Success = false;
                result.Error = $"vfs 目录不存在: {vfsPath}";
                return result;
            }

            try
            {
                // 1. 清理旧的 hook DLL
                var oldSppcs = Path.Combine(vfsPath, "sppcs.dll");
                var oldSppc = Path.Combine(vfsPath, "sppc.dll");

                foreach (var old in new[] { oldSppcs, oldSppc })
                {
                    try
                    {
                        if (File.Exists(old))
                            File.Delete(old);
                    }
                    catch (Exception ex)
                    {
                        // 如果删除失败，尝试移到临时目录
                        var tempDest = Path.Combine(
                            Path.GetTempPath(), $"goi_old_{Guid.NewGuid():N}.tmp");
                        File.Move(old, tempDest);
                        try { File.Delete(tempDest); } catch { }
                        result.Warnings.Add($"旧文件移动清理: {Path.GetFileName(old)} ({ex.Message})");
                    }

                    if (File.Exists(old))
                    {
                        result.Success = false;
                        result.Error = $"无法删除旧 hook DLL: {Path.GetFileName(old)}";
                        return result;
                    }
                }

                result.Steps.Add("已清理旧 hook DLL");

                // 2. 创建符号链接: sppcs.dll → C:\Windows\System32\sppc.dll
                var systemSppcPath = OhookPathResolver.GetSystemSppcPath(install.Is64BitOffice);
                if (!File.Exists(systemSppcPath))
                {
                    result.Success = false;
                    result.Error = $"系统 sppc.dll 不存在: {systemSppcPath}";
                    return result;
                }

                var symlinkPath = Path.Combine(vfsPath, "sppcs.dll");
                var success = CreateSymbolicLink(
                    symlinkPath, systemSppcPath, SymbolicLinkFlags.File);

                // 如果 mklink 失败（权限不足），尝试降级方案：直接复制
                if (!success)
                {
                    var lastErr = Marshal.GetLastWin32Error();
                    if (lastErr == 1314) // ERROR_PRIVILEGE_NOT_HELD
                    {
                        File.Copy(systemSppcPath, symlinkPath, overwrite: true);
                        result.Warnings.Add("符号链接权限不足，已降级为文件复制");
                    }
                    else
                    {
                        result.Success = false;
                        result.Error = $"创建符号链接失败 (win32 error: {lastErr})";
                        return result;
                    }
                }
                result.Steps.Add($"符号链接创建: sppcs.dll → {systemSppcPath}");

                // 3. 写入自定义 sppc.dll
                var targetSppc = Path.Combine(vfsPath, "sppc.dll");
                File.WriteAllBytes(targetSppc, dllBytes);
                result.Steps.Add($"写入自定义 DLL: {targetSppc}");

                result.Success = true;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Error = ex.Message;
            }

            return result;
        }

        /// <summary>卸载 C2R Ohook</summary>
        public static DeployResult Undeploy(OfficeInstallation install)
        {
            var result = new DeployResult { Phase = "C2R Undeploy" };
            var vfsPath = install.VfsPath;

            if (string.IsNullOrEmpty(vfsPath) || !Directory.Exists(vfsPath))
            {
                // 已经不存在了，算卸载成功
                result.Success = true;
                result.Steps.Add("vfs 目录已不存在，无需清理");
                return result;
            }

            try
            {
                foreach (var file in new[] { "sppc.dll", "sppcs.dll" })
                {
                    var fullPath = Path.Combine(vfsPath, file);
                    try
                    {
                        if (File.Exists(fullPath))
                        {
                            File.Delete(fullPath);
                            result.Steps.Add($"已删除: {Path.GetFileName(fullPath)}");
                        }
                    }
                    catch
                    {
                        // 尝试移动到临时目录后删除
                        var temp = Path.Combine(
                            Path.GetTempPath(), $"goi_cleanup_{Guid.NewGuid():N}.tmp");
                        File.Move(fullPath, temp);
                        try { File.Delete(temp); } catch { }
                        result.Steps.Add($"已移动清理: {Path.GetFileName(fullPath)}");
                    }
                }

                result.Success = true;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Error = ex.Message;
            }

            return result;
        }
    }
}
