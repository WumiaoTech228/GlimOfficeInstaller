using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using GOI.Helpers;

namespace GOI.Activation
{
    /// <summary>
    /// Ohook 激活器主入口。
    /// 自动探测已安装的 Office 版本，根据 C2R/MSI 类型选择对应的部署策略，
    /// 完成 DLL 提取、PE 修补、部署的全流程。
    /// </summary>
    public static class OhookActivator
    {
        private static bool IsUserAnAdmin()
        {
            try
            {
                using var identity = System.Security.Principal.WindowsIdentity.GetCurrent();
                var principal = new System.Security.Principal.WindowsPrincipal(identity);
                return principal.IsInRole(System.Security.Principal.WindowsBuiltInRole.Administrator);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 激活本机所有已安装的 Microsoft Office。
        /// </summary>
        /// <param name="progress">进度报告</param>
        /// <param name="ct">取消令牌</param>
        /// <returns>激活结果</returns>
        public static async Task<OhookResult> ActivateAsync(
            IProgress<string> progress = null,
            CancellationToken ct = default)
        {
            var result = new OhookResult();

            try
            {
                if (!IsUserAnAdmin())
                {
                    result.Success = false;
                    result.Error = "检测到当前未以管理员身份运行。请在程序图标上右键选择“以管理员身份运行”以进行激活。";
                    progress?.Report(result.Error);
                    return result;
                }

                progress?.Report("正在探测已安装 of Office 版本...");

                var installations = OhookPathResolver.FindAllInstallations();
                if (installations.Count == 0)
                {
                    result.Success = false;
                    result.Error = "未检测到已安装的 Microsoft Office。请先安装 Office 后再使用激活功能。";
                    Logger.Warn(result.Error);
                    return result;
                }

                progress?.Report($"检测到 {installations.Count} 个 Office 安装");

                foreach (var installation in installations)
                {
                    ct.ThrowIfCancellationRequested();

                    var label = $"Office {installation.Version} ({installation.Type}, {installation.Architecture})";
                    progress?.Report($"正在处理: {label}");

                    // 提取并获取对应架构的原始文件时间戳进行克隆修补，避免自曝
                    var dllBytes = OhookDllExtractor.ExtractForArch(installation.Is64BitOffice);
                    int timestamp;
                    if (installation.Type == OfficeInstallType.C2R)
                    {
                        var systemSppc = OhookPathResolver.GetSystemSppcPath(installation.Is64BitOffice);
                        timestamp = PeTimestampPatcher.ReadTimestamp(systemSppc);
                    }
                    else
                    {
                        var osppPath = OhookPathResolver.GetOsppPath() ?? OhookPathResolver.GetCommonOfficeSharedPath();
                        var originalMsiDll = Path.Combine(osppPath, "sppcs.dll");
                        if (!File.Exists(originalMsiDll))
                        {
                            originalMsiDll = Path.Combine(osppPath, "OSPPC.DLL");
                        }
                        timestamp = PeTimestampPatcher.ReadTimestamp(originalMsiDll);
                    }
                    dllBytes = PeTimestampPatcher.Patch(dllBytes, timestamp);

                    // 部署
                    DeployResult deployResult;
                    if (installation.Type == OfficeInstallType.C2R)
                    {
                        deployResult = OhookC2RDeployer.Deploy(installation, dllBytes);
                    }
                    else // MSI
                    {
                        // MSI 是通过 OSPP 路径操作，同时在 Office 根目录下建立 sppcs.dll 符号链接以供 Office 加载器重定向
                        var osppInst = new OfficeInstallation
                        {
                            Type = OfficeInstallType.MSI,
                            VfsPath = installation.RootPath, // 在 MSI 部署中，符号链接创建在 Office 安装的 RootPath 下
                            Is64BitOffice = installation.Is64BitOffice
                        };
                        deployResult = OhookMsiDeployer.Deploy(osppInst, dllBytes);
                    }

                    foreach (var step in deployResult.Steps)
                    {
                        Logger.Info($"[Ohook] {label}: {step}");
                        result.Steps.Add($"{label}: {step}");
                    }
                    foreach (var warn in deployResult.Warnings)
                    {
                        Logger.Warn($"[Ohook] {label}: {warn}");
                        result.Warnings.Add($"{label}: {warn}");
                    }

                    if (deployResult.Success)
                    {
                        result.Success = true;
                        result.ActivatedInstallations.Add(label);
                    }
                    else
                    {
                        result.FailedInstallations.Add($"{label}: {deployResult.Error}");
                        Logger.Error($"[Ohook] {label} 部署失败: {deployResult.Error}");
                    }
                }

                if (result.ActivatedInstallations.Count > 0)
                    progress?.Report($"激活完成！成功处理 {result.ActivatedInstallations.Count} 个 Office 安装。");
                else
                    progress?.Report("激活失败：未能处理任何 Office 安装。");
            }
            catch (OperationCanceledException)
            {
                result.Success = false;
                result.Error = "操作已取消";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Error = $"激活过程发生异常: {ex.Message}";
                Logger.Error("[Ohook] 激活异常", ex);
            }

            return result;
        }

        /// <summary>
        /// 卸载本机所有 Ohook 激活。
        /// </summary>
        public static async Task<OhookResult> DeactivateAsync(
            IProgress<string> progress = null,
            CancellationToken ct = default)
        {
            var result = new OhookResult();

            try
            {
                if (!IsUserAnAdmin())
                {
                    result.Success = false;
                    result.Error = "检测到当前未以管理员身份运行。请在程序图标上右键选择“以管理员身份运行”以进行卸载。";
                    progress?.Report(result.Error);
                    return result;
                }

                progress?.Report("正在探测已安装 of Office 版本...");

                var installations = OhookPathResolver.FindAllInstallations();
                if (installations.Count == 0)
                {
                    result.Success = true; // 没有 Office，卸载等于已成功
                    result.Steps.Add("未检测到 Office，无需清理");
                    return result;
                }

                foreach (var installation in installations)
                {
                    ct.ThrowIfCancellationRequested();

                    var label = $"Office {installation.Version} ({installation.Type})";
                    progress?.Report($"正在卸载: {label}");

                    DeployResult deployResult;
                    if (installation.Type == OfficeInstallType.C2R)
                    {
                        deployResult = OhookC2RDeployer.Undeploy(installation);
                    }
                    else
                    {
                        deployResult = OhookMsiDeployer.Undeploy(installation);
                    }

                    foreach (var step in deployResult.Steps)
                        result.Steps.Add($"{label}: {step}");

                    if (deployResult.Success)
                    {
                        result.Success = true;
                        result.ActivatedInstallations.Add(label);
                    }
                    else
                    {
                        result.FailedInstallations.Add($"{label}: {deployResult.Error}");
                    }
                }

                progress?.Report("Ohook 卸载完成。");
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Error = $"卸载过程发生异常: {ex.Message}";
                Logger.Error("[Ohook] 卸载异常", ex);
            }

            return result;
        }
    }

    /// <summary>Ohook 激活/卸载结果</summary>
    public class OhookResult
    {
        public bool Success { get; set; }
        public string Error { get; set; }
        public List<string> Steps { get; set; } = new List<string>();
        public List<string> Warnings { get; set; } = new List<string>();
        public List<string> ActivatedInstallations { get; set; } = new List<string>();
        public List<string> FailedInstallations { get; set; } = new List<string>();
    }
}
