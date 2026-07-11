using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Win32;
using GOI.Helpers;

namespace GOI.Activation
{
    public class OfficeProductKeyInfo
    {
        public string VolumeLicenseName { get; set; }
        public string GvlkKey { get; set; }
        public string SkuId { get; set; }
    }

    public static class OhookKeyMapper
    {
        public static OfficeProductKeyInfo GetKeyInfo(string productId)
        {
            string id = productId.ToLowerInvariant();

            // Microsoft 365 / Subscription
            if (id.Contains("o365homeprem"))
            {
                return new OfficeProductKeyInfo 
                { 
                    VolumeLicenseName = "O365HomePremRetail", 
                    GvlkKey = "3NMDC-G7C3W-68RGP-CB4MH-4CXCH", 
                    SkuId = "a96f8dae-da54-4fad-bdc6-108da592707a" 
                };
            }
            if (id.Contains("o365proplus") || id.Contains("o365enterprise"))
            {
                return new OfficeProductKeyInfo 
                { 
                    VolumeLicenseName = "O365ProPlusRetail", 
                    GvlkKey = "H8DN8-Y2YP3-CR9JT-DHDR9-C7GP3", 
                    SkuId = "e3dacc06-3bc2-4e13-8e59-8e05f3232325" 
                };
            }
            if (id.Contains("o365business"))
            {
                return new OfficeProductKeyInfo 
                { 
                    VolumeLicenseName = "O365BusinessRetail", 
                    GvlkKey = "Y9NF9-M2QWD-FF6RJ-QJW36-RRF2T", 
                    SkuId = "742178ed-6b28-42dd-b3d7-b7c0ea78741b" 
                };
            }
            if (id.Contains("o365educloud"))
            {
                return new OfficeProductKeyInfo 
                { 
                    VolumeLicenseName = "O365EduCloudRetail", 
                    GvlkKey = "W62NQ-267QR-RTF74-PF2MH-JQMTH", 
                    SkuId = "2f5c71b4-5b7a-4005-bb68-f9fac26f2ea3" 
                };
            }
            if (id.Contains("mondo"))
            {
                return new OfficeProductKeyInfo 
                { 
                    VolumeLicenseName = "MondoVolume", 
                    GvlkKey = "FMTQQ-84NR8-2744R-MXF4P-PGYR3", 
                    SkuId = "2cd0ea7e-749f-4288-a05e-567c573b2a6c" 
                };
            }

            // Office 2024
            if (id.Contains("2024"))
            {
                if (id.Contains("projectpro"))
                    return new OfficeProductKeyInfo { VolumeLicenseName = "ProjectPro2024Volume", GvlkKey = "G2NDM-JQRXF-TGWF6-Y7WTH-K3TC2", SkuId = "2141d341-41aa-4e45-9ca1-201e117d6495" };
                if (id.Contains("visiopro"))
                    return new OfficeProductKeyInfo { VolumeLicenseName = "VisioPro2024Volume", GvlkKey = "DR23B-NFB84-GF4CF-2B7GQ-X2234", SkuId = "4c2f32bf-9d0b-4d8c-8ab1-b4c6a0b9992d" };
                
                return new OfficeProductKeyInfo { VolumeLicenseName = "ProPlus2024Volume", GvlkKey = "2TVEK-N4Q8C-9W8XX-M9DFX-8V244", SkuId = "d77244dc-2b82-4f0a-b8ae-1fca00b7f3e2" };
            }

            // Office 2021
            if (id.Contains("2021"))
            {
                if (id.Contains("projectpro"))
                    return new OfficeProductKeyInfo { VolumeLicenseName = "ProjectPro2021Volume", GvlkKey = "HVC34-CVNPG-RVCMT-X2JRF-CR7RK", SkuId = "17739068-86c4-4924-8633-1e529abc7efc" };
                if (id.Contains("visiopro"))
                    return new OfficeProductKeyInfo { VolumeLicenseName = "VisioPro2021Volume", GvlkKey = "JNKBX-MH9P4-K8YYV-8CG2Y-VQ2C8", SkuId = "c590605a-a08a-4cc7-8dc2-f1ffb3d06949" };

                return new OfficeProductKeyInfo { VolumeLicenseName = "ProPlus2021Volume", GvlkKey = "FXYTK-NJJ8C-GB6XX-M87M8-8V2BP", SkuId = "3f180b30-9b05-4fe2-aa8d-0c1c4790f811" };
            }

            // Office 2019
            if (id.Contains("2019"))
            {
                if (id.Contains("projectpro"))
                    return new OfficeProductKeyInfo { VolumeLicenseName = "ProjectPro2019Volume", GvlkKey = "XM2V9-DNJM2-4B93W-3G3YC-CD8WW", SkuId = "fca53cfc-b26a-4c2f-8700-1c3905a92cfb" };
                if (id.Contains("visiopro"))
                    return new OfficeProductKeyInfo { VolumeLicenseName = "VisioPro2019Volume", GvlkKey = "9BGNQ-K37YR-RQHF2-38RQ3-7VCBB", SkuId = "f41abf81-f409-4b0d-889d-92b3e3d7d005" };

                return new OfficeProductKeyInfo { VolumeLicenseName = "ProPlus2019Volume", GvlkKey = "NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP", SkuId = "85dd8b2c-2387-4089-9d79-056e45a50f95" };
            }

            // Office 2016 / Default
            {
                if (id.Contains("projectpro"))
                    return new OfficeProductKeyInfo { VolumeLicenseName = "ProjectProVolume", GvlkKey = "YG9NW-3K39V-2T3HJ-93F3Q-G83KT", SkuId = "82f502b5-b0b0-4349-bd2c-c560df85b248" };
                if (id.Contains("visiopro"))
                    return new OfficeProductKeyInfo { VolumeLicenseName = "VisioProVolume", GvlkKey = "PD3PC-RHNGV-FXJ29-8JK7D-RJRJK", SkuId = "295b2c03-4b1c-4221-b292-1411f468bd02" };

                return new OfficeProductKeyInfo { VolumeLicenseName = "ProPlusVolume", GvlkKey = "XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99", SkuId = "c47456e3-265d-47b6-8ca0-c30abbd0ca36" };
            }
        }
    }

    /// <summary>
    /// Ohook 激活器主入口。
    /// 自动探测已安装的 Office 版本，根据 C2R/MSI 类型选择对应的部署策略，
    /// 完成 DLL 提取、PE 修补、部署、证书密钥转换安装与 KMS 激活的全流程。
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

        private static void ConfigureKmsHost()
        {
            try
            {
                foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
                {
                    using var baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view);
                    using var key = baseKey.CreateSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\0ff1ce15-a989-479d-af46-f275c6370663");
                    if (key != null)
                    {
                        key.SetValue("KeyManagementServiceName", "10.0.0.10", RegistryValueKind.String);
                    }
                }
                Logger.Info("[Ohook] 注册表 KMS 主机地址已配置为 10.0.0.10");
            }
            catch (Exception ex)
            {
                Logger.Error("[Ohook] 写入注册表 KMS 主机配置失败", ex);
            }
        }

        private static bool RunIntegrator(string rootPath, string volumeLicenseName, string gvlkKey)
        {
            try
            {
                string integratorPath = Path.Combine(rootPath, "integration", "integrator.exe");
                if (!File.Exists(integratorPath))
                {
                    Logger.Warn($"[Ohook] integrator.exe 未找到: {integratorPath}");
                    return false;
                }

                string args = $"/I /License PRIDName={volumeLicenseName}.16 PidKey={gvlkKey}";
                Logger.Info($"[Ohook] 运行 integrator.exe 转换授权证书: {args}");

                var psi = new ProcessStartInfo(integratorPath, args)
                {
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    WindowStyle = ProcessWindowStyle.Hidden
                };

                using (var proc = Process.Start(psi))
                {
                    if (proc != null)
                    {
                        proc.WaitForExit();
                        Logger.Info($"[Ohook] integrator.exe 执行完毕，退出码: {proc.ExitCode}");
                        return proc.ExitCode == 0;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("[Ohook] 运行 integrator.exe 发生异常", ex);
            }
            return false;
        }

        private class PowerShellResult
        {
            public int ExitCode { get; set; }
            public string Output { get; set; }
            public string Error { get; set; }
        }

        private static PowerShellResult RunPowerShell(string script)
        {
            var result = new PowerShellResult { ExitCode = -1, Output = string.Empty, Error = string.Empty };
            try
            {
                byte[] bytes = System.Text.Encoding.Unicode.GetBytes(script);
                string base64 = Convert.ToBase64String(bytes);

                var psi = new ProcessStartInfo("powershell.exe", $"-NoProfile -NonInteractive -EncodedCommand {base64}")
                {
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    WindowStyle = ProcessWindowStyle.Hidden,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };

                using (var proc = Process.Start(psi))
                {
                    if (proc != null)
                    {
                        result.Output = proc.StandardOutput.ReadToEnd();
                        result.Error = proc.StandardError.ReadToEnd();
                        proc.WaitForExit();
                        result.ExitCode = proc.ExitCode;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("[Ohook] 执行 PowerShell 发生异常", ex);
                result.Error = ex.Message;
            }
            return result;
        }

        /// <summary>
        /// 通过 WMI SoftwareLicensingService 安装产品密钥。
        /// </summary>
        private static bool WmiInstallProductKey(string productKey)
        {
            Logger.Info($"[Ohook] WMI 安装产品密钥: {productKey}");
            string script = $@"
try {{
    $service = Get-WmiObject -Class SoftwareLicensingService
    $service.InstallProductKey('{productKey}') | Out-Null
    exit 0
}} catch {{
    Write-Error $_.Exception.Message
    exit 1
}}";
            var res = RunPowerShell(script);
            Logger.Info($"[Ohook] WmiInstallProductKey 退出码: {res.ExitCode}");
            if (!string.IsNullOrWhiteSpace(res.Output)) Logger.Info($"[Ohook] stdout: {res.Output.Trim()}");
            if (!string.IsNullOrWhiteSpace(res.Error)) Logger.Warn($"[Ohook] stderr: {res.Error.Trim()}");
            return res.ExitCode == 0;
        }

        /// <summary>
        /// 刷新 SoftwareLicensingService 许可状态
        /// </summary>
        private static void WmiRefreshLicenseStatus()
        {
            Logger.Info("[Ohook] 刷新许可状态...");
            string script = @"
try {
    $service = Get-WmiObject -Class SoftwareLicensingService
    $service.RefreshLicenseStatus() | Out-Null
} catch {
    Write-Error $_.Exception.Message
}";
            var res = RunPowerShell(script);
            Logger.Info($"[Ohook] WmiRefreshLicenseStatus 退出码: {res.ExitCode}");
            if (!string.IsNullOrWhiteSpace(res.Error)) Logger.Warn($"[Ohook] stderr: {res.Error.Trim()}");
        }

        /// <summary>
        /// 卸载与当前激活 SKU ID 冲突的所有其它 Office 产品密钥，以保证激活成功
        /// </summary>
        private static void WmiUninstallConflictKeys(List<string> activeSkuIds)
        {
            Logger.Info("[Ohook] 正在卸载冲突的零售/订阅等其它 Office 产品密钥...");
            string joinedActiveSkuIds = string.Join("','", activeSkuIds);
            string script = $@"
$officeAppId = '0ff1ce15-a989-479d-af46-f275c6370663'
$activeSkuIds = @('{joinedActiveSkuIds}')
$products = Get-WmiObject -Query ""SELECT ID, Name, PartialProductKey FROM SoftwareLicensingProduct WHERE ApplicationId='$officeAppId' AND PartialProductKey IS NOT NULL""
foreach ($p in $products) {{
    if ($activeSkuIds -notcontains $p.ID) {{
        Write-Output ""Uninstalling conflict key for: $($p.Name) (ID: $($p.ID))""
        try {{
            $p.UninstallProductKey() | Out-Null
            Write-Output ""  -> Success""
        }} catch {{
            Write-Output ""  -> Failed: $($_.Exception.Message)""
        }}
    }}
}}";
            var res = RunPowerShell(script);
            Logger.Info($"[Ohook] WMI 卸载冲突密钥输出:\n{res.Output}");
            if (!string.IsNullOrWhiteSpace(res.Error)) Logger.Warn($"[Ohook] stderr: {res.Error.Trim()}");
        }

        /// <summary>
        /// 清理 vNext/订阅/设备等许可缓存与注册表项（等效于 MAS 的 oh_clearblock）
        /// 如果不清理，Office 365 会优先加载本地缓存的账户授权，而忽略 sppc.dll/Ohook 的本地 KMS 激活。
        /// </summary>
        private static void WmiClearVNextLicenseCache()
        {
            Logger.Info("[Ohook] 正在清理 vNext/订阅授权缓存与注册表项...");
            string script = @"
$regKeys = @(
    'HKLM:\SOFTWARE\Microsoft\Office\16.0\Common\Licensing',
    'HKLM:\SOFTWARE\Microsoft\Office\15.0\Common\Licensing',
    'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Office\16.0\Common\Licensing',
    'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Office\15.0\Common\Licensing'
)
foreach ($k in $regKeys) {
    if (Test-Path $k) { Remove-Item -Path $k -Recurse -Force -ErrorAction SilentlyContinue }
}

$pdFolder = ""$env:ProgramData\Microsoft\Office\Licenses""
if (Test-Path $pdFolder) { Remove-Item -Path $pdFolder -Recurse -Force -ErrorAction SilentlyContinue }

$usersFolder = 'C:\Users'
if (Test-Path $usersFolder) {
    $profiles = Get-ChildItem -Path $usersFolder -Directory
    foreach ($p in $profiles) {
        $paths = @(
            ""$($p.FullName)\AppData\Local\Microsoft\Office\Licenses"",
            ""$($p.FullName)\AppData\Local\Microsoft\Office\16.0\Licensing"",
            ""$($p.FullName)\AppData\Local\Microsoft\Office\15.0\Licensing""
        )
        foreach ($folder in $paths) {
            if (Test-Path $folder) { Remove-Item -Path $folder -Recurse -Force -ErrorAction SilentlyContinue }
        }
    }
}

try {
    $hkuPaths = Get-ChildItem -Path 'Registry::HKEY_USERS' -ErrorAction SilentlyContinue
    foreach ($userKey in $hkuPaths) {
        $uk = ""Registry::HKEY_USERS\$($userKey.PSChildName)\Software\Microsoft\Office""
        foreach ($ver in @('16.0', '15.0')) {
            $sub = ""$uk\$ver\Common\Licensing""
            if (Test-Path $sub) { Remove-Item -Path $sub -Recurse -Force -ErrorAction SilentlyContinue }
        }
    }
} catch {}
";
            var res = RunPowerShell(script);
            Logger.Info($"[Ohook] WMI 清理许可缓存完成，退出码: {res.ExitCode}");
            if (!string.IsNullOrWhiteSpace(res.Error)) Logger.Warn($"[Ohook] stderr: {res.Error.Trim()}");
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

                progress?.Report("正在探测已安装的 Office 版本...");

                var installations = OhookPathResolver.FindAllInstallations();
                if (installations.Count == 0)
                {
                    result.Success = false;
                    result.Error = "未检测到已安装的 Microsoft Office。请先安装 Office 后再使用激活功能。";
                    Logger.Warn(result.Error);
                    return result;
                }

                progress?.Report($"检测到 {installations.Count} 个 Office 安装");

                var activeSkuIds = new List<string>();

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

                    // 1. 部署钩子文件
                    DeployResult deployResult;
                    if (installation.Type == OfficeInstallType.C2R)
                    {
                        deployResult = OhookC2RDeployer.Deploy(installation, dllBytes);
                    }
                    else // MSI
                    {
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
                        progress?.Report($"{label}: 正在配置 KMS 密钥与授权...");

                        // 2. 配置 KMS 注册表主机地址为 10.0.0.10
                        ConfigureKmsHost();

                        // 3. 探测当前安装的产品 ID 并进行转换与密钥安装
                        var installedProductIds = OhookPathResolver.GetInstalledProductIds();
                        Logger.Info($"[Ohook] 探测到安装的 Office 产品 ID: {string.Join(", ", installedProductIds)}");

                        foreach (var productId in installedProductIds)
                        {
                            var keyInfo = OhookKeyMapper.GetKeyInfo(productId);
                            if (keyInfo != null)
                            {
                                if (installation.Type == OfficeInstallType.C2R)
                                {
                                    // 运行 integrator 转换零售版为量产版并安装证书与 GVLK/SubTest 密钥
                                    RunIntegrator(installation.RootPath, keyInfo.VolumeLicenseName, keyInfo.GvlkKey);
                                }

                                // 通过 WMI SoftwareLicensingService 注册产品密钥
                                WmiInstallProductKey(keyInfo.GvlkKey);

                                // 记录合法的 SKU ID 用于卸载冲突密钥
                                if (!activeSkuIds.Contains(keyInfo.SkuId))
                                {
                                    activeSkuIds.Add(keyInfo.SkuId);
                                }
                            }
                        }

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
                {
                    // 4. 清理残留与冲突的其它零售版密钥，以避免授权混乱
                    WmiUninstallConflictKeys(activeSkuIds);

                    // 新增：清理 vNext / 订阅授权缓存与注册表项以防拦截失效
                    WmiClearVNextLicenseCache();

                    // 5. 刷新许可状态以应用生效
                    WmiRefreshLicenseStatus();

                    progress?.Report($"激活完成！成功处理 {result.ActivatedInstallations.Count} 个 Office 安装。");
                }
                else
                {
                    progress?.Report("激活失败：未能处理任何 Office 安装。");
                }
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

                progress?.Report("正在探测已安装的 Office 版本...");

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

                // 清理注册表 KMS 主机地址
                try
                {
                    foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
                    {
                        using var baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view);
                        using var key = baseKey.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\0ff1ce15-a989-479d-af46-f275c6370663", writable: true);
                        if (key != null)
                        {
                            key.DeleteValue("KeyManagementServiceName", false);
                        }
                    }
                }
                catch { }

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
