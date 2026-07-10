# MAS Ohook 激活脚本技术分析

> 分析对象：`Resources/Ohook_Activation_AIO.cmd`（3362 行，源自 Microsoft-Activation-Scripts v3.12）
> 目标：将整份 167KB 的 CMD/PowerShell 混合脚本重写为纯 C# 实现

---

## 一、Ohook 的工作原理（一句话）

**DLL 劫持（DLL Hijacking）。** Office 启动时加载 `sppc.dll`（C2R 版本）或 `OSPPC.DLL`（MSI 版本）来做许可证验证。Ohook 用一个自定义的、签名校验被篡改过的 `sppc.dll` 替换原始 DLL，让 Office 相信"已激活"。

---

## 二、脚本核心流程拆解

### 2.1 主入口（第 162-166 行）

```
参数分发：
  /Ohook               → _act=1  → 激活流程
  /Ohook-Uninstall     → _rem=1  → 卸载流程
```

### 2.2 路径探测（`:oh_getpath`，第 958-981 行）

**这是整份脚本里跟 C# 重写最相关的一段。** 它探测 6 种 Office 安装类型：

| 类型 | 注册表路径 | 版本 |
|------|-----------|------|
| C2R Office 16.0 | `HKLM\SOFTWARE\Microsoft\Office\ClickToRun\InstallPath` | 2016/M365/2021/2024 |
| C2R Office 15.0 | `HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\InstallPath` | 2013 |
| MSI Office 16.0 | `HKLM\SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot\Path` | 2016+ MSI |
| MSI Office 15.0 | `HKLM\SOFTWARE\Microsoft\Office\15.0\Common\InstallRoot\Path` | 2013 MSI |
| MSI Office 14.0 | `HKLM\SOFTWARE\Microsoft\Office\14.0\Common\InstallRoot\Path` | 2010 MSI |

每种类型都同时查 `HKLM` 和 `HKLM\Wow6432Node`（x86/x64 双注册表视图）。

### 2.3 架构判定（`:oh_ppcpath`，第 1007-1047 行）

**C2R 版本（Office 2016+ 即点即用）：**
- Office x64 → Hook 路径：`{InstallRoot}\root\vfs\System`，DLL：`sppc64.dll`
- Office x86 → Hook 路径：`{InstallRoot}\root\vfs\SystemX86`，DLL：`sppc32.dll`

**MSI 版本（Office 2010/2013 传统安装）：**
- Hook 在 `Common Files\Microsoft Shared\OfficeSoftwareProtectionPlatform\` 里操作
- 重命名 `OSPPC.DLL` → `sppcs.dll`，写入自定义 `OSPPC.DLL`

### 2.4 C2R Hook 安装（`:oh_hookinstall`，第 1133-1173 行）

这是最核心的部署流程：

```
1. 删除旧的 sppcs.dll 和 sppc.dll（清理上次安装残留）
2. 创建符号链接：mklink "{vfs}\sppcs.dll" → "C:\Windows\System32\sppc.dll"
   - 这样真实系统 DLL 被重命名为 sppcs.dll，Office 仍然可以通过它验证签名
3. 从脚本自身提取 base64 编码的自定义 sppc64.dll/sppc32.dll
   - 解码 → 写入 "{vfs}\sppc.dll"
   - 修改 PE 文件时间戳和校验和（绕过签名检测）
4. 结果：Office 加载 sppc.dll（自定义）→ 依赖 sppcs.dll（系统原版）→ 激活成功
```

**关键细节：shellcode 偏移量**
- `sppc32.dll`（32位）：从脚本第 2564 字节开始提取
- `sppc64.dll`（64位）：从脚本第 3076 字节开始提取
- 这两个 DLL 被 base64 编码后嵌入在 CMD 文件内部，解码后直接写入磁盘

### 2.5 MSI Hook 安装（`:oh_hookinstall_ospp`，第 1177-1255 行）

MSI 版本的区别在于它操作 `Common Files` 目录而非 `vfs`：

```
1. 删除旧的 hook DLL（文件大小 < 100KB 的判定为 hook DLL）
2. 如果 sppcs.dll 存在（上次激活残留），移回 OSPPC.DLL
3. 重命名：OSPPC.DLL → sppcs.dll
4. 写入自定义 OSPPC.DLL（从脚本中提取 base64）
5. 创建符号链接：{vfs}\sppcs.dll → {CommonFiles}\sppcs.dll
```

### 2.6 Ohook 卸载（`:oh_uninstall`，第 800-930 行）

反向操作：
```
1. 删除所有 vfs\System*\sppc*.dll（清理 C2R hook）
2. 删除所有 Office*\sppc*.dll（清理 MSI hook）
3. 恢复 OSPPC.DLL：sppcs.dll（> 100KB）→ 重命名为 OSPPC.DLL
4. 删除 OfficeSoftwareProtectionPlatform 中的残留 hook DLL
5. 清理注册表中的 Resiliency 键
```

### 2.7 DLL 提取 + PE 篡改（`:oh_extractdll` + `:hexedit:`，第 3127-3200 行）

这是整份脚本里**唯一必须保留外部依赖**的部分：

```powershell
# 1. 从 bat 文件中读取 base64 编码的 DLL
# 2. 解码为 bytes
# 3. 修改 PE 文件中的时间戳（offset 136 和 2564/3076）
# 4. 使用 imagehlp.dll 的 MapFileAndCheckSum API 重新计算 PE 校验和
# 5. 将校验和写入 PE 文件（offset 216）
# 6. 写入最终 DLL 文件
```

---

## 三、C# 重写方案

### 3.1 文件结构

```
GOI/
├── Activation/
│   ├── OhookActivator.cs          # 主入口：ActivateAsync / DeactivateAsync
│   ├── OhookPathResolver.cs       # 注册表探测：找 Office 安装路径 + 架构
│   ├── OhookC2RDeployer.cs        # C2R 版本 Hook 部署（symlink + DLL）
│   ├── OhookMsiDeployer.cs        # MSI 版本 Hook 部署（OSPPC.DLL 替换）
│   ├── OhookDllExtractor.cs       # 从嵌入资源提取 sppc64/sppc32 DLL
│   ├── PeTimestampPatcher.cs      # 修改 PE 时间戳 + 校验和
│   └── OhookResult.cs             # 结果枚举/错误信息
```

### 3.2 核心类职责

#### `OhookActivator`
```csharp
public static class OhookActivator
{
    // 一步激活
    public static async Task<OhookResult> ActivateAsync(
        IProgress<string> progress = null,
        CancellationToken ct = default);

    // 一步卸载
    public static async Task<OhookResult> DeactivateAsync(
        IProgress<string> progress = null,
        CancellationToken ct = default);
}
```

#### `OhookPathResolver`
```csharp
public static class OhookPathResolver
{
    // 返回所有已安装的 Office 信息
    public static List<OfficeInstallation> FindAllInstallations();
    // 返回 C:\Windows\System32\sppc.dll 路径
    public static string GetSystemSppcPath(bool is64BitOffice);
    // 判断是 OSPP（MSI 传统激活）还是 SPP（现代激活）
    public static bool IsOsppMode(OfficeInstallation install);
}

public class OfficeInstallation
{
    public OfficeInstallType Type;     // C2R_16, C2R_15, MSI_16, MSI_15, MSI_14
    public string RootPath;            // Office 安装根目录
    public string VfsPath;             // vfs\System 或 vfs\SystemX86
    public string Architecture;        // "x64" or "x86"
    public string Version;             // "16.0", "15.0" etc
    public string LicensePath;         // Licenses 目录
    public List<string> ProductIds;    // ProPlus2024Volume, O365ProPlusRetail 等
}
```

#### `OhookC2RDeployer`
```csharp
// 执行 C2R 的 symlink + DLL 注入
// 1. 删除旧 sppc*.dll
// 2. mklink sppcs.dll → System32\sppc.dll  (C# 用 CreateSymbolicLink API)
// 3. 写入自定义 sppc.dll
public static bool Deploy(OfficeInstallation install, OhookDllBytes dll);
```

#### `OhookMsiDeployer`
```csharp
// 执行 MSI 的 OSPPC.DLL 替换
// 1. 恢复旧的 hook（sppcs.dll → OSPPC.DLL）
// 2. OSPPC.DLL → sppcs.dll（重命名原版）
// 3. 写入自定义 OSPPC.DLL
// 4. mklink sppcs.dll → CommonFiles\sppcs.dll
public static bool Deploy(OfficeInstallation install, OhookDllBytes dll);
```

#### `OhookDllExtractor`
```csharp
// 从嵌入资源中提取 sppc32.dll 和 sppc64.dll
// DLL 文件作为 EmbeddedResource 打包在 GOI.exe 中
// 或者可以从 Ohook_Activation_AIO.cmd 中解码 base64
public static class OhookDllExtractor
{
    public static OhookDllBytes ExtractSppc32();
    public static OhookDllBytes ExtractSppc64();
}
```

#### `PeTimestampPatcher`
```csharp
// 等效于 :hexedit: 标签中的 PowerShell 代码
// 使用 C# P/Invoke imagehlp.dll 的 MapFileAndCheckSum
public static class PeTimestampPatcher
{
    [DllImport("imagehlp.dll", CharSet = CharSet.Auto)]
    private static extern int MapFileAndCheckSum(
        string Filename, out int HeaderSum, out int CheckSum);

    public static byte[] PatchTimestampAndChecksum(
        byte[] dllBytes, int exportTimestampOffset);
}
```

### 3.3 保留的外部依赖

| 依赖 | 原因 |
|------|------|
| `imagehlp.dll`（Windows 自带）| PE 校验和计算，C# P/Invoke 调用，不需要额外文件 |
| `sppc32.dll` / `sppc64.dll`（二进制）| 这是 Ohook 的核心 hook DLL，不是代码能替代的。作为 EmbeddedResource 嵌入 GOI.exe |
| `sppc.dll`（系统自带）| `C:\Windows\System32\sppc.dll`，symlink 目标 |

### 3.4 不再需要的

- ❌ `Ohook_Activation_AIO.cmd`（167KB）
- ❌ `cmd.exe` 进程调用
- ❌ `ResourceHelper.cs` 中的 CMD 提取逻辑
- ❌ `InstallService.RunScriptAsync()` 中的 cmd 调用
- ❌ PowerShell 内联脚本（`:hexedit:` 标签里的 70 行 PS）

---

## 四、收益

| 维度 | 之前（CMD） | 之后（C#） |
|------|-----------|-----------|
| 调用方式 | `Process.Start("cmd.exe")`，异步等待、解析退出码 | 直接 `await OhookActivator.ActivateAsync()` |
| 错误处理 | 解析 CMD 输出字符串 | 强类型 `OhookResult` + Exception |
| 进度报告 | 无法获取 | `IProgress<string>` 每步可追踪 |
| 可调试性 | 黑盒 | 可以断点、逐步跟踪 |
| 代码体积 | 3362 行 CMD + 内嵌 PowerShell | ~400 行 C# + 2 个嵌入 DLL |
| 项目一致性 | 与 MVVM/WPF 架构割裂 | 统一 C# 体系 |
| 对 MAS 更新的依赖 | 整个 CMD 文件替换 | 只需更新 2 个二进制 DLL |
