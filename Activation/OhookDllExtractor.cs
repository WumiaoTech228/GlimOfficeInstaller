using System;
using System.IO;
using System.Reflection;

namespace GOI.Activation
{
    /// <summary>
    /// 从嵌入资源中提取 Ohook 钩子 DLL（sppc32.dll / sppc64.dll）。
    /// 这两个 DLL 是 MAS 项目的核心二进制 payload，负责拦截 Office 的许可证验证调用。
    /// 它们从 Ohook_Activation_AIO.cmd 中解码后作为 EmbeddedResource 编译进 GOI.exe。
    /// </summary>
    public static class OhookDllExtractor
    {
        private const string Sppc32ResourceName = "GOI.Activation.sppc32.dll";
        private const string Sppc64ResourceName = "GOI.Activation.sppc64.dll";

        /// <summary>提取 32 位 hook DLL</summary>
        public static byte[] ExtractSppc32()
        {
            return ExtractResource(Sppc32ResourceName, "sppc32.dll");
        }

        /// <summary>提取 64 位 hook DLL</summary>
        public static byte[] ExtractSppc64()
        {
            return ExtractResource(Sppc64ResourceName, "sppc64.dll");
        }

        /// <summary>根据架构选择正确的 DLL</summary>
        public static byte[] ExtractForArch(bool is64Bit)
        {
            return ExtractResource(
                is64Bit ? Sppc64ResourceName : Sppc32ResourceName,
                is64Bit ? "sppc64.dll" : "sppc32.dll");
        }

        private static byte[] ExtractResource(string resourceName, string displayName)
        {
            var asm = Assembly.GetExecutingAssembly();
            using var stream = asm.GetManifestResourceStream(resourceName);
            if (stream == null)
                throw new InvalidOperationException(
                    $"Ohook DLL 资源未找到: {resourceName}。请确保 Activation/sppc32.dll 和 sppc64.dll 的 Build Action 设为 EmbeddedResource。");

            var buffer = new byte[stream.Length];
            stream.Read(buffer, 0, buffer.Length);

            if (buffer.Length < 1024)
                throw new InvalidOperationException(
                    $"{displayName} 解码异常（仅 {buffer.Length} 字节），应为完整 PE 文件。");

            return buffer;
        }
    }
}
