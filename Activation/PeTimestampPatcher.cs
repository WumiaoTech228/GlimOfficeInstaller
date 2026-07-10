using System;
using System.IO;
using System.Runtime.InteropServices;

namespace GOI.Activation
{
    /// <summary>
    /// 修改 PE（Portable Executable）文件的时间戳和校验和。
    /// 等效于 MAS 脚本 :hexedit: 标签中的 PowerShell 代码。
    /// 不修改校验和的 DLL 会被 Windows 拒绝加载。
    /// </summary>
    public static class PeTimestampPatcher
    {
        // PE 文件标准结构:
        //   offset 136 (0x88):  COFF 文件头中的 TimeDateStamp (4 bytes)
        //   offset 216 (0xD8):  COFF 可选头中的 CheckSum (4 bytes)
        //   offset 2564/3076:   导出表时间戳（用于 sppc32/sppc64，具体位置由参数传入）
        private const int CoffTimestampOffset = 136;
        private const int CoffChecksumOffset = 216;

        [DllImport("imagehlp.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int MapFileAndCheckSum(
            string filename, out int headerSum, out int checkSum);

        /// <summary>
        /// 修改 DLL 字节的 PE 时间戳并重新计算校验和。
        /// </summary>
        /// <param name="dllBytes">原始 DLL 字节数组</param>
        /// <param name="exportTimestampOffset">
        /// 导出表时间戳的 PE 文件偏移量。
        /// sppc32.dll: 2564
        /// sppc64.dll: 3076
        /// </param>
        /// <returns>包含更新后时间戳和校验和的 DLL 字节数组</returns>
        public static byte[] Patch(byte[] dllBytes, int exportTimestampOffset)
        {
            if (dllBytes == null || dllBytes.Length < CoffChecksumOffset + 4)
                throw new ArgumentException("DLL 文件太小，无法进行 PE 修补");

            var result = new byte[dllBytes.Length];
            Array.Copy(dllBytes, result, dllBytes.Length);

            // 1. 生成新的时间戳
            var unixTimestamp = (int)(DateTime.UtcNow - new DateTime(1970, 1, 1)).TotalSeconds;

            // 2. 写入 COFF 时间戳
            BitConverter.GetBytes(unixTimestamp).CopyTo(result, CoffTimestampOffset);

            // 3. 写入导出表时间戳
            if (exportTimestampOffset > 0 && exportTimestampOffset + 4 <= result.Length)
            {
                BitConverter.GetBytes(unixTimestamp).CopyTo(result, exportTimestampOffset);
            }

            // 4. 写入临时文件以调用 imagehlp 计算校验和
            var tempPath = Path.Combine(
                Path.GetTempPath(), $"goi_pe_temp_{Guid.NewGuid():N}.dll");
            try
            {
                File.WriteAllBytes(tempPath, result);

                var ret = MapFileAndCheckSum(tempPath, out int headerSum, out int newChecksum);
                if (ret != 0)
                    throw new InvalidOperationException(
                        $"MapFileAndCheckSum 失败 (error code: {Marshal.GetLastWin32Error()})");

                if (headerSum != newChecksum)
                {
                    BitConverter.GetBytes(newChecksum).CopyTo(result, CoffChecksumOffset);
                }
                // 如果 headerSum == newChecksum，说明校验和已经是新的，不需要再写

                return result;
            }
            finally
            {
                try { File.Delete(tempPath); } catch { }
            }
        }
    }
}
