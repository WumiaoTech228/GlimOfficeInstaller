using System;
using System.IO;
using System.Runtime.InteropServices;
using GOI.Helpers;

namespace GOI.Activation
{
    /// <summary>
    /// 修改 PE（Portable Executable）文件的时间戳和校验和。
    /// 解析 PE 头部，动态查找 COFF 时间戳与导出表时间戳，并进行克隆修补。
    /// 之后重新计算 PE 校验和，躲过 Windows 的安全性检测。
    /// </summary>
    public static class PeTimestampPatcher
    {
        [DllImport("imagehlp.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int MapFileAndCheckSum(
            string filename, out int headerSum, out int checkSum);

        /// <summary>
        /// 读取已存在文件的 PE 时间戳
        /// </summary>
        public static int ReadTimestamp(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    byte[] header = new byte[512];
                    using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                    {
                        fs.Read(header, 0, header.Length);
                    }
                    int e_lfanew = BitConverter.ToInt32(header, 0x3C);
                    if (e_lfanew > 0 && e_lfanew + 12 <= header.Length)
                    {
                        return BitConverter.ToInt32(header, e_lfanew + 8);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"[Ohook] 读取原始文件时间戳失败: {filePath}", ex);
            }
            // 降级：如果文件不存在或读取失败，使用当前 Unix 时间戳
            return (int)(DateTime.UtcNow - new DateTime(1970, 1, 1)).TotalSeconds;
        }

        /// <summary>
        /// 动态解析 PE，将指定的时间戳打入 COFF 头部与导出表时间戳，并计算新校验和
        /// </summary>
        public static byte[] Patch(byte[] dllBytes, int timestamp)
        {
            if (dllBytes == null || dllBytes.Length < 512)
                throw new ArgumentException("DLL 字节数组太小，无法解析 PE");

            var result = new byte[dllBytes.Length];
            Array.Copy(dllBytes, result, dllBytes.Length);

            try
            {
                // 1. 获取 e_lfanew
                int e_lfanew = BitConverter.ToInt32(result, 0x3C);
                if (e_lfanew <= 0 || e_lfanew + 250 >= result.Length)
                    throw new InvalidOperationException("无效的 PE 格式: 错误的 e_lfanew 值");

                // 2. 写入 COFF 时间戳 (e_lfanew + 4 是 PE Signature, 后面是 COFF 头，TimeDateStamp 在 COFF 偏移 4)
                int coffTimestampOffset = e_lfanew + 4 + 4;
                BitConverter.GetBytes(timestamp).CopyTo(result, coffTimestampOffset);

                // 3. 动态解析可选头，获取导出表 RVA 偏移
                ushort magic = BitConverter.ToUInt16(result, e_lfanew + 24);
                int dataDirectoriesOffset;
                if (magic == 0x10B) // PE32 (32-bit)
                {
                    dataDirectoriesOffset = e_lfanew + 24 + 96;
                }
                else if (magic == 0x20B) // PE32+ (64-bit)
                {
                    dataDirectoriesOffset = e_lfanew + 24 + 112;
                }
                else
                {
                    throw new InvalidOperationException($"不支持的 PE 架构 Magic: 0x{magic:X}");
                }

                uint exportRva = BitConverter.ToUInt32(result, dataDirectoriesOffset);
                if (exportRva > 0)
                {
                    // 将 Export RVA 转换为文件偏移
                    int exportFileOffset = RvaToFileOffset(result, exportRva, e_lfanew);
                    if (exportFileOffset > 0 && exportFileOffset + 8 <= result.Length)
                    {
                        // 导出表的 TimeDateStamp 在导出表结构体的偏移 4 字节处
                        int exportTimestampOffset = exportFileOffset + 4;
                        BitConverter.GetBytes(timestamp).CopyTo(result, exportTimestampOffset);
                    }
                }

                // 4. 将修改后的字节写入临时文件，使用 imagehlp 重新计算校验和
                var tempPath = Path.Combine(
                    Path.GetTempPath(), $"goi_pe_temp_{Guid.NewGuid():N}.dll");
                try
                {
                    File.WriteAllBytes(tempPath, result);

                    var ret = MapFileAndCheckSum(tempPath, out int headerSum, out int newChecksum);
                    if (ret != 0)
                        throw new InvalidOperationException(
                            $"MapFileAndCheckSum 失败 (error code: {Marshal.GetLastWin32Error()})");

                    int coffChecksumOffset = e_lfanew + 24 + (magic == 0x10B ? 64 : 80);
                    if (headerSum != newChecksum)
                    {
                        BitConverter.GetBytes(newChecksum).CopyTo(result, coffChecksumOffset);
                    }

                    return result;
                }
                finally
                {
                    try { File.Delete(tempPath); } catch { }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("[Ohook] PE 时间戳与校验和修补失败，将返回未修补的字节", ex);
                return dllBytes;
            }
        }

        private static int RvaToFileOffset(byte[] bytes, uint rva, int e_lfanew)
        {
            ushort numberOfSections = BitConverter.ToUInt16(bytes, e_lfanew + 6);
            ushort sizeOfOptionalHeader = BitConverter.ToUInt16(bytes, e_lfanew + 20);
            int sectionHeadersOffset = e_lfanew + 24 + sizeOfOptionalHeader;

            for (int i = 0; i < numberOfSections; i++)
            {
                int sectionOffset = sectionHeadersOffset + (i * 40);
                if (sectionOffset + 40 > bytes.Length) break;

                uint virtualSize = BitConverter.ToUInt32(bytes, sectionOffset + 8);
                uint virtualAddress = BitConverter.ToUInt32(bytes, sectionOffset + 12);
                uint sizeOfRawData = BitConverter.ToUInt32(bytes, sectionOffset + 16);
                uint pointerToRawData = BitConverter.ToUInt32(bytes, sectionOffset + 20);

                if (rva >= virtualAddress && rva < virtualAddress + Math.Max(virtualSize, sizeOfRawData))
                {
                    return (int)(pointerToRawData + (rva - virtualAddress));
                }
            }
            return -1;
        }
    }
}
