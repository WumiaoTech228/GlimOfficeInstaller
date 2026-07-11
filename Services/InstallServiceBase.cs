using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using GOI.Helpers;

namespace GOI.Services
{
    public abstract class InstallServiceBase
    {
        /// <summary>在指定时间内将进度从 from 匀速推进到 to，可被取消</summary>
        protected static async Task FakeProgressAsync(
            IProgress<int> progress, int from, int to, int durationMs, CancellationToken ct)
        {
            int steps = to - from;
            int intervalMs = steps > 0 ? durationMs / steps : durationMs;
            for (int i = from; i <= to && !ct.IsCancellationRequested; i++)
            {
                progress?.Report(i);
                await Task.Delay(intervalMs, ct).ContinueWith(_ => { }); // 忽略取消异常
            }
        }

        /// <summary>安全删除临时文件，避免抛出异常</summary>
        protected static void SafeDeleteFile(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                }
            }
            catch (Exception ex)
            {
                Logger.Warn($"无法删除临时文件: {filePath}, 错误: {ex.Message}");
            }
        }

        /// <summary>安全删除临时文件夹，避免抛出异常</summary>
        protected static void SafeDeleteDirectory(string dirPath)
        {
            try
            {
                if (Directory.Exists(dirPath))
                {
                    Directory.Delete(dirPath, true);
                }
            }
            catch (Exception ex)
            {
                Logger.Warn($"无法删除临时文件夹: {dirPath}, 错误: {ex.Message}");
            }
        }
    }
}
