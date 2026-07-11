using System;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace GOI.Helpers
{
    public static class UrlConfigHelper
    {
        private static readonly string[] ConfigUrls = new string[]
        {
            "https://ghproxy.net/https://raw.githubusercontent.com/WumiaoTech228/GlimOfficeInstaller/main/config/urls.json",
            "https://raw.githubusercontent.com/WumiaoTech228/GlimOfficeInstaller/main/config/urls.json",
            "https://cdn.jsdelivr.net/gh/WumiaoTech228/GlimOfficeInstaller@main/config/urls.json"
        };

        // 默认下载直链 (已将 LibreOffice 纠正为中科大官方镜像源)
        public static string Wps2013Url { get; private set; } = "https://share.osbox.top/d/CloudService/WPS%20Pro/WPSPRO2013.exe";
        public static string Wps2016Url { get; private set; } = "https://share.osbox.top/d/CloudService/WPS%20Pro/WPSPRO2016.exe";
        public static string Wps2019Url { get; private set; } = "https://share.osbox.top/d/CloudService/WPS%20Pro/WPSPRO2019.exe";
        public static string Wps2023Url { get; private set; } = "https://share.osbox.top/d/CloudService/WPS%20Pro/WPSPRO2023.exe";
        public static string WpsLatestUrl { get; private set; } = "https://official-package.wpscdn.cn/wps/download/WPS_Setup_26899.exe";
        public static string LibreOfficeUrl { get; private set; } = "https://mirrors.ustc.edu.cn/tdf/libreoffice/stable/26.2.4/win/x86_64/LibreOffice_26.2.4_Win_x86-64.msi";
        public static string OnlyOfficeUrl { get; private set; } = "https://download.onlyoffice.com/install/desktop/editors/windows/distrib/onlyoffice/DesktopEditors_x64.exe";
        public static string YozoUrl { get; private set; } = "https://dl.yozosoft.com/yozo/project/file/20251224_131531_158622/9.0.6589.101ZH.S1.rar";

        // 默认版本描述 (云端可同步更新)
        public static string WpsLatestVersionLabel { get; private set; } = "最新版";
        public static string LibreOfficeVersionLabel { get; private set; } = "26.2.4";
        public static string OnlyOfficeVersionLabel { get; private set; } = "最新版";
        public static string YozoVersionLabel { get; private set; } = "9.0.6589";

        // 同步完成事件通知
        public static event Action SyncCompleted;

        public static async Task SyncAsync()
        {
            Logger.Info("[UrlConfig] 开始从云端多源同步直链及版本号...");
            
            foreach (var url in ConfigUrls)
            {
                try
                {
                    Logger.Info($"[UrlConfig] 正在尝试拉取配置源: {url}");
                    using (var client = new HttpClient())
                    {
                        client.Timeout = TimeSpan.FromSeconds(5);
                        string json = await client.GetStringAsync(url);
                        if (!string.IsNullOrWhiteSpace(json))
                        {
                            ParseConfigJson(json);
                            Logger.Info($"[UrlConfig] 云端同步成功，配置源已热更新。");
                            SyncCompleted?.Invoke();
                            return; // 成功后立即返回，不再尝试备用源
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.Warn($"[UrlConfig] 拉取配置源 {url} 失败: {ex.Message}");
                }
            }

            Logger.Warn("[UrlConfig] 所有云端配置源均同步失败，将继续使用本地硬编码默认值运行。");
        }

        private static void ParseConfigJson(string json)
        {
            var matches = Regex.Matches(json, "\"(\\w+)\"\\s*:\\s*\"([^\"]+)\"");
            foreach (Match m in matches)
            {
                string key = m.Groups[1].Value;
                string val = m.Groups[2].Value;

                switch (key)
                {
                    case "Wps2013Url": Wps2013Url = val; break;
                    case "Wps2016Url": Wps2016Url = val; break;
                    case "Wps2019Url": Wps2019Url = val; break;
                    case "Wps2023Url": Wps2023Url = val; break;
                    case "WpsLatestUrl": WpsLatestUrl = val; break;
                    case "LibreOfficeUrl": LibreOfficeUrl = val; break;
                    case "OnlyOfficeUrl": OnlyOfficeUrl = val; break;
                    case "YozoUrl": YozoUrl = val; break;

                    case "WpsLatestVersionLabel": WpsLatestVersionLabel = val; break;
                    case "LibreOfficeVersionLabel": LibreOfficeVersionLabel = val; break;
                    case "OnlyOfficeVersionLabel": OnlyOfficeVersionLabel = val; break;
                    case "YozoVersionLabel": YozoVersionLabel = val; break;
                }
            }
        }
    }
}
