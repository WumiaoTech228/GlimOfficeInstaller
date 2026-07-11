using System;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace GOI.Helpers
{
    public static class UrlConfigHelper
    {
        private const string ConfigUrl = "https://ghproxy.net/https://raw.githubusercontent.com/WumiaoTech228/GlimOfficeInstaller/main/config/urls.json";

        // Default Fallbacks
        public static string Wps2013Url { get; private set; } = "https://share.osbox.top/d/CloudService/WPS%20Pro/WPSPRO2013.exe";
        public static string Wps2016Url { get; private set; } = "https://share.osbox.top/d/CloudService/WPS%20Pro/WPSPRO2016.exe";
        public static string Wps2019Url { get; private set; } = "https://share.osbox.top/d/CloudService/WPS%20Pro/WPSPRO2019.exe";
        public static string Wps2023Url { get; private set; } = "https://share.osbox.top/d/CloudService/WPS%20Pro/WPSPRO2023.exe";
        public static string WpsLatestUrl { get; private set; } = "https://official-package.wpscdn.cn/wps/download/WPS_Setup_26899.exe";
        public static string LibreOfficeUrl { get; private set; } = "https://download.documentfoundation.org/libreoffice/stable/26.2.4/win/x86_64/LibreOffice_26.2.4_Win_x86-64.msi";
        public static string OnlyOfficeUrl { get; private set; } = "https://download.onlyoffice.com/install/desktop/editors/windows/distrib/onlyoffice/DesktopEditors_x64.exe";
        public static string YozoUrl { get; private set; } = "https://dl.yozosoft.com/yozo/project/file/20251224_131531_158622/9.0.6589.101ZH.S1.rar";

        public static async Task SyncAsync()
        {
            try
            {
                Logger.Info("[UrlConfig] 开始从云端同步下载直链...");
                using (var client = new HttpClient())
                {
                    client.Timeout = TimeSpan.FromSeconds(8);
                    string json = await client.GetStringAsync(ConfigUrl);
                    if (string.IsNullOrWhiteSpace(json)) return;

                    var matches = Regex.Matches(json, "\"(\\w+)\"\\s*:\\s*\"([^\"]+)\"");
                    int count = 0;
                    foreach (Match m in matches)
                    {
                        string key = m.Groups[1].Value;
                        string val = m.Groups[2].Value;

                        switch (key)
                        {
                            case "Wps2013Url": Wps2013Url = val; count++; break;
                            case "Wps2016Url": Wps2016Url = val; count++; break;
                            case "Wps2019Url": Wps2019Url = val; count++; break;
                            case "Wps2023Url": Wps2023Url = val; count++; break;
                            case "WpsLatestUrl": WpsLatestUrl = val; count++; break;
                            case "LibreOfficeUrl": LibreOfficeUrl = val; count++; break;
                            case "OnlyOfficeUrl": OnlyOfficeUrl = val; count++; break;
                            case "YozoUrl": YozoUrl = val; count++; break;
                        }
                    }
                    Logger.Info($"[UrlConfig] 云端同步完成，成功更新了 {count} 个直链。");
                }
            }
            catch (Exception ex)
            {
                Logger.Warn($"[UrlConfig] 同步云端直链失败，将继续使用本地硬编码默认链接: {ex.Message}");
            }
        }
    }
}
