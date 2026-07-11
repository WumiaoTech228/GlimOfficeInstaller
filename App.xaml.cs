using System;
using System.Windows;

namespace GOI
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            
            // 启用 TLS 1.2 和 TLS 1.3 支持，以解决现代 CDN (如 WPS 官方下载直链) 握手失败问题
            System.Net.ServicePointManager.SecurityProtocol = 
                System.Net.SecurityProtocolType.Tls11 | 
                System.Net.SecurityProtocolType.Tls12 | 
                (System.Net.SecurityProtocolType)12288; // TLS 1.3
            
            Helpers.AppConfig.Initialize();
            _ = Helpers.UrlConfigHelper.SyncAsync();
        }

        protected override void OnExit(ExitEventArgs e)
        {
            try
            {
                if (System.IO.Directory.Exists(Helpers.AppConfig.RootPath))
                {
                    System.IO.Directory.Delete(Helpers.AppConfig.RootPath, true);
                }
            }
            catch (Exception ex)
            {
                Helpers.Logger.Warn("退出清理临时目录失败: " + ex.Message);
            }
            base.OnExit(e);
        }
    }
}
