using System;
using System.Windows;

namespace GOI
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            Helpers.AppConfig.Initialize();
            Helpers.ResourceHelper.ExtractAllScripts();
        }
    }
}
