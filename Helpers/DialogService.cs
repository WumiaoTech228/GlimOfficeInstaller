using System;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using iNKORE.UI.WPF.Modern.Controls;
using GOI.Helpers;

namespace GOI.Helpers
{
    /// <summary>
    /// 封装基于 iNKORE ContentDialog 的确认/提示弹窗，避免 ViewModel 直接耦合 UI 控件。
    /// </summary>
    public static class DialogService
    {
        public static async Task<bool> ShowConfirmAsync(string title, string content, string primaryText, string closeText)
        {
            ContentDialog dialog = new ContentDialog
            {
                Title = title
            };
            ((ContentControl)dialog).Content = content;
            dialog.PrimaryButtonText = primaryText;
            dialog.CloseButtonText = closeText;
            dialog.DefaultButton = ContentDialogButton.Primary;
            if (Application.Current?.MainWindow != null)
            {
                dialog.Owner = Application.Current.MainWindow;
            }
            return (int)(await dialog.ShowAsync()) == 1;
        }

        public static async Task ShowMessageAsync(string title, string content, string closeButtonText = null)
        {
            ContentDialog dialog = new ContentDialog
            {
                Title = title
            };
            ((ContentControl)dialog).Content = content;
            dialog.CloseButtonText = closeButtonText ?? LocalizationStrings.Instance.BtnOk;
            if (Application.Current?.MainWindow != null)
            {
                dialog.Owner = Application.Current.MainWindow;
            }
            await dialog.ShowAsync();
        }

        public static async Task HandleFailureAsync(string title, string content)
        {
            await ShowMessageAsync(title, content + LocalizationStrings.Instance.ErrSubmitLogHint);
            try
            {
                Process.Start(new ProcessStartInfo(Logger.LogFilePath) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                Logger.Warn("启动日志文件查看失败: " + ex.Message);
            }
        }
    }
}
