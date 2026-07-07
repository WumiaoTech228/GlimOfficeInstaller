using System;
using System.Threading.Tasks;
using GOI.Helpers;

namespace GOI.Services
{
    public class LicenseService
    {
        /// <summary>修改 Office 授权人信息（用户名/公司）</summary>
        public bool ChangeOwner(string userName, string organization)
        {
            if (string.IsNullOrWhiteSpace(userName)) return false;

            var success = 0;
            var regPaths = new[]
            {
                @"Software\Microsoft\Office\Common\UserInfo",
                @"Software\Microsoft\Office\16.0\Common\UserInfo",
                @"Software\Microsoft\Office\15.0\Common\UserInfo",
                @"Software\Microsoft\Office\14.0\Common\UserInfo",
            };

            try
            {
                // HKCU 路径
                foreach (var path in regPaths)
                {
                    try
                    {
                        using (var key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(path))
                        {
                            if (key == null) continue;
                            key.SetValue("UserName", userName);
                            key.SetValue("Company", organization);
                            key.SetValue("UserInitials", GetInitials(userName));
                            success++;
                        }
                    }
                    catch { }
                }

                // HKLM Windows 注册信息
                try
                {
                    using (var key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(
                        @"SOFTWARE\Microsoft\Windows NT\CurrentVersion", writable: true))
                    {
                        if (key != null)
                        {
                            key.SetValue("RegisteredOwner", userName);
                            key.SetValue("RegisteredOrganization", organization);
                            success++;
                        }
                    }
                }
                catch { }

                Logger.Info($"授权人信息修改完成: {userName}, 成功修改 {success} 处。");
                return success > 0;
            }
            catch (Exception ex)
            {
                Logger.Error("修改授权人信息失败", ex);
                return false;
            }
        }

        private static string GetInitials(string name)
        {
            var parts = name.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length >= 2)
                return (parts[0][0].ToString() + parts[1][0].ToString()).ToUpper();
            return name.Length >= 2 ? name.Substring(0, 2).ToUpper() : name.ToUpper();
        }
    }
}
