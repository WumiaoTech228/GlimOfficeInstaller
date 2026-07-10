using System;
using System.IO;
using System.Diagnostics;
using Microsoft.Win32;
using GOI.Models;

namespace GOI.Helpers.Cleaners
{
    public class MsOfficeCleaner : IProductCleaner
    {
        public ProductType Product => ProductType.MsOffice;

        public void KillProcesses()
        {
            string[] names = new string[]
            {
                "winword", "excel", "powerpnt", "outlook", "onenote", "publisher", "infopath", "visio", "winproj", "msaccess",
                "lync", "groove", "teams", "officeclicktorun", "officeintegration", "setuphost", "msoev", "msosync", "msoia", "setup"
            };
            RegistryHelper.KillProcessesByName(names, IsOfficeRelatedProcess);
        }

        private bool IsOfficeRelatedProcess(Process process)
        {
            try
            {
                string name = process.ProcessName.ToLowerInvariant();
                if (name == "setup" || name == "setuphost" || name == "teams" || name == "lync")
                {
                    string filePath = process.MainModule?.FileName?.ToLowerInvariant();
                    if (string.IsNullOrEmpty(filePath)) return false;

                    if (filePath.Contains("microsoft office") || 
                        filePath.Contains("office16") || 
                        filePath.Contains("officeclicktorun") || 
                        filePath.Contains("common files\\microsoft shared\\office16") ||
                        filePath.Contains("odt"))
                    {
                        return true;
                    }

                    var versionInfo = FileVersionInfo.GetVersionInfo(filePath);
                    bool isMicrosoft = versionInfo.CompanyName != null && versionInfo.CompanyName.Contains("Microsoft");
                    bool isOfficeProduct = versionInfo.ProductName != null && 
                                          (versionInfo.ProductName.Contains("Office") || 
                                           versionInfo.ProductName.Contains("Word") || 
                                           versionInfo.ProductName.Contains("Excel") || 
                                           versionInfo.ProductName.Contains("PowerPoint") || 
                                           versionInfo.ProductName.Contains("Click-to-Run"));

                    if (isMicrosoft && isOfficeProduct)
                    {
                        return true;
                    }
                    
                    return false;
                }
            }
            catch
            {
                return false;
            }
            return true;
        }

        public string GetInstalledVersion()
        {
            string[] keywords = new string[] { "Microsoft Office", "Microsoft 365", "Office 16", "Office 15" };
            return RegistryHelper.GetInstalledVersionFromUninstallKeys(keywords, name => 
                !name.Contains("Access Runtime") && !name.Contains("Language Pack")
            );
        }

        public void CleanUninstallEntries()
        {
            RegistryHelper.CleanUninstallEntriesByFilter(IsUninstallItemMatch, backupRestoreFonts: true);
        }

        private bool IsUninstallItemMatch(string keyName, string displayName, string publisher)
        {
            string uninstallString = "";
            try
            {
                using var root = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\" + keyName);
                if (root != null)
                {
                    uninstallString = (root.GetValue("UninstallString") as string) ?? "";
                }
            }
            catch {}
            if (string.IsNullOrEmpty(uninstallString))
            {
                try
                {
                    using var root = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\" + keyName);
                    if (root != null)
                    {
                        uninstallString = (root.GetValue("UninstallString") as string) ?? "";
                    }
                }
                catch {}
            }

            return displayName.Contains("Microsoft Office") || 
                   displayName.Contains("Microsoft 365") || 
                   displayName.Contains("Office 16") || 
                   displayName.Contains("Office 15") || 
                   keyName.StartsWith("Office1") || 
                   uninstallString.Contains("OfficeClickToRun") || 
                   (publisher.Contains("Microsoft Corporation") && (displayName.Contains("Office") || displayName.Contains("365")));
        }

        public void CleanResidualFolders()
        {
            string[] processToWaitFor = new string[] { "winword", "excel", "powerpnt", "officeclicktorun", "setup" };
            string[] folders = new string[]
            {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Microsoft Office"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Microsoft Office"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "Microsoft\\Office"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft\\Office"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft\\Office"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Common Files\\microsoft shared\\OFFICE16"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Common Files\\microsoft shared\\OFFICE16"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Common Files\\Microsoft Shared\\ClickToRun"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "Microsoft\\ClickToRun")
            };
            string[] registryKeys = new string[] { "SOFTWARE\\Microsoft\\Office", "SOFTWARE\\WOW6432Node\\Microsoft\\Office" };
            RegistryHelper.CleanFoldersAndRegistryKeys(processToWaitFor, folders, registryKeys);
        }

        public void CleanShortcuts()
        {
            string[] names = new string[] { "Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Access", "Publisher", "Visio", "Project" };
            string[] targets = new string[] { "Microsoft Office", "Office16", "Office15" };
            string[] urls = new string[] { "office.com", "microsoft" };
            RegistryHelper.CleanShortcutsByFilter(names, targets, urls);
        }

        public void CleanFileAssociations()
        {
            string[] progIdPrefixes = new string[] { "Word.", "Excel.", "PowerPoint.", "Access.", "Outlook.", "OneNote." };
            string[] appExecutables = new string[0];
            RegistryHelper.CleanFileAssociationsByFilter(Product, progIdPrefixes, appExecutables);
        }

        public void CleanRegistryKeys()
        {
            string[] paths = new[]
            {
                @"SOFTWARE\Microsoft\Office",
                @"SOFTWARE\Microsoft\Office\ClickToRun",
                @"SOFTWARE\Microsoft\AppVisv",
                @"SOFTWARE\WOW6432Node\Microsoft\Office",
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe",
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Excel.exe",
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Powerpnt.exe",
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Outlook.exe",
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Onenote.exe",
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Visio.exe",
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winproj.exe",
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Msaccess.exe",
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Mspub.exe",
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Lync.exe",
                @"SOFTWARE\Microsoft\Office\16.0\Registration",
                @"SOFTWARE\Microsoft\Office\15.0\Registration",
                @"SOFTWARE\Microsoft\Office\14.0\Registration",
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\00006109C80000000000000000F01FEC",
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Office16.PROPLUS",
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Office16.VISIOPRO",
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Office16.PROJECTPRO",
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Office16.OUTLOOK"
            };
            foreach (var p in paths) RegistryHelper.DeleteKey(p);
        }
    }
}
