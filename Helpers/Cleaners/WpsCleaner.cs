using System;
using System.IO;
using GOI.Models;

namespace GOI.Helpers.Cleaners
{
    public class WpsCleaner : IProductCleaner
    {
        public ProductType Product => ProductType.Wps;

        public void KillProcesses()
        {
            string[] names = new string[] { "wps", "wpp", "et", "wpscloudsv", "wpscenter", "wpscloudsvr" };
            RegistryHelper.KillProcessesByName(names);
        }

        public string GetInstalledVersion()
        {
            string[] keywords = new string[] { "WPS Office" };
            return RegistryHelper.GetInstalledVersionFromUninstallKeys(keywords);
        }

        public void CleanUninstallEntries()
        {
            RegistryHelper.CleanUninstallEntriesByFilter(
                (keyName, displayName, publisher) => displayName.Contains("WPS Office") || keyName.Contains("WPS Office"),
                backupRestoreFonts: true
            );
        }

        public void CleanResidualFolders()
        {
            string[] processToWaitFor = new string[] { "wps", "wpp", "et", "wpsuninstall", "uninstall" };
            string[] folders = new string[]
            {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Kingsoft"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Kingsoft"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Kingsoft"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Kingsoft"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "Kingsoft"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "WPS Office"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "WPS Office"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "WPS Office"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "WPS Office"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "WPS Office"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "Microsoft\\Windows\\Start Menu\\Programs\\WPS Office"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft\\Windows\\Start Menu\\Programs\\WPS Office"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Personal), "WPS\\backup"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Personal), "WPS\\template"),
                Path.Combine(Path.GetTempPath(), "wps"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Common Files\\Kingsoft"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Common Files\\Kingsoft"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "AppData\\LocalLow\\Kingsoft")
            };
            string[] registryKeys = new string[]
            {
                "SOFTWARE\\Kingsoft", "SOFTWARE\\WOW6432Node\\Kingsoft", "Software\\Kingsoft",
                "SOFTWARE\\WPS", "SOFTWARE\\WOW6432Node\\WPS", "SOFTWARE\\WPS Office",
                "SOFTWARE\\WOW6432Node\\WPS Office", "Software\\WPS", "Software\\WPS Office"
            };
            RegistryHelper.CleanFoldersAndRegistryKeys(processToWaitFor, folders, registryKeys);
        }

        public void CleanShortcuts()
        {
            string[] names = new string[] { "WPS", "金山" };
            string[] targets = new string[] { "Kingsoft", "WPS Office", "WPSOffice" };
            string[] urls = new string[] { "wps.cn", "kingsoft" };
            RegistryHelper.CleanShortcutsByFilter(names, targets, urls);
        }

        public void CleanFileAssociations()
        {
            string[] progIdPrefixes = new string[]
            {
                "WPS.", "WPP.", "ET.", "KET.", "KWPP.", "KWPS.", "KPDF.", "Kingsoft", "wpsonline", "ksowps",
                "ksoWPSCloudSvr"
            };
            string[] appExecutables = new string[]
            {
                "wps.exe", "et.exe", "wpp.exe", "wpspdf.exe", "wpsoffice.exe", "ksolaunch.exe", "kso.exe", "wpsupdate.exe", "wpsofd.exe", "photolaunch.exe",
                "wpsphoto.exe", "wpsphotos.exe", "ksophoto.exe", "ksophotos.exe"
            };
            RegistryHelper.CleanFileAssociationsByFilter(Product, progIdPrefixes, appExecutables);
        }

        public void CleanRegistryKeys()
        {
            string[] paths = new[]
            {
                @"SOFTWARE\Kingsoft",
                @"SOFTWARE\WOW6432Node\Kingsoft"
            };
            foreach (var p in paths) RegistryHelper.DeleteKey(p);
        }
    }
}
