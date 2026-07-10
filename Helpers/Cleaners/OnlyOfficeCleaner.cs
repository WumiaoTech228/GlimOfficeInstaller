using System;
using System.IO;
using GOI.Models;

namespace GOI.Helpers.Cleaners
{
    public class OnlyOfficeCleaner : IProductCleaner
    {
        public ProductType Product => ProductType.OnlyOffice;

        public void KillProcesses()
        {
            string[] names = new string[] { "DesktopEditors", "ONLYOFFICE", "editors", "editors_helper", "updatesvc" };
            RegistryHelper.KillProcessesByName(names);
        }

        public string GetInstalledVersion()
        {
            string[] keywords = new string[] { "ONLYOFFICE" };
            return RegistryHelper.GetInstalledVersionFromUninstallKeys(keywords);
        }

        public void CleanUninstallEntries()
        {
            RegistryHelper.CleanUninstallEntriesByFilter(
                (keyName, displayName, publisher) => displayName.Contains("ONLYOFFICE") || keyName.Contains("ONLYOFFICE"),
                backupRestoreFonts: false
            );
        }

        public void CleanResidualFolders()
        {
            string[] processToWaitFor = new string[] { "DesktopEditors", "editors", "editors_helper", "updatesvc", "unins000" };
            string[] folders = new string[]
            {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "ONLYOFFICE"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "ONLYOFFICE"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Ascensio System SIA"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Ascensio System SIA"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ONLYOFFICE"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ONLYOFFICE"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "ONLYOFFICE"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "Microsoft\\Windows\\Start Menu\\Programs\\ONLYOFFICE"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft\\Windows\\Start Menu\\Programs\\ONLYOFFICE")
            };
            string[] registryKeys = new string[]
            {
                "SOFTWARE\\ONLYOFFICE", "SOFTWARE\\WOW6432Node\\ONLYOFFICE", "Software\\ONLYOFFICE",
                "SOFTWARE\\Ascensio System SIA", "SOFTWARE\\WOW6432Node\\Ascensio System SIA", "Software\\Ascensio System SIA"
            };
            RegistryHelper.CleanFoldersAndRegistryKeys(processToWaitFor, folders, registryKeys);
        }

        public void CleanShortcuts()
        {
            string[] names = new string[] { "ONLYOFFICE" };
            string[] targets = new string[] { "ONLYOFFICE" };
            string[] urls = new string[] { "onlyoffice" };
            RegistryHelper.CleanShortcutsByFilter(names, targets, urls);
        }

        public void CleanFileAssociations()
        {
            string[] progIdPrefixes = new string[] { "ONLYOFFICE.", "Ascensio" };
            string[] appExecutables = new string[] { "DesktopEditors.exe", "editors.exe", "editors_helper.exe", "updatesvc.exe" };
            RegistryHelper.CleanFileAssociationsByFilter(Product, progIdPrefixes, appExecutables);
        }

        public void CleanRegistryKeys()
        {
            string[] paths = new[]
            {
                @"SOFTWARE\ONLYOFFICE",
                @"SOFTWARE\WOW6432Node\ONLYOFFICE"
            };
            foreach (var p in paths) RegistryHelper.DeleteKey(p);
        }
    }
}
