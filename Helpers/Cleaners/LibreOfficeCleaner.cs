using System;
using System.IO;
using GOI.Models;

namespace GOI.Helpers.Cleaners
{
    public class LibreOfficeCleaner : IProductCleaner
    {
        public ProductType Product => ProductType.LibreOffice;

        public void KillProcesses()
        {
            string[] names = new string[] { "soffice.bin", "soffice.exe" };
            RegistryHelper.KillProcessesByName(names);
        }

        public string GetInstalledVersion()
        {
            string[] keywords = new string[] { "LibreOffice" };
            return RegistryHelper.GetInstalledVersionFromUninstallKeys(keywords);
        }

        public void CleanUninstallEntries()
        {
            RegistryHelper.CleanUninstallEntriesByFilter(
                (keyName, displayName, publisher) => displayName.Contains("LibreOffice") || keyName.Contains("LibreOffice"),
                backupRestoreFonts: false
            );
        }

        public void CleanResidualFolders()
        {
            string[] processToWaitFor = new string[] { "soffice", "soffice.bin", "uninstaller" };
            string[] folders = new string[]
            {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "LibreOffice"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "LibreOffice"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "LibreOffice"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "LibreOffice"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "LibreOffice"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "Microsoft\\Windows\\Start Menu\\Programs\\LibreOffice"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft\\Windows\\Start Menu\\Programs\\LibreOffice")
            };
            string[] registryKeys = new string[]
            {
                "SOFTWARE\\The Document Foundation", "SOFTWARE\\LibreOffice",
                "SOFTWARE\\WOW6432Node\\The Document Foundation", "SOFTWARE\\WOW6432Node\\LibreOffice",
                "Software\\The Document Foundation", "Software\\LibreOffice"
            };
            RegistryHelper.CleanFoldersAndRegistryKeys(processToWaitFor, folders, registryKeys);
        }

        public void CleanShortcuts()
        {
            string[] names = new string[] { "LibreOffice" };
            string[] targets = new string[] { "LibreOffice" };
            string[] urls = new string[] { "libreoffice" };
            RegistryHelper.CleanShortcutsByFilter(names, targets, urls);
        }

        public void CleanFileAssociations()
        {
            string[] progIdPrefixes = new string[] { "LibreOffice.", "soffice." };
            string[] appExecutables = new string[] { "soffice.exe", "scalc.exe", "swriter.exe", "simpress.exe" };
            RegistryHelper.CleanFileAssociationsByFilter(Product, progIdPrefixes, appExecutables);
        }

        public void CleanRegistryKeys()
        {
            string[] paths = new[]
            {
                @"SOFTWARE\The Document Foundation",
                @"SOFTWARE\LibreOffice",
                @"SOFTWARE\WOW6432Node\The Document Foundation",
                @"SOFTWARE\WOW6432Node\LibreOffice"
            };
            foreach (var p in paths) RegistryHelper.DeleteKey(p);
        }
    }
}
