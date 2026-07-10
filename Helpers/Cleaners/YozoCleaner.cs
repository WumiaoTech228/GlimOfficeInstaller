using System;
using System.IO;
using GOI.Models;

namespace GOI.Helpers.Cleaners
{
    public class YozoCleaner : IProductCleaner
    {
        public ProductType Product => ProductType.Yozo;

        public void KillProcesses()
        {
            string[] names = new string[] { "yozo_office", "yozo", "yozooffice", "yozoword", "yozosheet", "yozopresent", "yozopresentation", "yozo_binder" };
            RegistryHelper.KillProcessesByName(names);
        }

        public string GetInstalledVersion()
        {
            string[] keywords = new string[] { "永中Office", "Yozo Office", "Yozosoft" };
            return RegistryHelper.GetInstalledVersionFromUninstallKeys(keywords);
        }

        public void CleanUninstallEntries()
        {
            RegistryHelper.CleanUninstallEntriesByFilter(
                (keyName, displayName, publisher) => displayName.Contains("永中") || displayName.Contains("Yozo"),
                backupRestoreFonts: false
            );
        }

        public void CleanResidualFolders()
        {
            string[] processToWaitFor = new string[] { "yozo", "uninstall" };
            string path = Environment.GetEnvironmentVariable("USERPROFILE") ?? "";
            string[] folders = new string[]
            {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Yozosoft"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Yozosoft"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Yozosoft"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Yozosoft"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "Yozosoft"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "Microsoft\\Windows\\Start Menu\\Programs\\永中Office"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft\\Windows\\Start Menu\\Programs\\永中Office"),
                Path.Combine(path, "YozoOffice")
            };
            string[] registryKeys = new string[] { "SOFTWARE\\Yozosoft", "SOFTWARE\\WOW6432Node\\Yozosoft", "Software\\Yozosoft" };
            RegistryHelper.CleanFoldersAndRegistryKeys(processToWaitFor, folders, registryKeys);
        }

        public void CleanShortcuts()
        {
            string[] names = new string[] { "永中", "Yozo" };
            string[] targets = new string[] { "Yozosoft", "Yozo" };
            string[] urls = new string[] { "yozo", "yozosoft" };
            RegistryHelper.CleanShortcutsByFilter(names, targets, urls);
        }

        public void CleanFileAssociations()
        {
            string[] progIdPrefixes = new string[] { "Yozo", "yozoword", "yozosheet", "yozopresent", "yozobinder", "YOO." };
            string[] appExecutables = new string[]
            {
                "yozo.exe", "yozoword.exe", "yozosheet.exe", "yozopresent.exe", "yozobinder.exe", "yozolaunch.exe", "Yozo_Calc.exe", "Yozo_Impress.exe", "yozo_Ofd.exe", "Yozo_Office.exe",
                "Yozo_Writer.exe"
            };
            RegistryHelper.CleanFileAssociationsByFilter(Product, progIdPrefixes, appExecutables);
        }

        public void CleanRegistryKeys()
        {
            string[] paths = new[]
            {
                @"SOFTWARE\Yozosoft",
                @"SOFTWARE\WOW6432Node\Yozosoft"
            };
            foreach (var p in paths) RegistryHelper.DeleteKey(p);
        }
    }
}
