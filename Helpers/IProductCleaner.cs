using GOI.Models;

namespace GOI.Helpers
{
    public interface IProductCleaner
    {
        ProductType Product { get; }
        void KillProcesses();
        string GetInstalledVersion();
        void CleanUninstallEntries();
        void CleanResidualFolders();
        void CleanShortcuts();
        void CleanFileAssociations();
        void CleanRegistryKeys();
    }
}
