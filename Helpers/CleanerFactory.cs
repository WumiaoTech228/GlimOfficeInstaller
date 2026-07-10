using System.Collections.Generic;
using GOI.Models;
using GOI.Helpers.Cleaners;

namespace GOI.Helpers
{
    public static class CleanerFactory
    {
        private static readonly Dictionary<ProductType, IProductCleaner> Cleaners = new Dictionary<ProductType, IProductCleaner>
        {
            { ProductType.MsOffice, new MsOfficeCleaner() },
            { ProductType.Wps, new WpsCleaner() },
            { ProductType.Yozo, new YozoCleaner() },
            { ProductType.OnlyOffice, new OnlyOfficeCleaner() },
            { ProductType.LibreOffice, new LibreOfficeCleaner() }
        };

        public static IProductCleaner GetCleaner(ProductType product)
        {
            if (Cleaners.TryGetValue(product, out var cleaner))
            {
                return cleaner;
            }
            return null;
        }
    }
}
