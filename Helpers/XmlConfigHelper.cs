using System;
using System.Text;
using System.Linq;
using System.Collections.Generic;
using GOI.Models;

namespace GOI.Helpers
{
    /// <summary>生成 ODT 使用的 configuration.xml</summary>
    public static class XmlConfigHelper
    {
        /// <summary>根据版本、位数、通道、语言和用户勾选的组件生成 XML 内容</summary>
        public static string Generate(OfficeVersion version, string bitness, string channel, string lang, HashSet<OfficeComponent> selected)
        {
            var sb = new StringBuilder();
            var (_, productId) = GetProductInfo(version);

            sb.AppendLine("<Configuration>");
            sb.AppendLine($"  <Add OfficeClientEdition=\"{bitness}\" Channel=\"{channel}\">");
            sb.AppendLine($"    <Product ID=\"{productId}\">");
            sb.AppendLine($"      <Language ID=\"{lang}\" />");

            // 排除用户没勾选的组件
            var allComponents = Enum.GetValues(typeof(OfficeComponent)).Cast<OfficeComponent>();
            foreach (var c in allComponents)
            {
                if (!selected.Contains(c) && c != OfficeComponent.Visio && c != OfficeComponent.Project)
                    sb.AppendLine($"      <ExcludeApp ID=\"{MapComponentId(c)}\" />");
            }

            sb.AppendLine("    </Product>");

            // Visio / Project 作为独立 Product
            if (selected.Contains(OfficeComponent.Visio))
                sb.AppendLine($"    <Product ID=\"{GetVisioProjectId(version, true)}\"><Language ID=\"{lang}\" /></Product>");
            if (selected.Contains(OfficeComponent.Project))
                sb.AppendLine($"    <Product ID=\"{GetVisioProjectId(version, false)}\"><Language ID=\"{lang}\" /></Product>");

            sb.AppendLine("  </Add>");
            sb.AppendLine("  <Display Level=\"Full\" AcceptEULA=\"TRUE\" />");
            sb.AppendLine("  <Property Name=\"SharedComputerLicensing\" Value=\"0\" />");
            sb.AppendLine("  <Property Name=\"FORCEAPPSHUTDOWN\" Value=\"TRUE\" />");
            sb.AppendLine("  <Property Name=\"DeviceBasedLicensing\" Value=\"0\" />");
            sb.AppendLine("  <Updates Enabled=\"TRUE\" />");
            sb.AppendLine("</Configuration>");

            return sb.ToString();
        }

        /// <summary>根据版本返回 (Channel, ProductID)</summary>
        private static (string channel, string productId) GetProductInfo(OfficeVersion v)
        {
            switch (v)
            {
                case OfficeVersion.Office2024: return ("Current", "ProPlus2024Retail");
                case OfficeVersion.Office2021: return ("Current", "ProPlus2021Retail");
                case OfficeVersion.Office2019: return ("Current", "ProPlus2019Retail");
                case OfficeVersion.Office2016: return ("Current", "ProPlusRetail");
                case OfficeVersion.Microsoft365Pro: return ("Current", "O365ProPlusRetail");
                case OfficeVersion.Microsoft365Home: return ("Current", "O365HomePremRetail");
                default: return ("Current", "ProPlus2024Retail");
            }
        }

        /// <summary>组件枚举 → ODT ExcludeApp ID</summary>
        private static string MapComponentId(OfficeComponent c)
        {
            switch (c)
            {
                case OfficeComponent.Word: return "Word";
                case OfficeComponent.Excel: return "Excel";
                case OfficeComponent.PowerPoint: return "PowerPoint";
                case OfficeComponent.Outlook: return "Outlook";
                case OfficeComponent.OneNote: return "OneNote";
                case OfficeComponent.Access: return "Access";
                case OfficeComponent.Publisher: return "Publisher";
                case OfficeComponent.Lync: return "Lync";
                case OfficeComponent.Teams: return "Teams";
                case OfficeComponent.OneDrive: return "OneDrive";
                case OfficeComponent.Project: return "Project";
                case OfficeComponent.Visio: return "Visio";
                default: return "";
            }
        }

        private static string GetVisioProjectId(OfficeVersion v, bool isVisio)
        {
            var app = isVisio ? "VisioPro" : "ProjectPro";
            switch (v)
            {
                case OfficeVersion.Office2024: return $"{app}2024Retail";
                case OfficeVersion.Office2021: return $"{app}2021Retail";
                case OfficeVersion.Office2019: return $"{app}2019Retail";
                case OfficeVersion.Office2016: return $"{app}Retail";
                default: return $"{app}{GetYearSuffix(v)}Retail";
            }
        }

        private static string GetYearSuffix(OfficeVersion v)
        {
            switch (v)
            {
                case OfficeVersion.Office2024: return "2024";
                case OfficeVersion.Office2021: return "2021";
                case OfficeVersion.Office2019: return "2019";
                default: return "";
            }
        }
    }
}
