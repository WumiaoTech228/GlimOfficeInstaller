namespace GOI.Models
{
    /// <summary>产品类型（MS Office, WPS, 永中, OnlyOffice 或 LibreOffice）</summary>
    public enum ProductType
    {
        MsOffice,
        Wps,
        Yozo,
        OnlyOffice,
        LibreOffice
    }

    /// <summary>WPS Office 版本</summary>
    public enum WpsVersion
    {
        Wps2013,
        Wps2016,
        Wps2019,
        Wps2023
    }

    /// <summary>永中 Office 版本</summary>
    public enum YozoVersion
    {
        YozoPersonal
    }

    /// <summary>OnlyOffice 版本</summary>
    public enum OnlyOfficeVersion
    {
        OnlyOfficeDesktop
    }

    /// <summary>LibreOffice 版本</summary>
    public enum LibreOfficeVersion
    {
        LibreOfficeStable
    }

    /// <summary>Office 版本枚举</summary>
    public enum OfficeVersion
    {
        Office2024,
        Office2021,
        Office2019,
        Office2016,
        Microsoft365Pro,
        Microsoft365Home
    }

    /// <summary>Office 组件</summary>
    public enum OfficeComponent
    {
        Word,
        Excel,
        PowerPoint,
        Visio,
        Access,
        OneNote,
        Lync,
        Outlook,
        Teams,
        OneDrive,
        Publisher,
        Project
    }

    /// <summary>安装阶段</summary>
    public enum InstallPhase
    {
        Idle,
        Cleaning,
        Downloading,
        Installing,
        Activating,
        Completed,
        Failed
    }

    /// <summary>激活方式</summary>
    public enum ActivationMethod
    {
        Ohook,
        KMS,
        TsForge,
        HWID
    }

    /// <summary>系统架构</summary>
    public enum Architecture
    {
        x86,
        x64,
        ARM64
    }
}
