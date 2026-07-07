namespace GOI.Models
{
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
