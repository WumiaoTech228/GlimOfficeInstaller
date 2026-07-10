using System.Globalization;
using GOI.Models;

namespace GOI.Helpers
{
    /// <summary>
    /// 应用程序多语言支持。根据 Windows 系统语言自动选择界面语言：
    /// • zh-CN/zh-SG (简体中文 Windows) → 简体中文
    /// • zh-TW/zh-HK/zh-MO (繁體中文 Windows) → 繁體中文
    /// • 其他语言 → English
    /// </summary>
    public class LocalizationStrings
    {
        public enum AppLanguage { SimplifiedChinese, TraditionalChinese, English }
        private static readonly object _langLock = new object();
        private static AppLanguage _detected;
        public static AppLanguage Detected
        {
            get
            {
                lock (_langLock)
                {
                    return _detected;
                }
            }
            set
            {
                lock (_langLock)
                {
                    _detected = value;
                }
            }
        }

        static LocalizationStrings()
        {
            var name = CultureInfo.CurrentUICulture.Name;
            if (name.StartsWith("zh-TW") || name.StartsWith("zh-HK") || name.StartsWith("zh-MO")
                || name.StartsWith("zh-Hant"))
                Detected = AppLanguage.TraditionalChinese;
            else if (name.StartsWith("zh-CN") || name.StartsWith("zh-SG") || name.StartsWith("zh-Hans")
                     || name.StartsWith("zh"))
                Detected = AppLanguage.SimplifiedChinese;
            else
                Detected = AppLanguage.English;
        }

        private static string S(string sc, string tc, string en) => Detected switch
        {
            AppLanguage.SimplifiedChinese => sc,
            AppLanguage.TraditionalChinese => tc,
            _ => en
        };

        // ===================== 通用 =====================
        public string AppTitle => "Glim Office Installer";
        public string AppVersion => S("版本 2.1.4 Stable | 基于 Fluent 2.0 规范深度重构",
                                       "版本 2.1.4 Stable | 基於 Fluent 2.0 規範深度重構",
                                       "Version 2.1.4 Stable | Rebuilt on Fluent 2.0 Design");
        public string BtnDeploy => S("一键开始部署", "一鍵開始部署", "One-Click Deploy");
        public string BtnOk => S("确定", "確定", "OK");
        public string BtnContinue => S("继续部署", "繼續部署", "Continue");
        public string BtnCancel => S("取消", "取消", "Cancel");
        public string BtnUninstall => S("开始卸载", "開始解除安裝", "Uninstall");
        public string BtnClearLicense => S("开始清理", "開始清理", "Clear Now");
        public string BtnActivate => S("配置授权", "本機啟用", "Activate");
        public string BtnExportXml => S("导出 XML", "匯出 XML", "Export XML");

        // ===================== 导航栏 =====================
        public string NavMsOffice => "Microsoft Office";
        public string NavWps => "WPS Office";
        public string NavYozo => S("永中 Office 2024", "永中 Office 2024", "Yozo Office 2024");
        public string NavOnlyOffice => "OnlyOffice";
        public string NavLibreOffice => "LibreOffice";
        public string NavSettings => S("设置", "設定", "Settings");

        // ===================== 状态文字 =====================
        public string StatusReady => S("准备就绪", "準備就緒", "Ready");
        public string InstalledDetected => S("检测到已安装版本", "偵測到已安裝版本", "Installed Version Detected");

        // ===================== MS Office 页 =====================
        public string SelectVersion => S("选择部署版本", "選擇部署版本", "Select Version");
        public string SelectComponents => S("选择需要包含的组件", "選擇需要包含的元件", "Select Components");
        public string AdvancedOptions => S("高级部署选项", "進階部署選項", "Advanced Options");

        // 版本卡片
        public string Office2024Title => "Office 2024";
        public string Office2024Desc => S("最新功能 · 持续更新 (零售版)", "最新功能 · 持續更新 (零售版)", "Latest features · Continuous updates (Retail)");
        public string M365Title => "Microsoft 365";
        public string M365Desc => S("云端同步 · 订阅服务 (个人/家庭版)", "雲端同步 · 訂閱服務 (個人/家庭版)", "Cloud sync · Subscription (Personal/Family)");
        public string Office2021Title => "Office 2021";
        public string Office2021Desc => S("经典版本 · 永久授权 (零售版)", "經典版本 · 永久授權 (零售版)", "Classic · Perpetual license (Retail)");
        public string Office2019Title => "Office 2019";
        public string Office2019Desc => S("广泛兼容 · 稳定可靠 (零售版)", "廣泛相容 · 穩定可靠 (零售版)", "Broad compatibility · Reliable (Retail)");
        public string Office2016Title => "Office 2016";
        public string Office2016Desc => S("旧版兼容 · 极速轻量 (零售版)", "舊版相容 · 極速輕量 (零售版)", "Legacy compat · Lightweight (Retail)");

        // 高级选项标签
        public string LabelUpdateChannel => S("更新通道", "更新通道", "Update Channel");
        public string LabelUpdateChannelDesc => S("控制 Office 获取更新的节奏与版本", "控制 Office 取得更新的節奏與版本", "Controls how Office receives updates");
        public string LabelLanguage => S("Office 语言", "Office 語言", "Office Language");
        public string LabelLanguageDesc => S("Office 应用程序的显示语言", "Office 應用程式的顯示語言", "Display language for Office applications");
        public string LabelBitness => S("安装位数", "安裝位元數", "Architecture");
        public string LabelBitnessDesc => S("推荐 64 位；如需与老旧 32 位外接程序兼容请选 32 位", "建議 64 位元；若需與舊版 32 位元增益集相容請選 32 位元", "64-bit recommended; choose 32-bit for legacy add-in compatibility");
        public string Bit64 => S("64 位（推荐）", "64 位元（建議）", "64-bit (Recommended)");
        public string Bit32 => S("32 位", "32 位元", "32-bit");

        // 更新通道选项
        public string ChannelCurrent => S("当前通道（每月更新）", "目前通道（每月更新）", "Current Channel (Monthly)");
        public string ChannelCurrentPreview => S("当前通道预览", "目前通道預覽", "Current Channel Preview");
        public string ChannelMonthlyEnterprise => S("月度企业通道", "每月企業通道", "Monthly Enterprise Channel");
        public string ChannelSemiAnnual => S("半年企业通道", "半年企業通道", "Semi-Annual Enterprise Channel");
        public string ChannelSemiAnnualPreview => S("半年通道（预览）", "半年通道（預覽）", "Semi-Annual Channel Preview");
        public string ChannelBeta => S("Beta 通道（内部）", "Beta 通道（內部）", "Beta Channel (Insider)");

        // 组件名称
        public string CompWord => S("Word 文字", "Word 文字", "Word");
        public string CompExcel => S("Excel 表格", "Excel 試算表", "Excel");
        public string CompPowerPoint => S("PowerPoint 演示", "PowerPoint 簡報", "PowerPoint");
        public string CompOutlook => S("Outlook 邮箱", "Outlook 電子郵件", "Outlook");
        public string CompOneNote => S("OneNote 笔记", "OneNote 筆記", "OneNote");
        public string CompAccess => S("Access 数据库", "Access 資料庫", "Access");
        public string CompPublisher => S("Publisher 出版", "Publisher 出版", "Publisher");
        public string CompProject => S("Project 项目", "Project 專案", "Project");
        public string CompVisio => S("Visio 绘图", "Visio 繪圖", "Visio");
        public string CompTeams => S("Teams 协作", "Teams 協作", "Teams");
        public string CompOneDrive => S("OneDrive 网盘", "OneDrive 網路磁碟", "OneDrive");

        // ===================== WPS 页 =====================
        public string WpsLatestTitle => S("WPS Office 最新版", "WPS Office 最新版", "WPS Office Latest");
        public string WpsLatestDesc => S("最新官方原版，支持云端同步与团队协作", "最新官方原版，支援雲端同步與團隊協作", "Latest official release, supports cloud sync and collaboration");
        public string Wps2023Title => S("WPS Office 2023 (非官方版本)", "WPS Office 2023 (非官方版本)", "WPS Office 2023 (Non-official)");
        public string Wps2019Title => S("WPS Office 2019 (非官方版本)", "WPS Office 2019 (非官方版本)", "WPS Office 2019 (Non-official)");
        public string Wps2019Desc => S("主版本号 11.8 (经典无广告稳定版本)", "主版本號 11.8 (經典無廣告穩定版本)", "Version 11.8 (Classic ad-free stable version)");
        public string Wps2016Title => S("WPS Office 2016 (非官方版本)", "WPS Office 2016 (非官方版本)", "WPS Office 2016 (Non-official)");
        public string Wps2016Desc => S("主版本号 10.1 (早期轻量兼容版本)", "主版本號 10.1 (早期輕量相容版本)", "Version 10.1 (Early lightweight compatible version)");
        public string Wps2013Title => S("WPS Office 2013 (非官方版本)", "WPS Office 2013 (非官方版本)", "WPS Office 2013 (Non-official)");
        public string Wps2013Desc => S("主版本号 9.1 (低配轻量化版本)", "主版本號 9.1 (低配輕量化版本)", "Version 9.1 (Low-spec lightweight version)");

        // ===================== Yozo 页 =====================
        public string YozoTitle => S("永中 Office 2024", "永中 Office 2024", "Yozo Office 2024");
        public string YozoDesc => S("个人版 (官方渠道下载)", "個人版 (官方管道下載)", "Personal Version (Official Channel)");
        public string YozoFeat1Title => S("主流 Office 格式兼容", "主流 Office 格式相容", "Mainstream Office Format Compatibility");
        public string YozoFeat1Desc => S("支持解析并打开 .docx、.xlsx、.pptx 等格式文档，减少格式排版偏差", "支援解析並打開 .docx、.xlsx、.pptx 等格式文檔，減少格式排版偏差", "Supports parsing and opening .docx, .xlsx, .pptx and other formats with minimized layout deviation");
        public string YozoFeat2Title => S("本地化办公方案", "本地化辦公方案", "Localized Office Solution");
        public string YozoFeat2Desc => S("由永中软件自主研发，不进行境外数据传输，适用于国产化替代场景", "由永中軟體自主研發，不進行境外資料傳輸，適用於國產化替代場景", "Independently developed by Yozo; no overseas data transmission, suitable for domestic software substitution");

        // ===================== OnlyOffice 页 =====================
        public string OnlyOfficeTitle => "OnlyOffice";
        public string OnlyOfficeDesc => S("Desktop Editors (开源无广告官方通道)", "Desktop Editors (開源無廣告官方管道)", "Desktop Editors (Open-Source, Ad-free Official Channel)");
        public string OnlyOfficeFeat1Title => S("协同编辑支持", "協同編輯支援", "Collaborative Editing Support");
        public string OnlyOfficeFeat1Desc => S("支持集成云端协作模块，实现多用户在线协作处理文档", "支援整合雲端協作模組，實現多使用者線上協作處理文檔", "Supports integration with cloud collaboration modules for multi-user online document editing");
        public string OnlyOfficeFeat2Title => S("开源多标签页界面", "開源多標籤頁介面", "Open-Source Multi-Tab Interface");
        public string OnlyOfficeFeat2Desc => S("完全免费且开源的桌面编辑器，支持单窗口多标签文档处理", "完全免費且開源的桌面編輯器，支援單視窗多標籤文檔處理", "Completely free and open-source desktop editor with single-window multi-tab support");

        // ===================== LibreOffice 页 =====================
        public string LibreOfficeTitle => "LibreOffice";
        public string LibreOfficeDesc => S("稳定版 26.2.4 (中科大镜像源)", "穩定版 26.2.4 (中科大鏡像源)", "Stable 26.2.4 (USTC Mirror Source)");
        public string LibreOfficeFeat1Title => S("标准开放文档格式支持", "標準開放文件格式支援", "Standard OpenDocument Format Support");
        public string LibreOfficeFeat1Desc => S("由 Document Foundation 维护，提供对 ODF 国际标准文档格式 (ODT/ODS/ODP) 的完整支持", "由 Document Foundation 維護，提供對 ODF 國際標準文件格式 (ODT/ODS/ODP) 的完整支援", "Maintained by The Document Foundation, providing full support for international ODF standard formats (ODT/ODS/ODP)");
        public string LibreOfficeFeat2Title => S("国内高校镜像源加速", "國內高校鏡像源加速", "Domestic University Mirror Acceleration");
        public string LibreOfficeFeat2Desc => S("从中国科学技术大学 (USTC) 开源镜像站拉取安装包，提升国内网络下载速度", "從中國科學技術大學 (USTC) 開源鏡像站拉取安裝包，提升國內網路下載速度", "Downloads packages from the University of Science and Technology of China (USTC) open-source mirror to improve network speeds");

        // ===================== 设置页 =====================
        public string SettingsPersonalization => S("个性化设置", "個人化設定", "Personalization");
        public string SettingsThemeTitle => S("应用主题", "應用程式主題", "Application Theme");
        public string SettingsThemeDesc => S("更改 GOI 软件的视觉界面模式", "更改 GOI 軟體的視覺介面模式", "Change the visual interface mode of GOI");
        public string SettingsLanguageTitle => S("软件语言", "軟體語言", "App Language");
        public string SettingsLanguageDesc => S("更改 GOI 软件界面的显示语言", "更改 GOI 軟體介面的顯示語言", "Change the display language of GOI");
        public string ThemeSystem => S("跟随系统", "跟隨系統", "System Default");
        public string ThemeLight => S("浅色模式", "淺色模式", "Light Mode");
        public string ThemeDark => S("深色模式", "深色模式", "Dark Mode");
        public string SettingsCleanupSectionTitle => S("Office 深度清理与卸载", "Office 深度清理與解除安裝", "Office Cleanup & Uninstall");
        public string SettingsUninstallTitle => S("深度清理残留", "深度清理殘留", "Deep Cleanup");
        public string SettingsUninstallDesc => S("强制终止所选 Office 进程，并彻底清除注册表与残留文件", "強制終止所選 Office 程序，並徹底清除登錄機碼與殘留檔案", "Force-terminate selected Office processes and thoroughly clean registry and residual files");
        public string SettingsLicenseSectionTitle => S("Office 授权与注册表工具", "Office 授權與登錄機碼工具", "Office License & Registry Tools");
        public string SettingsClearLicenseTitle => S("清除 Office 授权许可证书", "清除 Office 啟用授權資訊", "Clear Office Activation License");
        public string SettingsClearLicenseDesc => S("卸载本地所有 Office 产品授权密钥，重置授权状态", "解除安裝本機所有 Office 產品啟用金鑰，重設啟用狀態", "Uninstall all local Office activation keys and reset activation status");
        public string SettingsOhookTitle => S("一键配置 Office 本地授权 (本地离线)", "一鍵設定 Office 學習啟用 (Ohook 本機離線)", "Configure Office Activation (Ohook Local Offline)");
        public string SettingsOhookDesc => S("部署本地 Ohook DLL 劫持模块，支持 M365 订阅版与零售版（无需网络）", "部署本機 Ohook DLL 攔截模組，支援 M365 訂閱版與零售版（无需網路）", "Deploy local Ohook DLL hook module, supports M365 subscription and retail (no internet)");
        
        public string SettingsFileAssociationSectionTitle => S("文件关联与图标净化", "文件關聯與圖標淨化", "File Association & Icon Cleanup");
        public string SettingsCleanAssociationsTitle => S("清除失效的 Office 文件关联", "清除失效的 Office 文件關聯", "Clean Invalid Office File Associations");
        public string SettingsCleanAssociationsDesc => S("扫描并清除未安装的 WPS、永中等 Office 品牌在注册表残留的 ProgID、右键菜单和死链快捷方式，恢复系统默认", "掃描並清除未安裝的 WPS、永中等 Office 品牌在註冊表殘留的 ProgID、右鍵選單和死鏈快捷方式，恢復系統預設", "Scan and clean leftover ProgIDs, context menus, and dead shortcuts for uninstalled Office brands, restoring system default");
        public string BtnCleanAssociations => S("开始净化", "開始淨化", "Start Purge");

        public string SettingsRefreshIconCacheTitle => S("重建并刷新系统图标缓存", "重建並重新整理系統圖示快取", "Rebuild & Refresh System Icon Cache");
        public string SettingsRefreshIconCacheDesc => S("强制重建 Windows 图标缓存并通知资源管理器，解决文件图标显示异常、变成白色空白纸张或文件关联更改后图标未即时更新的问题", "強制重建 Windows 圖示快取並通知資源管理器，解決檔案圖示顯示異常、變成白色空白紙張或檔案關聯變更後圖示未即時更新的問題", "Force rebuild Windows icon cache and notify Explorer, solving blank icons or delayed icon updates after association changes");
        public string BtnRefreshIconCache => S("立即刷新", "立即重新整理", "Refresh Now");

        public string SettingsRepairAssociationsTitle => S("修复已安装 Office 的文件关联", "修復已安裝 Office 的文件關聯", "Repair File Associations for Installed Office");
        public string SettingsRepairAssociationsDesc => S("自动检测本机当前已安装的 Office 品牌（如 MS Office/WPS），修复并重新关联它们的默认双击打开关系与文件图标", "自動檢測本機當前已安裝的 Office 品牌（如 MS Office/WPS），修復並重新激活他們的默認雙擊打開關係與文件圖標", "Auto-detect currently installed Office suites (e.g. MS Office/WPS) and repair/reactivate their default file association handlers and file icons");
        public string BtnRepairAssociations => S("立即修复", "立即修復", "Repair Now");

        public string SettingsRepairCOMTitle => S("修复并重新注册 Office COM 组件", "修復並重新註冊 Office COM 組件", "Repair & Re-register Office COM Components");
        public string SettingsRepairCOMDesc => S("修复由于组件缺失、卸载残留或路径损坏导致的第三方程序、插件、RPA 自动化工具无法调用 Office (Word/Excel) 接口的顽疾", "修復由於組件缺失、卸載殘留或路徑損壞導致的第三方程序、插件、RPA 自動化工具無法調用 Office (Word/Excel) 接口的頑疾", "Repair Component Object Model (COM) registration for Word, Excel, PowerPoint to fix integration errors with plugins and automation tools");
        public string BtnRepairCOM => S("重新注册", "重新註冊", "Re-register");

        public string SettingsDebugSectionTitle => S("调试工具", "除錯工具", "Debug Tools");
        public string SettingsExportXmlTitle => S("导出所选 Office 的 XML 配置", "匯出所選 Office 的 XML 設定", "Export Selected Office XML Configuration");
        public string SettingsExportXmlDesc => S("将主页中当前选择的 Office 部署设置导出为 XML 文件，用于调试或手动部署", "將主頁中目前選擇的 Office 部署設定匯出為 XML 檔案，用於除錯或手動部署", "Export the Office deployment settings currently selected on the main page to an XML file for debugging or manual deployment");
        public string SettingsEnableExportXmlTitle => S("在主页开启导出 XML", "在主頁開啟匯出 XML", "Enable Export XML on Home Page");
        public string SettingsEnableExportXmlDesc => S("开启后，将在 Microsoft Office 部署页面显示“导出 XML”按钮", "開啟後，將在 Microsoft Office 部署頁面顯示「匯出 XML」按鈕", "When enabled, the \"Export XML\" button will be shown on the Microsoft Office page.");
        public string SettingsAboutSectionTitle => S("关于软件", "關於軟體", "About");
        public string AboutDesc1 => S("本软件为完全免费开源项目，专为高效、静默部署各类常用办公套件设计。",
                                       "本軟體為完全免費開源專案，專為高效、靜默部署各類常用辦公套件設計。",
                                       "This software is a completely free open-source project for efficient, silent deployment of common office suites.");
        public string AboutDesc2 => S("版权所有 © 2025-2026 OSBoxTeam ＆ Glim。仅供网络技术学习与环境部署研究使用。",
                                       "版權所有 © 2025-2026 OSBoxTeam ＆ Glim。僅供網路技術學習與環境部署研究使用。",
                                       "Copyright © 2025-2026 OSBoxTeam ＆ Glim. For educational and deployment research purposes only.");
        public string AboutDialogDesc1 => S("本软件为完全免费开源项目，专为高效、静默部署各类常用办公套件设计。",
                                             "本軟體為完全免費開源專案，專為高效、靜默部署各類常用辦公套件設計。",
                                             "Free and open-source project for efficient, silent deployment of common office suites.");
        public string AboutDialogDesc2 => S("版权所有 © 2025-2026 OSBoxTeam ＆ Glim。保留所有权利。",
                                             "版權所有 © 2025-2026 OSBoxTeam ＆ Glim。保留所有權利。",
                                             "Copyright © 2025-2026 OSBoxTeam ＆ Glim. All rights reserved.");
        public string AboutAuthor => S("作者：WumiaoTech", "作者：WumiaoTech", "Author: WumiaoTech");


        // ===================== 非 MSO 办公套件页面文本 =====================
        public string Wps2023Desc => S("主版本号 12.1 (集成优化补丁版)", "主版本號 12.1 (整合優化補丁版)", "Version 12.1 (Optimized patch version)");
        public string DeployAndParams => S("部署及参数说明", "部署及參數說明", "Deployment & Parameter Notes");
        public string WpsParam1Title => S("官方直链高速下载", "官方直鏈高速下載", "Official Direct Link High-Speed Download");
        public string WpsParam1Desc => S("直接自 WPS 官方 CDN 节点下载原版安装包，保障网络传输速度与安全", "直接自 WPS 官方 CDN 節點下載原版安裝包，保障網路傳輸速度與安全", "Download original setup packages directly from official WPS CDN nodes");
        public string WpsParam2Title => S("静默无缝安装", "靜默無縫安裝", "Silent Seamless Installation");
        public string WpsParam2Desc => S("后台全自动运行部署，无需手动确认或人工交互，避免附带捆绑软件", "後台全自動執行部署，無需手動確認或人工交互，避免附帶綑綁軟體", "Runs fully silently in the background without user intervention, preventing bundled software");
        public string WpsParam3Title => S("自动清理安装缓存", "自動清理安裝快取", "Auto Clean Install Cache");
        public string WpsParam3Desc => S("部署完成后自动清理下载的临时安装文件与安装缓存，释放本地磁盘空间", "部署完成後自動清理下載的臨時安裝檔案與安裝快取，釋放本機磁碟空間", "Automatically deletes downloaded temporary files and installation cache after deployment to free up disk space");
        public string YozoParam1Title => S("电子公文标准适配", "電子公文標準適配", "Electronic Document Standard Support");
        public string YozoParam1Desc => S("符合国家电子公文规范，内置排版算法支持标准中文公文排版与格式要求", "符合國家電子公文規範，內建排版演算法支援標準中文公文排版與格式要求", "Complies with national electronic document specifications, with layout algorithms supporting standard Chinese public document formatting");
        public string YozoParam2Title => S("引导交互式安装", "引導互動式安裝", "Interactive Guided Setup");
        public string YozoParam2Desc => S("由于官方未提供静默参数，程序将在后台下载并解压官方安装包，并弹出官方安装向导引导您完成安装", "由於官方未提供靜默參數，程式將在後台下載並解壓官方安裝包，並彈出官方安裝精靈引導您完成安裝", "Since the official installer lacks silent switches, the program will download and extract the files, then open the official wizard to guide your setup.");
        public string OnlyOfficeParam1Title => S("高规格格式兼容性", "高規格格式相容性", "High-Specification Format Compatibility");
        public string OnlyOfficeParam1Desc => S("对 Office Open XML (.docx/.xlsx/.pptx) 格式进行优化，保持良好的排版兼容性", "對 Office Open XML (.docx/.xlsx/.pptx) 格式進行優化，保持良好的排版相容性", "Optimized for Office Open XML (.docx, .xlsx, .pptx) formats to maintain layout compatibility");
        public string OnlyOfficeParam2Title => S("无广告多标签页", "無廣告多標籤頁", "Ad-Free Multi-Tab Interface");
        public string OnlyOfficeParam2Desc => S("采用无广告开源分发，支持单窗口多标签处理以提升编辑体验", "採用無廣告開源分發，支援單視窗多標籤處理以提升編輯體驗", "Distributed as ad-free open source, supporting single-window multi-tab document editing");
        public string LibreOfficeParam1Title => S("标准 ODF 格式支持", "標準 ODF 格式支援", "Standard ODF Format Support");
        public string LibreOfficeParam1Desc => S("符合 ODF (OpenDocument Format) 国际开放标准规范，提供原生解析支持", "符合 ODF (OpenDocument Format) 國際開放標準規範，提供原生解析支援", "Complies with international ODF standards to offer native parsing support");
        public string LibreOfficeParam2Title => S("镜像源代理高速下载", "鏡像源代理高速下載", "Mirror Source High-Speed Download");
        public string LibreOfficeParam2Desc => S("通过国内科大开源镜像服务器下载安装文件，提升网络传输与下载连接稳定性", "透過國內科大開源鏡像伺服器下載安裝檔案，提升網路傳輸與下載連線穩定性", "Downloads install files via domestic USTC open-source mirror servers to improve network stability and speed");

        // ===================== 状态栏与开关标签 =====================
        public string LabelOn => S("开", "開", "On");
        public string LabelOff => S("关", "關", "Off");

        public string StatusScanningAssociations => S("正在扫描失效的文件关联与残留图标...", "正在掃描失效的文件關聯與殘留圖示...", "Scanning for invalid file associations and leftover icons...");
        public string StatusAssociationsCleaned => S("文件关联与图标净化完成！", "文件關聯與圖示淨化完成！", "File association and icon purge completed!");
        public string StatusCleanAssociationsFailed(string msg) => S("净化失败: " + msg, "淨化失敗: " + msg, "Purge failed: " + msg);

        public string StatusRebuildingIconCache => S("正在强制重建 Windows 图标缓存并通知资源管理器...", "正在強制重建 Windows 圖示快取並通知資源管理器...", "Rebuilding Windows icon cache and notifying Explorer...");
        public string StatusIconCacheRefreshed => S("系统图标缓存已成功刷新！", "系統圖示快取已成功重新整理！", "System icon cache successfully refreshed!");
        public string StatusRefreshIconCacheFailed(string msg) => S("刷新失败: " + msg, "重新整理失敗: " + msg, "Refresh failed: " + msg);

        public string StatusRepairingAssociations => S("正在检测并修复已安装办公套件的文件关联...", "正在檢測並修復已安裝辦公套件的文件關聯...", "Detecting and repairing file associations for installed suites...");
        public string StatusAssociationsRepaired => S("已安装办公套件的文件关联修复完成！", "已安裝辦公套件的文件關聯修復完成！", "File associations for installed suites successfully repaired!");
        public string StatusRepairAssociationsFailed(string msg) => S("文件关联修复失败: " + msg, "文件關聯修復失敗: " + msg, "File association repair failed: " + msg);

        public string StatusRepairingCOM => S("正在检测并重新注册已安装 Office 套件的 COM 组件...", "正在檢測並重新註冊已安裝 Office 套件的 COM 組件...", "Detecting and re-registering COM components for installed Office suites...");
        public string StatusCOMRepaired => S("COM 组件重新注册完成！", "COM 組件重新註冊完成！", "COM components successfully re-registered!");
        public string StatusRepairCOMFailed(string msg) => S("COM 组件修复失败: " + msg, "COM 組件修復失敗: " + msg, "COM component repair failed: " + msg);

        public string StatusGuidingDefaultApp(string name) => S("正在引导设置 " + name + " 默认打开程序...", "正在引導設置 " + name + " 預設開啟程式...", "Guiding default application setup for " + name + "...");
        public string StatusDefaultAppGuided => S("默认办公软件设置引导完成！", "預設辦公軟體設置引導完成！", "Default office application setup guide completed!");

        // ===================== 架构文字 =====================
        public string ArchX64 => S("系统架构：x64（64 位）", "系統架構：x64（64 位元）", "Architecture: x64 (64-bit)");
        public string ArchX86 => S("系统架构：x86（32 位）", "系統架構：x86（32 位元）", "Architecture: x86 (32-bit)");

        // ===================== 进度状态文字 =====================
        public string StatusClean => S("正在清理旧版本 Office 残留...", "正在清理舊版本 Office 殘留...", "Cleaning up old Office residuals...");
        public string StatusDownloading => S("正在下载安装组件...", "正在下載安裝元件...", "Downloading installation components...");
        public string StatusConfiguringXml => S("正在生成安装配置...", "正在產生安裝設定...", "Generating installation configuration...");
        public string StatusInstallingWizard => S("正在启动 Office 安装向导，请耐心等待部署完成...", "正在啟動 Office 安裝精靈，請耐心等待部署完成...", "Starting Office setup wizard, please wait for deployment to complete...");
        public string StatusActivating => S("正在配置授权...", "正在啟用 Office...", "Activating Office...");

        // ===================== 对话框（代码内使用）=====================
        public string DlgInstallSuccessTitle => S("部署成功", "部署成功", "Deployment Successful");
        public string DlgInstallSuccessMsg => S("Office 已成功部署并完成配置！", "Office 已成功部署並完成啟用！", "Office has been successfully deployed and activated!");
        
        public (string Title, string Msg) GetInstallSuccessInfo(ProductType product)
        {
            switch (product)
            {
                case ProductType.MsOffice:
                    return (
                        S("部署并配置成功", "部署並啟用成功", "Deployment & Activation Successful"),
                        S("Microsoft Office 已成功部署在您的计算机上，内置本地授权配置组件已自动为您配置完毕！",
                          "Microsoft Office 已成功部署在您的電腦上，內建學習啟用（Ohook）元件已自動為您配置完畢！",
                          "Microsoft Office has been successfully deployed and activated offline via Ohook!")
                    );
                case ProductType.Wps:
                    return (
                        S("WPS 部署成功", "WPS 部署成功", "WPS Deployment Successful"),
                        S("WPS Office 已成功部署在您的计算机上！\n\n提示：如果您的系统上同时安装了 Microsoft Office，WPS 的快捷方式已为您准备完毕，同时卸载时将自动恢复 Microsoft Office 默认关联。",
                          "WPS Office 已成功部署在您的電腦上！\n\n提示：如果您的系統上同時安裝了 Microsoft Office，WPS 的捷徑已為您準備完畢，同時解除安裝時將自動恢復 Microsoft Office 預設關聯。",
                          "WPS Office has been successfully deployed!\n\nNote: If Microsoft Office is also installed, associations will be restored to it automatically upon uninstalling WPS.")
                    );
                case ProductType.Yozo:
                    return (
                        S("永中 Office 2024 部署完成", "永中 Office 2024 部署完成", "Yozo Office 2024 Deployment Complete"),
                        S("永中 Office 2024 个人版已成功安装在您的计算机上！\n\n提示：由于永中官方暂未提供静默安装，您已通过手动引导向导完成了全部安装流程。",
                          "永中 Office 2024 個人版已成功安裝在您的電腦上！\n\n提示：由於永中官方暫未提供靜默安裝，您已透過手動引導精靈完成了全部安裝流程。",
                          "Yozo Office 2024 Personal has been successfully installed!\n\nNote: As Yozo does not support silent install, you have completed the installation via manual wizard.")
                    );
                case ProductType.OnlyOffice:
                    return (
                        S("ONLYOFFICE 部署成功", "ONLYOFFICE 部署成功", "ONLYOFFICE Deployment Successful"),
                        S("ONLYOFFICE Desktop Editors 已成功安装！\n\n这是一款完全开源、无广告的安全办公套件，支持单窗口多标签文档协作编辑。",
                          "ONLYOFFICE Desktop Editors 已成功安裝！\n\n這是一款完全開源、無廣告的安全辦公套件，支援單視窗多標籤文件協作編輯。",
                          "ONLYOFFICE Desktop Editors has been successfully installed!\n\nIt is a completely open-source, ad-free secure office suite featuring tabbed editing.")
                    );
                case ProductType.LibreOffice:
                    return (
                        S("LibreOffice 部署成功", "LibreOffice 部署成功", "LibreOffice Deployment Successful"),
                        S("LibreOffice 稳定版已成功部署！\n\n这是一款强力、免费开源的办公套件，基于国际开放文档格式（ODF）标准设计。",
                          "LibreOffice 穩定版已成功部署！\n\n這是一款強力、免費開源的辦公套件，基於國際開放文件格式（ODF）標準設計。",
                          "LibreOffice stable has been successfully deployed!\n\nIt is a powerful, free and open-source office suite based on the Open Document Format (ODF) standard.")
                    );
                default:
                    return (DlgInstallSuccessTitle, DlgInstallSuccessMsg);
            }
        }

        public string DlgInstallFailTitle => S("部署失败", "部署失敗", "Deployment Failed");
        public string DlgUninstallSuccessTitle => S("卸载完成", "解除安裝完成", "Uninstall Complete");
        public string DlgUninstallSuccessMsg => S("所选 Office 产品已成功从本机清除。", "所选 Office 產品已成功從本機清除。", "The selected Office product has been successfully removed.");
        public string DlgUninstallFailTitle => S("卸载失败", "解除安裝失敗", "Uninstall Failed");
        public string DlgClearLicenseTitle => S("清理完成", "清理完成", "Cleanup Complete");
        public string DlgClearLicenseMsg => S("已成功卸载所有 Office 授权许可，授权状态已重置。", "已成功解除安裝所有 Office 啟用授權，啟用狀態已重設。", "All Office activation licenses have been successfully removed.");
        public string DlgClearLicenseFailTitle => S("清理失败", "清理失敗", "Cleanup Failed");
        public string DlgActivateSuccessTitle => S("配置成功", "啟用成功", "Activation Successful");
        public string DlgActivateSuccessMsg => S("Office 本地离线授权已成功配置！", "Office 本機離線啟用已成功設定！", "Office local offline activation configured successfully!");
        public string DlgActivateFailTitle => S("配置失败", "啟用失敗", "Activation Failed");
        public string DlgConfirmInstallTitle => S("检测到已安装版本", "偵測到已安裝版本", "Installed Version Detected");
        public string DlgConfirmInstallMsg(string detected) => Detected switch
        {
            AppLanguage.TraditionalChinese => $"偵測到已安裝：{detected}\n\n建議先解除安裝舊版本再繼續，否則可能產生衝突。是否繼續部署？",
            AppLanguage.English => $"Detected installed: {detected}\n\nIt is recommended to uninstall first to avoid conflicts. Continue deployment?",
            _ => $"检测到已安装：{detected}\n\n建议先卸载旧版本再继续，否则可能产生冲突。是否继续部署？"
        };
        public string DlgExportXmlTitle => S("导出成功", "匯出成功", "Export Successful");
        public string DlgExportXmlMsg(string path) => Detected switch
        {
            AppLanguage.TraditionalChinese => $"XML 設定已匯出至：\n{path}",
            AppLanguage.English => $"XML configuration exported to:\n{path}",
            _ => $"XML 配置已导出至：\n{path}"
        };
        public string DlgExportXmlFailTitle => S("导出失败", "匯出失敗", "Export Failed");

        // ===================== 部署错误提示 =====================
        public string ErrDownloadFailed => S("下载失败，请检查网络连接。", "下載失敗，請檢查網路連線。", "Download failed, please check your network connection.");
        public string ErrCannotStartInstaller => S("无法启动安装程序。", "無法啟動安裝程序。", "Could not start installer.");
        public string ErrInstallerExitCode(int code) => Detected switch
        {
            AppLanguage.TraditionalChinese => $"安裝程序退出，錯誤代碼: {code}",
            AppLanguage.English => $"Installer exited with code: {code}",
            _ => $"安装程序退出，错误代码: {code}"
        };
        public string ErrInstallFailed(string msg) => Detected switch
        {
            AppLanguage.TraditionalChinese => $"安裝失敗: {msg}",
            AppLanguage.English => $"Installation failed: {msg}",
            _ => $"安装失败: {msg}"
        };

        // ===================== 部署状态文字 =====================
        public string StatusDeploymentCancelled => S("已取消部署", "已取消部署", "Deployment cancelled.");
        public string StatusCleaningOldVersions(string company) => Detected switch
        {
            AppLanguage.TraditionalChinese => $"正在安全清理 {company} 舊版本...",
            AppLanguage.English => $"Cleaning up old {company} versions...",
            _ => $"正在安全清理 {company} 旧版本..."
        };
        public string StatusDeploySuccess => S("部署完成！软件已成功安装。", "部署完成！軟體已成功安裝。", "Deployment completed successfully!");
        public string StatusDeployFail => S("安装失败，请查看日志了解详情。", "安裝失敗，請檢視記錄了解詳情。", "Installation failed. See logs for details.");

        // ===================== 激活与清除许可相关文字 =====================
        public string StatusScanningActivationKeys => S("正在扫描本地 Office 授权密钥...", "正在掃描本機 Office 啟用金鑰...", "Scanning local Office activation keys...");
        public string ErrOsppNotFound => S("未在系统中找到 OSPP.VBS 许可管理脚本，可能由于未安装 Microsoft Office。", "未在系統中找到 OSPP.VBS 授權管理指令碼，可能由於未安裝 Microsoft Office。", "OSPP.VBS license management script not found. Microsoft Office might not be installed.");
        public string StatusPathNotFound => S("未检测到 Office 许可路径", "未偵測到 Office 授權路徑", "Office license path not detected");
        public string DlgScanNoKeysTitle => S("扫描结果", "掃描結果", "Scan Results");
        public string DlgScanNoKeysMsg => S("未检测到本地有已安装的 Office 密钥授权信息。", "未偵測到本機有已安裝的 Office 金鑰啟用資訊。", "No installed Office activation keys detected on this machine.");
        public string StatusNoKeysFound => S("未找到已授权的密钥", "未找到已啟用的金鑰", "No activated keys found");
        
        public string DlgConfirmClearTitle => S("确认删除授权信息", "確認刪除啟用資訊", "Confirm Deleting Activation Info");
        public string DlgConfirmClearMsg(int count, string keys) => Detected switch
        {
            AppLanguage.TraditionalChinese => $"偵測到本機存在 {count} 個 Office 啟用金鑰：{keys}\n\n清除啟用資訊後，Office 將會處於未啟用狀態。確定要繼續刪除嗎？",
            AppLanguage.English => $"Detected {count} local Office activation keys: {keys}\n\nOffice will be deactivated after clearing. Do you want to continue?",
            _ => $"检测到本地存在 {count} 个 Office 授权密钥：{keys}\n\n清除授权信息后，Office 将会处于未授权状态。确定要继续删除吗？"
        };
        public string BtnDeleteConfirm => S("确定删除", "確定刪除", "Delete");
        public string StatusClearCancelled => S("已取消删除", "已取消刪除", "Deletion cancelled");
        public string StatusClearingKeys => S("正在删除授权密钥...", "正在刪除啟用金鑰...", "Deleting activation keys...");
        public string StatusClearSuccess(int count) => Detected switch
        {
            AppLanguage.TraditionalChinese => $"成功清除 {count} 個啟用金鑰！",
            AppLanguage.English => $"Successfully removed {count} activation keys!",
            _ => $"成功清除 {count} 个激活密钥！"
        };
        public string DlgClearSuccessMsg(int count) => Detected switch
        {
            AppLanguage.TraditionalChinese => $"已成功清除 {count} 個 Office 啟用資訊！\n重新開啟 Office 元件（如 Word）後可進行重新啟用或綁定。",
            AppLanguage.English => $"Successfully cleared {count} Office activation keys!\nLaunch Office applications (e.g. Word) to reactivate or bind license.",
            _ => $"已成功清除 {count} 个 Office 授权信息！\n重新打开 Office 组件（如 Word）后可进行重新配置或绑定。"
        };

        public string DlgConfirmOhookTitle => S("Office 本地授权配置确认", "Office Ohook 啟用確認", "Office Ohook Activation Confirmation");
        public string DlgConfirmOhookMsg => S("本地授权配置功能将使用内置模块在本地配置所有 Microsoft Office 版本（包括零售版与 365 订阅版）的授权状态。\n\n确认要开始一键配置吗？",
                                               "學習啟用（Ohook）功能將使用內建 Ohook 模組在本機啟用所有 Microsoft Office 版本（包括零售版與 365 訂閱版）。\n\n確認要開始一鍵啟用嗎？",
                                               "The offline learning activation (Ohook) will configure the local Ohook module to activate all Microsoft Office versions (including Retail and 365 subscriptions).\n\nDo you want to start?");
        public string BtnOhookConfirm => S("一键配置", "一鍵啟用", "Activate");
        public string StatusReleasingOhook => S("正在释放本地授权组件...", "正在釋放 Ohook 啟用元件...", "Releasing Ohook activation components...");
        public string ErrReleaseFailed(string msg) => Detected switch
        {
            AppLanguage.TraditionalChinese => $"釋放 Ohook 資源失敗: {msg}",
            AppLanguage.English => $"Failed to release Ohook resource: {msg}",
            _ => $"释放 Ohook 资源失败: {msg}"
        };
        public string StatusRunningOhook => S("正在后台配置本地授权环境...", "正在背景執行 Ohook 啟用指令碼...", "Running Ohook activation script in background...");
        public string StatusOhookSuccess => S("本地授权部署完成！", "Ohook 啟用部署完成！", "Ohook activation deployment completed!");
        public string DlgOhookSuccessMsg => S("本地授权配置已完成！您现在即可打开 Office 正常使用。",
                                               "Ohook 啟用部署完成！您現在即可開啟 Office 正常使用。",
                                               "Ohook activation deployed! You can now launch and use Office applications.");
        public string StatusOhookFail => S("Ohook 部署异常中途退出。", "Ohook 部署異常中途退出。", "Ohook deployment aborted due to errors.");
        public string DlgOhookFailMsg => S("Ohook 部署脚本运行失败，请确认系统环境。",
                                            "Ohook 部署指令碼執行失敗，請確認系統環境。",
                                            "Ohook deployment script failed to run. Please check your system environment.");

        // Static Instance for easy access in non-viewmodel context
        public static LocalizationStrings Instance { get; } = new LocalizationStrings();

        // Deep Cleanup Dialog
        public string DlgCleanAssociationsTitle => S("净化完成", "淨化完成", "Purge Completed");
        public string DlgCleanAssociationsMsg => S(
            "系统已成功扫描并清除所有已卸载 Office 品牌（WPS、永中、OnlyOffice、LibreOffice）在注册表残留的文件关联、右键菜单项及快捷图标，关联环境已深度净化！",
            "系統已成功掃描並清除所有已解除安裝 Office 品牌（WPS、永中、OnlyOffice、LibreOffice）在註冊表殘留的文件關聯、右鍵選單項目及快捷圖示，關聯環境已深度淨化！",
            "Successfully scanned and cleaned leftover file associations, context menus, and shortcut icons in the registry for uninstalled Office suites (WPS, Yozo, OnlyOffice, LibreOffice).");
        public string DlgCleanAssociationsFailTitle => S("净化失败", "淨化失敗", "Purge Failed");

        // Rebuild Icon Cache Dialog
        public string DlgRefreshIconCacheTitle => S("刷新完成", "重新整理完成", "Refresh Completed");
        public string DlgRefreshIconCacheMsg => S(
            "已成功强制重建系统图标缓存并向 Windows Shell 发送全局重绘通知！桌面及文件夹下的所有文件图标已恢复正常显示。",
            "已成功強制重建系統圖示快取並向 Windows Shell 傳送全域重繪通知！桌面及資料夾下的所有檔案圖示已恢復正常顯示。",
            "Successfully forced a rebuild of the system icon cache and sent a global redraw notification to Windows Shell. All file icons have been restored.");
        public string DlgRefreshIconCacheFailTitle => S("刷新失败", "重新整理失敗", "Refresh Failed");

        // Repair Associations Dialog
        public string DlgRepairAssociationsTitle => S("修复完成", "修復完成", "Repair Completed");
        public string DlgRepairAssociationsMsg => S(
            "已检测到系统中的已安装办公套件并恢复了它们的文件关联默认值与 ProgID，且成功刷新了系统图标缓存！",
            "已偵測到系統中的已安裝辦公套件並恢復了它們的文件關聯預設值與 ProgID，且成功重新整理了系統圖示快取！",
            "Detected installed office suites on your system, restored their default file associations/ProgIDs, and successfully refreshed the system icon cache!");
        public string DlgRepairAssociationsFailTitle => S("修复失败", "修復失敗", "Repair Failed");
        public string DlgRepairAssociationsFailMsg(string msg) => S("修复过程中发生错误: " + msg, "修復過程中發生錯誤: " + msg, "An error occurred during repair: " + msg);

        // Default App Setup Dialog
        public string DlgDefaultAppTitle => S("设置默认打开程序", "設定預設開啟程式", "Set Default Programs");
        public string DlgDefaultAppMsg => S(
            "检测到您的系统上安装了多款 Office 办公软件。\n为了方便您自主选择默认打开的软件，程序将依次为您弹出 Word（.docx）和 Excel（.xlsx）文件的“打开方式”选择框。\n\n请在弹出的窗口中选择您首选的默认软件，并勾选“始终使用此应用打开”。\n\n是否现在开始设置？",
            "偵測到您的系統上安裝了多款 Office 辦公軟體。\n為了方便您自主選擇預設開啟的軟體，程式將依次為您彈出 Word（.docx）和 Excel（.xlsx）檔案的「開啟方式」選擇框。\n\n請在彈出的視窗中選擇您首選的預設軟體，並勾選「始終使用此應用程式開啟」。\n\n是否現在開始設定？",
            "Multiple Office suites detected on your system.\nTo help you set default file associations, the program will open the 'Open With' dialog for Word (.docx) and Excel (.xlsx) files.\n\nPlease select your preferred software and check 'Always use this app to open files'.\n\nDo you want to start setting up now?");
        public string BtnDefaultAppStart => S("开始设置", "開始設定", "Start Setup");
        public string BtnDefaultAppSkip => S("跳过", "跳過", "Skip");

        // Format names
        public string FormatWord => S("Word 文档 (.docx)", "Word 文件 (.docx)", "Word Document (.docx)");
        public string FormatWordLegacy => S("旧版 Word 文档 (.doc)", "舊版 Word 文件 (.doc)", "Legacy Word Document (.doc)");
        public string FormatExcel => S("Excel 工作簿 (.xlsx)", "Excel 活頁簿 (.xlsx)", "Excel Workbook (.xlsx)");
        public string FormatExcelLegacy => S("旧版 Excel 工作表 (.xls)", "舊版 Excel 工作表 (.xls)", "Legacy Excel Sheet (.xls)");
        public string FormatPowerPoint => S("PowerPoint 演示文稿 (.pptx)", "PowerPoint 簡報 (.pptx)", "PowerPoint Presentation (.pptx)");
        public string FormatPowerPointLegacy => S("旧版 PowerPoint 演示文稿 (.ppt)", "舊版 PowerPoint 簡報 (.ppt)", "Legacy PowerPoint Presentation (.ppt)");
        public string FormatPdf => S("PDF 文档 (.pdf)", "PDF 文件 (.pdf)", "PDF Document (.pdf)");

        public string DlgDefaultAppSuccessTitle => S("设置完成", "設定完成", "Setup Completed");
        public string DlgDefaultAppSuccessMsg => S("您已成功为各办公格式设置了首选的默认打开程序！", "您已成功為各辦公格式設定了首選的預設開啟程式！", "You have successfully configured the preferred default programs for each office format!");

        // COM Repair Dialog
        public string DlgRepairComTitle => S("修复完成", "修復完成", "Repair Completed");
        public string DlgRepairComMsg => S(
            "已成功为系统中安装的办公套件（Word/Excel/WPS/ET 等）执行 COM 服务组件自我注册，解决了第三方调用及插件报错的顽疾！",
            "已成功為系統中安裝的辦公套件（Word/Excel/WPS/ET 等）執行 COM 服務元件自我註冊，解決了第三方呼叫及增益集報報錯的頑疾！",
            "Successfully performed COM component self-registration for installed office suites (Word, Excel, WPS, ET, etc.), resolving issues with plugin integrations and third-party API calls.");
        public string DlgRepairComFailTitle => S("修复失败", "修復失敗", "Repair Failed");
        public string DlgRepairComFailMsg(string msg) => S("修复过程中发生错误: " + msg, "修復過程中發生錯誤: " + msg, "An error occurred during repair: " + msg);

        // General Error
        public string DlgDeployFailTitle => S("部署失败", "部署失敗", "Deployment Failed");
        public string DlgDeployFailMsg => S("办公套件在部署/安装过程中发生了错误。", "辦公套件在部署/安裝過程中發生了錯誤。", "An error occurred during office suite deployment/installation.");

        // Installer services status reports
        public string StatusDownloadingProduct(string name) => S($"正在下载 {name}...", $"正在下載 {name}...", $"Downloading {name}...");
        public string StatusInstallingProduct(string name) => S($"正在静默安装 {name}...", $"正在靜默安裝 {name}...", $"Silently installing {name}...");
        public string StatusInstallingProductGuide(string name) => S($"正在引导安装 {name}...", $"正在引導安裝 {name}...", $"Guiding installation of {name}...");
        public string StatusProductInstalled(string name) => S($"{name} 安装完成！", $"{name} 安裝完成！", $"{name} installation completed!");
        public string StatusProductInstallFailed(string name, string msg) => S($"{name} 安装失败: {msg}", $"{name} 安裝失敗: {msg}", $"{name} installation failed: {msg}");
        public string ErrDownloadFailedWithMsg => S("下载失败，请检查网络连接。", "下載失敗，請檢查網路連線。", "Download failed, please check your network connection.");
        public string ErrCannotStartInstallerWithMsg => S("无法启动安装程序。", "無法啟動安裝程序。", "Could not start installer.");
        public string ErrCannotStartMsiWithMsg => S("无法启动 MSI 安装引擎。", "無法啟動 MSI 安裝引擎。", "Could not start MSI installation engine.");
        public string ErrInstallerAbortedWithCode(string name, int code) => S($"{name} 安装程序异常退出，错误码: {code}", $"{name} 安裝程序異常退出，錯誤碼: {code}", $"{name} installer exited abnormally with code: {code}");
        
        public string DlgConfirmYozoTitle => S("安装提示", "安裝提示", "Installation Hint");
        public string DlgConfirmYozoMsg => S("因永中软件官方暂未提供静默安装，请您按照永中软件官方的安装程序引导进行安装永中Office，若关闭弹出的安装窗口则安装失败", "因永中軟體官方暫未提供靜默安裝，請您按照永中軟體官方的安裝程序引導進行安裝永中Office，若關閉彈出的安裝視窗則安裝失敗", "As Yozo Office does not support silent installation, please follow the official installer wizard. Closing the wizard will result in installation failure.");
        public string StatusExtractingProduct(string name) => S($"正在解压 {name} 安装包...", $"正在解壓 {name} 安裝包...", $"Extracting {name} installation files...");
        public string StatusDownloadYozoRar => S("正在下载 永中Office 官方压缩包...", "正在下載 永中Office 官方壓縮包...", "Downloading Yozo Office official package...");
        public string ErrExtractFailed(string msg) => S($"安装包解压失败: {msg}", $"安裝包解壓失敗: {msg}", $"Failed to extract installation package: {msg}");
        public string ErrYozoExeNotFound => S("未能在压缩包内找到安装执行文件。", "未能在壓縮包內找到安裝執行文件。", "Installation executable file not found in the archive.");

        // WPS versions labels
        public string WpsVersionLatestLabel => S("官方最新版", "官方最新版", "Official Latest");

        // Cleanup service strings
        public string StatusCleanStartC2R => S("正在启动 Microsoft Office 官方卸载程序...", "正在啟動 Microsoft Office 官方解除安裝程式...", "Starting Microsoft Office official uninstaller...");
        public string StatusCleanKillProcesses => S("正在终止相关进程...", "正在終止相關程序...", "Terminating related processes...");
        public string StatusCleanC2RService => S("正在清理 ClickToRun 服务...", "正在清理 ClickToRun 服務...", "Cleaning up ClickToRun service...");
        public string StatusCleanRegistry => S("正在清理注册表...", "正在清理登錄機碼...", "Cleaning up registry...");
        public string StatusCleanUninstallEntries => S("正在清理卸载记录...", "正在清理解除安裝記錄...", "Cleaning up uninstall records...");
        public string StatusCleanResidualFiles => S("正在清理残留文件...", "正在清理殘留檔案...", "Cleaning up residual files...");
        public string StatusCleanAssociations => S("正在清理快捷方式与文件关联...", "正在清理捷徑與文件關聯...", "Cleaning up shortcuts and file associations...");

        // Uninstall prompt details
        public string DlgConfirmUninstallMsg(string name) => S(
            $"深度卸载将强制终止所有正在运行的 {name} 进程，并彻底清除注册表与残留文件夹。\n\n确认要继续吗？请务必先保存正在编辑的文档。",
            $"深度解除安裝將強制終止所有正在執行的 {name} 程序，並徹底清除登錄機碼與殘留資料夾。\n\n確認要繼續嗎？請務必先儲存正在編輯的文件。",
            $"Deep uninstallation will force-terminate all running {name} processes and completely clean up registry entries and residual folders.\n\nDo you want to continue? Please make sure to save your work first.");

        // Dots-based progress labels
        public string StatusDownloadingOfficeFiles(string dots) => S("正在下载 Office 文件" + dots, "正在下載 Office 檔案" + dots, "Downloading Office files" + dots);
        public string StatusInstallingOfficeComponents(string dots) => S("正在安装 Office 组件" + dots, "正在安裝 Office 元件" + dots, "Installing Office components" + dots);
        public string StatusAlmostDone(string dots) => S("即将完成，请稍候" + dots, "即將完成，請稍候" + dots, "Almost done, please wait" + dots);

        // Registry and Operation logs
        public string ErrSubmitLogHint => S("\n\n操作发生错误！为协助排查该问题，建议您将日志文件提交给开发者。\n点击“确定”将自动打开该日志文件以供查看。", "\n\n操作發生錯誤！為協助排查該問題，建議您將記錄檔提交給開發者。\n點擊「確定」將自動開啟該記錄檔以供查看。", "\n\nAn error occurred! To help diagnose this issue, you are advised to submit the log file to the developer.\nClicking \"OK\" will automatically open the log file for viewing.");
        public string StatusDetectMsOffice(string ver) => S("检测到 Microsoft Office " + ver + " 已安装，正在修复关联并重写键值...", "偵測到 Microsoft Office " + ver + " 已安裝，正在修復關聯並重寫鍵值...", "Detected Microsoft Office " + ver + " installed. Repairing file associations...");
        public string StatusDetectWps(string ver) => S("检测到 WPS Office " + ver + " 已安装，正在修复其特有关联并刷新图标...", "偵測到 WPS Office " + ver + " 已安裝，正在修復其特有關聯並重新整理圖示...", "Detected WPS Office " + ver + " installed. Repairing associations and refreshing icons...");
        public string StatusPurgingProduct(string name) => S("检测到 " + name + " 处于未安装状态，正在深度净化其余留关联、右键菜单与死链快捷键...", "偵測到 " + name + " 處於未安裝狀態，正在深度淨化其餘留關聯、右鍵選單與捷徑死鏈...", "Detected " + name + " is not installed. Purging registry leftovers and shortcuts...");
        public string StatusRestoringCOM(string name) => S("正在重新配置并恢复 " + name + " 的 COM 注册信息...", "正在重新配置並恢復 " + name + " 的 COM 註冊資訊...", "Restoring COM registration details for " + name + "...");
    }
}
