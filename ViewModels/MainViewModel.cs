using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using GOI.Models;
using GOI.Services;
using GOI.Helpers;

namespace GOI.ViewModels
{
    public class MainViewModel : ObservableObject
    {
        private readonly InstallService _installService;
        private readonly WpsInstallService _wpsService;
        private readonly YozoInstallService _yozoService;
        private readonly OnlyOfficeInstallService _onlyOfficeService;
        private readonly LibreOfficeInstallService _libreOfficeService;
        private readonly CleanupService _cleanupService;

        public MainViewModel(InstallService installService)
        {
            _installService = installService;
            _wpsService = new WpsInstallService();
            _yozoService = new YozoInstallService();
            _onlyOfficeService = new OnlyOfficeInstallService();
            _libreOfficeService = new LibreOfficeInstallService();
            _cleanupService = new CleanupService();

            DetectedArch = Environment.Is64BitOperatingSystem ? Architecture.x64 : Architecture.x86;
            ArchText = DetectedArch == Architecture.x64 ? "系统架构：x64（64 位）" : "系统架构：x86（32 位）";

            // 初始化 Office 2024 作为默认选择，并计算初始翻页逻辑
            SelectLeft();
            RefreshCards();

            // 首次刷新检测安装版本
            RefreshInstalledVersion();
        }

        // ========== 架构 ==========
        private Architecture DetectedArch { get; } = Architecture.x64;
        private string _archText = "";
        public string ArchText { get => _archText; set => Set(ref _archText, value); }

        // ========== 产品类型切换（MS Office / WPS / 永中 / OnlyOffice / LibreOffice）==========
        private ProductType _productType = ProductType.MsOffice;
        public ProductType CurrentProductType
        {
            get => _productType;
            set
            {
                Set(ref _productType, value);
                OnPropertyChanged(nameof(IsMsOffice));
                OnPropertyChanged(nameof(IsWps));
                OnPropertyChanged(nameof(IsYozo));
                OnPropertyChanged(nameof(IsOnlyOffice));
                OnPropertyChanged(nameof(IsLibreOffice));

                // 切换产品时，刷新已安装版本状态
                RefreshInstalledVersion();
            }
        }
        public bool IsMsOffice => CurrentProductType == ProductType.MsOffice;
        public bool IsWps => CurrentProductType == ProductType.Wps;
        public bool IsYozo => CurrentProductType == ProductType.Yozo;
        public bool IsOnlyOffice => CurrentProductType == ProductType.OnlyOffice;
        public bool IsLibreOffice => CurrentProductType == ProductType.LibreOffice;

        public ICommand SelectMsOfficeCommand => new RelayCommand(() => CurrentProductType = ProductType.MsOffice);
        public ICommand SelectWpsCommand => new RelayCommand(() => CurrentProductType = ProductType.Wps);
        public ICommand SelectYozoCommand => new RelayCommand(() => CurrentProductType = ProductType.Yozo);
        public ICommand SelectOnlyOfficeCommand => new RelayCommand(() => CurrentProductType = ProductType.OnlyOffice);
        public ICommand SelectLibreOfficeCommand => new RelayCommand(() => CurrentProductType = ProductType.LibreOffice);

        // ========== 已安装版本检测与警告属性 ==========
        private string _installedVersionText = "";
        public string InstalledVersionText { get => _installedVersionText; set => Set(ref _installedVersionText, value); }

        private bool _isInstalledWarningVisible;
        public bool IsInstalledWarningVisible { get => _isInstalledWarningVisible; set => Set(ref _isInstalledWarningVisible, value); }

        public void RefreshInstalledVersion()
        {
            var version = RegistryHelper.GetInstalledProductVersion(CurrentProductType);
            if (!string.IsNullOrEmpty(version))
            {
                InstalledVersionText = $"系统检测到已安装版本: {version}。建议先进行卸载以防冲突。";
                IsInstalledWarningVisible = true;
            }
            else
            {
                InstalledVersionText = "";
                IsInstalledWarningVisible = false;
            }
        }

        // ========== WPS 版本选择（仅官方最新版） ==========
        private WpsVersion _selectedWpsVersion = WpsVersion.WpsOfficial;
        public WpsVersion SelectedWpsVersion { get => _selectedWpsVersion; set { Set(ref _selectedWpsVersion, value); OnPropertyChanged(nameof(WpsOfficialSelected)); } }

        public bool WpsOfficialSelected => SelectedWpsVersion == WpsVersion.WpsOfficial;
        public ICommand SelectWpsOfficialCommand => new RelayCommand(() => SelectedWpsVersion = WpsVersion.WpsOfficial);

        // ========== 永中版本选择（仅1个） ==========
        private YozoVersion _selectedYozoVersion = YozoVersion.YozoPersonal;
        public YozoVersion SelectedYozoVersion { get => _selectedYozoVersion; set { Set(ref _selectedYozoVersion, value); OnPropertyChanged(nameof(YozoPersonalSelected)); } }
        public bool YozoPersonalSelected => SelectedYozoVersion == YozoVersion.YozoPersonal;
        public ICommand SelectYozoPersonalCommand => new RelayCommand(() => SelectedYozoVersion = YozoVersion.YozoPersonal);

        // ========== OnlyOffice版本选择（仅1个） ==========
        private OnlyOfficeVersion _selectedOnlyOfficeVersion = OnlyOfficeVersion.OnlyOfficeDesktop;
        public OnlyOfficeVersion SelectedOnlyOfficeVersion { get => _selectedOnlyOfficeVersion; set { Set(ref _selectedOnlyOfficeVersion, value); OnPropertyChanged(nameof(OnlyOfficeSelected)); } }
        public bool OnlyOfficeSelected => SelectedOnlyOfficeVersion == OnlyOfficeVersion.OnlyOfficeDesktop;
        public ICommand SelectOnlyOfficeDesktopCommand => new RelayCommand(() => SelectedOnlyOfficeVersion = OnlyOfficeVersion.OnlyOfficeDesktop);

        // ========== LibreOffice版本选择（仅1个） ==========
        private LibreOfficeVersion _selectedLibreOfficeVersion = LibreOfficeVersion.LibreOfficeStable;
        public LibreOfficeVersion SelectedLibreOfficeVersion { get => _selectedLibreOfficeVersion; set { Set(ref _selectedLibreOfficeVersion, value); OnPropertyChanged(nameof(LibreOfficeSelected)); } }
        public bool LibreOfficeSelected => SelectedLibreOfficeVersion == LibreOfficeVersion.LibreOfficeStable;
        public ICommand SelectLibreOfficeStableCommand => new RelayCommand(() => SelectedLibreOfficeVersion = LibreOfficeVersion.LibreOfficeStable);

        // ========== MS Office 激活方式与配置 ==========
        private OfficeVersion _currentVersion = OfficeVersion.Office2024;
        private ActivationMethod _actMethod = ActivationMethod.Ohook;
        public ActivationMethod CurrentActivationMethod { get => _actMethod; set { Set(ref _actMethod, value); OnPropertyChanged(nameof(IsOhook)); OnPropertyChanged(nameof(IsKMS)); } }

        public bool IsOhook { get => CurrentActivationMethod == ActivationMethod.Ohook; set { if (value) CurrentActivationMethod = ActivationMethod.Ohook; } }
        public bool IsKMS { get => CurrentActivationMethod == ActivationMethod.KMS; set { if (value) CurrentActivationMethod = ActivationMethod.KMS; } }

        // ========== MS Office 组件列表 (三列布局 UniformGrid) ==========
        public ObservableCollection<ComponentItem> Components { get; } = new ObservableCollection<ComponentItem>
        {
            new ComponentItem("Word 文字", OfficeComponent.Word, true),
            new ComponentItem("Excel 表格", OfficeComponent.Excel, true),
            new ComponentItem("PowerPoint 演示", OfficeComponent.PowerPoint, true),
            new ComponentItem("Outlook 邮箱", OfficeComponent.Outlook, false),
            new ComponentItem("OneNote 笔记", OfficeComponent.OneNote, false),
            new ComponentItem("Access 数据库", OfficeComponent.Access, false),
            new ComponentItem("Publisher 出版", OfficeComponent.Publisher, false),
            new ComponentItem("Project 项目", OfficeComponent.Project, false),
            new ComponentItem("Visio 绘图", OfficeComponent.Visio, false),
            new ComponentItem("Teams 协作", OfficeComponent.Teams, false),
            new ComponentItem("OneDrive 网盘", OfficeComponent.OneDrive, false)
        };

        // ========== MS Office 翻页与卡片属性 ==========
        private int _versionGroup = 0; // 0: 2024/M365, 1: 2021/2019, 2: 2016
        public int VersionGroup { get => _versionGroup; set { Set(ref _versionGroup, value); RefreshCards(); } }

        private string _leftTitle; public string LeftTitle { get => _leftTitle; set => Set(ref _leftTitle, value); }
        private string _leftSub; public string LeftSub { get => _leftSub; set => Set(ref _leftSub, value); }
        private string _leftDesc; public string LeftDesc { get => _leftDesc; set => Set(ref _leftDesc, value); }
        private bool _leftSelected; public bool LeftSelected { get => _leftSelected; set => Set(ref _leftSelected, value); }

        private string _rightTitle; public string RightTitle { get => _rightTitle; set => Set(ref _rightTitle, value); }
        private string _rightSub; public string RightSub { get => _rightSub; set => Set(ref _rightSub, value); }
        private string _rightDesc; public string RightDesc { get => _rightDesc; set => Set(ref _rightDesc, value); }
        private bool _rightSelected; public bool RightSelected { get => _rightSelected; set => Set(ref _rightSelected, value); }

        private bool _rightVisible = true; public bool RightVisible { get => _rightVisible; set => Set(ref _rightVisible, value); }
        private bool _leftArrowVisible; public bool LeftArrowVisible { get => _leftArrowVisible; set => Set(ref _leftArrowVisible, value); }
        private bool _rightArrowVisible; public bool RightArrowVisible { get => _rightArrowVisible; set => Set(ref _rightArrowVisible, value); }

        // ========== 安装状态、进度 ==========
        private InstallPhase _phase = InstallPhase.Idle;
        public InstallPhase Phase
        {
            get => _phase;
            set
            {
                Set(ref _phase, value);
                OnPropertyChanged(nameof(StatusText));
                OnPropertyChanged(nameof(CanInstall));
                CommandManager.InvalidateRequerySuggested();
            }
        }

        private string _statusText = "准备就绪";
        public string StatusText { get => _statusText; set => Set(ref _statusText, value); }

        private int _downloadProgress;
        public int DownloadProgress { get => _downloadProgress; set => Set(ref _downloadProgress, value); }

        private bool _isProgressVisible;
        public bool IsProgressVisible { get => _isProgressVisible; set => Set(ref _isProgressVisible, value); }

        public bool CanInstall => Phase != InstallPhase.Cleaning && Phase != InstallPhase.Downloading
                               && Phase != InstallPhase.Installing && Phase != InstallPhase.Activating;

        // ========== 标题点击（关于） ==========
        public ICommand TitleClickCommand => new RelayCommand(() =>
        {
            MessageBox.Show("GOI - Glim Office Installer\n\n版本 2.0.0\n© 2025-2026 GlimStudio\n\n本软件仅供学习研究使用。",
                "关于 GOI");
        });

        // ========== 导航命令 ==========
        public ICommand InstallCommand => new RelayCommand(async () => await InstallAsync(), () => CanInstall);
        public ICommand UninstallCurrentCommand => new RelayCommand(async () => await UninstallCurrentAsync(), () => CanInstall);
        public ICommand SelectLeftCommand => new RelayCommand(SelectLeft);
        public ICommand SelectRightCommand => new RelayCommand(SelectRight);
        public ICommand PrevGroupCommand => new RelayCommand(() => { if (VersionGroup > 0) VersionGroup--; });
        public ICommand NextGroupCommand => new RelayCommand(() => { if (VersionGroup < 2) VersionGroup++; });

        private void SelectLeft()
        {
            LeftSelected = true; RightSelected = false;
            _currentVersion = VersionGroup switch
            {
                0 => OfficeVersion.Office2024, 1 => OfficeVersion.Office2021,
                2 => OfficeVersion.Office2016, _ => OfficeVersion.Office2024
            };
            IsM365 = false;
        }

        private void SelectRight()
        {
            LeftSelected = false; RightSelected = true;
            _currentVersion = OfficeVersion.Microsoft365Pro;
            IsM365 = true;
        }

        private bool _isM365;
        public bool IsM365 { get => _isM365; set => Set(ref _isM365, value); }

        private void RefreshCards()
        {
            switch (VersionGroup)
            {
                case 0:
                    LeftTitle = "Office 2024"; LeftSub = "最新功能 · 持续更新"; LeftDesc = "零售版";
                    RightTitle = "Microsoft 365"; RightSub = "云端同步 · 订阅服务"; RightDesc = "个人/家庭版";
                    RightVisible = true; break;
                case 1:
                    LeftTitle = "Office 2021"; LeftSub = "买断制 · 永久授权"; LeftDesc = "零售版";
                    RightTitle = "Office 2019"; RightSub = "经典稳定 · 广泛兼容"; RightDesc = "零售版";
                    RightVisible = true; break;
                case 2:
                    LeftTitle = "Office 2016"; LeftSub = "经典版本 · 兼容性好"; LeftDesc = "零售版";
                    RightTitle = ""; RightSub = ""; RightDesc = "";
                    RightVisible = false; break;
            }
            LeftArrowVisible = VersionGroup > 0;
            RightArrowVisible = VersionGroup < 2;
            SelectLeft();
        }

        // ========== 安装入口（自动分流不同 Office 软件）==========
        private CancellationTokenSource _installCts;

        private async Task InstallAsync()
        {
            Phase = InstallPhase.Downloading;
            IsProgressVisible = true;
            DownloadProgress = 0;
            _installCts = new CancellationTokenSource();

            var phases = new Progress<string>(msg =>
            {
                StatusText = msg;
                if (msg.Contains("清理")) Phase = InstallPhase.Cleaning;
                else if (msg.Contains("下载")) Phase = InstallPhase.Downloading;
                else if (msg.Contains("安装")) Phase = InstallPhase.Installing;
                else if (msg.Contains("激活")) Phase = InstallPhase.Activating;
            });
            var dl = new Progress<int>(p => DownloadProgress = p);

            bool ok;
            if (CurrentProductType == ProductType.Wps)
            {
                ok = await _wpsService.InstallAsync(SelectedWpsVersion, phases, dl, _installCts.Token);
            }
            else if (CurrentProductType == ProductType.Yozo)
            {
                ok = await _yozoService.InstallAsync(SelectedYozoVersion, phases, dl, _installCts.Token);
            }
            else if (CurrentProductType == ProductType.OnlyOffice)
            {
                ok = await _onlyOfficeService.InstallAsync(SelectedOnlyOfficeVersion, phases, dl, _installCts.Token);
            }
            else if (CurrentProductType == ProductType.LibreOffice)
            {
                ok = await _libreOfficeService.InstallAsync(SelectedLibreOfficeVersion, phases, dl, _installCts.Token);
            }
            else
            {
                var selected = new HashSet<OfficeComponent>(Components.Where(c => c.IsSelected).Select(c => c.Component));
                if (selected.Count == 0) { Phase = InstallPhase.Idle; IsProgressVisible = false; return; }
                ok = await _installService.RunAsync(_currentVersion, DetectedArch, selected, phases, dl);
            }

            Phase = ok ? InstallPhase.Completed : InstallPhase.Failed;
            StatusText = ok
                ? "部署完成！软件已成功安装。"
                : "安装失败，请查看日志了解详情。";
            IsProgressVisible = false;

            // 安装完成刷新已安装版本状态
            RefreshInstalledVersion();
        }

        // ========== 独立深层卸载当前软件 ==========
        private async Task UninstallCurrentAsync()
        {
            string productName = CurrentProductType switch
            {
                ProductType.MsOffice => "MS Office",
                ProductType.Wps => "WPS Office",
                ProductType.Yozo => "永中 Office",
                ProductType.OnlyOffice => "OnlyOffice",
                ProductType.LibreOffice => "LibreOffice",
                _ => "Office"
            };

            var result = MessageBox.Show(
                $"深度卸载将强制终止所有正在运行的 {productName} 进程，并删除相关的注册表及残留文件夹。\n\n确认要继续吗？请务必先保存正在编辑的文档。",
                $"{productName} 深度卸载确认",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (result != MessageBoxResult.Yes) return;

            Phase = InstallPhase.Cleaning;
            IsProgressVisible = true;
            DownloadProgress = 10;
            StatusText = $"正在清理 {productName} 残留...";

            var phases = new Progress<string>(msg => StatusText = msg);

            try
            {
                await _cleanupService.CleanAsync(CurrentProductType, phases);
                DownloadProgress = 100;
                Phase = InstallPhase.Completed;
                StatusText = $"{productName} 残留清理完成！";

                // 刷新已安装版本状态
                RefreshInstalledVersion();
            }
            catch (Exception ex)
            {
                Logger.Error($"{productName} 深度卸载失败", ex);
                Phase = InstallPhase.Failed;
                StatusText = "清理失败: " + ex.Message;
            }
            finally
            {
                IsProgressVisible = false;
            }
        }
    }

    public class ComponentItem : ObservableObject
    {
        public string Name { get; }
        public OfficeComponent Component { get; }
        private bool _isSelected;
        public bool IsSelected { get => _isSelected; set => Set(ref _isSelected, value); }
        public ComponentItem(string name, OfficeComponent c, bool sel = false)
        { Name = name; Component = c; _isSelected = sel; }
    }
}
