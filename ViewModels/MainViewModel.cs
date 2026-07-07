using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using GOI.Models;
using GOI.Services;

namespace GOI.ViewModels
{
    public class MainViewModel : ObservableObject
    {
        private readonly InstallService _installService;

        public MainViewModel(InstallService installService)
        {
            _installService = installService;
            DetectedArch = Environment.Is64BitOperatingSystem ? Architecture.x64 : Architecture.x86;
            ArchText = DetectedArch == Architecture.x64 ? "系统架构：x64（64 位）" : "系统架构：x86（32 位）";
        }

        // ========== 架构 ==========
        private Architecture DetectedArch { get; } = Architecture.x64;
        private string _archText = "";
        public string ArchText { get => _archText; set => Set(ref _archText, value); }

        // ========== 版本卡片 ==========
        private int _versionGroup;
        public int VersionGroup { get => _versionGroup; set { Set(ref _versionGroup, value); RefreshCards(); } }

        private string _leftTitle = "Office 2024";
        public string LeftTitle { get => _leftTitle; set => Set(ref _leftTitle, value); }
        private string _leftSub = "最新功能 · 持续更新";
        public string LeftSub { get => _leftSub; set => Set(ref _leftSub, value); }
        private string _leftDesc = "零售版";
        public string LeftDesc { get => _leftDesc; set => Set(ref _leftDesc, value); }

        private string _rightTitle = "Microsoft 365";
        public string RightTitle { get => _rightTitle; set => Set(ref _rightTitle, value); }
        private string _rightSub = "云端同步 · 订阅服务";
        public string RightSub { get => _rightSub; set => Set(ref _rightSub, value); }
        private string _rightDesc = "个人/家庭版";
        public string RightDesc { get => _rightDesc; set => Set(ref _rightDesc, value); }

        private bool _leftSelected = true;
        public bool LeftSelected { get => _leftSelected; set => Set(ref _leftSelected, value); }
        private bool _rightSelected;
        public bool RightSelected { get => _rightSelected; set => Set(ref _rightSelected, value); }
        private bool _rightVisible = true;
        public bool RightVisible { get => _rightVisible; set => Set(ref _rightVisible, value); }
        private bool _leftArrowVisible;
        public bool LeftArrowVisible { get => _leftArrowVisible; set => Set(ref _leftArrowVisible, value); }
        private bool _rightArrowVisible = true;
        public bool RightArrowVisible { get => _rightArrowVisible; set => Set(ref _rightArrowVisible, value); }
        private OfficeVersion _currentVersion = OfficeVersion.Office2024;

        // ========== 组件选择 ==========
        public ObservableCollection<ComponentItem> Components { get; } = new ObservableCollection<ComponentItem>
        {
            new ComponentItem("PowerPoint", OfficeComponent.PowerPoint, true),
            new ComponentItem("Word",       OfficeComponent.Word,       true),
            new ComponentItem("Excel",      OfficeComponent.Excel,      true),
            new ComponentItem("Visio",      OfficeComponent.Visio,      false),
            new ComponentItem("Access",     OfficeComponent.Access,     false),
            new ComponentItem("OneNote",    OfficeComponent.OneNote,    false),
            new ComponentItem("Lync",       OfficeComponent.Lync,       false),
            new ComponentItem("Outlook",    OfficeComponent.Outlook,    false),
            new ComponentItem("Teams",      OfficeComponent.Teams,      false),
            new ComponentItem("OneDrive",   OfficeComponent.OneDrive,   false),
            new ComponentItem("Publisher",  OfficeComponent.Publisher,  false),
            new ComponentItem("Project",    OfficeComponent.Project,    false),
        };

        private bool _isM365;
        public bool IsM365 { get => _isM365; set => Set(ref _isM365, value); }

        // ========== 安装状态 ==========
        private InstallPhase _phase = InstallPhase.Idle;
        public InstallPhase Phase { get => _phase; set { Set(ref _phase, value); OnPropertyChanged(nameof(CanInstall)); } }

        private string _statusText = "建议保持网络连接，以便完成激活过程";
        public string StatusText { get => _statusText; set => Set(ref _statusText, value); }

        private int _downloadProgress;
        public int DownloadProgress { get => _downloadProgress; set => Set(ref _downloadProgress, value); }

        private bool _isProgressVisible;
        public bool IsProgressVisible { get => _isProgressVisible; set => Set(ref _isProgressVisible, value); }

        public bool CanInstall => Phase != InstallPhase.Cleaning && Phase != InstallPhase.Downloading
                               && Phase != InstallPhase.Installing && Phase != InstallPhase.Activating;

        // ========== 产品选择标题点击 ==========
        public ICommand TitleClickCommand => new RelayCommand(() =>
        {
            MessageBox.Show("GOI - Glim Office Installer\n\n版本 2.0.0\n© 2025-2026 GlimStudio\n\n本软件仅供学习研究使用。",
                "关于 GOI");
        });

        // ========== 命令 ==========
        public ICommand InstallCommand => new RelayCommand(async () => await InstallAsync(), () => CanInstall);
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
            _currentVersion = VersionGroup switch
            {
                0 => OfficeVersion.Microsoft365Pro, 1 => OfficeVersion.Office2019,
                _ => OfficeVersion.Microsoft365Pro
            };
            IsM365 = true;
        }

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

        // ========== 安装 ==========
        private async Task InstallAsync()
        {
            var selected = new HashSet<OfficeComponent>(Components.Where(c => c.IsSelected).Select(c => c.Component));
            if (selected.Count == 0) return;

            Phase = InstallPhase.Cleaning;
            IsProgressVisible = true;
            DownloadProgress = 0;

            var phases = new Progress<string>(msg =>
            {
                StatusText = msg;
                if (msg.Contains("清理")) Phase = InstallPhase.Cleaning;
                else if (msg.Contains("下载")) Phase = InstallPhase.Downloading;
                else if (msg.Contains("安装")) Phase = InstallPhase.Installing;
                else if (msg.Contains("激活")) Phase = InstallPhase.Activating;
            });
            var dl = new Progress<int>(p => DownloadProgress = p);

            bool ok = await _installService.RunAsync(_currentVersion, DetectedArch, selected, phases, dl);
            Phase = ok ? InstallPhase.Completed : InstallPhase.Failed;
            StatusText = ok ? "激活完成！Office已成功安装并激活。" : "安装失败，请查看日志了解详情。";
            IsProgressVisible = false;
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
