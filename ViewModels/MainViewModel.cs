using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.IO;
using System.Text;
using System.Diagnostics;
using Microsoft.Win32;
using iNKORE.UI.WPF.Modern;
using iNKORE.UI.WPF.Modern.Controls;
using System.Windows.Controls;
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

		private string _archText = "";

		private ProductType _productType = ProductType.MsOffice;

		private string _installedVersionText = "";

		private bool _isInstalledWarningVisible;

		private WpsVersion _selectedWpsVersion = WpsVersion.WpsLatest;

		private OfficeVersion _currentVersion = OfficeVersion.Office2024;

		private bool _isM365;

		private string _selectedBitness = "64";

		private string _selectedUpdateChannel = "Current";

		private string _selectedOfficeLanguage = "zh-cn (简体中文)";

		private InstallPhase _phase = InstallPhase.Idle;

		private string _statusText = "准备就绪";

		private int _downloadProgress;

		private bool _isProgressVisible;

		private bool _isProgressIndeterminate;

		private string _selectedTheme = "Default";

		private string _selectedAppLanguage = "SimplifiedChinese";

		private bool _isExportXmlEnabled = false;

		private ProductType _productTypeToUninstall = ProductType.MsOffice;

		private CancellationTokenSource _installCts;

		public LocalizationStrings Loc { get; } = new LocalizationStrings();

		private GOI.Models.Architecture DetectedArch { get; } = Environment.Is64BitOperatingSystem ? GOI.Models.Architecture.x64 : GOI.Models.Architecture.x86;

		public string ArchText
		{
			get
			{
				return _archText;
			}
			set
			{
				Set(ref _archText, value, "ArchText");
			}
		}

		public ProductType CurrentProductType
		{
			get
			{
				return _productType;
			}
			set
			{
				if (Set(ref _productType, value, "CurrentProductType"))
				{
					OnPropertyChanged("IsMsOffice");
					OnPropertyChanged("IsWps");
					OnPropertyChanged("IsYozo");
					OnPropertyChanged("IsOnlyOffice");
					OnPropertyChanged("IsLibreOffice");
					OnPropertyChanged("IsSettings");
					OnPropertyChanged("IsActionButtonsVisible");
					OnPropertyChanged("HeaderIconPath");
					OnPropertyChanged("HeaderTitleText");
					OnPropertyChanged("IsExportXmlButtonVisible");
					RefreshInstalledVersion();
				}
			}
		}

		public bool IsMsOffice => CurrentProductType == ProductType.MsOffice;

		public bool IsWps => CurrentProductType == ProductType.Wps;

		public bool IsYozo => CurrentProductType == ProductType.Yozo;

		public bool IsOnlyOffice => CurrentProductType == ProductType.OnlyOffice;

		public bool IsLibreOffice => CurrentProductType == ProductType.LibreOffice;

		public bool IsSettings => CurrentProductType == ProductType.Settings;

		public bool IsActionButtonsVisible => CurrentProductType != ProductType.Settings;

		public string HeaderIconPath
		{
			get
			{
				ProductType currentProductType = CurrentProductType;
				
				string result = currentProductType switch
				{
					ProductType.MsOffice => "pack://application:,,,/GOI;component/Resources/ms_office_logo.png", 
					ProductType.Wps => "pack://application:,,,/GOI;component/Resources/wps_office_logo.png", 
					ProductType.Yozo => "pack://application:,,,/GOI;component/Resources/yozo_office_logo.png", 
					ProductType.OnlyOffice => "pack://application:,,,/GOI;component/Resources/onlyoffice_logo.png", 
					ProductType.LibreOffice => "pack://application:,,,/GOI;component/Resources/libreoffice_logo.png", 
					ProductType.Settings => "pack://application:,,,/GOI;component/Resources/logo.png", 
					_ => "pack://application:,,,/GOI;component/Resources/logo.png", 
				};
				
				return result;
			}
		}

		public string GetProductDisplayName(ProductType type)
		{
			return type switch
			{
				ProductType.MsOffice => "Microsoft Office",
				ProductType.Wps => "WPS Office",
				ProductType.Yozo => Loc.YozoTitle,
				ProductType.OnlyOffice => "OnlyOffice",
				ProductType.LibreOffice => "LibreOffice",
				ProductType.Settings => Loc.NavSettings,
				_ => "Office"
			};
		}

		public string HeaderTitleText
		{
			get
			{
				string name = GetProductDisplayName(CurrentProductType);
				return (name == "Office") ? Loc.AppTitle : name;
			}
		}

		public ICommand SelectMsOfficeCommand => new RelayCommand(delegate
		{
			CurrentProductType = ProductType.MsOffice;
		});

		public ICommand SelectWpsCommand => new RelayCommand(delegate
		{
			CurrentProductType = ProductType.Wps;
		});

		public ICommand SelectYozoCommand => new RelayCommand(delegate
		{
			CurrentProductType = ProductType.Yozo;
		});

		public ICommand SelectOnlyOfficeCommand => new RelayCommand(delegate
		{
			CurrentProductType = ProductType.OnlyOffice;
		});

		public ICommand SelectLibreOfficeCommand => new RelayCommand(delegate
		{
			CurrentProductType = ProductType.LibreOffice;
		});

		public string InstalledVersionText
		{
			get
			{
				return _installedVersionText;
			}
			set
			{
				Set(ref _installedVersionText, value, "InstalledVersionText");
			}
		}

		public bool IsInstalledWarningVisible
		{
			get
			{
				return _isInstalledWarningVisible;
			}
			set
			{
				Set(ref _isInstalledWarningVisible, value, "IsInstalledWarningVisible");
			}
		}

		public WpsVersion SelectedWpsVersion
		{
			get
			{
				return _selectedWpsVersion;
			}
			set
			{
				if (Set(ref _selectedWpsVersion, value, "SelectedWpsVersion"))
				{
					OnPropertyChanged("Wps2013Selected");
					OnPropertyChanged("Wps2016Selected");
					OnPropertyChanged("Wps2019Selected");
					OnPropertyChanged("Wps2023Selected");
					OnPropertyChanged("WpsLatestSelected");
				}
			}
		}

		public bool Wps2013Selected => SelectedWpsVersion == WpsVersion.Wps2013;

		public bool Wps2016Selected => SelectedWpsVersion == WpsVersion.Wps2016;

		public bool Wps2019Selected => SelectedWpsVersion == WpsVersion.Wps2019;

		public bool Wps2023Selected => SelectedWpsVersion == WpsVersion.Wps2023;

		public bool WpsLatestSelected => SelectedWpsVersion == WpsVersion.WpsLatest;

		public ICommand SelectWps2013Command => new RelayCommand(delegate
		{
			SelectedWpsVersion = WpsVersion.Wps2013;
		});

		public ICommand SelectWps2016Command => new RelayCommand(delegate
		{
			SelectedWpsVersion = WpsVersion.Wps2016;
		});

		public ICommand SelectWps2019Command => new RelayCommand(delegate
		{
			SelectedWpsVersion = WpsVersion.Wps2019;
		});

		public ICommand SelectWps2023Command => new RelayCommand(delegate
		{
			SelectedWpsVersion = WpsVersion.Wps2023;
		});

		public ICommand SelectWpsLatestCommand => new RelayCommand(delegate
		{
			SelectedWpsVersion = WpsVersion.WpsLatest;
		});

		public ICommand OpenUrlCommand => new RelayCommand<string>(delegate(string url)
		{
			try
			{
				System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(url) { UseShellExecute = true });
			}
			catch (Exception ex)
			{
				GOI.Helpers.Logger.Error("Failed to open URL: " + url, ex);
			}
		});

		private bool _yozoPersonalSelected = true;
		public bool YozoPersonalSelected
		{
			get => _yozoPersonalSelected;
			set
			{
				_yozoPersonalSelected = value;
				OnPropertyChanged("YozoPersonalSelected");
			}
		}
		public ICommand SelectYozoPersonalCommand => new RelayCommand(delegate {});

		private bool _onlyOfficeSelected = true;
		public bool OnlyOfficeSelected
		{
			get => _onlyOfficeSelected;
			set
			{
				_onlyOfficeSelected = value;
				OnPropertyChanged("OnlyOfficeSelected");
			}
		}
		public ICommand SelectOnlyOfficeDesktopCommand => new RelayCommand(delegate {});

		private bool _libreOfficeSelected = true;
		public bool LibreOfficeSelected
		{
			get => _libreOfficeSelected;
			set
			{
				_libreOfficeSelected = value;
				OnPropertyChanged("LibreOfficeSelected");
			}
		}
		public ICommand SelectLibreOfficeStableCommand => new RelayCommand(delegate {});

		public bool IsM365
		{
			get
			{
				return _isM365;
			}
			set
			{
				Set(ref _isM365, value, "IsM365");
			}
		}

		public bool Office2024Selected => _currentVersion == OfficeVersion.Office2024 && !IsM365;

		public bool M365Selected => IsM365;

		public bool Office2021Selected => _currentVersion == OfficeVersion.Office2021 && !IsM365;

		public bool Office2019Selected => _currentVersion == OfficeVersion.Office2019 && !IsM365;

		public bool Office2016Selected => _currentVersion == OfficeVersion.Office2016 && !IsM365;

		public ICommand SelectOffice2024Command => new RelayCommand(delegate
		{
			SetOfficeVersion(OfficeVersion.Office2024, isM365: false);
		});

		public ICommand SelectM365Command => new RelayCommand(delegate
		{
			SetOfficeVersion(OfficeVersion.Microsoft365Pro, isM365: true);
		});

		public ICommand SelectOffice2021Command => new RelayCommand(delegate
		{
			SetOfficeVersion(OfficeVersion.Office2021, isM365: false);
		});

		public ICommand SelectOffice2019Command => new RelayCommand(delegate
		{
			SetOfficeVersion(OfficeVersion.Office2019, isM365: false);
		});

		public ICommand SelectOffice2016Command => new RelayCommand(delegate
		{
			SetOfficeVersion(OfficeVersion.Office2016, isM365: false);
		});

		public string SelectedBitness
		{
			get
			{
				return _selectedBitness;
			}
			set
			{
				Set(ref _selectedBitness, value, "SelectedBitness");
			}
		}

		public string SelectedUpdateChannel
		{
			get
			{
				return _selectedUpdateChannel;
			}
			set
			{
				Set(ref _selectedUpdateChannel, value, "SelectedUpdateChannel");
			}
		}

		public string SelectedOfficeLanguage
		{
			get
			{
				return _selectedOfficeLanguage;
			}
			set
			{
				Set(ref _selectedOfficeLanguage, value, "SelectedOfficeLanguage");
			}
		}

		public string ResolvedOfficeLanguage
		{
			get
			{
				if (string.IsNullOrWhiteSpace(SelectedOfficeLanguage))
				{
					return "zh-cn";
				}
				string raw = SelectedOfficeLanguage.Trim();
				int spaceIdx = raw.IndexOf(' ');
				string langCode = (spaceIdx >= 0) ? raw.Substring(0, spaceIdx) : raw;
				return langCode.ToLowerInvariant().Trim();
			}
		}

		public ObservableCollection<ComponentItem> Components { get; }

		public InstallPhase Phase
		{
			get
			{
				return _phase;
			}
			set
			{
				if (Set(ref _phase, value, "Phase"))
				{
					OnPropertyChanged("StatusText");
					OnPropertyChanged("CanInstall");
					IsProgressIndeterminate = (value == InstallPhase.Cleaning || value == InstallPhase.Installing || value == InstallPhase.Activating);
					CommandManager.InvalidateRequerySuggested();
				}
			}
		}

		public string StatusText
		{
			get
			{
				return _statusText;
			}
			set
			{
				Set(ref _statusText, value, "StatusText");
			}
		}

		public int DownloadProgress
		{
			get
			{
				return _downloadProgress;
			}
			set
			{
				Set(ref _downloadProgress, value, "DownloadProgress");
			}
		}

		public bool IsProgressVisible
		{
			get
			{
				return _isProgressVisible;
			}
			set
			{
				Set(ref _isProgressVisible, value, "IsProgressVisible");
			}
		}

		public bool IsProgressIndeterminate
		{
			get
			{
				return _isProgressIndeterminate;
			}
			set
			{
				Set(ref _isProgressIndeterminate, value, "IsProgressIndeterminate");
			}
		}

		public bool CanInstall => Phase != InstallPhase.Cleaning && Phase != InstallPhase.Downloading && Phase != InstallPhase.Installing && Phase != InstallPhase.Activating;

		public string SelectedTheme
		{
			get
			{
				return _selectedTheme;
			}
			set
			{
				if (Set(ref _selectedTheme, value, "SelectedTheme"))
				{
					UpdateAppTheme(value);
				}
			}
		}

		public string SelectedAppLanguage
		{
			get
			{
				return _selectedAppLanguage;
			}
			set
			{
				if (Set(ref _selectedAppLanguage, value, "SelectedAppLanguage"))
				{
					LocalizationStrings.AppLanguage appLanguage = value switch
					{
						"TraditionalChinese" => LocalizationStrings.AppLanguage.TraditionalChinese,
						"English" => LocalizationStrings.AppLanguage.English,
						_ => LocalizationStrings.AppLanguage.SimplifiedChinese
					};
					SetAppLanguage(appLanguage);
				}
			}
		}

		public bool IsExportXmlEnabled
		{
			get
			{
				return _isExportXmlEnabled;
			}
			set
			{
				if (Set(ref _isExportXmlEnabled, value, "IsExportXmlEnabled"))
				{
					OnPropertyChanged("IsExportXmlButtonVisible");
				}
			}
		}

		public bool IsExportXmlButtonVisible => IsMsOffice && IsExportXmlEnabled;

		public ProductType ProductTypeToUninstall
		{
			get
			{
				return _productTypeToUninstall;
			}
			set
			{
				Set(ref _productTypeToUninstall, value, "ProductTypeToUninstall");
			}
		}

		public ICommand UninstallSelectedCommand => new RelayCommand(async delegate
		{
			await UninstallSelectedAsync();
		}, () => CanInstall);

		public ICommand ClearOfficeActivationCommand => new RelayCommand(async delegate
		{
			await ClearOfficeActivationAsync();
		}, () => CanInstall);

		public ICommand ActivateOfficeOhookCommand => new RelayCommand(async delegate
		{
			await ActivateOfficeOhookAsync();
		}, () => CanInstall);

		public ICommand CleanOrphanedAssociationsCommand => new RelayCommand(async delegate
		{
			await CleanOrphanedAssociationsAsync();
		}, () => CanInstall);

		public ICommand RefreshIconCacheCommand => new RelayCommand(async delegate
		{
			await RefreshIconCacheAsync();
		}, () => CanInstall);

		public ICommand RepairAssociationsCommand => new RelayCommand(async delegate
		{
			await RepairAssociationsAsync();
		}, () => CanInstall);

		public ICommand RepairCOMCommand => new RelayCommand(async delegate
		{
			await RepairCOMAsync();
		}, () => CanInstall);

		public ICommand ExportXmlCommand => new RelayCommand(ExportXml);

		public ICommand InstallCommand => new RelayCommand(async delegate
		{
			await InstallAsync();
		}, () => CanInstall);

		public MainViewModel(InstallService installService)
		{
			_installService = installService;
			_wpsService = new WpsInstallService();
			_yozoService = new YozoInstallService();
			_onlyOfficeService = new OnlyOfficeInstallService();
			_libreOfficeService = new LibreOfficeInstallService();
			_cleanupService = new CleanupService();
			ArchText = ((DetectedArch == GOI.Models.Architecture.x64) ? Loc.ArchX64 : Loc.ArchX86);
			SelectedBitness = (Environment.Is64BitOperatingSystem ? "64" : "32");
			LocalizationStrings.AppLanguage detected = LocalizationStrings.Detected;
			
			string selectedOfficeLanguage = detected switch
			{
				LocalizationStrings.AppLanguage.TraditionalChinese => "zh-tw (繁體中文)", 
				LocalizationStrings.AppLanguage.English => "en-us (English - US)", 
				_ => "zh-cn (简体中文)", 
			};
			
			SelectedOfficeLanguage = selectedOfficeLanguage;
			SelectedUpdateChannel = "Current";
			LocalizationStrings.AppLanguage detected2 = LocalizationStrings.Detected;
			
			selectedOfficeLanguage = detected2 switch
			{
				LocalizationStrings.AppLanguage.TraditionalChinese => "TraditionalChinese", 
				LocalizationStrings.AppLanguage.English => "English", 
				_ => "SimplifiedChinese", 
			};
			
			_selectedAppLanguage = selectedOfficeLanguage;
			Components = new ObservableCollection<ComponentItem>
			{
				new ComponentItem(Loc.CompWord, OfficeComponent.Word, sel: true),
				new ComponentItem(Loc.CompExcel, OfficeComponent.Excel, sel: true),
				new ComponentItem(Loc.CompPowerPoint, OfficeComponent.PowerPoint, sel: true),
				new ComponentItem(Loc.CompOutlook, OfficeComponent.Outlook),
				new ComponentItem(Loc.CompOneNote, OfficeComponent.OneNote),
				new ComponentItem(Loc.CompAccess, OfficeComponent.Access),
				new ComponentItem(Loc.CompPublisher, OfficeComponent.Publisher),
				new ComponentItem(Loc.CompProject, OfficeComponent.Project),
				new ComponentItem(Loc.CompVisio, OfficeComponent.Visio),
				new ComponentItem(Loc.CompTeams, OfficeComponent.Teams),
				new ComponentItem(Loc.CompOneDrive, OfficeComponent.OneDrive)
			};
			SetOfficeVersion(OfficeVersion.Office2024, isM365: false);
			RefreshInstalledVersion();
			StatusText = Loc.StatusReady;
		}

		public void RefreshInstalledVersion()
		{
			System.Threading.Tasks.Task.Run(() =>
			{
				string installedProductVersion = RegistryHelper.GetInstalledProductVersion(CurrentProductType);
				System.Windows.Application.Current?.Dispatcher?.Invoke(() =>
				{
					if (!string.IsNullOrEmpty(installedProductVersion))
					{
						InstalledVersionText = Loc.DlgConfirmInstallMsg(installedProductVersion);
						IsInstalledWarningVisible = true;
					}
					else
					{
						InstalledVersionText = "";
						IsInstalledWarningVisible = false;
					}
				});
			});
		}

		private void SetOfficeVersion(OfficeVersion version, bool isM365)
		{
			_currentVersion = version;
			IsM365 = isM365;
			OnPropertyChanged("Office2024Selected");
			OnPropertyChanged("M365Selected");
			OnPropertyChanged("Office2021Selected");
			OnPropertyChanged("Office2019Selected");
			OnPropertyChanged("Office2016Selected");
		}

		private void UpdateAppTheme(string themeStr)
		{
			if (themeStr == "Light")
			{
				ThemeManager.Current.ApplicationTheme = ApplicationTheme.Light;
			}
			else if (themeStr == "Dark")
			{
				ThemeManager.Current.ApplicationTheme = ApplicationTheme.Dark;
			}
			else
			{
				ThemeManager.Current.ApplicationTheme = null;
			}
		}

		private void SetAppLanguage(LocalizationStrings.AppLanguage lang)
		{
			LocalizationStrings.Detected = lang;
			OnPropertyChanged(string.Empty);
			UpdateComponentsLanguage();
			ArchText = ((DetectedArch == GOI.Models.Architecture.x64) ? Loc.ArchX64 : Loc.ArchX86);
			RefreshInstalledVersion();
		}

		private void UpdateComponentsLanguage()
		{
			if (Components == null)
			{
				return;
			}
			foreach (var item in Components)
			{
				switch (item.Component)
				{
					case OfficeComponent.Word: item.Name = Loc.CompWord; break;
					case OfficeComponent.Excel: item.Name = Loc.CompExcel; break;
					case OfficeComponent.PowerPoint: item.Name = Loc.CompPowerPoint; break;
					case OfficeComponent.Outlook: item.Name = Loc.CompOutlook; break;
					case OfficeComponent.OneNote: item.Name = Loc.CompOneNote; break;
					case OfficeComponent.Access: item.Name = Loc.CompAccess; break;
					case OfficeComponent.Publisher: item.Name = Loc.CompPublisher; break;
					case OfficeComponent.Project: item.Name = Loc.CompProject; break;
					case OfficeComponent.Visio: item.Name = Loc.CompVisio; break;
					case OfficeComponent.Teams: item.Name = Loc.CompTeams; break;
					case OfficeComponent.OneDrive: item.Name = Loc.CompOneDrive; break;
				}
			}
		}

		private async Task InstallAsync()
		{
			if (_installCts != null)
			{
				try
				{
					_installCts.Cancel();
					_installCts.Dispose();
				}
				catch {}
				_installCts = null;
			}

			_installCts = new CancellationTokenSource();
			try
			{
				Phase = InstallPhase.Downloading;
				IsProgressVisible = true;
				DownloadProgress = 0;
				Progress<string> phases = new Progress<string>(delegate(string msg)
				{
					StatusText = msg;
				});
				Progress<InstallPhase> phaseProgress = new Progress<InstallPhase>(delegate(InstallPhase p)
				{
					Phase = p;
				});
				Progress<int> dl = new Progress<int>(delegate(int p)
				{
					DownloadProgress = p;
				});
				string installedVersion = RegistryHelper.GetInstalledProductVersion(CurrentProductType);
				if (!string.IsNullOrEmpty(installedVersion))
				{
					string companyName = GetProductDisplayName(CurrentProductType);
					if (!(await DialogService.ShowConfirmAsync(Loc.DlgConfirmInstallTitle, Loc.DlgConfirmInstallMsg(installedVersion), Loc.BtnContinue, Loc.BtnCancel)))
					{
						Phase = InstallPhase.Idle;
						IsProgressVisible = false;
						StatusText = Loc.StatusDeploymentCancelled;
						return;
					}
					Phase = InstallPhase.Cleaning;
					DownloadProgress = 10;
					StatusText = Loc.StatusCleaningOldVersions(companyName);
					await _cleanupService.CleanAsync(CurrentProductType, phases);
				}
				bool ok;
				if (CurrentProductType == ProductType.Wps)
				{
					ok = await _wpsService.InstallAsync(SelectedWpsVersion, phases, dl, phaseProgress, _installCts.Token);
				}
				else if (CurrentProductType == ProductType.Yozo)
				{
					ok = await _yozoService.InstallAsync(phases, dl, phaseProgress, _installCts.Token);
				}
				else if (CurrentProductType == ProductType.OnlyOffice)
				{
					ok = await _onlyOfficeService.InstallAsync(phases, dl, phaseProgress, _installCts.Token);
				}
				else if (CurrentProductType == ProductType.LibreOffice)
				{
					ok = await _libreOfficeService.InstallAsync(phases, dl, phaseProgress, _installCts.Token);
				}
				else
				{
					HashSet<OfficeComponent> selected = new HashSet<OfficeComponent>(from c in Components
						where c.IsSelected
						select c.Component);
					if (selected.Count == 0)
					{
						Phase = InstallPhase.Idle;
						IsProgressVisible = false;
						return;
					}
					ok = await _installService.RunAsync(_currentVersion, SelectedBitness, SelectedUpdateChannel, ResolvedOfficeLanguage, selected, autoActivate: true, phases, dl, phaseProgress);
				}
				Phase = (ok ? InstallPhase.Completed : InstallPhase.Failed);
				StatusText = (ok ? Loc.StatusDeploySuccess : Loc.StatusDeployFail);
				IsProgressVisible = false;
				if (ok)
				{
					var info = Loc.GetInstallSuccessInfo(CurrentProductType);
					await DialogService.ShowMessageAsync(info.Title, info.Msg);
				}
				else
				{
					await DialogService.HandleFailureAsync(Loc.DlgDeployFailTitle, Loc.DlgDeployFailMsg);
				}
				RefreshInstalledVersion();
			}
			finally
			{
				if (_installCts != null)
				{
					_installCts.Dispose();
					_installCts = null;
				}
			}
		}

		private async Task UninstallSelectedAsync()
		{
			string productName = GetProductDisplayName(ProductTypeToUninstall);
			if (!(await DialogService.ShowConfirmAsync(Loc.SettingsUninstallTitle, Loc.DlgConfirmUninstallMsg(productName), Loc.BtnUninstall, Loc.BtnCancel)))
			{
				return;
			}
			Phase = InstallPhase.Cleaning;
			IsProgressVisible = true;
			DownloadProgress = 10;
			StatusText = Loc.StatusCleaningOldVersions(productName);
			Progress<string> phases = new Progress<string>(delegate(string msg)
			{
				StatusText = msg;
			});
			try
			{
				await _cleanupService.CleanAsync(ProductTypeToUninstall, phases);
				DownloadProgress = 100;
				Phase = InstallPhase.Completed;
				StatusText = Loc.DlgUninstallSuccessMsg;
				await DialogService.ShowMessageAsync(Loc.DlgUninstallSuccessTitle, Loc.DlgUninstallSuccessMsg);
			}
			catch (Exception ex)
			{
				Logger.Error(productName + " 卸载失败", ex);
				Phase = InstallPhase.Failed;
				StatusText = Loc.ErrInstallFailed(ex.Message);
				await DialogService.HandleFailureAsync(Loc.DlgUninstallFailTitle, ex.Message);
			}
			finally
			{
				IsProgressVisible = false;
				RefreshInstalledVersion();
			}
		}

		private async Task CleanOrphanedAssociationsAsync()
		{
			Phase = InstallPhase.Cleaning;
			IsProgressVisible = true;
			DownloadProgress = 20;
			StatusText = Loc.StatusScanningAssociations;
			Progress<string> phases = new Progress<string>(delegate(string msg)
			{
				StatusText = msg;
			});
			try
			{
				int cleanedCount = 0;
				await Task.Run(delegate
				{
					cleanedCount = RegistryHelper.CleanOrphanedFileAssociations(phases);
				});
				DownloadProgress = 100;
				Phase = InstallPhase.Completed;
				StatusText = Loc.StatusAssociationsCleaned;
				await DialogService.ShowMessageAsync(Loc.DlgCleanAssociationsTitle, Loc.DlgCleanAssociationsMsg);
			}
			catch (Exception ex)
			{
				Logger.Error("文件关联净化失败", ex);
				Phase = InstallPhase.Failed;
				StatusText = Loc.StatusCleanAssociationsFailed(ex.Message);
				await DialogService.HandleFailureAsync(Loc.DlgCleanAssociationsFailTitle, ex.Message);
			}
			finally
			{
				IsProgressVisible = false;
				RefreshInstalledVersion();
			}
		}

		private async Task RefreshIconCacheAsync()
		{
			Phase = InstallPhase.Cleaning;
			IsProgressVisible = true;
			DownloadProgress = 30;
			StatusText = Loc.StatusRebuildingIconCache;
			try
			{
				await Task.Run(delegate
				{
					RegistryHelper.RefreshIconCache();
				});
				DownloadProgress = 100;
				Phase = InstallPhase.Completed;
				StatusText = Loc.StatusIconCacheRefreshed;
				await DialogService.ShowMessageAsync(Loc.DlgRefreshIconCacheTitle, Loc.DlgRefreshIconCacheMsg);
			}
			catch (Exception ex)
			{
				Logger.Error("刷新图标缓存失败", ex);
				Phase = InstallPhase.Failed;
				StatusText = Loc.StatusRefreshIconCacheFailed(ex.Message);
				await DialogService.HandleFailureAsync(Loc.DlgRefreshIconCacheFailTitle, ex.Message);
			}
			finally
			{
				IsProgressVisible = false;
				RefreshInstalledVersion();
			}
		}

		private async Task RepairAssociationsAsync()
		{
			Phase = InstallPhase.Cleaning;
			IsProgressVisible = true;
			DownloadProgress = 30;
			StatusText = Loc.StatusRepairingAssociations;
			try
			{
				Progress<string> progressReporter = new Progress<string>(delegate(string msg)
				{
					StatusText = msg;
				});
				await Task.Run(delegate
				{
					RegistryHelper.RestoreInstalledProductAssociations(progressReporter);
					RegistryHelper.RefreshIconCache();
				});
				DownloadProgress = 100;
				Phase = InstallPhase.Completed;
				StatusText = Loc.StatusAssociationsRepaired;
				await DialogService.ShowMessageAsync(Loc.DlgRepairAssociationsTitle, Loc.DlgRepairAssociationsMsg);
				await PromptUserToSelectDefaultOfficeAppsAsync();
			}
			catch (Exception ex)
			{
				DownloadProgress = 100;
				Phase = InstallPhase.Completed;
				StatusText = Loc.StatusRepairAssociationsFailed(ex.Message);
				Logger.Error("独立修复文件关联失败", ex);
				await DialogService.HandleFailureAsync(Loc.DlgRepairAssociationsFailTitle, Loc.DlgRepairAssociationsFailMsg(ex.Message));
			}
			finally
			{
				IsProgressVisible = false;
				RefreshInstalledVersion();
			}
		}

		private async Task PromptUserToSelectDefaultOfficeAppsAsync()
		{
			List<ProductType> installedProducts = new List<ProductType>();
			foreach (ProductType pt in Enum.GetValues(typeof(ProductType)))
			{
				if (pt != ProductType.MsOffice && !string.IsNullOrEmpty(RegistryHelper.GetInstalledProductVersion(pt)))
				{
					installedProducts.Add(pt);
				}
			}
			bool isMsInstalled = !string.IsNullOrEmpty(RegistryHelper.GetInstalledProductVersion(ProductType.MsOffice));
			bool hasConflict = (isMsInstalled && installedProducts.Count >= 1) || (installedProducts.Count >= 2);

			if (!hasConflict)
			{
				return;
			}

			bool confirmed = await DialogService.ShowConfirmAsync(Loc.DlgDefaultAppTitle, Loc.DlgDefaultAppMsg, Loc.BtnDefaultAppStart, Loc.BtnDefaultAppSkip);
			if (!confirmed)
			{
				return;
			}
			string tempDir = Path.GetTempPath();
			List<(string ext, string name)> formats = new List<(string, string)>
			{
				(".docx", Loc.FormatWord),
				(".doc", Loc.FormatWordLegacy),
				(".xlsx", Loc.FormatExcel),
				(".xls", Loc.FormatExcelLegacy)
			};
			if (isMsInstalled || installedProducts.Contains(ProductType.Wps) || installedProducts.Contains(ProductType.Yozo) || installedProducts.Contains(ProductType.LibreOffice))
			{
				formats.Add((".pptx", Loc.FormatPowerPoint));
				formats.Add((".ppt", Loc.FormatPowerPointLegacy));
			}
			formats.Add((".pdf", Loc.FormatPdf));
			foreach (var item in formats)
			{
				string tempFile = Path.Combine(tempDir, "GOI_Sample_Document" + item.ext);
				try
				{
					if (!File.Exists(tempFile))
					{
						File.WriteAllBytes(tempFile, new byte[0]);
					}
					StatusText = Loc.StatusGuidingDefaultApp(item.name);
					ProcessStartInfo psi = new ProcessStartInfo
					{
						FileName = "rundll32.exe",
						Arguments = "shell32.dll,OpenAs_RunDLL \"" + tempFile + "\"",
						UseShellExecute = false,
						CreateNoWindow = true
					};
					Process process = Process.Start(psi);
					try
					{
						if (process != null)
						{
							await Task.Run(delegate
							{
								process.WaitForExit();
							});
						}
					}
					finally
					{
						if (process != null)
						{
							((IDisposable)process).Dispose();
						}
					}
					await Task.Delay(2000);
				}
				catch (Exception ex)
				{
					Exception ex2 = ex;
					Logger.Error("引导设置 " + item.name + " 默认关联失败", ex2);
				}
				finally
				{
					try
					{
						if (File.Exists(tempFile))
						{
							File.Delete(tempFile);
						}
					}
					catch (Exception ex)
					{
						Logger.Warn("删除默认应用临时配置失败: " + ex.Message);
					}
				}
			}
			StatusText = Loc.StatusDefaultAppGuided;
			await DialogService.ShowMessageAsync(Loc.DlgDefaultAppSuccessTitle, Loc.DlgDefaultAppSuccessMsg);
		}

		private async Task RepairCOMAsync()
		{
			Phase = InstallPhase.Cleaning;
			IsProgressVisible = true;
			DownloadProgress = 20;
			StatusText = Loc.StatusRepairingCOM;
			try
			{
				Progress<string> progressReporter = new Progress<string>(delegate(string msg)
				{
					StatusText = msg;
				});
				await Task.Run(delegate
				{
					RegistryHelper.RepairOfficeComComponents(progressReporter);
				});
				DownloadProgress = 100;
				Phase = InstallPhase.Completed;
				StatusText = Loc.StatusCOMRepaired;
				await DialogService.ShowMessageAsync(Loc.DlgRepairComTitle, Loc.DlgRepairComMsg);
			}
			catch (Exception ex)
			{
				DownloadProgress = 100;
				Phase = InstallPhase.Completed;
				StatusText = Loc.StatusRepairCOMFailed(ex.Message);
				Logger.Error("独立修复 COM 组件失败", ex);
				await DialogService.HandleFailureAsync(Loc.DlgRepairComFailTitle, Loc.DlgRepairComFailMsg(ex.Message));
			}
			finally
			{
				IsProgressVisible = false;
				RefreshInstalledVersion();
			}
		}

		private async Task ClearOfficeActivationAsync()
		{
			StatusText = Loc.StatusScanningActivationKeys;
			IsProgressVisible = true;
			DownloadProgress = 20;
			string osppPath = RegistryHelper.LocateOsppVbs();
			if (osppPath == null)
			{
				await DialogService.ShowMessageAsync(Loc.DlgClearLicenseFailTitle, Loc.ErrOsppNotFound);
				IsProgressVisible = false;
				StatusText = Loc.StatusPathNotFound;
				return;
			}
			ProcessStartInfo startInfo = new ProcessStartInfo
			{
				FileName = "cscript.exe",
				Arguments = "//NoLogo \"" + osppPath + "\" /dstatus",
				RedirectStandardOutput = true,
				UseShellExecute = false,
				CreateNoWindow = true
			};
			List<string> keys = new List<string>();
			await Task.Run(delegate
			{
				try
				{
					using Process process = Process.Start(startInfo);
					if (process != null)
					{
						string text = process.StandardOutput.ReadToEnd();
						process.WaitForExit();
						Logger.Info("OSPP /dstatus 扫描输出:\n" + text);
						string[] array = text.Split(new string[2] { "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
						string[] array2 = array;
						foreach (string text2 in array2)
						{
							if (text2.Contains("Last 5 characters of installed product key:"))
							{
								string[] array3 = text2.Split(':');
								if (array3.Length > 1)
								{
									string text3 = array3[1].Trim();
									if (!string.IsNullOrEmpty(text3) && !keys.Contains(text3))
									{
										keys.Add(text3);
									}
								}
							}
						}
					}
				}
				catch (Exception ex)
				{
					Logger.Error("执行 cscript /dstatus 出错", ex);
				}
			});
			if (keys.Count == 0)
			{
				await DialogService.ShowMessageAsync(Loc.DlgScanNoKeysTitle, Loc.DlgScanNoKeysMsg);
				IsProgressVisible = false;
				StatusText = Loc.StatusNoKeysFound;
				return;
			}
			if (!(await DialogService.ShowConfirmAsync(Loc.DlgConfirmClearTitle, Loc.DlgConfirmClearMsg(keys.Count, string.Join(", ", keys)), Loc.BtnDeleteConfirm, Loc.BtnCancel)))
			{
				StatusText = Loc.StatusClearCancelled;
				IsProgressVisible = false;
				return;
			}
			StatusText = Loc.StatusClearingKeys;
			int deletedCount = 0;
			await Task.Run(delegate
			{
				foreach (string item in keys)
				{
					ProcessStartInfo startInfo2 = new ProcessStartInfo
					{
						FileName = "cscript.exe",
						Arguments = "//NoLogo \"" + osppPath + "\" /unpkey:" + item,
						RedirectStandardOutput = true,
						UseShellExecute = false,
						CreateNoWindow = true
					};
					try
					{
						using Process process = Process.Start(startInfo2);
						if (process != null)
						{
							process.WaitForExit();
							if (process.ExitCode == 0)
							{
								deletedCount++;
								Logger.Info("成功清除密钥: " + item);
							}
							else
							{
								Logger.Warn($"清除密钥失败: {item}, ExitCode: {process.ExitCode}");
							}
						}
					}
					catch (Exception ex)
					{
						Logger.Error("清除密钥异常: " + item, ex);
					}
				}
			});
			DownloadProgress = 100;
			IsProgressVisible = false;
			StatusText = Loc.StatusClearSuccess(deletedCount);
			await DialogService.ShowMessageAsync(Loc.DlgClearLicenseTitle, Loc.DlgClearSuccessMsg(deletedCount));
		}

		private async Task ActivateOfficeOhookAsync()
		{
			if (await DialogService.ShowConfirmAsync(Loc.DlgConfirmOhookTitle, Loc.DlgConfirmOhookMsg, Loc.BtnOhookConfirm, Loc.BtnCancel))
			{
				StatusText = Loc.StatusReleasingOhook;
				IsProgressVisible = true;
				DownloadProgress = 30;
				bool ok = await Task.Run(async () => await _installService.ActivateOhookAsync());
				DownloadProgress = 100;
				IsProgressVisible = false;
				if (ok)
				{
					StatusText = Loc.StatusOhookSuccess;
					await DialogService.ShowMessageAsync(Loc.DlgActivateSuccessTitle, Loc.DlgOhookSuccessMsg);
				}
				else
				{
					StatusText = Loc.StatusOhookFail;
					await DialogService.HandleFailureAsync(Loc.DlgActivateFailTitle, Loc.DlgOhookFailMsg);
				}
				RefreshInstalledVersion();
			}
		}

		private void ExportXml()
		{
			HashSet<OfficeComponent> selected = new HashSet<OfficeComponent>(from c in Components
				where c.IsSelected
				select c.Component);
			string contents = XmlConfigHelper.Generate(_currentVersion, SelectedBitness, SelectedUpdateChannel, ResolvedOfficeLanguage, selected);
			SaveFileDialog saveFileDialog = new SaveFileDialog
			{
				Filter = "XML Files (*.xml)|*.xml",
				FileName = "configuration.xml",
				Title = Loc.BtnExportXml
			};
			if (saveFileDialog.ShowDialog() == true)
			{
				try
				{
					File.WriteAllText(saveFileDialog.FileName, contents, Encoding.UTF8);
					_ = DialogService.ShowMessageAsync(Loc.DlgExportXmlTitle, Loc.DlgExportXmlMsg(saveFileDialog.FileName));
				}
				catch (Exception ex)
				{
					Logger.Error("导出 XML 失败", ex);
					_ = DialogService.ShowMessageAsync(Loc.DlgExportXmlFailTitle, ex.Message);
				}
			}
		}

	}

	public class ComponentItem : ObservableObject
	{
		private bool _isSelected;
		private string _name;

		public string Name
		{
			get { return _name; }
			set { Set(ref _name, value, "Name"); }
		}

		public OfficeComponent Component { get; }

		public bool IsSelected
		{
			get
			{
				return _isSelected;
			}
			set
			{
				Set(ref _isSelected, value, "IsSelected");
			}
		}

		public ComponentItem(string name, OfficeComponent c, bool sel = false)
		{
			_name = name;
			Component = c;
			_isSelected = sel;
		}
	}
}
