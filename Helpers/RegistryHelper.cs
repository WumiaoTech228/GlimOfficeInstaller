using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Threading;
using Microsoft.Win32;
using GOI.Models;

namespace GOI.Helpers
{
	public static class RegistryHelper
	{
		public class FontBackupInfo
		{
			public string Name { get; set; }
			public string FileName { get; set; }
			public string BackupPath { get; set; }
			public bool IsUserFont { get; set; }
		}

		private const int SHCNE_ASSOCCHANGED = 134217728;
		private const int SHCNF_IDLIST = 0;
		private const int SHCNF_FLUSH = 4096;
		private const int WM_FONTCHANGE = 29;
		private static readonly IntPtr HWND_BROADCAST = new IntPtr(65535);

		private static readonly Dictionary<string, string> MsOfficeProgIds = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
		{
			{ ".doc", "Word.Document.8" },
			{ ".docx", "Word.Document.12" },
			{ ".docm", "Word.DocumentMacroEnabled.12" },
			{ ".dot", "Word.Template.8" },
			{ ".dotx", "Word.Template.12" },
			{ ".rtf", "Word.RTF.8" },
			{ ".xls", "Excel.Sheet.8" },
			{ ".xlsx", "Excel.Sheet.12" },
			{ ".xlsm", "Excel.SheetMacroEnabled.12" },
			{ ".xlsb", "Excel.SheetBinaryMacroEnabled.12" },
			{ ".csv", "Excel.CSV" },
			{ ".ppt", "PowerPoint.Show.8" },
			{ ".pptx", "PowerPoint.Show.12" },
			{ ".pptm", "PowerPoint.ShowMacroEnabled.12" },
			{ ".pps", "PowerPoint.SlideShow.8" },
			{ ".ppsx", "PowerPoint.SlideShow.12" }
		};

		private static readonly Dictionary<string, string> WpsProgIds = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
		{
			{ ".doc", "WPS.Doc.8" },
			{ ".docx", "WPS.Doc.12" },
			{ ".xls", "ET.Xls.8" },
			{ ".xlsx", "ET.Xlsx.12" },
			{ ".ppt", "WPP.Ppt.8" },
			{ ".pptx", "WPP.Pptx.12" }
		};

		private static readonly Dictionary<string, (string exeName, int iconIndex)> MsOfficeProgIdIcons = new Dictionary<string, (string, int)>(StringComparer.OrdinalIgnoreCase)
		{
			{ "Word.Document.8", ("wordicon.exe", 1) },
			{ "Word.Document.12", ("wordicon.exe", 13) },
			{ "Word.DocumentMacroEnabled.12", ("wordicon.exe", 15) },
			{ "Word.Template.8", ("wordicon.exe", 2) },
			{ "Word.Template.12", ("wordicon.exe", 14) },
			{ "Word.RTF.8", ("wordicon.exe", 1) },
			{ "Excel.Sheet.8", ("xlicons.exe", 1) },
			{ "Excel.Sheet.12", ("xlicons.exe", 1) },
			{ "Excel.SheetMacroEnabled.12", ("xlicons.exe", 2) },
			{ "Excel.SheetBinaryMacroEnabled.12", ("xlicons.exe", 3) },
			{ "Excel.CSV", ("xlicons.exe", 1) },
			{ "PowerPoint.Show.8", ("pptico.exe", 1) },
			{ "PowerPoint.Show.12", ("pptico.exe", 1) },
			{ "PowerPoint.ShowMacroEnabled.12", ("pptico.exe", 2) },
			{ "PowerPoint.SlideShow.8", ("pptico.exe", 1) },
			{ "PowerPoint.SlideShow.12", ("pptico.exe", 1) }
		};

		#region Facade Methods Delegating to IProductCleaner strategies

		public static void KillOfficeProcesses(ProductType product)
		{
			CleanerFactory.GetCleaner(product)?.KillProcesses();
		}

		public static string GetInstalledProductVersion(ProductType product)
		{
			return CleanerFactory.GetCleaner(product)?.GetInstalledVersion() ?? "";
		}

		public static void CleanUninstallEntries(ProductType product)
		{
			CleanerFactory.GetCleaner(product)?.CleanUninstallEntries();
		}

		public static void CleanResidualFolders(ProductType product)
		{
			CleanerFactory.GetCleaner(product)?.CleanResidualFolders();
		}

		public static void CleanShortcuts(ProductType product)
		{
			CleanerFactory.GetCleaner(product)?.CleanShortcuts();
		}

		public static void CleanFileAssociations(ProductType product)
		{
			CleanerFactory.GetCleaner(product)?.CleanFileAssociations();
		}

		#endregion

		#region Low-Level Utility Helpers

		public static void DeleteKey(string subKeyPath)
		{
			DeleteKey(Registry.CurrentUser, subKeyPath);
			DeleteKey(Registry.LocalMachine, subKeyPath);
		}

		public static void DeleteKey(RegistryKey rootKey, string subKeyPath)
		{
			try
			{
				rootKey.DeleteSubKeyTree(subKeyPath, throwOnMissingSubKey: false);
			}
			catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs at DeleteKey", ex_captured); }
		}

		public static bool IsCurrentUserAdmin()
		{
			try
			{
				using WindowsIdentity identity = WindowsIdentity.GetCurrent();
				WindowsPrincipal principal = new WindowsPrincipal(identity);
				return principal.IsInRole(WindowsBuiltInRole.Administrator);
			}
			catch
			{
				return false;
			}
		}

		public static void ForceDeleteRegistryKey(RegistryKey parentKey, string subKeyName)
		{
			try
			{
				parentKey.DeleteSubKeyTree(subKeyName, throwOnMissingSubKey: false);
				Logger.Info("已直接清理注册表项: " + subKeyName);
				return;
			}
			catch (UnauthorizedAccessException)
			{
			}
			catch (Exception ex)
			{
				Logger.Warn("直接清理注册表项 " + subKeyName + " 异常 (将尝试夺权修复): " + ex.Message);
			}

			if (!IsCurrentUserAdmin())
			{
				Logger.Warn("当前运行权限不足且非管理员，跳过注册表夺权操作: " + subKeyName);
				return;
			}

			try
			{
				using RegistryKey registryKey = parentKey.OpenSubKey(subKeyName, RegistryRights.TakeOwnership);
				if (registryKey != null)
				{
					RegistrySecurity accessControl = registryKey.GetAccessControl();
					accessControl.SetOwner(WindowsIdentity.GetCurrent().User);
					registryKey.SetAccessControl(accessControl);
				}
			}
			catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs at ForceDeleteRegistryKey (TakeOwnership)", ex_captured); }

			try
			{
				using RegistryKey registryKey2 = parentKey.OpenSubKey(subKeyName, RegistryRights.ChangePermissions);
				if (registryKey2 != null)
				{
					RegistrySecurity accessControl2 = registryKey2.GetAccessControl();
					SecurityIdentifier user = WindowsIdentity.GetCurrent().User;
					AuthorizationRuleCollection accessRules = accessControl2.GetAccessRules(includeExplicit: true, includeInherited: true, typeof(NTAccount));
					foreach (RegistryAccessRule item in accessRules)
					{
						if (item.AccessControlType == AccessControlType.Deny)
						{
							accessControl2.RemoveAccessRule(item);
						}
					}
					RegistryAccessRule accessRule = new RegistryAccessRule(user, RegistryRights.FullControl, AccessControlType.Allow);
					accessControl2.SetAccessRule(accessRule);
					registryKey2.SetAccessControl(accessControl2);
				}
			}
			catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs at ForceDeleteRegistryKey (ChangePermissions)", ex_captured); }

			try
			{
				using RegistryKey registryKey3 = parentKey.OpenSubKey(subKeyName, writable: true);
				if (registryKey3 != null)
				{
					string[] subKeyNames = registryKey3.GetSubKeyNames();
					foreach (string subKeyName2 in subKeyNames)
					{
						ForceDeleteRegistryKey(registryKey3, subKeyName2);
					}
				}
			}
			catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs at ForceDeleteRegistryKey (Recurse)", ex_captured); }

			try
			{
				parentKey.DeleteSubKeyTree(subKeyName, throwOnMissingSubKey: false);
				Logger.Info("已通过夺权强制清理受保护注册表项: " + subKeyName);
			}
			catch (Exception ex)
			{
				Logger.Warn("夺权强制清理受保护注册表项 " + subKeyName + " 失败: " + ex.Message);
			}
		}

		public static void KillProcessesByName(string[] processNames, Func<Process, bool> filter = null)
		{
			foreach (string processName in processNames)
			{
				try
				{
					Process[] processesByName = Process.GetProcessesByName(processName);
					foreach (Process process in processesByName)
					{
						if (filter == null || filter(process))
						{
							process.Kill();
							process.WaitForExit(2000);
							Logger.Info("已终止进程: " + process.ProcessName);
						}
					}
				}
				catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs at KillProcessesByName", ex_captured); }
			}
		}

		public static string GetInstalledVersionFromUninstallKeys(string[] keywords, Func<string, bool> nameFilter = null)
		{
			string[] array2 = new string[2] { "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall", "SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall" };
			RegistryKey[] array3 = new RegistryKey[2] { Registry.LocalMachine, Registry.CurrentUser };
			foreach (RegistryKey registryKey in array3)
			{
				foreach (string name in array2)
				{
					try
					{
						using RegistryKey registryKey2 = registryKey.OpenSubKey(name);
						if (registryKey2 == null) continue;

						string[] subKeyNames = registryKey2.GetSubKeyNames();
						foreach (string name2 in subKeyNames)
						{
							try
							{
								using RegistryKey registryKey3 = registryKey2.OpenSubKey(name2);
								if (registryKey3 == null) continue;

								string text = registryKey3.GetValue("DisplayName") as string;
								if (string.IsNullOrEmpty(text)) continue;

								if (nameFilter != null && !nameFilter(text)) continue;

								foreach (string text2 in keywords)
								{
									if (text.ToLower().Contains(text2.ToLower()))
									{
										string text3 = registryKey3.GetValue("DisplayVersion") as string;
										return string.IsNullOrEmpty(text3) ? text : (text + " (" + text3 + ")");
									}
								}
							}
							catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs at GetInstalledVersionFromUninstallKeys", ex_captured); }
						}
					}
					catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs at GetInstalledVersionFromUninstallKeys", ex_captured); }
				}
			}
			return null;
		}

		public static void CleanUninstallEntriesByFilter(Func<string, string, string, bool> filter, bool backupRestoreFonts = false)
		{
			List<FontBackupInfo> list = null;
			if (backupRestoreFonts)
			{
				list = BackupFonts();
			}
			string[] array = new string[2] { "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall", "SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall" };
			RegistryKey[] array2 = new RegistryKey[2] { Registry.LocalMachine, Registry.CurrentUser };
			foreach (RegistryKey registryKey in array2)
			{
				foreach (string name in array)
				{
					try
					{
						using RegistryKey registryKey2 = registryKey.OpenSubKey(name, writable: true);
						if (registryKey2 == null) continue;

						string[] subKeyNames = registryKey2.GetSubKeyNames();
						foreach (string text in subKeyNames)
						{
							try
							{
								using RegistryKey registryKey3 = registryKey2.OpenSubKey(text);
								if (registryKey3 == null) continue;

								string displayName = (registryKey3.GetValue("DisplayName") as string) ?? "";
								string publisher = (registryKey3.GetValue("Publisher") as string) ?? "";
								string uninstallString = (registryKey3.GetValue("UninstallString") as string) ?? "";
								if (filter(text, displayName, publisher))
								{
									Logger.Info("发现卸载项: " + displayName + "，准备静默调用卸载器...");
									if (!string.IsNullOrEmpty(uninstallString))
									{
										RunUninstaller(uninstallString);
									}
									registryKey2.DeleteSubKeyTree(text, throwOnMissingSubKey: false);
									Logger.Info("已清理注册表卸载项: " + displayName);
								}
							}
							catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs at CleanUninstallEntriesByFilter", ex_captured); }
						}
					}
					catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs at CleanUninstallEntriesByFilter", ex_captured); }
				}
			}
			if (list != null && list.Count > 0)
			{
				Logger.Info("等待卸载程序释放文件锁，准备恢复字体...");
				Thread.Sleep(3000);
				RestoreFonts(list);
			}
		}

		public static void CleanFoldersAndRegistryKeys(string[] processToWaitFor, string[] folderPaths, string[] registryKeys)
		{
			if (processToWaitFor != null && processToWaitFor.Length > 0)
			{
				WaitForProcessesToExit(processToWaitFor);
			}
			foreach (string item in folderPaths)
			{
				for (int i = 0; i < 3; i++)
				{
					try
					{
						if (Directory.Exists(item))
						{
							Directory.Delete(item, recursive: true);
							Logger.Info("已删除残留目录: " + item);
						}
					}
					catch (Exception ex)
					{
						if (i == 2)
						{
							Logger.Warn("删除残留目录失败 (三次尝试均受阻): " + item + ", 错误: " + ex.Message);
							continue;
						}
						Logger.Info($"目录 {item} 锁定中，等待 2 秒后进行第 {i + 2} 次重试...");
						Thread.Sleep(2000);
						continue;
					}
					break;
				}
			}
			foreach (string item2 in registryKeys)
			{
				DeleteKey(item2);
			}
		}

		public static void CleanShortcutsByFilter(string[] nameKeywords, string[] targetKeywords, string[] urlKeywords)
		{
			List<string> list = new List<string>();
			try
			{
				list.Add(Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
				list.Add(Environment.GetFolderPath(Environment.SpecialFolder.CommonDesktopDirectory));
				list.Add(Environment.GetFolderPath(Environment.SpecialFolder.Programs));
				list.Add(Environment.GetFolderPath(Environment.SpecialFolder.CommonPrograms));
				list.Add(Environment.GetFolderPath(Environment.SpecialFolder.SendTo));
				list.Add(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft\\Internet Explorer\\Quick Launch\\User Pinned\\TaskBar"));
				list.Add(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft\\Internet Explorer\\Quick Launch"));
				string directoryName = Path.GetDirectoryName(Environment.GetFolderPath(Environment.SpecialFolder.Programs));
				if (!string.IsNullOrEmpty(directoryName))
				{
					list.Add(directoryName);
				}
				string directoryName2 = Path.GetDirectoryName(Environment.GetFolderPath(Environment.SpecialFolder.CommonPrograms));
				if (!string.IsNullOrEmpty(directoryName2))
				{
					list.Add(directoryName2);
				}
			}
			catch (Exception ex)
			{
				Logger.Warn("获取快捷方式目录失败: " + ex.Message);
			}
			foreach (string item in list)
			{
				try
				{
					if (!Directory.Exists(item)) continue;

					string[] files = Directory.GetFiles(item, "*.*", SearchOption.AllDirectories);
					foreach (string text in files)
					{
						string text2 = Path.GetExtension(text).ToLower();
						if (text2 != ".lnk" && text2 != ".url") continue;

						bool flag = false;
						string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(text);
						foreach (string value in nameKeywords)
						{
							if (fileNameWithoutExtension.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0)
							{
								flag = true;
								break;
							}
						}
						if (!flag && text2 == ".lnk" && targetKeywords != null)
						{
							string shortcutTarget = GetShortcutTarget(text);
							if (!string.IsNullOrEmpty(shortcutTarget))
							{
								foreach (string value2 in targetKeywords)
								{
									if (shortcutTarget.IndexOf(value2, StringComparison.OrdinalIgnoreCase) >= 0)
									{
										flag = true;
										break;
									}
								}
							}
						}
						if (!flag && text2 == ".url" && urlKeywords != null)
						{
							try
							{
								string text3 = File.ReadAllText(text);
								foreach (string value3 in urlKeywords)
								{
									if (text3.IndexOf(value3, StringComparison.OrdinalIgnoreCase) >= 0)
									{
										flag = true;
										break;
									}
								}
							}
							catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs at CleanShortcutsByFilter (read url)", ex_captured); }
						}
						if (flag)
						{
							try
							{
								File.Delete(text);
								Logger.Info("已删除快捷方式残留: " + text);
							}
							catch (Exception ex2)
							{
								Logger.Warn("删除快捷方式失败: " + text + ", 错误: " + ex2.Message);
							}
						}
					}
					if (!item.Contains("Start Menu") && !item.Contains("Programs")) continue;

					string[] directories = Directory.GetDirectories(item, "*", SearchOption.AllDirectories);
					Array.Sort(directories, (string a, string b) => b.Length.CompareTo(a.Length));
					foreach (string text4 in directories)
					{
						if (!Directory.Exists(text4)) continue;

						string fileName = Path.GetFileName(text4);
						bool flag2 = false;
						foreach (string value4 in nameKeywords)
						{
							if (fileName.IndexOf(value4, StringComparison.OrdinalIgnoreCase) >= 0)
							{
								flag2 = true;
								break;
							}
						}
						if (!flag2)
						{
							try
							{
								string[] files2 = Directory.GetFiles(text4, "*.*", SearchOption.AllDirectories);
								if (files2.Length == 0)
								{
									flag2 = true;
								}
							}
							catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs at CleanShortcutsByFilter (check empty dir)", ex_captured); }
						}
						if (flag2)
						{
							try
							{
								Directory.Delete(text4, recursive: true);
								Logger.Info("已删除开始菜单目录或空残留文件夹: " + text4);
							}
							catch (Exception ex3)
							{
								Logger.Warn("删除开始菜单残留文件夹失败: " + text4 + ", 错误: " + ex3.Message);
							}
						}
					}
				}
				catch (Exception ex4)
				{
					Logger.Warn("扫描/清理快捷方式与开始菜单残留目录失败: " + item + ", 错误: " + ex4.Message);
				}
			}
		}

		public static void CleanFileAssociationsByFilter(ProductType product, string[] progIdPrefixes, string[] appExecutables)
		{
			if (appExecutables != null)
			{
				foreach (string text in appExecutables)
				{
					try
					{
						Registry.CurrentUser.DeleteSubKeyTree("Software\\Classes\\Applications\\" + text, throwOnMissingSubKey: false);
						Registry.LocalMachine.DeleteSubKeyTree("SOFTWARE\\Classes\\Applications\\" + text, throwOnMissingSubKey: false);
						Registry.CurrentUser.DeleteSubKeyTree("Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\" + text, throwOnMissingSubKey: false);
						Registry.LocalMachine.DeleteSubKeyTree("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\" + text, throwOnMissingSubKey: false);
						Logger.Info("已清除注册的 Applications\\" + text + " 和 App Paths\\" + text + " 关联节点");
					}
					catch (Exception ex)
					{
						Logger.Warn("清除 Applications/App Paths 节点 " + text + " 失败: " + ex.Message);
					}
				}
			}
			RegistryKey[] array3 = new RegistryKey[2] { Registry.LocalMachine, Registry.CurrentUser };
			foreach (RegistryKey registryKey in array3)
			{
				try
				{
					using RegistryKey registryKey2 = registryKey.OpenSubKey("SOFTWARE\\Classes", writable: true);
					if (registryKey2 == null) continue;

					string[] subKeyNames = registryKey2.GetSubKeyNames();
					foreach (string text2 in subKeyNames)
					{
						bool flag = false;
						foreach (string value2 in progIdPrefixes)
						{
							if (text2.StartsWith(value2, StringComparison.OrdinalIgnoreCase))
							{
								flag = true;
								break;
							}
						}
						if (flag)
						{
							try
							{
								ForceDeleteRegistryKey(registryKey2, text2);
								Logger.Info("已强力删除 ProgID 注册表关联项 (" + registryKey2.Name + "\\" + text2 + ")");
							}
							catch (Exception ex3)
							{
								Logger.Warn("清理 Classes 关联失败 (" + registryKey2.Name + "): " + ex3.Message);
							}
						}
					}
				}
				catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs at CleanFileAssociationsByFilter (Classes)", ex_captured); }
			}
			try
			{
				string[] array11 = GetProductKeywords(product, includeKso: false);
				string[] array12 = array11;
				using RegistryKey registryKey6 = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts", writable: true);
				if (registryKey6 != null)
				{
					string[] subKeyNames4 = registryKey6.GetSubKeyNames();
					foreach (string text8 in subKeyNames4)
					{
						try
						{
							string[] array13 = new string[2] { "UserChoice", "UserChoiceLatest" };
							foreach (string text9 in array13)
							{
								bool flag5 = false;
								using (RegistryKey registryKey7 = registryKey6.OpenSubKey(text8 + "\\" + text9))
								{
									if (registryKey7 != null)
									{
										string text10 = registryKey7.GetValue("ProgId") as string;
										if (!string.IsNullOrEmpty(text10))
										{
											foreach (string value6 in progIdPrefixes)
											{
												if (text10.StartsWith(value6, StringComparison.OrdinalIgnoreCase))
												{
													flag5 = true;
													break;
												}
											}
										}
									}
								}
								if (flag5)
								{
									try
									{
										ForceDeleteRegistryKey(registryKey6, text8 + "\\" + text9);
										Logger.Info("已强制删除残留的 UserChoice 指针: " + text8 + "\\" + text9);
									}
									catch (Exception ex4)
									{
										Logger.Warn("强制清理 UserChoice 残留指针失败: " + text8 + "\\" + text9 + ", 错误: " + ex4.Message);
									}
								}
							}
						}
						catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs at CleanFileAssociationsByFilter (FileExts)", ex_captured); }
						try
						{
							using RegistryKey registryKey8 = registryKey6.OpenSubKey(text8 + "\\OpenWithList", writable: true);
							if (registryKey8 != null)
							{
								string[] valueNames2 = registryKey8.GetValueNames();
								foreach (string text11 in valueNames2)
								{
									string text12 = registryKey8.GetValue(text11) as string;
									if (string.IsNullOrEmpty(text12)) continue;

									bool flag6 = false;
									foreach (string value7 in array12)
									{
										if (text12.ToLower().Contains(value7.ToLower()))
										{
											flag6 = true;
											break;
										}
									}
									if (flag6)
									{
										try
										{
											registryKey8.DeleteValue(text11, throwOnMissingValue: false);
											Logger.Info("已清理 " + text8 + "\\OpenWithList 中残留的 " + text12);
										}
										catch (Exception ex5)
										{
											Logger.Warn("清理 " + text8 + "\\OpenWithList 残留失败 (" + text11 + "): " + ex5.Message);
										}
									}
								}
							}
						}
						catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs at CleanFileAssociationsByFilter (OpenWithList)", ex_captured); }
						try
						{
							using RegistryKey registryKey9 = registryKey6.OpenSubKey(text8 + "\\OpenWithProgids", writable: true);
							if (registryKey9 != null)
							{
								string[] valueNames3 = registryKey9.GetValueNames();
								foreach (string text13 in valueNames3)
								{
									bool flag7 = false;
									foreach (string value8 in progIdPrefixes)
									{
										if (text13.StartsWith(value8, StringComparison.OrdinalIgnoreCase))
										{
											flag7 = true;
											break;
										}
									}
									if (flag7)
									{
										try
										{
											registryKey9.DeleteValue(text13, throwOnMissingValue: false);
											Logger.Info("已从 " + text8 + "\\OpenWithProgids 移除了 ProgID: " + text13);
										}
										catch (Exception ex6)
										{
											Logger.Warn("清理 " + text8 + "\\OpenWithProgids 关联失败 (" + text13 + "): " + ex6.Message);
										}
									}
								}
							}
						}
						catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs at CleanFileAssociationsByFilter (OpenWithProgids)", ex_captured); }
					}
				}
			}
			catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs at CleanFileAssociationsByFilter (Explorer FileExts)", ex_captured); }
			try
			{
				RefreshIconCache();
			}
			catch (Exception ex6)
			{
				Logger.Error("自动刷新图标缓存失败", ex6);
			}
			try
			{
				string[] array11 = GetProductKeywords(product, includeKso: true);
				string[] array24 = array11;
				using RegistryKey registryKey12 = Registry.CurrentUser.OpenSubKey("Software\\Classes\\Local Settings\\Software\\Microsoft\\Windows\\Shell\\MuiCache", writable: true);
				if (registryKey12 != null)
				{
					string[] valueNames5 = registryKey12.GetValueNames();
					foreach (string text21 in valueNames5)
					{
						bool flag10 = false;
						foreach (string value12 in array24)
						{
							if (text21.ToLower().Contains(value12.ToLower()))
							{
								flag10 = true;
								break;
							}
						}
						if (flag10)
						{
							try
							{
								registryKey12.DeleteValue(text21, throwOnMissingValue: false);
								Logger.Info("已从 MuiCache 历史痕迹中强力擦除友好名称: " + text21);
							}
							catch (Exception ex7)
							{
								Logger.Warn("从 MuiCache 中清理友好名称 " + text21 + " 失败: " + ex7.Message);
							}
						}
					}
				}
			}
			catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs at CleanFileAssociationsByFilter (MuiCache)", ex_captured); }
		}

		private static void WaitForProcessesToExit(string[] processNames, int timeoutMs = 15000)
		{
			Stopwatch stopwatch = Stopwatch.StartNew();
			while (stopwatch.ElapsedMilliseconds < timeoutMs)
			{
				bool flag = false;
				foreach (string processName in processNames)
				{
					try
					{
						if (Process.GetProcessesByName(processName).Length != 0)
						{
							flag = true;
							break;
						}
					}
					catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs at WaitForProcessesToExit", ex_captured); }
				}
				if (!flag)
				{
					break;
				}
				Thread.Sleep(500);
			}
		}

		private static void RunUninstaller(string uninstallString)
		{
			try
			{
				string text = uninstallString.Trim();
				string text2 = "";
				if (text.ToLower().Contains("msiexec"))
				{
					text = "msiexec.exe";
					string[] source = uninstallString.Split(' ');
					string text3 = source.FirstOrDefault(p => p.Contains("{") && p.Contains("}"));
					if (!string.IsNullOrEmpty(text3))
					{
						text2 = "/x " + text3 + " /qn /norestart";
					}
					else
					{
						text2 = uninstallString.Replace("/I", "/X").Replace("/i", "/x") + " /qn /norestart";
						text2 = text2.Replace("msiexec.exe", "").Replace("msiexec", "").Trim();
					}
				}
				else
				{
					if (text.StartsWith("\""))
					{
						int num = text.IndexOf("\"", 1);
						if (num > 0)
						{
							text2 = text.Substring(num + 1).Trim();
							text = text.Substring(1, num - 1);
						}
					}
					else
					{
						int num2 = text.IndexOf(' ');
						if (num2 > 0)
						{
							text2 = text.Substring(num2 + 1);
							text = text.Substring(0, num2);
						}
					}
					if (!text2.Contains("/S") && !text2.Contains("/s") && !text2.Contains("/silent") && !text2.Contains("/verysilent"))
					{
						text2 = ((!uninstallString.ToLower().Contains("wps") && !uninstallString.ToLower().Contains("yozo")) ? (text2 + " /VERYSILENT /NORESTART") : (text2 + " /S"));
					}
				}
				Logger.Info("执行静默卸载命令: " + text + " " + text2);
				ProcessStartInfo startInfo = new ProcessStartInfo(text, text2)
				{
					CreateNoWindow = true,
					UseShellExecute = false
				};
				using (Process process = Process.Start(startInfo))
				{
					process?.WaitForExit(30000);
				}
			}
			catch (Exception ex)
			{
				Logger.Warn("调用卸载命令失败: " + uninstallString + ", 错误: " + ex.Message);
			}
		}

		private static string GetShortcutTarget(string lnkPath)
		{
			try
			{
				Type typeFromProgID = Type.GetTypeFromProgID("WScript.Shell");
				if (typeFromProgID == null) return "";

				object target = Activator.CreateInstance(typeFromProgID);
				object obj = typeFromProgID.InvokeMember("CreateShortcut", BindingFlags.InvokeMethod, null, target, new object[1] { lnkPath });
				if (obj == null) return "";

				string text = obj.GetType().InvokeMember("TargetPath", BindingFlags.GetProperty, null, obj, null) as string;
				return text ?? "";
			}
			catch (Exception ex)
			{
				GOI.Helpers.Logger.Error("Failed to resolve shortcut target for: " + lnkPath, ex);
				return "";
			}
		}

		private static string[] GetProductKeywords(ProductType product, bool includeKso = false)
		{
			return product switch
			{
				ProductType.Wps => includeKso ? new string[5] { "wps", "et", "wpp", "金山", "kso" } : new string[4] { "wps", "et", "wpp", "金山" },
				ProductType.Yozo => new string[2] { "yozo", "永中" },
				ProductType.OnlyOffice => new string[1] { "onlyoffice" },
				ProductType.LibreOffice => new string[2] { "libreoffice", "soffice" },
				_ => new string[0]
			};
		}

		private static List<FontBackupInfo> BackupFonts()
		{
			List<FontBackupInfo> list = new List<FontBackupInfo>();
			string text = Path.Combine(Path.GetTempPath(), "GOIFontBackup");
			try
			{
				if (!Directory.Exists(text))
				{
					Directory.CreateDirectory(text);
				}
				string folderPath = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
				(RegistryKey, bool)[] array = new(RegistryKey, bool)[2]
				{
					(Registry.LocalMachine, false),
					(Registry.CurrentUser, true)
				};
				foreach (var item in array)
				{
					RegistryKey registryKey = item.Item1;
					bool isUserFont = item.Item2;
					using RegistryKey registryKey2 = registryKey.OpenSubKey("SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Fonts");
					if (registryKey2 == null) continue;

					string[] valueNames = registryKey2.GetValueNames();
					foreach (string text2 in valueNames)
					{
						try
						{
							bool flag = text2.Contains("方正") || text2.IndexOf("fz", StringComparison.OrdinalIgnoreCase) >= 0;
							string text3 = registryKey2.GetValue(text2) as string;
							if (string.IsNullOrEmpty(text3)) continue;

							if (!flag && (text3.StartsWith("fz", StringComparison.OrdinalIgnoreCase) || text3.StartsWith("FZ", StringComparison.OrdinalIgnoreCase)))
							{
								flag = true;
							}
							if (flag)
							{
								string text4 = (Path.IsPathRooted(text3) ? text3 : Path.Combine(folderPath, text3));
								if (File.Exists(text4))
								{
									string text5 = Path.Combine(text, Path.GetFileName(text4));
									File.Copy(text4, text5, overwrite: true);
									list.Add(new FontBackupInfo
									{
										Name = text2,
										FileName = text3,
										BackupPath = text5,
										IsUserFont = isUserFont
									});
									Logger.Info("已备份字体: " + text2 + " -> " + text5);
								}
							}
						}
						catch (Exception ex)
						{
							Logger.Warn("备份单个字体失败: " + text2 + ", 错误: " + ex.Message);
						}
					}
				}
			}
			catch (Exception ex2)
			{
				Logger.Warn("备份字体流程遇到异常: " + ex2.Message);
			}
			return list;
		}

		private static void RestoreFonts(List<FontBackupInfo> backupList)
		{
			if (backupList == null || backupList.Count == 0) return;

			string folderPath = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
			bool flag = false;
			foreach (FontBackupInfo backup in backupList)
			{
				try
				{
					if (!File.Exists(backup.BackupPath)) continue;

					string text = (Path.IsPathRooted(backup.FileName) ? backup.FileName : Path.Combine(folderPath, backup.FileName));
					if (!File.Exists(text))
					{
						string directoryName = Path.GetDirectoryName(text);
						if (!string.IsNullOrEmpty(directoryName) && !Directory.Exists(directoryName))
						{
							Directory.CreateDirectory(directoryName);
						}
						File.Copy(backup.BackupPath, text, overwrite: true);
						flag = true;
						Logger.Info("已恢复被删除的字体文件: " + backup.Name + " -> " + text);
					}
					try
					{
						AddFontResource(text);
					}
					catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs at RestoreFonts", ex_captured); }

					RegistryKey registryKey = (backup.IsUserFont ? Registry.CurrentUser : Registry.LocalMachine);
					using RegistryKey registryKey2 = registryKey.OpenSubKey("SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Fonts", writable: true);
					if (registryKey2 != null)
					{
						string text2 = registryKey2.GetValue(backup.Name) as string;
						if (text2 == null)
						{
							registryKey2.SetValue(backup.Name, backup.FileName, RegistryValueKind.String);
							flag = true;
							Logger.Info("已恢复字体注册表项: " + backup.Name);
						}
					}
				}
				catch (Exception ex)
				{
					Logger.Warn("恢复字体失败: " + backup.Name + ", " + ex.Message);
				}
			}
			if (flag)
			{
				try
				{
					SendMessage(HWND_BROADCAST, WM_FONTCHANGE, IntPtr.Zero, IntPtr.Zero);
					Logger.Info("已广播 WM_FONTCHANGE 字体变更消息通知系统");
				}
				catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs at RestoreFonts", ex_captured); }
			}
			try
			{
				string path = Path.Combine(Path.GetTempPath(), "GOIFontBackup");
				if (Directory.Exists(path))
				{
					Directory.Delete(path, recursive: true);
				}
			}
			catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs at RestoreFonts", ex_captured); }
		}

		private static string GetAppPathFromRegistry(string exeName)
		{
			try
			{
				using (RegistryKey registryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\" + exeName))
				{
					if (registryKey != null)
					{
						string text = registryKey.GetValue("") as string;
						if (!string.IsNullOrEmpty(text))
						{
							return text.Trim('"');
						}
					}
				}
				using (RegistryKey registryKey2 = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\" + exeName))
				{
					if (registryKey2 != null)
					{
						string text2 = registryKey2.GetValue("") as string;
						if (!string.IsNullOrEmpty(text2))
						{
							return text2.Trim('"');
						}
					}
				}
			}
			catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs at GetAppPathFromRegistry", ex_captured); }
			return "";
		}

		public static bool IsOfficeProgId(string progId)
		{
			if (string.IsNullOrEmpty(progId)) return false;
			string[] prefixes = { "WPS.", "WPP.", "ET.", "Word.", "Excel.", "PowerPoint.", "Yozo", "ONLYOFFICE", "LibreOffice", "soffice" };
			foreach (var prefix in prefixes)
			{
				if (progId.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
					return true;
			}
			return false;
		}

		private static bool ShouldRestoreExtensionAssociation(string ext)
		{
			using (RegistryKey rkExt = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts\\" + ext + "\\UserChoice"))
			{
				if (rkExt != null)
				{
					string userChoiceProgId = rkExt.GetValue("ProgId") as string;
					if (!string.IsNullOrEmpty(userChoiceProgId))
					{
						if (!IsOfficeProgId(userChoiceProgId)) return false;
					}
				}
			}

			string currentProgId = null;
			using (RegistryKey rk = Registry.CurrentUser.OpenSubKey("Software\\Classes\\" + ext))
			{
				currentProgId = rk?.GetValue("") as string;
			}
			if (string.IsNullOrEmpty(currentProgId))
			{
				using (RegistryKey rk = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Classes\\" + ext))
				{
					currentProgId = rk?.GetValue("") as string;
				}
			}

			if (!string.IsNullOrEmpty(currentProgId))
			{
				if (!IsOfficeProgId(currentProgId)) return false;
			}

			return true;
		}

		public static void RestoreInstalledProductAssociations(IProgress<string> progress = null)
		{
			string installedProductVersion = GetInstalledProductVersion(ProductType.MsOffice);
			if (!string.IsNullOrEmpty(installedProductVersion))
			{
				var dictionary = new Dictionary<string, string[]>
				{
					{ "winword.exe", new string[5] { ".doc", ".docx", ".docm", ".dot", ".dotx" } },
					{ "excel.exe", new string[4] { ".xls", ".xlsx", ".xlsm", ".xlsb" } },
					{ "powerpnt.exe", new string[5] { ".ppt", ".pptx", ".pptm", ".pps", ".ppsx" } }
				};
				foreach (KeyValuePair<string, string[]> item in dictionary)
				{
					string key = item.Key;
					string[] value = item.Value;
					string text = GetAppPathFromRegistry(key);
					if (string.IsNullOrEmpty(text) || !File.Exists(text))
					{
						string[] array = new string[6]
						{
							"C:\\Program Files\\Microsoft Office\\root\\Office16\\" + key,
							"C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\" + key,
							"C:\\Program Files\\Microsoft Office\\Office16\\" + key,
							"C:\\Program Files (x86)\\Microsoft Office\\Office16\\" + key,
							"C:\\Program Files\\Microsoft Office\\Office15\\" + key,
							"C:\\Program Files (x86)\\Microsoft Office\\Office15\\" + key
						};
						foreach (string text2 in array)
						{
							if (File.Exists(text2))
							{
								text = text2;
								break;
							}
						}
					}
					if (string.IsNullOrEmpty(text) || !File.Exists(text)) continue;

					string text3 = key switch
					{
						"winword.exe" => "Word",
						"excel.exe" => "Excel",
						"powerpnt.exe" => "PowerPoint",
						_ => key,
					};
					progress?.Report(LocalizationStrings.Instance.StatusDetectMsOffice(text3));
					foreach (string text5 in value)
					{
						if (!MsOfficeProgIds.TryGetValue(text5, out var value2)) continue;

						if (!ShouldRestoreExtensionAssociation(text5))
						{
							Logger.Info("用户已为 " + text5 + " 配置了自定义打开程序，跳过强制修复以防覆盖用户偏好。");
							continue;
						}
						try
						{
							using (RegistryKey registryKey = Registry.CurrentUser.CreateSubKey("Software\\Classes\\" + text5))
							{
								registryKey?.SetValue("", value2);
							}
							using (RegistryKey registryKey2 = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Classes\\" + text5))
							{
								registryKey2?.SetValue("", value2);
							}
							try
							{
								using RegistryKey registryKey3 = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts\\" + text5, writable: true);
								if (registryKey3 != null)
								{
									ForceDeleteRegistryKey(registryKey3, "UserChoice");
									ForceDeleteRegistryKey(registryKey3, "UserChoiceLatest");
									Logger.Info("已强力清除 " + text5 + " 的 UserChoice 与 UserChoiceLatest 以激活 MS Office " + text3 + " 默认关联");
								}
							}
							catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs at RestoreInstalledProductAssociations", ex_captured); }
							using (RegistryKey registryKey4 = Registry.CurrentUser.CreateSubKey("Software\\Classes\\" + text5 + "\\OpenWithProgids"))
							{
								if (registryKey4 != null && registryKey4.GetValue(value2) == null)
								{
									registryKey4.SetValue(value2, new byte[0], RegistryValueKind.Binary);
								}
							}
							using (RegistryKey registryKey5 = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Classes\\" + text5 + "\\OpenWithProgids"))
							{
								if (registryKey5 != null && registryKey5.GetValue(value2) == null)
								{
									registryKey5.SetValue(value2, new byte[0], RegistryValueKind.Binary);
								}
							}
						}
						catch (Exception ex)
						{
							Logger.Warn("修复 " + text5 + " 的 MS Office 关联 " + value2 + " 失败: " + ex.Message);
						}
					}
					foreach (string key2 in value)
					{
						if (!MsOfficeProgIds.TryGetValue(key2, out var value3)) continue;

						try
						{
							Registry.CurrentUser.DeleteSubKeyTree("Software\\Classes\\" + value3, throwOnMissingSubKey: false);
							Logger.Info("已删除 HKCU 层的 ProgID 劫持键: " + value3);
						}
						catch (Exception ex2)
						{
							Logger.Warn("清理 HKCU 层的 ProgID 劫持键 " + value3 + " 失败: " + ex2.Message);
						}
						if (MsOfficeProgIdIcons.TryGetValue(value3, out (string, int) value4))
						{
							string text6 = Path.Combine(Path.GetDirectoryName(text), value4.Item1);
							if (File.Exists(text6))
							{
								string text7 = $"\"{text6}\",{value4.Item2}";
								try
								{
									using (RegistryKey registryKey6 = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Classes\\" + value3 + "\\DefaultIcon"))
									{
										registryKey6?.SetValue("", text7);
									}
									Logger.Info("已修复 HKLM 层 " + value3 + " 的 DefaultIcon 为: " + text7);
								}
								catch (Exception ex3)
								{
									Logger.Warn("写入 " + value3 + " DefaultIcon 失败: " + ex3.Message);
								}
							}
						}

						string text8 = key == "winword.exe" ? "\"" + text + "\" /n \"%1\" /o \"%u\"" : "\"" + text + "\" \"%1\"";
						try
						{
							using (RegistryKey registryKey7 = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Classes\\" + value3 + "\\shell\\Open\\command"))
							{
								registryKey7?.SetValue("", text8);
							}
							Logger.Info("已修复 HKLM 层 " + value3 + " 的 Open Command 为: " + text8);
						}
						catch (Exception ex4)
						{
							Logger.Warn("写入 " + value3 + " Open Command 失败: " + ex4.Message);
						}
						try
						{
							using (RegistryKey registryKey8 = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Classes\\" + value3 + "\\protocol\\StdFileEditing\\server"))
							{
								registryKey8?.SetValue("", text);
							}
							Logger.Info("已修复 HKLM 层 " + value3 + " 的 StdFileEditing\\server 为: " + text);
						}
						catch (Exception ex5)
						{
							Logger.Warn("写入 " + value3 + " StdFileEditing\\server 失败: " + ex5.Message);
						}
					}
				}
			}
			string installedProductVersion2 = GetInstalledProductVersion(ProductType.Wps);
			if (string.IsNullOrEmpty(installedProductVersion2)) return;

			var dictionaryWps = new Dictionary<string, string[]>
			{
				{ "wps.exe", new string[2] { ".doc", ".docx" } },
				{ "et.exe", new string[2] { ".xls", ".xlsx" } },
				{ "wpp.exe", new string[2] { ".ppt", ".pptx" } }
			};
			foreach (KeyValuePair<string, string[]> item2 in dictionaryWps)
			{
				string key3 = item2.Key;
				string[] value5 = item2.Value;
				string appPathFromRegistry = GetAppPathFromRegistry(key3);
				if (string.IsNullOrEmpty(appPathFromRegistry) || !File.Exists(appPathFromRegistry)) continue;

				string text3 = key3 switch
				{
					"wps.exe" => "WPS 文字",
					"et.exe" => "WPS 表格",
					"wpp.exe" => "WPS 演示",
					_ => key3,
				};
				progress?.Report(LocalizationStrings.Instance.StatusDetectWps(text3));
				foreach (string text10 in value5)
				{
					if (!WpsProgIds.TryGetValue(text10, out var value6)) continue;

					try
					{
						using (RegistryKey registryKey9 = Registry.CurrentUser.CreateSubKey("Software\\Classes\\" + text10))
						{
							registryKey9?.SetValue("", value6);
						}
						using (RegistryKey registryKey10 = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Classes\\" + text10))
						{
							registryKey10?.SetValue("", value6);
						}
						try
						{
							using RegistryKey registryKey11 = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts\\" + text10, writable: true);
							if (registryKey11 != null)
							{
								ForceDeleteRegistryKey(registryKey11, "UserChoice");
								ForceDeleteRegistryKey(registryKey11, "UserChoiceLatest");
								Logger.Info("已强力清除 " + text10 + " 的 UserChoice 与 UserChoiceLatest 以激活 WPS " + text3 + " 默认关联");
							}
						}
						catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs at RestoreInstalledProductAssociations", ex_captured); }
						using (RegistryKey registryKey12 = Registry.CurrentUser.CreateSubKey("Software\\Classes\\" + text10 + "\\OpenWithProgids"))
						{
							if (registryKey12 != null && registryKey12.GetValue(value6) == null)
							{
								registryKey12.SetValue(value6, new byte[0], RegistryValueKind.Binary);
							}
						}
						using (RegistryKey registryKey13 = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Classes\\" + text10 + "\\OpenWithProgids"))
						{
							if (registryKey13 != null && registryKey13.GetValue(value6) == null)
							{
								registryKey13.SetValue(value6, new byte[0], RegistryValueKind.Binary);
							}
						}
					}
					catch (Exception ex6)
					{
						Logger.Warn("修复 " + text10 + " 的 WPS 关联 " + value6 + " 失败: " + ex6.Message);
					}
				}
			}
		}

		public static int CleanOrphanedFileAssociations(IProgress<string> progress = null)
		{
			int num = 0;
			ProductType[] array = new ProductType[4]
			{
				ProductType.Wps,
				ProductType.Yozo,
				ProductType.OnlyOffice,
				ProductType.LibreOffice
			};
			foreach (ProductType productType in array)
			{
				string installedProductVersion = GetInstalledProductVersion(productType);
				if (string.IsNullOrEmpty(installedProductVersion))
				{
					string text = productType switch
					{
						ProductType.Wps => "WPS Office",
						ProductType.Yozo => LocalizationStrings.Instance.YozoTitle,
						ProductType.OnlyOffice => "OnlyOffice",
						ProductType.LibreOffice => "LibreOffice",
						_ => "",
					};
					progress?.Report(LocalizationStrings.Instance.StatusPurgingProduct(text));
					try
					{
						CleanFileAssociations(productType);
						CleanShortcuts(productType);
						num++;
					}
					catch (Exception ex)
					{
						Logger.Error("清除 " + text + " 残留文件关联失败", ex);
					}
				}
			}
			try
			{
				RestoreInstalledProductAssociations(progress);
			}
			catch (Exception ex2)
			{
				Logger.Error("自动修复当前已安装 Office 关联失败", ex2);
			}
			try
			{
				RepairOfficeComComponents(progress);
			}
			catch (Exception ex3)
			{
				Logger.Error("自动修复已安装 Office COM 组件注册失败", ex3);
			}
			try
			{
				progress?.Report(LocalizationStrings.Instance.StatusRebuildingIconCache);
				RefreshIconCache();
			}
			catch (Exception ex4)
			{
				Logger.Error("自动刷新图标缓存失败", ex4);
			}
			return num;
		}

		public static void RepairOfficeComComponents(IProgress<string> progress = null)
		{
			List<string> list = new List<string>();
			string installedProductVersion = GetInstalledProductVersion(ProductType.MsOffice);
			if (!string.IsNullOrEmpty(installedProductVersion))
			{
				list.AddRange(new string[3] { "winword.exe", "excel.exe", "powerpnt.exe" });
			}
			string installedProductVersion2 = GetInstalledProductVersion(ProductType.Wps);
			if (!string.IsNullOrEmpty(installedProductVersion2))
			{
				list.AddRange(new string[3] { "wps.exe", "et.exe", "wpp.exe" });
			}
			foreach (string item in list)
			{
				try
				{
					string text = null;
					using (RegistryKey registryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\" + item))
					{
						if (registryKey != null)
						{
							text = registryKey.GetValue("") as string;
						}
					}
					if (string.IsNullOrEmpty(text))
					{
						using (RegistryKey registryKey2 = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\" + item))
						{
							if (registryKey2 != null)
							{
								text = registryKey2.GetValue("") as string;
							}
						}
					}
					if (string.IsNullOrEmpty(text))
					{
						string folderPath = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
						string folderPath2 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
						string[] array = new string[5]
						{
							Path.Combine(folderPath, "Microsoft Office\\root\\Office16\\" + item),
							Path.Combine(folderPath2, "Microsoft Office\\root\\Office16\\" + item),
							Path.Combine(folderPath, "Microsoft Office\\Office16\\" + item),
							Path.Combine(folderPath2, "Microsoft Office\\Office16\\" + item),
							Path.Combine(folderPath, "Microsoft Office\\Office15\\" + item)
						};
						foreach (string text2 in array)
						{
							if (File.Exists(text2))
							{
								text = text2;
								break;
							}
						}
					}
					if (string.IsNullOrEmpty(text)) continue;

					text = text.Trim('"');
					if (File.Exists(text))
					{
						progress?.Report(LocalizationStrings.Instance.StatusRestoringCOM(item));
						Logger.Info("执行 COM 组件注册: " + text + " /regserver");
						ProcessStartInfo startInfo = new ProcessStartInfo(text, "/regserver")
						{
							CreateNoWindow = true,
							UseShellExecute = false
						};
						using (Process process = Process.Start(startInfo))
						{
							process?.WaitForExit(10000);
						}
					}
				}
				catch (Exception ex)
				{
					Logger.Warn("修复 " + item + " 的 COM 组件注册失败: " + ex.Message);
				}
			}
		}

		public static void RemoveClickToRunService()
		{
			try
			{
				RunSc("stop ClickToRunSvc");
				RunSc("delete ClickToRunSvc");
			}
			catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs at RemoveClickToRunService", ex_captured); }
		}

		public static void RefreshIconCache()
		{
			try
			{
				SHChangeNotify(134217728, 4096, IntPtr.Zero, IntPtr.Zero);
				Logger.Info("已向 Windows Shell 发送 ASSOCCHANGED 关联变动通知");
				try
				{
					ProcessStartInfo startInfo = new ProcessStartInfo("ie4uinit.exe", "-show")
					{
						CreateNoWindow = true,
						UseShellExecute = false
					};
					using Process process = Process.Start(startInfo);
					process?.WaitForExit(3000);
				}
				catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs at RefreshIconCache", ex_captured); }
				try
				{
					ProcessStartInfo startInfo2 = new ProcessStartInfo("ie4uinit.exe", "-ClearIconCache")
					{
						CreateNoWindow = true,
						UseShellExecute = false
					};
					using Process process2 = Process.Start(startInfo2);
					process2?.WaitForExit(3000);
				}
				catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs at RefreshIconCache", ex_captured); }
			}
			catch (Exception ex)
			{
				Logger.Warn("刷新图标缓存失败: " + ex.Message);
			}
		}

		private static void RunSc(string args)
		{
			try
			{
				ProcessStartInfo startInfo = new ProcessStartInfo("sc.exe", args)
				{
					CreateNoWindow = true,
					UseShellExecute = false
				};
				using (Process process = Process.Start(startInfo))
				{
					process?.WaitForExit(5000);
				}
			}
			catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs at RunSc", ex_captured); }
		}

		/// <summary>
		/// 定位 OSPP.VBS 许可管理脚本：依次尝试 HKLM/HKCU App Paths、ClickToRun 安装路径、以及一组已知的静态安装路径。
		/// </summary>
		public static string LocateOsppVbs()
		{
			string[] officeExes = new string[3] { "winword.exe", "excel.exe", "powerpnt.exe" };
			string[] appPathsRoots = new string[2] { "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\", "SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\App Paths\\" };
			foreach (string exe in officeExes)
			{
				foreach (string root in appPathsRoots)
				{
					try
					{
						RegistryView[] views = new RegistryView[2] { RegistryView.Registry64, RegistryView.Registry32 };
						foreach (RegistryView view in views)
						{
							using RegistryKey baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view);
							using RegistryKey subKey = baseKey.OpenSubKey(root + exe);
							if (subKey == null)
							{
								continue;
							}
							string exePath = subKey.GetValue("") as string;
							if (!string.IsNullOrEmpty(exePath) && File.Exists(exePath))
							{
								string candidate = Path.Combine(Path.GetDirectoryName(exePath), "OSPP.VBS");
								if (File.Exists(candidate))
								{
									Logger.Info("通过 App Paths 注册表成功定位 OSPP.VBS: " + candidate);
									return candidate;
								}
							}
						}
					}
					catch (Exception ex)
					{
						Logger.Warn("读取 App Paths 注册表失败 (" + root + exe + "): " + ex.Message);
					}
				}
			}
			try
			{
				RegistryView[] views = new RegistryView[2] { RegistryView.Registry64, RegistryView.Registry32 };
				foreach (RegistryView view in views)
				{
					using RegistryKey baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view);
					using RegistryKey cfg = baseKey.OpenSubKey("SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration");
					if (cfg == null)
					{
						continue;
					}
					string installPath = cfg.GetValue("InstallPath") as string;
					if (!string.IsNullOrEmpty(installPath) && Directory.Exists(installPath))
					{
						string office16 = Path.Combine(installPath, "root\\Office16\\OSPP.VBS");
						if (File.Exists(office16))
						{
							return office16;
						}
						string office15 = Path.Combine(installPath, "root\\Office15\\OSPP.VBS");
						if (File.Exists(office15))
						{
							return office15;
						}
					}
				}
			}
			catch (Exception ex)
			{
				Logger.Warn("读取 ClickToRun 注册表路径失败: " + ex.Message);
			}
			foreach (string exe in officeExes)
			{
				try
				{
					using RegistryKey subKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\" + exe);
					if (subKey == null)
					{
						continue;
					}
					string exePath = subKey.GetValue("") as string;
					if (!string.IsNullOrEmpty(exePath) && File.Exists(exePath))
					{
						string candidate = Path.Combine(Path.GetDirectoryName(exePath), "OSPP.VBS");
						if (File.Exists(candidate))
						{
							Logger.Info("通过 HKCU App Paths 注册表成功定位 OSPP.VBS: " + candidate);
							return candidate;
						}
					}
				}
				catch (Exception ex)
				{
					Logger.Warn("从 HKCU App Paths 读取 OSPP 失败: " + ex.Message);
				}
			}
			string[] staticPaths = new string[10]
			{
				"C:\\Program Files\\Microsoft Office\\root\\Office16\\OSPP.VBS",
				"C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\OSPP.VBS",
				"C:\\Program Files\\Microsoft Office\\root\\Office15\\OSPP.VBS",
				"C:\\Program Files (x86)\\Microsoft Office\\root\\Office15\\OSPP.VBS",
				"C:\\Program Files\\Microsoft Office\\Office16\\OSPP.VBS",
				"C:\\Program Files (x86)\\Microsoft Office\\Office16\\OSPP.VBS",
				"C:\\Program Files\\Microsoft Office\\Office15\\OSPP.VBS",
				"C:\\Program Files (x86)\\Microsoft Office\\Office15\\OSPP.VBS",
				"C:\\Program Files\\Microsoft Office\\Office14\\OSPP.VBS",
				"C:\\Program Files (x86)\\Microsoft Office\\Office14\\OSPP.VBS"
			};
			foreach (string path in staticPaths)
			{
				if (File.Exists(path))
				{
					Logger.Info("找到 OSPP.VBS 静态路径: " + path);
					return path;
				}
			}
			return null;
		}

		[DllImport("shell32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		private static extern void SHChangeNotify(int wEventId, int uFlags, IntPtr dwItem1, IntPtr dwItem2);

		[DllImport("gdi32.dll", CharSet = CharSet.Unicode, EntryPoint = "AddFontResourceW")]
		private static extern int AddFontResource(string lpFileName);

		[DllImport("user32.dll", CharSet = CharSet.Auto)]
		private static extern int SendMessage(IntPtr hWnd, int wMsg, IntPtr wParam, IntPtr lParam);

		#endregion
	}
}
