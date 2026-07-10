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

		private static readonly Dictionary<ProductType, string[]> ProductExecutables = new Dictionary<ProductType, string[]>
		{
			{
				ProductType.Wps,
				new string[14]
				{
					"wps.exe", "et.exe", "wpp.exe", "wpspdf.exe", "wpsoffice.exe", "ksolaunch.exe", "kso.exe", "wpsupdate.exe", "wpsofd.exe", "photolaunch.exe",
					"wpsphoto.exe", "wpsphotos.exe", "ksophoto.exe", "ksophotos.exe"
				}
			},
			{
				ProductType.Yozo,
				new string[11]
				{
					"yozo.exe", "yozoword.exe", "yozosheet.exe", "yozopresent.exe", "yozobinder.exe", "yozolaunch.exe", "Yozo_Calc.exe", "Yozo_Impress.exe", "yozo_Ofd.exe", "Yozo_Office.exe",
					"Yozo_Writer.exe"
				}
			},
			{
				ProductType.OnlyOffice,
				new string[4] { "DesktopEditors.exe", "editors.exe", "editors_helper.exe", "updatesvc.exe" }
			},
			{
				ProductType.LibreOffice,
				new string[4] { "soffice.exe", "scalc.exe", "swriter.exe", "simpress.exe" }
			}
		};

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
			{
				"Word.Document.8",
				("wordicon.exe", 1)
			},
			{
				"Word.Document.12",
				("wordicon.exe", 13)
			},
			{
				"Word.DocumentMacroEnabled.12",
				("wordicon.exe", 15)
			},
			{
				"Word.Template.8",
				("wordicon.exe", 2)
			},
			{
				"Word.Template.12",
				("wordicon.exe", 14)
			},
			{
				"Word.RTF.8",
				("wordicon.exe", 1)
			},
			{
				"Excel.Sheet.8",
				("xlicons.exe", 1)
			},
			{
				"Excel.Sheet.12",
				("xlicons.exe", 1)
			},
			{
				"Excel.SheetMacroEnabled.12",
				("xlicons.exe", 2)
			},
			{
				"Excel.SheetBinaryMacroEnabled.12",
				("xlicons.exe", 3)
			},
			{
				"Excel.CSV",
				("xlicons.exe", 1)
			},
			{
				"PowerPoint.Show.8",
				("pptico.exe", 1)
			},
			{
				"PowerPoint.Show.12",
				("pptico.exe", 1)
			},
			{
				"PowerPoint.ShowMacroEnabled.12",
				("pptico.exe", 2)
			},
			{
				"PowerPoint.SlideShow.8",
				("pptico.exe", 1)
			},
			{
				"PowerPoint.SlideShow.12",
				("pptico.exe", 1)
			}
		};

		private const int SHCNE_ASSOCCHANGED = 134217728;

		private const int SHCNF_IDLIST = 0;

		private const int SHCNF_FLUSH = 4096;

		private const int WM_FONTCHANGE = 29;

		private static readonly IntPtr HWND_BROADCAST = new IntPtr(65535);

		public static void DeleteKey(string subKeyPath)
		{
			DeleteKey(Registry.CurrentUser, subKeyPath);
			DeleteKey(Registry.LocalMachine, subKeyPath);
		}

		private static void DeleteKey(RegistryKey root, string path)
		{
			try
			{
				root.DeleteSubKeyTree(path, throwOnMissingSubKey: false);
			}
			catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs", ex_captured); }
		}

		private static void WaitForProcessesToExit(string[] processNames, int timeoutMs = 15000)
		{
			Stopwatch stopwatch = Stopwatch.StartNew();
			while (stopwatch.ElapsedMilliseconds < timeoutMs)
			{
				bool flag = false;
				foreach (string processName in processNames)
				{
					Process[] processesByName = Process.GetProcessesByName(processName);
					if (processesByName.Length != 0)
					{
						flag = true;
						Process[] array = processesByName;
						foreach (Process process in array)
						{
							process.Dispose();
						}
						break;
					}
				}
				if (!flag)
				{
					break;
				}
				Thread.Sleep(500);
			}
		}

		private static void ForceDeleteRegistryKey(RegistryKey parentKey, string subKeyName)
		{
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
			catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs", ex_captured); }
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
			catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs", ex_captured); }
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
			catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs", ex_captured); }
			try
			{
				parentKey.DeleteSubKeyTree(subKeyName, throwOnMissingSubKey: false);
				Logger.Info("已强制清理受保护注册表项: " + subKeyName);
			}
			catch (Exception ex)
			{
				Logger.Warn("强制清理受保护注册表项 " + subKeyName + " 失败: " + ex.Message);
			}
		}

		public static void KillOfficeProcesses(ProductType product)
		{
			string[] array;
			switch (product)
			{
			default:
				return;
			case ProductType.MsOffice:
				array = new string[20]
				{
					"winword", "excel", "powerpnt", "outlook", "onenote", "publisher", "infopath", "visio", "winproj", "msaccess",
					"lync", "groove", "teams", "officeclicktorun", "officeintegration", "setuphost", "msoev", "msosync", "msoia", "setup"
				};
				break;
			case ProductType.Wps:
				array = new string[6] { "wps", "wpp", "et", "wpscloudsv", "wpscenter", "wpscloudsvr" };
				break;
			case ProductType.Yozo:
				array = new string[8] { "yozo_office", "yozo", "yozooffice", "yozoword", "yozosheet", "yozopresent", "yozopresentation", "yozo_binder" };
				break;
			case ProductType.OnlyOffice:
				array = new string[5] { "DesktopEditors", "ONLYOFFICE", "editors", "editors_helper", "updatesvc" };
				break;
			case ProductType.LibreOffice:
				array = new string[2] { "soffice.bin", "soffice.exe" };
				break;
			}
			string[] array2 = array;
			foreach (string processName in array2)
			{
				try
				{
					Process[] processesByName = Process.GetProcessesByName(processName);
					foreach (Process process in processesByName)
					{
						process.Kill();
						process.WaitForExit(2000);
						Logger.Info("已终止进程: " + process.ProcessName);
					}
				}
				catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs", ex_captured); }
			}
		}

		public static string GetInstalledProductVersion(ProductType product)
		{
			string[] array;
			switch (product)
			{
			case ProductType.MsOffice:
				array = new string[4] { "Microsoft Office", "Microsoft 365", "Office 16", "Office 15" };
				break;
			case ProductType.Wps:
				array = new string[1] { "WPS Office" };
				break;
			case ProductType.Yozo:
				array = new string[3] { "永中Office", "Yozo Office", "Yozosoft" };
				break;
			case ProductType.OnlyOffice:
				array = new string[1] { "ONLYOFFICE" };
				break;
			case ProductType.LibreOffice:
				array = new string[1] { "LibreOffice" };
				break;
			default:
				return null;
			}
			string[] array2 = new string[2] { "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall", "SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall" };
			RegistryKey[] array3 = new RegistryKey[2]
			{
				Registry.LocalMachine,
				Registry.CurrentUser
			};
			RegistryKey[] array4 = array3;
			foreach (RegistryKey registryKey in array4)
			{
				string[] array5 = array2;
				foreach (string name in array5)
				{
					try
					{
						using RegistryKey registryKey2 = registryKey.OpenSubKey(name);
						if (registryKey2 == null)
						{
							continue;
						}
						string[] subKeyNames = registryKey2.GetSubKeyNames();
						foreach (string name2 in subKeyNames)
						{
							try
							{
								using RegistryKey registryKey3 = registryKey2.OpenSubKey(name2);
								if (registryKey3 == null)
								{
									continue;
								}
								string text = registryKey3.GetValue("DisplayName") as string;
								if (string.IsNullOrEmpty(text) || (product == ProductType.MsOffice && (text.Contains("Access Runtime") || text.Contains("Language Pack"))))
								{
									continue;
								}
								string[] array6 = array;
								foreach (string text2 in array6)
								{
									if (text.ToLower().Contains(text2.ToLower()))
									{
										string text3 = registryKey3.GetValue("DisplayVersion") as string;
										return string.IsNullOrEmpty(text3) ? text : (text + " (" + text3 + ")");
									}
								}
							}
							catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs", ex_captured); }
						}
					}
					catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs", ex_captured); }
				}
			}
			return null;
		}

		public static void RemoveClickToRunService()
		{
			RunSc("stop ClickToRunSvc");
			RunSc("delete ClickToRunSvc");
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
				Process.Start(startInfo)?.WaitForExit(5000);
			}
			catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs", ex_captured); }
		}

		public static void CleanUninstallEntries(ProductType product)
		{
			List<FontBackupInfo> list = null;
			if (product == ProductType.Wps || product == ProductType.MsOffice)
			{
				list = BackupFonts();
			}
			string[] array = new string[2] { "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall", "SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall" };
			RegistryKey[] array2 = new RegistryKey[2]
			{
				Registry.LocalMachine,
				Registry.CurrentUser
			};
			RegistryKey[] array3 = array2;
			foreach (RegistryKey registryKey in array3)
			{
				string[] array4 = array;
				foreach (string name in array4)
				{
					try
					{
						using RegistryKey registryKey2 = registryKey.OpenSubKey(name, writable: true);
						if (registryKey2 == null)
						{
							continue;
						}
						string[] subKeyNames = registryKey2.GetSubKeyNames();
						foreach (string text in subKeyNames)
						{
							try
							{
								using RegistryKey registryKey3 = registryKey2.OpenSubKey(text);
								if (registryKey3 == null)
								{
									continue;
								}
								string text2 = (registryKey3.GetValue("DisplayName") as string) ?? "";
								string text3 = (registryKey3.GetValue("Publisher") as string) ?? "";
								string text4 = (registryKey3.GetValue("UninstallString") as string) ?? "";
								bool flag = text2.Contains("Microsoft Office") || text2.Contains("Microsoft 365") || text2.Contains("Office 16") || text2.Contains("Office 15") || text.StartsWith("Office1") || text4.Contains("OfficeClickToRun") || (text3.Contains("Microsoft Corporation") && (text2.Contains("Office") || text2.Contains("365")));
								bool flag2 = text2.Contains("WPS Office") || text.Contains("WPS Office");
								bool flag3 = text2.Contains("永中") || text2.Contains("Yozo");
								bool flag4 = text2.Contains("ONLYOFFICE") || text.Contains("ONLYOFFICE");
								bool flag5 = text2.Contains("LibreOffice") || text.Contains("LibreOffice");
								bool flag6 = false;
								switch (product)
								{
								case ProductType.MsOffice:
									flag6 = flag;
									break;
								case ProductType.Wps:
									flag6 = flag2;
									break;
								case ProductType.Yozo:
									flag6 = flag3;
									break;
								case ProductType.OnlyOffice:
									flag6 = flag4;
									break;
								case ProductType.LibreOffice:
									flag6 = flag5;
									break;
								}
								if (flag6)
								{
									Logger.Info("发现卸载项: " + text2 + "，准备静默调用卸载器...");
									if (!string.IsNullOrEmpty(text4))
									{
										RunUninstaller(text4);
									}
									registryKey2.DeleteSubKeyTree(text, throwOnMissingSubKey: false);
									Logger.Info("已清理注册表卸载项: " + text2);
								}
							}
							catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs", ex_captured); }
						}
					}
					catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs", ex_captured); }
				}
			}
			if (list != null && list.Count > 0)
			{
				Logger.Info("等待卸载程序释放文件锁，准备恢复字体...");
				Thread.Sleep(3000);
				RestoreFonts(list);
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
					string text3 = source.FirstOrDefault((string p) => p.Contains("{") && p.Contains("}"));
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
				Process.Start(startInfo)?.WaitForExit(30000);
			}
			catch (Exception ex)
			{
				Logger.Warn("调用卸载命令失败: " + uninstallString + ", 错误: " + ex.Message);
			}
		}

		public static void CleanResidualFolders(ProductType product)
		{
			
			string[] array = product switch
			{
				ProductType.MsOffice => new string[5] { "winword", "excel", "powerpnt", "officeclicktorun", "setup" }, 
				ProductType.Wps => new string[5] { "wps", "et", "wpp", "wpsuninstall", "uninstall" }, 
				ProductType.Yozo => new string[2] { "yozo", "uninstall" }, 
				ProductType.OnlyOffice => new string[5] { "DesktopEditors", "editors", "editors_helper", "updatesvc", "unins000" }, 
				ProductType.LibreOffice => new string[3] { "soffice", "soffice.bin", "uninstaller" }, 
				_ => new string[0], 
			};
			
			string[] array2 = array;
			if (array2.Length != 0)
			{
				WaitForProcessesToExit(array2);
			}
			List<string> list = new List<string>();
			List<string> list2 = new List<string>();
			string path = Environment.GetEnvironmentVariable("USERPROFILE") ?? "";
			switch (product)
			{
			case ProductType.MsOffice:
				list.AddRange(new string[9]
				{
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Microsoft Office"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Microsoft Office"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "Microsoft\\Office"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft\\Office"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft\\Office"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Common Files\\microsoft shared\\OFFICE16"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Common Files\\microsoft shared\\OFFICE16"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Common Files\\Microsoft Shared\\ClickToRun"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "Microsoft\\ClickToRun")
				});
				list2.AddRange(new string[2] { "SOFTWARE\\Microsoft\\Office", "SOFTWARE\\WOW6432Node\\Microsoft\\Office" });
				break;
			case ProductType.Wps:
				list.AddRange(new string[18]
				{
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Kingsoft"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Kingsoft"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Kingsoft"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Kingsoft"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "Kingsoft"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "WPS Office"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "WPS Office"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "WPS Office"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "WPS Office"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "WPS Office"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "Microsoft\\Windows\\Start Menu\\Programs\\WPS Office"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft\\Windows\\Start Menu\\Programs\\WPS Office"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Personal), "WPS\\backup"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Personal), "WPS\\template"),
					Path.Combine(Path.GetTempPath(), "wps"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Common Files\\Kingsoft"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Common Files\\Kingsoft"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "AppData\\LocalLow\\Kingsoft")
				});
				list2.AddRange(new string[9] { "SOFTWARE\\Kingsoft", "SOFTWARE\\WOW6432Node\\Kingsoft", "Software\\Kingsoft", "SOFTWARE\\WPS", "SOFTWARE\\WOW6432Node\\WPS", "SOFTWARE\\WPS Office", "SOFTWARE\\WOW6432Node\\WPS Office", "Software\\WPS", "Software\\WPS Office" });
				break;
			case ProductType.Yozo:
				list.AddRange(new string[8]
				{
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Yozosoft"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Yozosoft"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Yozosoft"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Yozosoft"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "Yozosoft"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "Microsoft\\Windows\\Start Menu\\Programs\\永中Office"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft\\Windows\\Start Menu\\Programs\\永中Office"),
					Path.Combine(path, "YozoOffice")
				});
				list2.AddRange(new string[3] { "SOFTWARE\\Yozosoft", "SOFTWARE\\WOW6432Node\\Yozosoft", "Software\\Yozosoft" });
				break;
			case ProductType.OnlyOffice:
				list.AddRange(new string[9]
				{
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "ONLYOFFICE"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "ONLYOFFICE"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Ascensio System SIA"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Ascensio System SIA"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ONLYOFFICE"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ONLYOFFICE"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "ONLYOFFICE"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "Microsoft\\Windows\\Start Menu\\Programs\\ONLYOFFICE"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft\\Windows\\Start Menu\\Programs\\ONLYOFFICE")
				});
				list2.AddRange(new string[6] { "SOFTWARE\\ONLYOFFICE", "SOFTWARE\\WOW6432Node\\ONLYOFFICE", "Software\\ONLYOFFICE", "SOFTWARE\\Ascensio System SIA", "SOFTWARE\\WOW6432Node\\Ascensio System SIA", "Software\\Ascensio System SIA" });
				break;
			case ProductType.LibreOffice:
				list.AddRange(new string[7]
				{
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "LibreOffice"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "LibreOffice"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "LibreOffice"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "LibreOffice"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "LibreOffice"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "Microsoft\\Windows\\Start Menu\\Programs\\LibreOffice"),
					Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft\\Windows\\Start Menu\\Programs\\LibreOffice")
				});
				list2.AddRange(new string[6] { "SOFTWARE\\The Document Foundation", "SOFTWARE\\LibreOffice", "SOFTWARE\\WOW6432Node\\The Document Foundation", "SOFTWARE\\WOW6432Node\\LibreOffice", "Software\\The Document Foundation", "Software\\LibreOffice" });
				break;
			}
			foreach (string item in list)
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
						Logger.Info($"目录 {item} 锁定中或正在被卸载程序使用，等待 2 秒后进行第 {i + 2} 次重试...");
						Thread.Sleep(2000);
						continue;
					}
					break;
				}
			}
			foreach (string item2 in list2)
			{
				DeleteKey(item2);
			}
		}

		public static void CleanShortcuts(ProductType product)
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
			string[] array;
			switch (product)
			{
			default:
				return;
			case ProductType.MsOffice:
				array = new string[9] { "Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Access", "Publisher", "Visio", "Project" };
				break;
			case ProductType.Wps:
				array = new string[2] { "WPS", "金山" };
				break;
			case ProductType.Yozo:
				array = new string[2] { "永中", "Yozo" };
				break;
			case ProductType.OnlyOffice:
				array = new string[1] { "ONLYOFFICE" };
				break;
			case ProductType.LibreOffice:
				array = new string[1] { "LibreOffice" };
				break;
			}
			foreach (string item in list)
			{
				try
				{
					if (!Directory.Exists(item))
					{
						continue;
					}
					string[] files = Directory.GetFiles(item, "*.*", SearchOption.AllDirectories);
					string[] array2 = files;
					foreach (string text in array2)
					{
						string text2 = Path.GetExtension(text).ToLower();
						if (text2 != ".lnk" && text2 != ".url")
						{
							continue;
						}
						bool flag = false;
						string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(text);
						string[] array3 = array;
						foreach (string value in array3)
						{
							if (fileNameWithoutExtension.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0)
							{
								flag = true;
								break;
							}
						}
						if (!flag && text2 == ".lnk")
						{
							string shortcutTarget = GetShortcutTarget(text);
							if (!string.IsNullOrEmpty(shortcutTarget))
							{
								
								string[] array4 = product switch
								{
									ProductType.MsOffice => new string[3] { "Microsoft Office", "Office16", "Office15" }, 
									ProductType.Wps => new string[3] { "Kingsoft", "WPS Office", "WPSOffice" }, 
									ProductType.Yozo => new string[2] { "Yozosoft", "Yozo" }, 
									ProductType.OnlyOffice => new string[1] { "ONLYOFFICE" }, 
									ProductType.LibreOffice => new string[1] { "LibreOffice" }, 
									_ => new string[0], 
								};
								
								string[] array5 = array4;
								string[] array6 = array5;
								foreach (string value2 in array6)
								{
									if (shortcutTarget.IndexOf(value2, StringComparison.OrdinalIgnoreCase) >= 0)
									{
										flag = true;
										break;
									}
								}
							}
						}
						if (!flag && text2 == ".url")
						{
							try
							{
								string text3 = File.ReadAllText(text);
								
								string[] array4 = product switch
								{
									ProductType.MsOffice => new string[2] { "office.com", "microsoft" }, 
									ProductType.Wps => new string[2] { "wps.cn", "kingsoft" }, 
									ProductType.Yozo => new string[2] { "yozo", "yozosoft" }, 
									ProductType.OnlyOffice => new string[1] { "onlyoffice" }, 
									ProductType.LibreOffice => new string[1] { "libreoffice" }, 
									_ => new string[0], 
								};
								
								string[] array7 = array4;
								string[] array8 = array7;
								foreach (string value3 in array8)
								{
									if (text3.IndexOf(value3, StringComparison.OrdinalIgnoreCase) >= 0)
									{
										flag = true;
										break;
									}
								}
							}
							catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs", ex_captured); }
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
					if (!item.Contains("Start Menu") && !item.Contains("Programs"))
					{
						continue;
					}
					string[] directories = Directory.GetDirectories(item, "*", SearchOption.AllDirectories);
					Array.Sort(directories, (string a, string b) => b.Length.CompareTo(a.Length));
					string[] array9 = directories;
					foreach (string text4 in array9)
					{
						if (!Directory.Exists(text4))
						{
							continue;
						}
						string fileName = Path.GetFileName(text4);
						bool flag2 = false;
						string[] array10 = array;
						foreach (string value4 in array10)
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
							catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs", ex_captured); }
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

		private static string GetShortcutTarget(string lnkPath)
		{
			try
			{
				Type typeFromProgID = Type.GetTypeFromProgID("WScript.Shell");
				if (typeFromProgID == null)
				{
					return "";
				}
				object target = Activator.CreateInstance(typeFromProgID);
				object obj = typeFromProgID.InvokeMember("CreateShortcut", BindingFlags.InvokeMethod, null, target, new object[1] { lnkPath });
				if (obj == null)
				{
					return "";
				}
				string text = obj.GetType().InvokeMember("TargetPath", BindingFlags.GetProperty, null, obj, null) as string;
				return text ?? "";
			}
			catch
			{
				return "";
			}
		}

		public static void CleanFileAssociations(ProductType product)
		{
			string[] array;
			switch (product)
			{
			default:
				return;
			case ProductType.MsOffice:
				array = new string[6] { "Word.", "Excel.", "PowerPoint.", "Access.", "Outlook.", "OneNote." };
				break;
			case ProductType.Wps:
				array = new string[11]
				{
					"WPS.", "WPP.", "ET.", "KET.", "KWPP.", "KWPS.", "KPDF.", "Kingsoft", "wpsonline", "ksowps",
					"ksoWPSCloudSvr"
				};
				break;
			case ProductType.Yozo:
				array = new string[6] { "Yozo", "yozoword", "yozosheet", "yozopresent", "yozobinder", "YOO." };
				break;
			case ProductType.OnlyOffice:
				array = new string[2] { "ONLYOFFICE.", "Ascensio" };
				break;
			case ProductType.LibreOffice:
				array = new string[2] { "LibreOffice.", "soffice." };
				break;
			}
			if (ProductExecutables.TryGetValue(product, out var value))
			{
				string[] array2 = value;
				foreach (string text in array2)
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
			RegistryKey[] array3 = new RegistryKey[2]
			{
				Registry.LocalMachine,
				Registry.CurrentUser
			};
			RegistryKey[] array4 = array3;
			foreach (RegistryKey registryKey in array4)
			{
				try
				{
					using RegistryKey registryKey2 = registryKey.OpenSubKey("SOFTWARE\\Classes", writable: true);
					if (registryKey2 == null)
					{
						continue;
					}
					string[] subKeyNames = registryKey2.GetSubKeyNames();
					foreach (string text2 in subKeyNames)
					{
						bool flag = false;
						string[] array5 = array;
						foreach (string value2 in array5)
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
								registryKey2.DeleteSubKeyTree(text2, throwOnMissingSubKey: false);
								Logger.Info("已删除 ProgID 关联键: " + registryKey.Name + "\\SOFTWARE\\Classes\\" + text2);
							}
							catch (Exception ex2)
							{
								Logger.Warn("删除 ProgID 键失败: " + text2 + ", 错误: " + ex2.Message);
							}
						}
					}
					string[] array6 = new string[3] { "*\\shellex\\ContextMenuHandlers", "Directory\\Background\\shellex\\ContextMenuHandlers", "Folder\\shellex\\ContextMenuHandlers" };
					string[] array7 = array6;
					foreach (string text3 in array7)
					{
						try
						{
							using RegistryKey registryKey3 = registryKey2.OpenSubKey(text3, writable: true);
							if (registryKey3 == null)
							{
								continue;
							}
							string[] subKeyNames2 = registryKey3.GetSubKeyNames();
							foreach (string text4 in subKeyNames2)
							{
								bool flag2 = false;
								string[] array8 = array;
								foreach (string value3 in array8)
								{
									if (text4.IndexOf(value3, StringComparison.OrdinalIgnoreCase) >= 0)
									{
										flag2 = true;
										break;
									}
								}
								if (flag2)
								{
									registryKey3.DeleteSubKeyTree(text4, throwOnMissingSubKey: false);
									Logger.Info("已删除右键菜单残留: " + registryKey.Name + "\\SOFTWARE\\Classes\\" + text3 + "\\" + text4);
								}
							}
						}
						catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs", ex_captured); }
					}
					string[] subKeyNames3 = registryKey2.GetSubKeyNames();
					foreach (string text5 in subKeyNames3)
					{
						if (!text5.StartsWith("."))
						{
							continue;
						}
						try
						{
							using RegistryKey registryKey4 = registryKey2.OpenSubKey(text5, writable: true);
							if (registryKey4 == null)
							{
								continue;
							}
							string text6 = registryKey4.GetValue("") as string;
							if (!string.IsNullOrEmpty(text6))
							{
								bool flag3 = false;
								string[] array9 = array;
								foreach (string value4 in array9)
								{
									if (text6.StartsWith(value4, StringComparison.OrdinalIgnoreCase))
									{
										flag3 = true;
										break;
									}
								}
								if (flag3)
								{
									registryKey4.SetValue("", "");
									Logger.Info("已清除扩展名默认 ProgID 指向: " + registryKey.Name + "\\SOFTWARE\\Classes\\" + text5);
								}
							}
							using RegistryKey registryKey5 = registryKey4.OpenSubKey("OpenWithProgids", writable: true);
							if (registryKey5 == null)
							{
								continue;
							}
							string[] valueNames = registryKey5.GetValueNames();
							foreach (string text7 in valueNames)
							{
								bool flag4 = false;
								string[] array10 = array;
								foreach (string value5 in array10)
								{
									if (text7.StartsWith(value5, StringComparison.OrdinalIgnoreCase))
									{
										flag4 = true;
										break;
									}
								}
								if (flag4)
								{
									registryKey5.DeleteValue(text7, throwOnMissingValue: false);
									Logger.Info("已清除 OpenWithProgids 关联: " + registryKey.Name + "\\SOFTWARE\\Classes\\" + text5 + "\\OpenWithProgids\\" + text7);
								}
							}
						}
						catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs", ex_captured); }
					}
				}
				catch (Exception ex3)
				{
					Logger.Warn("清理 Classes 关联失败 (" + registryKey.Name + "): " + ex3.Message);
				}
			}
			try
			{
				
				string[] array11 = product switch
				{
					ProductType.Wps => new string[4] { "wps", "et", "wpp", "金山" }, 
					ProductType.Yozo => new string[2] { "yozo", "永中" }, 
					ProductType.OnlyOffice => new string[1] { "onlyoffice" }, 
					ProductType.LibreOffice => new string[2] { "libreoffice", "soffice" }, 
					_ => new string[0], 
				};
				
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
							string[] array14 = array13;
							foreach (string text9 in array14)
							{
								bool flag5 = false;
								using (RegistryKey registryKey7 = registryKey6.OpenSubKey(text8 + "\\" + text9))
								{
									if (registryKey7 != null)
									{
										string text10 = null;
										if (text9 == "UserChoice")
										{
											text10 = registryKey7.GetValue("ProgId") as string;
										}
										else
										{
											using RegistryKey registryKey8 = registryKey7.OpenSubKey("ProgId");
											if (registryKey8 != null)
											{
												text10 = registryKey8.GetValue("ProgId") as string;
											}
										}
										if (!string.IsNullOrEmpty(text10))
										{
											string[] array15 = array;
											foreach (string value6 in array15)
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
									ForceDeleteRegistryKey(registryKey6, text8 + "\\" + text9);
								}
							}
							using (RegistryKey registryKey9 = registryKey6.OpenSubKey(text8 + "\\OpenWithProgids", writable: true))
							{
								if (registryKey9 != null)
								{
									string[] valueNames2 = registryKey9.GetValueNames();
									foreach (string text11 in valueNames2)
									{
										bool flag6 = false;
										string[] array16 = array;
										foreach (string value7 in array16)
										{
											if (text11.StartsWith(value7, StringComparison.OrdinalIgnoreCase))
											{
												flag6 = true;
												break;
											}
										}
										if (flag6)
										{
											registryKey9.DeleteValue(text11, throwOnMissingValue: false);
											Logger.Info("已清除 Explorer FileExts 历史 ProgID: FileExts\\" + text8 + "\\OpenWithProgids\\" + text11);
										}
									}
								}
							}
							using RegistryKey registryKey10 = registryKey6.OpenSubKey(text8 + "\\OpenWithList", writable: true);
							if (registryKey10 == null)
							{
								continue;
							}
							string text12 = (registryKey10.GetValue("MRUList") as string) ?? "";
							string text13 = "";
							List<string> list = new List<string>();
							string[] valueNames3 = registryKey10.GetValueNames();
							foreach (string text14 in valueNames3)
							{
								if (text14.Equals("MRUList", StringComparison.OrdinalIgnoreCase))
								{
									continue;
								}
								string text15 = registryKey10.GetValue(text14) as string;
								if (string.IsNullOrEmpty(text15))
								{
									continue;
								}
								bool flag7 = false;
								string[] array17 = array12;
								foreach (string value8 in array17)
								{
									if (text15.IndexOf(value8, StringComparison.OrdinalIgnoreCase) >= 0)
									{
										flag7 = true;
										break;
									}
								}
								if (flag7)
								{
									list.Add(text14);
								}
								else if (text14.Length == 1 && text12.Contains(text14))
								{
									text13 += text14;
								}
							}
							foreach (string item3 in list)
							{
								registryKey10.DeleteValue(item3, throwOnMissingValue: false);
								Logger.Info("已清除 Explorer FileExts 历史打开方式: FileExts\\" + text8 + "\\OpenWithList\\" + item3);
							}
							if (text13 != text12)
							{
								registryKey10.SetValue("MRUList", text13);
							}
						}
						catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs", ex_captured); }
					}
				}
			}
			catch (Exception ex4)
			{
				Logger.Warn("清理 Explorer FileExts 关联失败: " + ex4.Message);
			}
			if (ProductExecutables.TryGetValue(product, out var value9))
			{
				string[] array18 = new string[25]
				{
					".doc", ".docx", ".docm", ".dot", ".dotx", ".rtf", ".xls", ".xlsx", ".xlsm", ".xlsb",
					".csv", ".ppt", ".pptx", ".pptm", ".pps", ".ppsx", ".pdf", ".jpg", ".jpeg", ".png",
					".bmp", ".gif", ".webp", ".tif", ".tiff"
				};
				(RegistryKey, string)[] array19 = new(RegistryKey, string)[4]
				{
					(Registry.ClassesRoot, ""),
					(Registry.ClassesRoot, "SystemFileAssociations"),
					(Registry.LocalMachine, "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts"),
					(Registry.CurrentUser, "Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts")
				};
				string[] array20 = array18;
				foreach (string text16 in array20)
				{
					(RegistryKey, string)[] array21 = array19;
					for (int num14 = 0; num14 < array21.Length; num14++)
					{
						(RegistryKey, string) tuple = array21[num14];
						RegistryKey item = tuple.Item1;
						string item2 = tuple.Item2;
						string text17 = (string.IsNullOrEmpty(item2) ? (text16 + "\\OpenWithList") : (item2 + "\\" + text16 + "\\OpenWithList"));
						try
						{
							using RegistryKey registryKey11 = item.OpenSubKey(text17, writable: true);
							if (registryKey11 == null)
							{
								continue;
							}
							string[] subKeyNames5 = registryKey11.GetSubKeyNames();
							foreach (string text18 in subKeyNames5)
							{
								bool flag8 = false;
								string[] array22 = value9;
								foreach (string value10 in array22)
								{
									if (text18.Equals(value10, StringComparison.OrdinalIgnoreCase))
									{
										flag8 = true;
										break;
									}
								}
								if (flag8)
								{
									registryKey11.DeleteSubKeyTree(text18, throwOnMissingSubKey: false);
									Logger.Info("已清理 OpenWithList 子键关联: " + item.Name + "\\" + text17 + "\\" + text18);
								}
							}
							string[] valueNames4 = registryKey11.GetValueNames();
							foreach (string text19 in valueNames4)
							{
								if (text19.Equals("MRUList", StringComparison.OrdinalIgnoreCase))
								{
									continue;
								}
								string text20 = registryKey11.GetValue(text19) as string;
								if (string.IsNullOrEmpty(text20))
								{
									continue;
								}
								bool flag9 = false;
								string[] array23 = value9;
								foreach (string value11 in array23)
								{
									if (text20.Equals(value11, StringComparison.OrdinalIgnoreCase) || text19.Equals(value11, StringComparison.OrdinalIgnoreCase))
									{
										flag9 = true;
										break;
									}
								}
								if (flag9)
								{
									registryKey11.DeleteValue(text19, throwOnMissingValue: false);
									Logger.Info("已清理 OpenWithList 值关联: " + item.Name + "\\" + text17 + "\\" + text19 + " -> " + text20);
								}
							}
						}
						catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs", ex_captured); }
					}
				}
			}
			try
			{
				RestoreInstalledProductAssociations();
			}
			catch (Exception ex5)
			{
				Logger.Warn("修复其余已安装产品关联失败: " + ex5.Message);
			}
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
				
				string[] array11 = product switch
				{
					ProductType.Wps => new string[5] { "wps", "et", "wpp", "金山", "kso" }, 
					ProductType.Yozo => new string[2] { "yozo", "永中" }, 
					ProductType.OnlyOffice => new string[1] { "onlyoffice" }, 
					ProductType.LibreOffice => new string[2] { "libreoffice", "soffice" }, 
					_ => new string[0], 
				};
				
				string[] array24 = array11;
				using RegistryKey registryKey12 = Registry.CurrentUser.OpenSubKey("Software\\Classes\\Local Settings\\Software\\Microsoft\\Windows\\Shell\\MuiCache", writable: true);
				if (registryKey12 != null)
				{
					string[] valueNames5 = registryKey12.GetValueNames();
					foreach (string text21 in valueNames5)
					{
						bool flag10 = false;
						string[] array25 = array24;
						foreach (string value12 in array25)
						{
							if (text21.IndexOf(value12, StringComparison.OrdinalIgnoreCase) >= 0)
							{
								flag10 = true;
								break;
							}
						}
						if (flag10)
						{
							registryKey12.DeleteValue(text21, throwOnMissingValue: false);
							Logger.Info("已清理 MuiCache 关联缓存: " + text21);
						}
					}
				}
			}
			catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs", ex_captured); }
			try
			{
				RestoreInstalledProductAssociations();
			}
			catch (Exception ex7)
			{
				Logger.Warn("修复其余已安装产品关联失败: " + ex7.Message);
			}
			try
			{
				RefreshIconCache();
			}
			catch (Exception ex8)
			{
				Logger.Error("自动刷新图标缓存失败", ex8);
			}
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
				using RegistryKey registryKey2 = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\" + exeName);
				if (registryKey2 != null)
				{
					string text2 = registryKey2.GetValue("") as string;
					if (!string.IsNullOrEmpty(text2))
					{
						return text2.Trim('"');
					}
				}
			}
			catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs", ex_captured); }
			return "";
		}

		private static bool ShouldRestoreExtensionAssociation(string ext)
		{
			// 检查当前的 UserChoice
			using (RegistryKey rkExt = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts\\" + ext + "\\UserChoice"))
			{
				if (rkExt != null)
				{
					string userChoiceProgId = rkExt.GetValue("ProgId") as string;
					if (!string.IsNullOrEmpty(userChoiceProgId))
					{
						// 如果 UserChoice 指向一个有效的非 Office 应用（例如 txtfile, Applications\notepad++.exe 等），不应该覆盖它
						// 只有当它指向已卸载的/损坏的 Office ProgID 或者为空时，我们才应该修复它
						bool isOfficeProgId = userChoiceProgId.StartsWith("WPS.", StringComparison.OrdinalIgnoreCase) ||
											  userChoiceProgId.StartsWith("WPP.", StringComparison.OrdinalIgnoreCase) ||
											  userChoiceProgId.StartsWith("ET.", StringComparison.OrdinalIgnoreCase) ||
											  userChoiceProgId.StartsWith("Word.", StringComparison.OrdinalIgnoreCase) ||
											  userChoiceProgId.StartsWith("Excel.", StringComparison.OrdinalIgnoreCase) ||
											  userChoiceProgId.StartsWith("PowerPoint.", StringComparison.OrdinalIgnoreCase) ||
											  userChoiceProgId.StartsWith("Yozo", StringComparison.OrdinalIgnoreCase) ||
											  userChoiceProgId.StartsWith("ONLYOFFICE", StringComparison.OrdinalIgnoreCase) ||
											  userChoiceProgId.StartsWith("LibreOffice", StringComparison.OrdinalIgnoreCase) ||
											  userChoiceProgId.StartsWith("soffice", StringComparison.OrdinalIgnoreCase);

						if (!isOfficeProgId)
						{
							// 用户显式设置了其他自定义非 Office 程序（如 Notepad++），绝不覆盖！
							return false;
						}
					}
				}
			}

			// 检查 Classes 下的默认 ProgID
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
				bool isOfficeProgId = currentProgId.StartsWith("WPS.", StringComparison.OrdinalIgnoreCase) ||
									  currentProgId.StartsWith("WPP.", StringComparison.OrdinalIgnoreCase) ||
									  currentProgId.StartsWith("ET.", StringComparison.OrdinalIgnoreCase) ||
									  currentProgId.StartsWith("Word.", StringComparison.OrdinalIgnoreCase) ||
									  currentProgId.StartsWith("Excel.", StringComparison.OrdinalIgnoreCase) ||
									  currentProgId.StartsWith("PowerPoint.", StringComparison.OrdinalIgnoreCase) ||
									  currentProgId.StartsWith("Yozo", StringComparison.OrdinalIgnoreCase) ||
									  currentProgId.StartsWith("ONLYOFFICE", StringComparison.OrdinalIgnoreCase) ||
									  currentProgId.StartsWith("LibreOffice", StringComparison.OrdinalIgnoreCase) ||
									  currentProgId.StartsWith("soffice", StringComparison.OrdinalIgnoreCase);

				if (!isOfficeProgId)
				{
					// 用户设置了其他非 Office 默认程序，不覆盖
					return false;
				}
			}

			return true;
		}

		public static void RestoreInstalledProductAssociations(IProgress<string> progress = null)
		{
			string installedProductVersion = GetInstalledProductVersion(ProductType.MsOffice);
			Dictionary<string, string[]> dictionary;
			if (!string.IsNullOrEmpty(installedProductVersion))
			{
				dictionary = new Dictionary<string, string[]>();
				dictionary.Add("winword.exe", new string[5] { ".doc", ".docx", ".docm", ".dot", ".dotx" });
				dictionary.Add("excel.exe", new string[4] { ".xls", ".xlsx", ".xlsm", ".xlsb" });
				dictionary.Add("powerpnt.exe", new string[5] { ".ppt", ".pptx", ".pptm", ".pps", ".ppsx" });
				Dictionary<string, string[]> dictionary2 = dictionary;
				foreach (KeyValuePair<string, string[]> item in dictionary2)
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
						string[] array2 = array;
						foreach (string text2 in array2)
						{
							if (File.Exists(text2))
							{
								text = text2;
								break;
							}
						}
					}
					if (string.IsNullOrEmpty(text) || !File.Exists(text))
					{
						continue;
					}
					
					string text3 = key switch
					{
						"winword.exe" => "Word", 
						"excel.exe" => "Excel", 
						"powerpnt.exe" => "PowerPoint", 
						_ => key, 
					};
					
					string text4 = text3;
					progress?.Report(LocalizationStrings.Instance.StatusDetectMsOffice(text4));
					string[] array3 = value;
					foreach (string text5 in array3)
					{
						if (!MsOfficeProgIds.TryGetValue(text5, out var value2))
						{
							continue;
						}
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
									Logger.Info("已强力清除 " + text5 + " 的 UserChoice 与 UserChoiceLatest 以激活 MS Office " + text4 + " 默认关联");
								}
							}
							catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs", ex_captured); }
							using (RegistryKey registryKey4 = Registry.CurrentUser.CreateSubKey("Software\\Classes\\" + text5 + "\\OpenWithProgids"))
							{
								if (registryKey4 != null && registryKey4.GetValue(value2) == null)
								{
									registryKey4.SetValue(value2, new byte[0], RegistryValueKind.Binary);
								}
							}
							using RegistryKey registryKey5 = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Classes\\" + text5 + "\\OpenWithProgids");
							if (registryKey5 != null && registryKey5.GetValue(value2) == null)
							{
								registryKey5.SetValue(value2, new byte[0], RegistryValueKind.Binary);
							}
						}
						catch (Exception ex)
						{
							Logger.Warn("修复 " + text5 + " 的 MS Office 关联 " + value2 + " 失败: " + ex.Message);
						}
					}
					string[] array4 = value;
					foreach (string key2 in array4)
					{
						if (!MsOfficeProgIds.TryGetValue(key2, out var value3))
						{
							continue;
						}
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
						
						text3 = ((!(key == "winword.exe")) ? ("\"" + text + "\" \"%1\"") : ("\"" + text + "\" /n \"%1\" /o \"%u\""));
						
						string text8 = text3;
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
			if (string.IsNullOrEmpty(installedProductVersion2))
			{
				return;
			}
			dictionary = new Dictionary<string, string[]>();
			dictionary.Add("wps.exe", new string[2] { ".doc", ".docx" });
			dictionary.Add("et.exe", new string[2] { ".xls", ".xlsx" });
			dictionary.Add("wpp.exe", new string[2] { ".ppt", ".pptx" });
			Dictionary<string, string[]> dictionary3 = dictionary;
			foreach (KeyValuePair<string, string[]> item2 in dictionary3)
			{
				string key3 = item2.Key;
				string[] value5 = item2.Value;
				string appPathFromRegistry = GetAppPathFromRegistry(key3);
				if (string.IsNullOrEmpty(appPathFromRegistry) || !File.Exists(appPathFromRegistry))
				{
					continue;
				}
				
				string text3 = key3 switch
				{
					"wps.exe" => "WPS 文字", 
					"et.exe" => "WPS 表格", 
					"wpp.exe" => "WPS 演示", 
					_ => key3, 
				};
				
				string text9 = text3;
				progress?.Report(LocalizationStrings.Instance.StatusDetectWps(text9));
				string[] array5 = value5;
				foreach (string text10 in array5)
				{
					if (!WpsProgIds.TryGetValue(text10, out var value6))
					{
						continue;
					}
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
								Logger.Info("已强力清除 " + text10 + " 的 UserChoice 与 UserChoiceLatest 以激活 WPS " + text9 + " 默认关联");
							}
						}
						catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs", ex_captured); }
						using (RegistryKey registryKey12 = Registry.CurrentUser.CreateSubKey("Software\\Classes\\" + text10 + "\\OpenWithProgids"))
						{
							if (registryKey12 != null && registryKey12.GetValue(value6) == null)
							{
								registryKey12.SetValue(value6, new byte[0], RegistryValueKind.Binary);
							}
						}
						using RegistryKey registryKey13 = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Classes\\" + text10 + "\\OpenWithProgids");
						if (registryKey13 != null && registryKey13.GetValue(value6) == null)
						{
							registryKey13.SetValue(value6, new byte[0], RegistryValueKind.Binary);
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
			ProductType[] array2 = array;
			foreach (ProductType productType in array2)
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
					
					string text2 = text;
					progress?.Report(LocalizationStrings.Instance.StatusPurgingProduct(text2));
					try
					{
						CleanFileAssociations(productType);
						CleanShortcuts(productType);
						num++;
					}
					catch (Exception ex)
					{
						Logger.Error("清除 " + text2 + " 残留文件关联失败", ex);
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
						using RegistryKey registryKey2 = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\" + item);
						if (registryKey2 != null)
						{
							text = registryKey2.GetValue("") as string;
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
						string[] array2 = array;
						foreach (string text2 in array2)
						{
							if (File.Exists(text2))
							{
								text = text2;
								break;
							}
						}
					}
					if (string.IsNullOrEmpty(text))
					{
						continue;
					}
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
						using Process process = Process.Start(startInfo);
						process?.WaitForExit(10000);
					}
				}
				catch (Exception ex)
				{
					Logger.Warn("修复 " + item + " 的 COM 组件注册失败: " + ex.Message);
				}
			}
		}

		[DllImport("shell32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		private static extern void SHChangeNotify(int wEventId, int uFlags, IntPtr dwItem1, IntPtr dwItem2);

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
				catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs", ex_captured); }
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
				catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs", ex_captured); }
			}
			catch (Exception ex)
			{
				Logger.Warn("刷新图标缓存失败: " + ex.Message);
			}
		}

		[DllImport("gdi32.dll", CharSet = CharSet.Unicode, EntryPoint = "AddFontResourceW")]
		private static extern int AddFontResource(string lpFileName);

		[DllImport("user32.dll", CharSet = CharSet.Auto)]
		private static extern int SendMessage(IntPtr hWnd, int wMsg, IntPtr wParam, IntPtr lParam);

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
				(RegistryKey, bool)[] array2 = array;
				for (int i = 0; i < array2.Length; i++)
				{
					var (registryKey, isUserFont) = array2[i];
					using RegistryKey registryKey2 = registryKey.OpenSubKey("SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Fonts");
					if (registryKey2 == null)
					{
						continue;
					}
					string[] valueNames = registryKey2.GetValueNames();
					foreach (string text2 in valueNames)
					{
						try
						{
							bool flag = text2.Contains("方正") || text2.IndexOf("fz", StringComparison.OrdinalIgnoreCase) >= 0;
							string text3 = registryKey2.GetValue(text2) as string;
							if (string.IsNullOrEmpty(text3))
							{
								continue;
							}
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
			if (backupList == null || backupList.Count == 0)
			{
				return;
			}
			string folderPath = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
			bool flag = false;
			foreach (FontBackupInfo backup in backupList)
			{
				try
				{
					if (!File.Exists(backup.BackupPath))
					{
						continue;
					}
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
					catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs", ex_captured); }
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
					SendMessage(HWND_BROADCAST, 29, IntPtr.Zero, IntPtr.Zero);
					Logger.Info("已广播 WM_FONTCHANGE 字体变更消息通知系统");
				}
				catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs", ex_captured); }
			}
			try
			{
				string path = Path.Combine(Path.GetTempPath(), "GOIFontBackup");
				if (Directory.Exists(path))
				{
					Directory.Delete(path, recursive: true);
				}
			}
			catch (Exception ex_captured) { GOI.Helpers.Logger.Error("Silent exception in RegistryHelper.cs", ex_captured); }
		}
	}
}
