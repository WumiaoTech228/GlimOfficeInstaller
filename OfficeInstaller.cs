// ============================================================
// Glim Office Installer - Office 一键安装工具
// 功能：自动下载、安装、激活 Microsoft Office
// 作者：Glim
// ============================================================

// ========== 引用命名空间 ==========
// System: 基础类库，包含基本数据类型和常用功能
using System;
// System.Collections.Generic: 泛型集合类，如List<T>、Dictionary<K,V>
using System.Collections.Generic;
// System.Diagnostics: 进程管理和调试功能，用于启动外部程序
using System.Diagnostics;
// System.Drawing: 图形绘制功能，用于绘制界面元素
using System.Drawing;
// System.Drawing.Drawing2D: 高级2D图形功能，如渐变、抗锯齿
using System.Drawing.Drawing2D;
// System.Drawing.Text: 字体和文本渲染功能
using System.Drawing.Text;
// System.IO: 文件和目录操作，如读写文件、创建文件夹
using System.IO;
// System.Linq: LINQ查询功能，用于集合数据查询
using System.Linq;
// System.Net: 网络功能，如WebClient下载文件
using System.Net;
// System.Reflection: 反射功能，用于获取程序集信息和嵌入资源
using System.Reflection;
// System.Runtime.InteropServices: 调用Windows API功能
using System.Runtime.InteropServices;
// System.Text: 文本编码功能，如UTF-8编码
using System.Text;
// System.Threading.Tasks: 异步编程功能，如async/await
using System.Threading.Tasks;
// System.Windows.Forms: Windows窗体应用程序，创建GUI界面
using System.Windows.Forms;
// System.Security.Principal: 用户身份验证，用于检查管理员权限
using System.Security.Principal;

// ========== 命名空间定义 ==========
// namespace: 代码的组织单位，防止类名冲突
// Office2024Installer: 本项目的命名空间名称
namespace Office2024Installer
{
    // ========== 全局配置类 ==========
    // static class: 静态类，不能实例化，所有成员都是静态的
    // GOIConfig: Glim Office Installer 的配置管理类
    // 作用：统一管理程序运行时的路径配置
    public static class GOIConfig
    {
        // readonly: 只读字段，只能在构造函数或声明时赋值
        // RootPath: 程序根目录，所有文件都在这个目录下
        // AppDomain.CurrentDomain.BaseDirectory: 获取程序所在目录
        // Path.Combine: 安全地拼接路径，自动处理斜杠
        public static readonly string RootPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "GOI");
        
        // LogPath: 日志文件夹路径
        public static readonly string LogPath = Path.Combine(RootPath, "logs");
        
        // LogFile: 当前日志文件的完整路径
        // DateTime.Now.ToString("yyyyMMdd_HHmmss"): 格式化当前时间为文件名
        public static readonly string LogFile = Path.Combine(LogPath, string.Format("install_{0}.log", DateTime.Now.ToString("yyyyMMdd_HHmmss")));
        
        // DownloadPath: 下载文件存放路径
        public static readonly string DownloadPath = Path.Combine(RootPath, "downloads");
        
        // ToolsPath: 工具文件存放路径（如激活工具）
        public static readonly string ToolsPath = Path.Combine(RootPath, "tools");
        
        // SetupPath: Office安装程序路径
        public static readonly string SetupPath = Path.Combine(RootPath, "setup.exe");
        
        // Initialize(): 初始化方法，创建必要的文件夹
        public static void Initialize()
        {
            try
            {
                // Directory.Exists: 检查目录是否存在
                // Directory.CreateDirectory: 创建目录（如果不存在）
                if (!Directory.Exists(RootPath)) Directory.CreateDirectory(RootPath);
                if (!Directory.Exists(LogPath)) Directory.CreateDirectory(LogPath);
                if (!Directory.Exists(DownloadPath)) Directory.CreateDirectory(DownloadPath);
                if (!Directory.Exists(ToolsPath)) Directory.CreateDirectory(ToolsPath);
                // 记录初始化完成的日志
                Logger.Info("GOI 运行环境初始化完成。根路径: " + RootPath);
            }
            catch (Exception ex)
            {
                // 异常处理：显示错误消息
                // GlimMessageBox: 自定义消息框，使用程序图标
                GlimMessageBox.Show("无法创建 GOI 工作目录，请确保程序有写入权限。\n错误: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    // ========== 日志系统类 ==========
    // Logger: 日志记录类，用于记录程序运行信息
    // 作用：方便调试和追踪程序运行状态
    public static class Logger
    {
        // lockObj: 锁对象，用于多线程同步
        // 防止多个线程同时写入日志文件导致冲突
        private static readonly object lockObj = new object();

        // Info: 记录普通信息日志
        public static void Info(string message) 
        { 
            Log("INFO", message); 
        }
        
        // Warn: 记录警告日志
        public static void Warn(string message) 
        { 
            Log("WARN", message); 
        }
        
        // Error: 记录错误日志（带异常信息）
        public static void Error(string message, Exception ex) 
        {
            string msg = message;
            if (ex != null)
            {
                // 将异常信息和堆栈跟踪添加到日志
                msg = msg + " | 异常: " + ex.Message + "\n" + ex.StackTrace;
            }
            Log("ERROR", msg);
        }
        
        // Error: 记录错误日志（不带异常信息）
        public static void Error(string message)
        {
            Log("ERROR", message);
        }

        // Log: 核心日志写入方法
        // level: 日志级别（INFO/WARN/ERROR）
        // message: 日志内容
        private static void Log(string level, string message)
        {
            try
            {
                // lock: 线程锁，确保同一时间只有一个线程能执行这段代码
                lock (lockObj)
                {
                    // 格式化日志行：[时间] [级别] 消息
                    string logLine = string.Format("[{0}] [{1}] {2}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"), level, message);
                    // Debug.WriteLine: 输出到调试窗口
                    Debug.WriteLine(logLine);
                    // 检查日志目录是否存在
                    if (Directory.Exists(GOIConfig.LogPath))
                    {
                        // File.AppendAllText: 追加文本到文件
                        // Environment.NewLine: 换行符（Windows是\r\n）
                        File.AppendAllText(GOIConfig.LogFile, logLine + Environment.NewLine, Encoding.UTF8);
                    }
                }
            }
            catch { } // 忽略日志写入错误，避免影响主程序
        }

        // GetLogs: 读取当前日志文件内容
        public static string GetLogs()
        {
            try
            {
                // 检查日志文件是否存在
                if (File.Exists(GOIConfig.LogFile))
                    // File.ReadAllText: 读取文件全部内容
                    return File.ReadAllText(GOIConfig.LogFile, Encoding.UTF8);
            }
            catch { }
            return "暂无日志信息。";
        }

        // ShowLogViewer: 显示日志查看器窗口
        // 作用：创建一个独立的窗口来查看日志内容
        public static void ShowLogViewer()
        {
            // Form: Windows窗体类，用于创建窗口
            // new Form { ... }: 对象初始化器，设置窗体属性
            Form logForm = new Form
            {
                Text = "Glim Office Installer - 日志查看器",  // 窗口标题
                Size = new Size(800, 600),  // 窗口大小（宽x高）
                StartPosition = FormStartPosition.CenterScreen,  // 窗口居中显示
                BackColor = Color.White  // 背景色为白色
            };

            // 加载窗口图标
            // Assembly: 程序集类，用于获取嵌入资源
            try
            {
                Assembly assembly = Assembly.GetExecutingAssembly();
                // GetManifestResourceStream: 获取嵌入的资源流
                // "OfficeInstaller.icon.ico": 资源的完整名称
                using (Stream stream = assembly.GetManifestResourceStream("OfficeInstaller.icon.ico"))
                {
                    if (stream != null) logForm.Icon = new Icon(stream);
                }
            }
            catch { }

            // TextBox: 文本框控件，用于显示和编辑文本
            TextBox txtLogs = new TextBox
            {
                Multiline = true,  // 多行模式
                ScrollBars = ScrollBars.Both,  // 显示水平和垂直滚动条
                Dock = DockStyle.Fill,  // 填充整个父容器
                ReadOnly = true,  // 只读，不允许编辑
                Font = new Font("Consolas", 10),  // 使用等宽字体
                Text = Logger.GetLogs(),  // 设置文本内容为日志
                BackColor = Color.FromArgb(30, 30, 30),  // 深色背景
                ForeColor = Color.LightGray,  // 浅灰色文字
                BorderStyle = BorderStyle.None  // 无边框
            };

            // 自动滚动到底部
            // SelectionStart: 选区起始位置
            txtLogs.SelectionStart = txtLogs.Text.Length;
            // ScrollToCaret: 滚动到光标位置
            txtLogs.ScrollToCaret();

            // Panel: 面板控件，用于容纳其他控件
            // 底部控制面板
            Panel pnlBottom = new Panel
            {
                Dock = DockStyle.Bottom,  // 停靠在底部
                Height = 50,  // 高度50像素
                BackColor = Color.FromArgb(240, 240, 240)  // 浅灰色背景
            };

            // Button: 按钮控件
            // 刷新按钮
            Button btnRefresh = new Button
            {
                Text = "刷新日志",  // 按钮文字
                Location = new Point(580, 10),  // 位置（x, y）
                Size = new Size(100, 30),  // 大小（宽x高）
                FlatStyle = FlatStyle.System  // 使用系统样式
            };
            // Click事件：点击按钮时触发
            // +=: 订阅事件
            // (s, e) => { ... }: Lambda表达式，事件处理程序
            btnRefresh.Click += (s, e) => {
                txtLogs.Text = Logger.GetLogs();  // 重新加载日志
                txtLogs.SelectionStart = txtLogs.Text.Length;
                txtLogs.ScrollToCaret();
            };

            // 关闭按钮
            Button btnClose = new Button
            {
                Text = "关闭",
                Location = new Point(690, 10),
                Size = new Size(90, 30),
                FlatStyle = FlatStyle.System
            };
            btnClose.Click += (s, e) => logForm.Close();  // 点击关闭窗口

            // Controls.Add: 向容器添加控件
            pnlBottom.Controls.Add(btnRefresh);
            pnlBottom.Controls.Add(btnClose);
            logForm.Controls.Add(txtLogs);
            logForm.Controls.Add(pnlBottom);

            // ShowDialog: 显示模态对话框
            // 模态对话框：必须关闭才能操作其他窗口
            logForm.ShowDialog();
        }
    }

    // ========== 嵌入字体管理类 ==========
    // EmbeddedFont: 管理嵌入到程序中的自定义字体
    // 作用：让程序使用自定义字体，而不需要用户安装字体
    public static class EmbeddedFont
    {
        // DllImport: 调用Windows API函数
        // gdi32.dll: Windows图形设备接口库
        // AddFontMemResourceEx: 从内存加载字体
        [DllImport("gdi32.dll")]
        private static extern IntPtr AddFontMemResourceEx(IntPtr pbFont, uint cbFont, IntPtr pdv, [In] ref uint pcFonts);

        // PrivateFontCollection: 私有字体集合，存储加载的字体
        private static PrivateFontCollection privateFonts = null;
        // FontFamily: 字体家族，代表一种字体
        private static FontFamily customFontFamily = null;
        
        // GetFontFamily: 获取自定义字体家族
        public static FontFamily GetFontFamily()
        {
            // 如果已经加载过，直接返回
            if (customFontFamily != null)
                return customFontFamily;
            
            try
            {
                // Assembly: 获取当前程序集
                Assembly assembly = Assembly.GetExecutingAssembly();
                // 尝试不同的资源名称格式
                string resourceName = "OfficeInstaller.font.ttf";
                Stream stream = assembly.GetManifestResourceStream(resourceName);
                // 如果第一个资源名称找不到，尝试简化的名称
                if (stream == null)
                {
                    resourceName = "font.ttf";
                    stream = assembly.GetManifestResourceStream(resourceName);
                }

                // using: 自动释放资源，确保流被正确关闭
                using (stream)
                {
                    if (stream != null)
                    {
                        // PrivateFontCollection: 私有字体集合
                        // 用于存储从内存加载的字体
                        privateFonts = new PrivateFontCollection();
                        
                        // 读取字体数据到内存
                        // byte[]: 字节数组，存储二进制数据
                        // stream.Length: 流的长度（字节数）
                        byte[] fontData = new byte[stream.Length];
                        // stream.Read: 从流中读取数据到数组
                        // 参数：目标数组、起始位置、读取长度
                        stream.Read(fontData, 0, (int)stream.Length);
                        
                        // 分配非托管内存
                        // IntPtr: 指针类型，表示内存地址
                        // Marshal.AllocCoTaskMem: 分配COM任务内存
                        IntPtr fontPtr = Marshal.AllocCoTaskMem(fontData.Length);
                        // Marshal.Copy: 将托管数组复制到非托管内存
                        // 参数：源数组、源起始索引、目标内存地址、复制长度
                        Marshal.Copy(fontData, 0, fontPtr, fontData.Length);
                        
                        // 1. 添加到 PrivateFontCollection (GDI+)
                        // AddMemoryFont: 从内存添加字体
                        privateFonts.AddMemoryFont(fontPtr, fontData.Length);
                        
                        // 2. 注册到系统字体表 (GDI) - 这对 TextRenderer 很重要
                        // AddFontMemResourceEx: Windows API，注册内存字体
                        uint dummy = 0;
                        AddFontMemResourceEx(fontPtr, (uint)fontData.Length, IntPtr.Zero, ref dummy);

                        // 释放非托管内存
                        // FreeCoTaskMem: 释放COM任务内存
                        Marshal.FreeCoTaskMem(fontPtr);
                        
                        // Families: 获取加载的字体家族数组
                        if (privateFonts.Families.Length > 0)
                        {
                            customFontFamily = privateFonts.Families[0];
                            return customFontFamily;
                        }
                    }
                    else
                    {
                        // 仅在调试时启用，或者如果真的找不到资源
                        // MessageBox.Show("未找到字体资源: " + resourceName);
                    }
                }
            }
            catch
            {
                // MessageBox.Show("加载字体失败");
            }
            
            // 如果加载失败，返回微软雅黑作为备用字体
            // FontFamily: 字体家族类
            // GenericSansSerif: 通用无衬线字体
            try { return new FontFamily("Microsoft YaHei UI"); } catch { return FontFamily.GenericSansSerif; }
        }
        
        // ========== 全局UI缩放系数 ==========
        // Scale: 默认1.0，不再动态修改（改由AutoScaleMode.Dpi处理）
        public static float Scale = 1.0f;

        // GetFont: 获取指定大小的字体对象
        public static Font GetFont(float size, FontStyle style = FontStyle.Bold)
        {
            FontFamily family = GetFontFamily();
            return new Font(family, size, style);
        }
    }
    
    // ========== 美化的圆角按钮类 ==========
    // ModernButton: 自定义按钮控件，继承自Button
    // 特点：圆角、悬停效果、按下效果
    public class ModernButton : Button
    {
        // isHovered: 鼠标是否悬停在按钮上
        private bool isHovered = false;
        // isPressed: 鼠标是否按下
        private bool isPressed = false;

        // 构造函数：初始化按钮属性
        public ModernButton()
        {
            // FlatStyle.Flat: 扁平化样式
            this.FlatStyle = FlatStyle.Flat;
            // BorderSize = 0: 无边框
            this.FlatAppearance.BorderSize = 0;
            // 默认大小
            this.Size = new Size(200, 50);
            // Glim蓝色背景
            this.BackColor = Color.FromArgb(0, 122, 204);
            // 白色文字
            this.ForeColor = Color.White;
            // 手型光标
            this.Cursor = Cursors.Hand;
            // 使用嵌入字体
            this.Font = EmbeddedFont.GetFont(12, FontStyle.Bold);
            
            // 鼠标事件处理
            // MouseEnter: 鼠标进入控件
            this.MouseEnter += (s, e) => { isHovered = true; this.Invalidate(); };
            // MouseLeave: 鼠标离开控件
            this.MouseLeave += (s, e) => { isHovered = false; isPressed = false; this.Invalidate(); };
            // MouseDown: 鼠标按下
            this.MouseDown += (s, e) => { isPressed = true; this.Invalidate(); };
            // MouseUp: 鼠标释放
            this.MouseUp += (s, e) => { isPressed = false; this.Invalidate(); };
        }

        // OnPaint: 重写绘制方法，自定义按钮外观
        // PaintEventArgs: 包含绘制所需的Graphics对象
        protected override void OnPaint(PaintEventArgs pevent)
        {
            // Graphics: 绘图对象，用于绑制图形
            Graphics g = pevent.Graphics;
            // SmoothingMode.AntiAlias: 抗锯齿，使边缘更平滑
            g.SmoothingMode = SmoothingMode.AntiAlias;
            // Clear: 清除画布，使用父控件背景色填充
            g.Clear(this.Parent.BackColor);

            // 根据状态选择颜色
            Color bgColor;
            if (isPressed)
                bgColor = Color.FromArgb(0, 80, 150); // 按下时更深
            else if (isHovered)
                bgColor = Color.FromArgb(0, 140, 220); // 悬停时更亮
            else
                bgColor = this.BackColor; // 默认

            // 圆角半径
            int radius = 12;
            // 使用更大的边距避免黑边
            // Rectangle: 矩形结构，表示位置和大小
            Rectangle drawRect = new Rectangle(2, 2, this.Width - 5, this.Height - 5);
            // GraphicsPath: 图形路径，用于绑制复杂形状
            GraphicsPath path = new GraphicsPath();
            // AddArc: 添加圆弧到路径
            // 参数：x, y, width, height, startAngle, sweepAngle
            // 左上角圆弧：从180度开始，扫过90度
            path.AddArc(drawRect.X, drawRect.Y, radius * 2, radius * 2, 180, 90);
            // 右上角圆弧：从270度开始，扫过90度
            path.AddArc(drawRect.Right - radius * 2, drawRect.Y, radius * 2, radius * 2, 270, 90);
            // 右下角圆弧：从0度开始，扫过90度
            path.AddArc(drawRect.Right - radius * 2, drawRect.Bottom - radius * 2, radius * 2, radius * 2, 0, 90);
            // 左下角圆弧：从90度开始，扫过90度
            path.AddArc(drawRect.X, drawRect.Bottom - radius * 2, radius * 2, radius * 2, 90, 90);
            // CloseFigure: 闭合当前图形，连接起点和终点
            path.CloseFigure();

            // 绘制按钮背景
            // SolidBrush: 实心画刷，用于填充图形
            // using: 自动释放画刷资源
            using (SolidBrush brush = new SolidBrush(bgColor))
            {
                // FillPath: 使用画刷填充路径
                g.FillPath(brush, path);
            }

            // 绘制文字
            // 文字绘制区域：整个按钮区域
            Rectangle textRect = new Rectangle(0, 0, this.Width, this.Height);
            // TextRenderer.DrawText: 绘制文本（比Graphics.DrawString更好）
            // 参数：Graphics, 文本内容, 字体, 区域, 颜色, 格式标志
            // HorizontalCenter: 水平居中
            // VerticalCenter: 垂直居中
            TextRenderer.DrawText(g, this.Text, this.Font, textRect, this.ForeColor, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        }
    }

    // ========== 美化的复选框类 ==========
    // ModernCheckBox: 自定义复选框控件，继承自CheckBox
    public class ModernCheckBox : CheckBox
    {
        public ModernCheckBox()
        {
            // 使用嵌入字体
            this.Font = EmbeddedFont.GetFont(10.5f, FontStyle.Regular);
            // 深灰色文字
            this.ForeColor = Color.FromArgb(64, 64, 64);
            // 手型光标
            this.Cursor = Cursors.Hand;
            // 自动调整大小以适应内容
            this.AutoSize = true;
        }
    }

    // ========== 主窗体类 ==========
    // MainForm: 程序的主窗口，继承自Form
    // 这是用户交互的主要界面
    public class MainForm : Form
    {
        // ========== 私有字段 ==========
        // glimBlue: Glim品牌蓝色
        private Color glimBlue = Color.FromArgb(0, 122, 204);
        // officeOrange: Office品牌橙色
        private Color officeOrange = Color.FromArgb(232, 65, 37);
        // productChecks: 产品复选框字典，键是产品名称，值是CheckBox控件
        // Dictionary<K, V>: 泛型字典，存储键值对
        private Dictionary<string, CheckBox> productChecks = new Dictionary<string, CheckBox>();
        // btnInstall: 安装按钮
        private ModernButton btnInstall;
        // lblStatus: 状态标签，显示当前操作状态
        private Label lblStatus;
        // lblGroupTitle: 分组标题标签
        private Label lblGroupTitle;
        // lblWarn: 警告标签
        private Label lblWarn;
        // lblM365Info: Microsoft 365信息标签
        private Label lblM365Info;
        // lblSecretMenu: 隐藏菜单标签
        private Label lblSecretMenu;
        // productPanel: 产品列表面板
        private Panel productPanel;
        
        // 是否正在部署Office
        // 用于关闭窗口时的确认提示
        private bool isDeploying = false;

        // ========== 构造函数 ==========
        public MainForm()
        {
            // 窗口标题
            this.Text = "Glim Office Installer - Office 一键安装工具";

            // AutoScaleMode.Dpi: 这是Windows处理“相同物理屏幕、不同分辨率下等大”的正确机制
            // 它注册了基准DPI=96，当用户设置125%/150%等DPI时WinForms自动等比缩放所有控件
            this.AutoScaleMode = AutoScaleMode.Dpi;
            this.AutoScaleDimensions = new SizeF(96F, 96F);

            // 固定边框和尺寸（设计基准: 900宽 × 760高，内容却需约696px客户区高度）
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.ClientSize = new Size(900, 720);

            // 居中显示
            this.StartPosition = FormStartPosition.CenterScreen;
            // 背景色
            this.BackColor = Color.FromArgb(242, 242, 247);

            // 加载嵌入的图标资源
            try
            {
                Assembly assembly = Assembly.GetExecutingAssembly();
                using (Stream stream = assembly.GetManifestResourceStream("OfficeInstaller.icon.ico"))
                {
                    if (stream != null) this.Icon = new Icon(stream);
                    else this.Icon = SystemIcons.Application;
                }
            }
            catch { this.Icon = SystemIcons.Application; }

            // 初始化界面控件
            InitializeComponents();

            // 提醒用户不要删除 GOI 文件夹
            this.Shown += (s, e) => {
                Logger.Info("主窗口显示完成。");
                ShowToast("重要提示", "已创建 [GOI] 工作文件夹，请勿在安装完成前删除。", ToastType.Warning);
            };
        }
        
        // ShowToast: 显示右下角Toast通知
        // title: 通知标题
        // message: 通知内容
        // ToastType: 通知类型（Info/Success/Warning/Error）
        public void ShowToast(string title, string message, ToastType type = ToastType.Info)
        {
            try
            {
                // this.Invoke: 在UI线程上执行委托
                // 因为Toast可能在非UI线程中调用
                // MethodInvoker: 无参数无返回值的委托类型
                this.Invoke((MethodInvoker)delegate {
                    // ToastNotification: Toast通知类
                    ToastNotification toast = new ToastNotification(title, message, type);
                    toast.ShowToast();
                });
                Logger.Info("显示Toast通知: " + title + " - " + message);
            }
            catch (Exception ex)
            {
                Logger.Error("显示Toast通知失败", ex);
            }
        }

        // ========== 安装类型枚举 ==========
        // enum: 枚举类型，定义一组命名常量
        // InstallType: 安装类型枚举
        private enum InstallType { Office2024, Office2021, Office2019, Office2016, Microsoft365Pro, Microsoft365Home }
        // currentInstallType: 当前选择的安装类型
        private InstallType currentInstallType = InstallType.Office2024;
        
        // ========== Office 下载链接常量 ==========
        // const: 常量，值不可修改
        // Office 2024 专业增强版批量许可版下载链接
        private const string OFFICE_2024_URL = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?productReleaseID=ProPlus2024Volume&platform=X64&language=zh-cn";
        // Microsoft 365 商业版下载链接
        private const string M365_PRO_URL = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?productReleaseID=O365ProPlusRetail&platform=X64&language=zh-cn";
        // Microsoft 365 家庭版下载链接
        private const string M365_HOME_URL = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?productReleaseID=O365HomePremRetail&platform=X64&language=zh-cn";

        // InitializeComponents: 初始化界面组件
        // 这是创建所有UI控件的主要方法
        private void InitializeComponents()
        {
            // currentY: 当前Y坐标，用于垂直布局
            int currentY = 0;
            // centerX: 中心X坐标，用于水平居中
            int centerX = 450;

            // Banner比例 5.25:1，宽度900，高度约170
            int bannerWidth = 900;
            int bannerHeight = (int)(bannerWidth / 5.25);
            // Panel: 面板控件，用于容纳其他控件
            // header: 顶部横幅面板
            Panel header = new Panel { 
                Location = new Point(0, 0),  // 位置：左上角
                Size = new Size(bannerWidth, bannerHeight),  // 大小
                BackColor = glimBlue,  // 背景色
                Cursor = Cursors.Hand  // 手型光标
            };
            
            // 从嵌入资源加载banner图片
            // Image: 图像类，表示位图或矢量图
            Image bannerImage = null;
            try
            {
                Assembly assembly = Assembly.GetExecutingAssembly();
                // GetManifestResourceStream: 获取嵌入资源流
                using (Stream stream = assembly.GetManifestResourceStream("OfficeInstaller.banner.png"))
                {
                    if (stream != null)
                    {
                        // Image.FromStream: 从流创建图像
                        bannerImage = Image.FromStream(stream);
                    }
                }
            }
            catch { }
            
            if (bannerImage != null)
            {
                // Paint事件：控件需要重绘时触发
                // +=: 订阅事件
                header.Paint += (s, e) => {
                    // e.Graphics: 获取绑图对象
                    Graphics g = e.Graphics;
                    // DrawImage: 绘制图像
                    // 参数：图像、目标矩形
                    g.DrawImage(bannerImage, 0, 0, header.Width, header.Height);
                };
            }
            else
            {
                // 如果图片不存在，使用默认渐变背景
                header.Paint += (s, e) => {
                    Graphics g = e.Graphics;
                    // LinearGradientBrush: 线性渐变画刷
                    // 参数：矩形、起始颜色、结束颜色、角度
                    using (LinearGradientBrush brush = new LinearGradientBrush(header.ClientRectangle, 
                        Color.FromArgb(0, 122, 204), Color.FromArgb(0, 100, 170), 90f))
                    {
                        // FillRectangle: 填充矩形
                        g.FillRectangle(brush, header.ClientRectangle);
                    }
                };
            }
            
            // ========== 隐藏功能触发逻辑 ==========
            // Banner连续点击5次触发隐藏功能
            // bannerClickCount: 点击计数器
            int bannerClickCount = 0;
            // lastClickTime: 上次点击时间
            DateTime lastClickTime = DateTime.Now;
            // header.Click: 点击Banner的事件处理
            header.Click += (s, e) => {
                DateTime now = DateTime.Now;
                // TotalSeconds: 时间差的秒数
                // 如果超过2秒，重置计数器
                if ((now - lastClickTime).TotalSeconds > 2)
                {
                    // 超过2秒重置计数
                    bannerClickCount = 1;
                }
                else
                {
                    // 2秒内连续点击，计数加1
                    bannerClickCount++;
                }
                lastClickTime = now;
                
                // 点击5次触发隐藏窗口
                if (bannerClickCount >= 5)
                {
                    bannerClickCount = 0;
                    // ShowSecretWindow: 显示隐藏功能窗口
                    ShowSecretWindow();
                }
            };
            
            // Controls.Add: 向窗体添加控件
            this.Controls.Add(header);
            // 更新当前Y坐标，为下一个控件定位
            currentY = bannerHeight + 5;  // 减小间距

            // ========== 版本选择标题 ==========
            // Label: 标签控件，用于显示文本
            Label lblTypeTitle = new Label {
                Text = "选择要安装的版本",  // 标题文本
                ForeColor = Color.Black, // 改为纯黑色文字
                Location = new Point(0, currentY),  // 位置
                Size = new Size(900, 40),  // 大小
                AutoSize = false,  // 禁用自动大小
                TextAlign = ContentAlignment.MiddleCenter,  // 文字居中
                Font = EmbeddedFont.GetFont(15, FontStyle.Bold) // 稍微加大并保持加粗
            };
            this.Controls.Add(lblTypeTitle);
            // 更新Y坐标，向下移动35像素（原55）
            currentY += 35;

            // ========== 版本选择卡片容器 ==========
            // 创建一个容器来放置两个卡片，确保 RadioButton 互斥
            Panel cardsContainer = new Panel {
                Location = new Point(0, currentY),  // 位置
                Size = new Size(900, 150),  // 大小
                BackColor = this.BackColor  // 背景色与主窗体相同
            };
            this.Controls.Add(cardsContainer);

            // ========== 卡片尺寸计算 ==========
            // 两个选项卡片 - 水平排列，优化尺寸和间距
            int cardWidth = 320;  // 卡片宽度
            int cardHeight = 120;  // 卡片高度（原140）
            int cardSpacing = 40;  // 卡片间距
            // 计算起始X坐标，使两个卡片居中
            // centerX - (cardWidth * 2 + cardSpacing) / 2: 居中计算
            int startCardX = centerX - (cardWidth * 2 + cardSpacing) / 2;
            
            // 创建版本选择卡片 - 使用自定义绘制
            int cardRadius = 16;  // 卡片圆角半径
            int cardBorderWidth = 2;  // 卡片边框宽度
            
            // 当前显示的版本组：0=第一组(2024/M365), 1=第二组(2021/2019), 2=第三组(2016/M365家庭)
            int currentVersionGroup = 0;
            
            // ========== Office 2024 卡片 ==========
            Panel cardOffice2024 = new Panel {
                Location = new Point(startCardX, 10),  // 左侧卡片位置
                Size = new Size(cardWidth, cardHeight),
                BackColor = Color.Transparent,  // 透明背景，便于自定义绘制
                Cursor = Cursors.Hand  // 手型光标
            };
            
            // ========== Microsoft 365 专业版卡片 ==========
            Panel cardM365Pro = new Panel {
                Location = new Point(startCardX + cardWidth + cardSpacing, 10),  // 右侧卡片位置
                Size = new Size(cardWidth, cardHeight),
                BackColor = Color.Transparent,
                Cursor = Cursors.Hand
            };
            
            // ========== 版本组标签数组 ==========
            // 二维数组：3组版本，每组6个标签
            // Label[,]: 二维数组，第一个维度是组，第二个维度是标签索引
            Label[,] versionLabels = new Label[3, 6];
            // 版本文本二维数组
            string[,] versionTexts = new string[,] {
                { "Office 2024", "零售版", "最新功能 · 持续更新", "Microsoft 365", "个人/家庭版", "云端同步 · 订阅服务" },
                { "Office 2021", "零售版", "买断制 · 永久授权", "Office 2019", "零售版", "经典稳定 · 广泛兼容" },
                { "Office 2016", "零售版", "经典版本 · 兼容性好", "", "", "" }
            };
            
            // ========== 标签创建函数 ==========
            // Func<T1, T2, T3, T4, TResult>: 泛型委托，表示一个函数
            // 参数：文本、字体大小、Y位置、颜色，返回：Label对象
            // 这是一个局部函数，用于简化标签创建
            Func<string, int, int, Color, Label> createLabel = (text, fontSize, posY, color) => {
                Label lbl = new Label {
                    Text = text,  // 标签文本
                    // 字体大小>=14时使用粗体
                    Font = EmbeddedFont.GetFont(fontSize, fontSize >= 14 ? FontStyle.Bold : FontStyle.Regular),
                    ForeColor = color,  // 文字颜色
                    Location = new Point(60, posY),  // 位置
                    Size = new Size(cardWidth - 70, 32),  // 大小
                    AutoSize = false,  // 禁用自动大小
                    TextAlign = ContentAlignment.MiddleLeft,  // 文字左对齐
                    BackColor = Color.Transparent  // 透明背景
                };
                // Enabled = false: 禁用标签，使其不响应鼠标事件
                // 这样点击标签时，事件会传递到父控件（卡片）
                lbl.Enabled = false;
                return lbl;
            };

            // ========== 创建第一组标签（Office 2024 / M365 Pro）==========
            // versionLabels[0, 0]: 第一组左卡片的第一个标签（主标题）
            // createLabel: 使用之前定义的局部函数创建标签
            // 参数：文本、字体大小、Y位置、颜色
            versionLabels[0, 0] = createLabel(versionTexts[0, 0], 14, 28, Color.FromArgb(0, 90, 158));
            versionLabels[0, 1] = createLabel(versionTexts[0, 1], 11, 58, Color.FromArgb(60, 60, 60));
            versionLabels[0, 2] = createLabel(versionTexts[0, 2], 11, 86, Color.FromArgb(100, 100, 100));
            // 将标签添加到左卡片
            cardOffice2024.Controls.Add(versionLabels[0, 0]);
            cardOffice2024.Controls.Add(versionLabels[0, 1]);
            cardOffice2024.Controls.Add(versionLabels[0, 2]);

            // ========== 创建第二组标签（Office 2021 / Office 2019）初始隐藏 ==========
            versionLabels[1, 0] = createLabel(versionTexts[1, 0], 13, 22, Color.FromArgb(0, 90, 158));
            versionLabels[1, 1] = createLabel(versionTexts[1, 1], 10, 50, Color.FromArgb(60, 60, 60));
            versionLabels[1, 2] = createLabel(versionTexts[1, 2], 10, 74, Color.FromArgb(100, 100, 100));
            // Visible = false: 初始隐藏，点击切换按钮时显示
            versionLabels[1, 0].Visible = false;
            versionLabels[1, 1].Visible = false;
            versionLabels[1, 2].Visible = false;
            cardOffice2024.Controls.Add(versionLabels[1, 0]);
            cardOffice2024.Controls.Add(versionLabels[1, 1]);
            cardOffice2024.Controls.Add(versionLabels[1, 2]);

            // ========== M365 Pro 标签（第一组）==========
            // versionLabels[0, 3]: 第一组右卡片的第一个标签
            versionLabels[0, 3] = createLabel(versionTexts[0, 3], 14, 28, Color.FromArgb(0, 90, 158));
            versionLabels[0, 4] = createLabel(versionTexts[0, 4], 11, 58, Color.FromArgb(60, 60, 60));
            versionLabels[0, 5] = createLabel(versionTexts[0, 5], 11, 86, Color.FromArgb(100, 100, 100));
            cardM365Pro.Controls.Add(versionLabels[0, 3]);
            cardM365Pro.Controls.Add(versionLabels[0, 4]);
            cardM365Pro.Controls.Add(versionLabels[0, 5]);

            // ========== 第二组标签（初始隐藏）==========
            versionLabels[1, 3] = createLabel(versionTexts[1, 3], 14, 28, Color.FromArgb(0, 90, 158));
            versionLabels[1, 4] = createLabel(versionTexts[1, 4], 11, 58, Color.FromArgb(60, 60, 60));
            versionLabels[1, 5] = createLabel(versionTexts[1, 5], 11, 86, Color.FromArgb(100, 100, 100));
            versionLabels[1, 3].Visible = false;
            versionLabels[1, 4].Visible = false;
            versionLabels[1, 5].Visible = false;
            cardM365Pro.Controls.Add(versionLabels[1, 3]);
            cardM365Pro.Controls.Add(versionLabels[1, 4]);
            cardM365Pro.Controls.Add(versionLabels[1, 5]);

            // ========== 第三组标签（Office 2016）初始隐藏 ==========
            versionLabels[2, 0] = createLabel(versionTexts[2, 0], 14, 28, Color.FromArgb(0, 90, 158));
            versionLabels[2, 1] = createLabel(versionTexts[2, 1], 11, 58, Color.FromArgb(60, 60, 60));
            versionLabels[2, 2] = createLabel(versionTexts[2, 2], 11, 86, Color.FromArgb(100, 100, 100));
            versionLabels[2, 0].Visible = false;
            versionLabels[2, 1].Visible = false;
            versionLabels[2, 2].Visible = false;
            cardOffice2024.Controls.Add(versionLabels[2, 0]);
            cardOffice2024.Controls.Add(versionLabels[2, 1]);
            cardOffice2024.Controls.Add(versionLabels[2, 2]);

            // ========== 第三组右卡片标签（初始隐藏）==========
            // 使用橙色突出显示（officeOrange）
            versionLabels[2, 3] = createLabel(versionTexts[2, 3], 14, 28, Color.FromArgb(232, 65, 37));
            versionLabels[2, 4] = createLabel(versionTexts[2, 4], 11, 58, Color.FromArgb(60, 60, 60));
            versionLabels[2, 5] = createLabel(versionTexts[2, 5], 11, 86, Color.FromArgb(100, 100, 100));
            versionLabels[2, 3].Visible = false;
            versionLabels[2, 4].Visible = false;
            versionLabels[2, 5].Visible = false;
            cardM365Pro.Controls.Add(versionLabels[2, 3]);
            cardM365Pro.Controls.Add(versionLabels[2, 4]);
            cardM365Pro.Controls.Add(versionLabels[2, 5]);

            // ========== 绘制卡片方法 ==========
            // Action<T>: 泛型委托，表示一个无返回值的方法
            // setupCardPaint: 设置卡片的Paint事件处理
            // 使用Tag存储选中状态
            Action<Panel> setupCardPaint = (card) => {
                // Paint事件：控件需要重绘时触发
                card.Paint += (s, e) => {
                    // Graphics: 绑图对象
                    Graphics g = e.Graphics;
                    // SmoothingMode.AntiAlias: 抗锯齿模式
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    
                    // Tag: 控件的标签属性，可用于存储任意数据
                    // 这里存储的是bool类型，表示是否被选中
                    bool isSelected = (bool)card.Tag;
                    
                    // 边距计算
                    int padding = cardBorderWidth + 3;
                    
                    // 绘制区域矩形
                    // Rectangle: 矩形结构，表示位置和大小
                    Rectangle rect = new Rectangle(
                        padding, 
                        padding, 
                        card.Width - padding * 2 - 1, 
                        card.Height - padding * 2 - 1
                    );
                    
                    // ========== 绘制阴影（未选中时）==========
                    if (!isSelected)
                    {
                        // 阴影矩形：向右下偏移2像素
                        Rectangle shadowRect = new Rectangle(rect.X + 2, rect.Y + 2, rect.Width, rect.Height);
                        // CreateRoundedPath: 创建圆角路径（自定义方法）
                        using (GraphicsPath shadowPath = CreateRoundedPath(shadowRect.X, shadowRect.Y, shadowRect.Width, shadowRect.Height, cardRadius))
                        {
                            // 半透明黑色阴影（透明度30）
                            using (SolidBrush shadowBrush = new SolidBrush(Color.FromArgb(30, 0, 0, 0)))
                            {
                                // FillPath: 填充路径
                                g.FillPath(shadowBrush, shadowPath);
                            }
                        }
                    }
                    
                    // ========== 填充背景 ==========
                    // 选中时浅蓝色背景，未选中时白色背景
                    Color bgColor = isSelected ? Color.FromArgb(240, 248, 255) : Color.White;
                    using (GraphicsPath path = CreateRoundedPath(rect.X, rect.Y, rect.Width, rect.Height, cardRadius))
                    {
                        using (SolidBrush brush = new SolidBrush(bgColor))
                        {
                            g.FillPath(brush, path);
                        }
                    }
                    
                    // ========== 绘制边框 - 选中时使用渐变效果 ==========
                    if (isSelected)
                    {
                        // 选中状态：蓝色渐变边框效果
                        // CreateRoundedPath: 创建圆角路径
                        using (GraphicsPath path = CreateRoundedPath(rect.X, rect.Y, rect.Width, rect.Height, cardRadius))
                        {
                            // Pen: 画笔类，用于绘制线条
                            // 参数：颜色、宽度
                            using (Pen pen = new Pen(glimBlue, cardBorderWidth))
                            {
                                // DrawPath: 绘制路径轮廓
                                g.DrawPath(pen, path);
                            }
                        }
                        // 内发光效果：在内部绘制一个半透明边框
                        Rectangle innerRect = new Rectangle(rect.X + 2, rect.Y + 2, rect.Width - 4, rect.Height - 4);
                        using (GraphicsPath innerPath = CreateRoundedPath(innerRect.X, innerRect.Y, innerRect.Width, innerRect.Height, cardRadius - 2))
                        {
                            // Color.FromArgb(100, glimBlue): 半透明蓝色
                            using (Pen innerPen = new Pen(Color.FromArgb(100, glimBlue), 1))
                            {
                                g.DrawPath(innerPen, innerPath);
                            }
                        }
                    }
                    else
                    {
                        // 未选中状态：灰色边框
                        using (GraphicsPath path = CreateRoundedPath(rect.X, rect.Y, rect.Width, rect.Height, cardRadius))
                        {
                            // 浅灰色边框
                            using (Pen pen = new Pen(Color.FromArgb(200, 200, 200), cardBorderWidth))
                            {
                                g.DrawPath(pen, path);
                            }
                        }
                    }
                    
                    // ========== 绘制单选按钮指示器（现代风格）==========
                    // 指示器位置和大小
                    int indicatorX = 20;  // X坐标：左侧20像素
                    int indicatorY = card.Height / 2 - 12;  // Y坐标：垂直居中
                    int indicatorSize = 24;  // 大小：24x24像素
                    Rectangle indicatorRect = new Rectangle(indicatorX, indicatorY, indicatorSize, indicatorSize);
                    
                    // 外圈背景
                    // 选中时白色，未选中时浅灰色
                    using (SolidBrush bgBrush = new SolidBrush(isSelected ? Color.White : Color.FromArgb(245, 245, 245)))
                    {
                        // FillEllipse: 填充椭圆（圆形）
                        g.FillEllipse(bgBrush, indicatorRect);
                    }
                    
                    // 外圈边框
                    // 选中时蓝色，未选中时灰色
                    using (Pen pen = new Pen(isSelected ? glimBlue : Color.FromArgb(180, 180, 180), 2))
                    {
                        // DrawEllipse: 绘制椭圆轮廓
                        g.DrawEllipse(pen, indicatorRect);
                    }
                    
                    // 内圈（选中时填充）
                    if (isSelected)
                    {
                        int innerSize = 14;  // 内圈大小
                        // 计算内圈位置，使其居中于外圈
                        int innerX = indicatorX + (indicatorSize - innerSize) / 2;
                        int innerY = indicatorY + (indicatorSize - innerSize) / 2;
                        // 填充蓝色内圈
                        using (SolidBrush brush = new SolidBrush(glimBlue))
                        {
                            g.FillEllipse(brush, innerX, innerY, innerSize, innerSize);
                        }
                    }
                };
            };

            // ========== 设置初始选中状态 ==========
            // Tag: 用于存储选中状态
            // true: 选中，false: 未选中
            cardOffice2024.Tag = true;  // 左卡片默认选中
            cardM365Pro.Tag = false;  // 右卡片默认未选中
            
            // 调用绘制方法设置Paint事件
            setupCardPaint(cardOffice2024);
            setupCardPaint(cardM365Pro);

            // ========== 刷新卡片显示 ==========
            // Action: 无参数无返回值的委托
            // Invalidate: 使控件无效，触发重绘
            Action refreshCards = () => {
                cardOffice2024.Invalidate();
                cardM365Pro.Invalidate();
            };

            // ========== 切换版本组的方法 ==========
            // Action<int>: 带一个int参数的无返回值委托
            // groupIndex: 版本组索引（0/1/2）
            Action<int> switchVersionGroup = (groupIndex) => {
                // 隐藏当前组的所有标签
                // for循环：遍历当前组的6个标签
                for (int i = 0; i < 3; i++)
                {
                    // 左卡片标签（索引0-2）
                    versionLabels[currentVersionGroup, i].Visible = false;
                    // 右卡片标签（索引3-5）
                    versionLabels[currentVersionGroup, i + 3].Visible = false;
                }
                
                // 显示新组的标签
                currentVersionGroup = groupIndex;
                for (int i = 0; i < 3; i++)
                {
                    versionLabels[currentVersionGroup, i].Visible = true;
                    versionLabels[currentVersionGroup, i + 3].Visible = true;
                }
                
                    // 第三组时隐藏右卡片（M365Pro/Office2019位置）
                    // 第三组只有一个选项（Office 2016）
                    if (currentVersionGroup == 2)
                    {
                        // 隐藏右卡片
                        cardM365Pro.Visible = false;
                        // 设置安装类型为Office 2016
                        currentInstallType = InstallType.Office2016;
                    }
                    else
                    {
                        // 其他组显示右卡片
                        cardM365Pro.Visible = true;
                    }
                    
                    // 重置选中状态：默认选中左卡片
                    if (currentVersionGroup == 0)
                    {
                        // 第一组：Office 2024
                        currentInstallType = InstallType.Office2024;
                    }
                    else if (currentVersionGroup == 1)
                    {
                        // 第二组：Office 2021
                        currentInstallType = InstallType.Office2021;
                    }
                    
                    // 设置选中状态
                    cardOffice2024.Tag = true;  // 左卡片选中
                    cardM365Pro.Tag = false;  // 右卡片未选中
                    // 刷新卡片显示
                    refreshCards();
                    
                    // 记录日志
                    Logger.Info("切换到版本组: " + (currentVersionGroup + 1));
            };

            // ========== 点击事件处理 ==========
            // selectLeftCard: 选择左卡片的方法
            Action selectLeftCard = () => {
                // 第三组左卡片是Office 2016
                if (currentVersionGroup == 2)
                {
                    currentInstallType = InstallType.Office2016;
                    cardOffice2024.Tag = true;
                    cardM365Pro.Tag = false;
                    // 显示产品面板和相关标签
                    productPanel.Visible = true;
                    lblGroupTitle.Visible = true;
                    lblWarn.Visible = true;
                    // 隐藏M365信息标签
                    if (lblM365Info != null) lblM365Info.Visible = false;
                    refreshCards();
                }
                else
                {
                    // 其他组：设置选中状态
                    cardOffice2024.Tag = true;
                    cardM365Pro.Tag = false;
                    productPanel.Visible = true;
                    lblGroupTitle.Visible = true;
                    lblWarn.Visible = true;
                    if (lblM365Info != null) lblM365Info.Visible = false;
                    refreshCards();
                    
                    // 根据当前组设置正确的InstallType
                    if (currentVersionGroup == 0) currentInstallType = InstallType.Office2024;
                    else if (currentVersionGroup == 1) currentInstallType = InstallType.Office2021;
                }
            };

            // selectRightCard: 选择右卡片的方法
            Action selectRightCard = () => {
                // 设置选中状态
                cardOffice2024.Tag = false;
                cardM365Pro.Tag = true;
                productPanel.Visible = true;
                lblGroupTitle.Visible = true;
                lblWarn.Visible = true;
                if (lblM365Info != null) lblM365Info.Visible = false;
                refreshCards();
                
                // 根据当前组设置正确的InstallType
                if (currentVersionGroup == 0) currentInstallType = InstallType.Microsoft365Pro;
                else if (currentVersionGroup == 1) currentInstallType = InstallType.Office2019;
            };

            // ========== 绑定点击事件 ==========
            // Click事件：点击卡片时触发
            cardOffice2024.Click += (s, e) => selectLeftCard();
            cardM365Pro.Click += (s, e) => selectRightCard();
            
            // 为所有标签绑定点击事件
            // 双重for循环：遍历所有标签
            for (int g = 0; g < 3; g++)
            {
                // 闭包问题：需要捕获变量
                // 如果直接使用g，闭包会捕获引用而不是值
                int group = g; // 捕获变量，避免闭包问题
                for (int i = 0; i < 3; i++)
                {
                    // 左卡片标签点击事件
                    versionLabels[group, i].Click += (s, e) => selectLeftCard();
                    // 右卡片标签点击事件
                    versionLabels[group, i + 3].Click += (s, e) => selectRightCard();
                }
            }

            // 将卡片添加到容器
            cardsContainer.Controls.Add(cardOffice2024);
            cardsContainer.Controls.Add(cardM365Pro);
            
            // ========== 添加左右切换按钮 ==========
            // 1:2比例，浅灰色背景，浅蓝色箭头
            int btnWidth = 28;  // 宽度（增大）
            int btnHeight = 56; // 高度 (1:2比例，增大)
            // 计算按钮Y坐标，使其垂直居中于卡片
            int switchBtnY = 10 + (cardHeight - btnHeight) / 2;
            
            // 左切换按钮位置（在左卡片左侧）
            int leftBtnX = startCardX - btnWidth - 10;
            // 右切换按钮位置（在右卡片右侧）
            int rightBtnX = startCardX + cardWidth * 2 + cardSpacing + 10;
            
            // ========== 创建浅灰色圆角长方形切换按钮绘制方法 ==========
            // Action<Panel, bool>: 带Panel和bool参数的委托
            // panel: 要绑制的面板
            // isRight: 是否是右箭头（false=左箭头）
            Action<Panel, bool> drawSwitchButton = (panel, isRight) => {
                // Paint事件：面板需要重绘时触发
                panel.Paint += (s, e) => {
                    // Graphics: 绑图对象
                    Graphics g = e.Graphics;
                    // 抗锯齿模式
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    
                    // 判断按钮是否可用
                    // 右按钮：当前组<2时可用（可以前进）
                    // 左按钮：当前组>0时可用（可以后退）
                    bool isEnabled = isRight ? (currentVersionGroup < 2) : (currentVersionGroup > 0);
                    // 根据可用状态设置颜色
                    // 可用：较亮的灰色，不可用：较暗的灰色
                    Color btnColor = isEnabled ? Color.FromArgb(230, 230, 230) : Color.FromArgb(200, 200, 200);
                    
                    // 圆角半径
                    int radius = 3;
                    // 绘制区域矩形
                    Rectangle rect = new Rectangle(1, 1, panel.Width - 3, panel.Height - 3);
                    
                    // 创建圆角矩形路径
                    using (GraphicsPath path = new GraphicsPath())
                    {
                        // 绘制圆角矩形 - 四个角的圆弧
                        // 左上角圆弧
                        path.AddArc(rect.X, rect.Y, radius * 2, radius * 2, 180, 90);
                        // 右上角圆弧
                        path.AddArc(rect.Right - radius * 2, rect.Y, radius * 2, radius * 2, 270, 90);
                        // 右下角圆弧
                        path.AddArc(rect.Right - radius * 2, rect.Bottom - radius * 2, radius * 2, radius * 2, 0, 90);
                        // 左下角圆弧
                        path.AddArc(rect.X, rect.Bottom - radius * 2, radius * 2, radius * 2, 90, 90);
                        // 闭合路径
                        path.CloseFigure();
                        
                        // 填充浅灰色背景
                        using (SolidBrush brush = new SolidBrush(btnColor))
                        {
                            g.FillPath(brush, path);
                        }
                    }
                    
                    // ========== 绘制方向三角形 - 浅蓝色，更明显 ==========
                    int triangleSize = 7; // 三角形大小
                    // 三角形中心坐标
                    int triCenterX = rect.X + rect.Width / 2;
                    int triCenterY = rect.Y + rect.Height / 2;
                    // 箭头颜色：可用时浅蓝色，不可用时灰色
                    Color arrowColor = isEnabled ? Color.FromArgb(80, 160, 210) : Color.FromArgb(180, 180, 180);
                    
                    // Point[]: 点数组，用于定义多边形顶点
                    Point[] trianglePoints;
                    if (isRight)
                    {
                        // 右向三角形：三个顶点
                        // 左顶点、右顶点（尖端）、左下顶点
                        trianglePoints = new Point[] {
                            new Point(triCenterX - triangleSize/2 + 1, triCenterY - triangleSize),
                            new Point(triCenterX + triangleSize + 1, triCenterY),
                            new Point(triCenterX - triangleSize/2 + 1, triCenterY + triangleSize)
                        };
                    }
                    else
                    {
                        // 左向三角形：三个顶点
                        // 右顶点、左顶点（尖端）、右下顶点
                        trianglePoints = new Point[] {
                            new Point(triCenterX + triangleSize/2 - 1, triCenterY - triangleSize),
                            new Point(triCenterX - triangleSize - 1, triCenterY),
                            new Point(triCenterX + triangleSize/2 - 1, triCenterY + triangleSize)
                        };
                    }
                    
                    // FillPolygon: 填充多边形
                    using (SolidBrush brush = new SolidBrush(arrowColor))
                    {
                        g.FillPolygon(brush, trianglePoints);
                    }
                };
            };
            
            // ========== 创建左切换按钮 ==========
            // 只在可用时显示（当前组>0时）
            Panel btnLeft = new Panel {
                Location = new Point(leftBtnX, switchBtnY),  // 位置
                Size = new Size(btnWidth, btnHeight),  // 大小
                BackColor = Color.Transparent,  // 透明背景
                Cursor = Cursors.Hand,  // 手型光标
                Visible = currentVersionGroup > 0 // 只在非第一组时显示
            };
            // 调用绘制方法（false表示左箭头）
            drawSwitchButton(btnLeft, false);
            
            // ========== 创建右切换按钮 ==========
            // 只在可用时显示（当前组<2时）
            Panel btnRight = new Panel {
                Location = new Point(rightBtnX, switchBtnY),
                Size = new Size(btnWidth, btnHeight),
                BackColor = Color.Transparent,
                Cursor = Cursors.Hand,
                Visible = currentVersionGroup < 2 // 只在非最后一组时显示
            };
            // 调用绘制方法（true表示右箭头）
            drawSwitchButton(btnRight, true);
            
            // ========== 左按钮点击事件 ==========
            // 返回上一组
            btnLeft.Click += (s, e) => {
                // 检查是否可以后退
                if (currentVersionGroup > 0)
                {
                    // 切换到上一组
                    switchVersionGroup(currentVersionGroup - 1);
                    // 更新按钮可见性
                    btnLeft.Visible = (currentVersionGroup > 0);
                    btnRight.Visible = (currentVersionGroup < 2);
                    // 触发重绘
                    btnLeft.Invalidate();
                    btnRight.Invalidate();
                }
            };
            
            // ========== 右按钮点击事件 ==========
            // 前进下一组（不循环）
            btnRight.Click += (s, e) => {
                // 检查是否可以前进（有三组，最大索引为2）
                if (currentVersionGroup < 2)
                {
                    // 切换到下一组
                    switchVersionGroup(currentVersionGroup + 1);
                    // 更新按钮可见性
                    btnLeft.Visible = (currentVersionGroup > 0);
                    btnRight.Visible = (currentVersionGroup < 2);
                    // 触发重绘
                    btnLeft.Invalidate();
                    btnRight.Invalidate();
                }
            };
            
            // 将按钮添加到容器
            cardsContainer.Controls.Add(btnLeft);
            cardsContainer.Controls.Add(btnRight);
            
            // 更新Y坐标，为下一个控件定位
            // 120是卡片容器高度，10是额外间距，10是底部间距
            currentY += 120 + 10 + 10;

            // ========== 产品选择标题 ==========
            // lblGroupTitle: 产品选择分组标题
            lblGroupTitle = new Label { 
                Text = "产品选择",  // 标题文本
                ForeColor = glimBlue,  // Glim蓝色文字
                Location = new Point(0, currentY),  // 位置
                Size = new Size(900, 35),  // 大小
                AutoSize = false,  // 禁用自动大小
                TextAlign = ContentAlignment.MiddleCenter,  // 文字居中
                Font = EmbeddedFont.GetFont(12, FontStyle.Bold),  // 粗体字体
                BackColor = Color.Transparent,  // 透明背景
                Cursor = Cursors.Hand  // 手型光标（可点击）
            };
            // 点击标题显示版权信息窗口
            lblGroupTitle.Click += (s, e) => {
                // CopyrightWindow: 版权信息窗口类
                CopyrightWindow copyrightWindow = new CopyrightWindow();
                // ShowDialog: 显示模态对话框
                copyrightWindow.ShowDialog(this);
            };
            // 将标题添加到窗体
            this.Controls.Add(lblGroupTitle);
            // 更新Y坐标
            currentY += 30;

            // ========== 产品列表面板 ==========
            // productPanel: 包含产品复选框的面板
            productPanel = new Panel {
                Location = new Point(centerX - 320, currentY),  // 居中位置
                Size = new Size(640, 180),  // 大小（原230）
                BackColor = Color.Transparent  // 透明背景
            };
            // Paint事件：绘制面板边框
            productPanel.Paint += (s, e) => {
                // Graphics: 绑图对象
                Graphics g = e.Graphics;
                // 抗锯齿模式
                g.SmoothingMode = SmoothingMode.AntiAlias;
                
                // 边框宽度
                int panelBorderWidth = 1;
                // 圆角半径
                int panelRadius = 15;
                
                // 关键：Pen的宽度会向外扩展，需要留出足够的边距
                // halfBorder: 边框宽度的一半加1，用于计算绘制区域
                int halfBorder = panelBorderWidth / 2 + 1;
                
                // 绘制区域 - 确保边框完全在控件内部
                // Rectangle: 矩形结构，表示绘制区域
                Rectangle rect = new Rectangle(
                    halfBorder,  // X坐标
                    halfBorder,  // Y坐标
                    productPanel.Width - halfBorder * 2,  // 宽度
                    productPanel.Height - halfBorder * 2  // 高度
                );
                
                // 先填充白色背景
                // CreateRoundedPath: 创建圆角路径（自定义方法）
                using (GraphicsPath path = CreateRoundedPath(rect.X, rect.Y, rect.Width, rect.Height, panelRadius))
                {
                    // SolidBrush: 实心画刷
                    using (SolidBrush brush = new SolidBrush(Color.White))
                    {
                        // FillPath: 填充路径
                        g.FillPath(brush, path);
                    }
                }
                
                // 再绘制边框
                using (GraphicsPath path = CreateRoundedPath(rect.X, rect.Y, rect.Width, rect.Height, panelRadius))
                {
                    // Pen: 画笔，用于绘制线条
                    // Color.FromArgb(200, 200, 200): 浅灰色
                    using (Pen pen = new Pen(Color.FromArgb(200, 200, 200), panelBorderWidth))
                    {
                        // DrawPath: 绘制路径轮廓
                        g.DrawPath(pen, path);
                    }
                }
            };

            // ========== 产品列表 ==========
            // string[]: 字符串数组，存储产品名称
            string[] products = { 
                "PowerPoint", "Word", "Excel",  // 第一行
                "Visio", "Access", "OneNote",   // 第二行
                "Lync", "Outlook", "Teams",     // 第三行
                "OneDrive", "Publisher", "Project"  // 第四行
            };

            // ========== 布局参数 ==========
            // 三列布局，每列宽度固定，确保对齐
            int colWidth = 190;  // 列宽
            int rowHeight = 35;  // 行高（原40）
            int startX = 60;     // 起始X坐标
            int startY = 25;     // 起始Y坐标（原45）
            
            // ========== 创建产品复选框 ==========
            // for循环：遍历产品数组
            for (int i = 0; i < products.Length; i++)
            {
                // 获取当前产品名称
                string p = products[i];
                // 计算列索引（0, 1, 2）
                // %: 取模运算符，获取余数
                int col = i % 3;
                // 计算行索引（0, 1, 2, 3）
                // /: 整数除法
                int row = i / 3;
                
                // 创建复选框
                // var: 隐式类型，编译器自动推断类型
                var cb = new ModernCheckBox { 
                    Text = p,  // 复选框文本
                    // 计算位置：起始位置 + 列偏移 + 行偏移
                    Location = new Point(startX + col * colWidth, startY + row * rowHeight),
                    Font = EmbeddedFont.GetFont(11, FontStyle.Regular),  // 常规字体
                    Size = new Size(170, 30)  // 大小
                };
                // 默认选中Word、Excel、PowerPoint
                // ||: 逻辑或运算符
                if (p == "Word" || p == "Excel" || p == "PowerPoint") cb.Checked = true;
                // 将复选框添加到面板
                productPanel.Controls.Add(cb);
                // productChecks: 字典，存储产品名称和对应的复选框
                // Add: 向字典添加键值对
                productChecks.Add(p, cb);
            }
            // 将产品面板添加到窗体
            this.Controls.Add(productPanel);
            // 更新Y坐标：180是面板高度，15是额外间距
            currentY += 180 + 15;

            // ========== Office 2024 的提示信息 ==========
            // lblWarn: 警告标签，提示用户选择产品
            // 居中显示
            lblWarn = new Label {
                Text = "请选择所有需要的产品！程序将卸载所有已安装的产品，并安装勾选的产品！",  // 提示文本
                ForeColor = Color.FromArgb(100, 100, 100),  // 深灰色文字
                Font = EmbeddedFont.GetFont(10, FontStyle.Regular),  // 常规字体
                AutoSize = false,  // 禁用自动大小
                Size = new Size(900, 25),  // 大小
                TextAlign = ContentAlignment.MiddleCenter,  // 文字居中
                Location = new Point(0, currentY)  // 位置
            };
            // 将警告标签添加到窗体
            this.Controls.Add(lblWarn);

            // ========== Microsoft 365 的提示信息 ==========
            // lblM365Info: M365信息标签（默认隐藏）
            // 居中显示
            lblM365Info = new Label {
                Text = "Microsoft 365 包含完整的 Office 组件套装，无需单独选择",  // 提示文本
                ForeColor = Color.FromArgb(100, 100, 100),  // 深灰色文字
                Font = EmbeddedFont.GetFont(11, FontStyle.Regular),  // 常规字体
                AutoSize = false,  // 禁用自动大小
                Size = new Size(900, 25),  // 大小
                TextAlign = ContentAlignment.MiddleCenter,  // 文字居中
                Location = new Point(0, currentY),  // 位置
                Visible = false  // 默认隐藏
            };
            // 将M365信息标签添加到窗体
            this.Controls.Add(lblM365Info);
            
            // 更新Y坐标
            currentY += 30;

            // ========== 安装按钮 ==========
            // btnInstall: 一键安装按钮
            // 居中显示
            btnInstall = new ModernButton { 
                Text = "一键安装",  // 按钮文本
                Location = new Point(centerX - 75, currentY),  // 居中位置（宽度150，偏移75）
                Size = new Size(150, 55),  // 大小
                Font = EmbeddedFont.GetFont(14, FontStyle.Bold)  // 粗体字体
            };
            // Click事件：点击按钮时触发
            // +=: 订阅事件
            // BtnInstall_Click: 事件处理方法
            btnInstall.Click += BtnInstall_Click;
            // 将安装按钮添加到窗体
            this.Controls.Add(btnInstall);
            // 更新Y坐标
            currentY += 60;

            // ========== 状态标签 ==========
            // lblStatus: 显示当前状态的标签
            lblStatus = new Label {
                Text = "建议保持网络连接，以便完成激活过程",  // 提示文本
                TextAlign = ContentAlignment.MiddleCenter,  // 文字居中
                ForeColor = Color.FromArgb(150, 150, 150),  // 浅灰色文字
                Font = EmbeddedFont.GetFont(10, FontStyle.Regular),  // 常规字体
                AutoSize = false,  // 禁用自动大小
                Size = new Size(900, 30),  // 大小
                Location = new Point(0, currentY)  // 位置
            };
            // 将状态标签添加到窗体
            this.Controls.Add(lblStatus);

            // ========== 隐藏功能菜单入口 ==========
            // lblSecretMenu: 隐藏菜单标签
            // 透明且小巧，放在左下角避免显示问题
            lblSecretMenu = new Label {
                Text = "隐藏功能菜单",  // 标签文本
                ForeColor = Color.FromArgb(240, 240, 240), // 更浅的灰色，几乎看不见
                Font = new Font("Microsoft YaHei UI", 6f), // 更小的字体（6号）
                AutoSize = true,  // 自动大小
                Cursor = Cursors.Hand,  // 手型光标
                Location = new Point(5, currentY + 5), // 放在左下角，使用固定位置
                BackColor = Color.Transparent  // 透明背景
            };
            
            // ========== 隐藏菜单点击计数 ==========
            // secretClickCount: 点击计数器
            int secretClickCount = 0;
            // lastSecretClickTime: 上次点击时间
            DateTime lastSecretClickTime = DateTime.Now;
            
            // Click事件：连续点击5次显示日志查看器
            lblSecretMenu.Click += (s, e) => {
                // 获取当前时间
                DateTime now = DateTime.Now;
                // 判断是否超过2秒
                // TotalSeconds: 时间差的秒数
                if ((now - lastSecretClickTime).TotalSeconds > 2)
                    // 超过2秒重置计数
                    secretClickCount = 1;
                else
                    // 2秒内连续点击，计数加1
                    secretClickCount++;
                
                // 更新上次点击时间
                lastSecretClickTime = now;
                
                // 点击5次触发日志查看器
                if (secretClickCount >= 5)
                {
                    // 重置计数器
                    secretClickCount = 0;
                    // 显示日志查看器
                    Logger.ShowLogViewer();
                }
            };
            // 将隐藏菜单标签添加到窗体
            this.Controls.Add(lblSecretMenu);
        }

        // ========== 创建圆角路径方法 ==========
        // CreateRoundedPath: 创建圆角矩形路径
        // 参数：x, y - 左上角坐标；width, height - 宽高；radius - 圆角半径
        // 返回：GraphicsPath - 圆角矩形路径
        private GraphicsPath CreateRoundedPath(int x, int y, int width, int height, int radius)
        {
            // GraphicsPath: 图形路径类，用于定义复杂形状
            GraphicsPath path = new GraphicsPath();
            // AddArc: 添加圆弧
            // 参数：x, y, width, height - 圆弧的边界矩形
            // startAngle - 起始角度，sweepAngle - 扫过角度
            // 左上角圆弧：从180度开始，扫过90度
            path.AddArc(x, y, radius * 2, radius * 2, 180, 90);
            // 右上角圆弧：从270度开始，扫过90度
            path.AddArc(x + width - radius * 2, y, radius * 2, radius * 2, 270, 90);
            // 右下角圆弧：从0度开始，扫过90度
            path.AddArc(x + width - radius * 2, y + height - radius * 2, radius * 2, radius * 2, 0, 90);
            // 左下角圆弧：从90度开始，扫过90度
            path.AddArc(x, y + height - radius * 2, radius * 2, radius * 2, 90, 90);
            // CloseFigure: 闭合当前图形，连接起点和终点
            path.CloseFigure();
            // 返回创建的路径
            return path;
        }

        // ========== 安装按钮点击事件 ==========
        // async: 异步方法关键字，允许在方法内使用await
        // await: 等待异步操作完成，但不会阻塞UI线程
        // BtnInstall_Click: 安装按钮的点击事件处理方法
        // sender: 事件发送者（这里是btnInstall按钮）
        // EventArgs e: 事件参数，包含事件相关信息
        private async void BtnInstall_Click(object sender, EventArgs e)
        {
            // 禁用安装按钮，防止用户重复点击
            btnInstall.Enabled = false;
            // 设置部署状态为true，表示正在部署
            isDeploying = true;
            // 记录日志：用户点击了安装按钮
            Logger.Info("用户点击一键安装按钮。");
            
            // ========== 阶段1：确认开始 ==========
            // GetInstallTypeText(): 获取当前选择的Office版本文本
            string installTypeText = GetInstallTypeText();
            // GlimMessageBox.Show: 显示自定义消息框
            // string.Format: 格式化字符串，将{0}替换为installTypeText
            // MessageBoxButtons.YesNo: 显示"是"和"否"按钮
            // MessageBoxIcon.Warning: 显示警告图标
            DialogResult confirmResult = GlimMessageBox.Show(
                string.Format("即将开始安装 {0}\n\n注意：程序将先彻底清理系统中现有的所有 Office 残留（包括软件、注册表、文件和日志），确保安装环境纯净。\n\n安装过程包括：\n1. 彻底清理旧版本\n2. 下载安装程序\n3. 安装 Office\n4. 自动激活\n\n是否继续？", installTypeText), 
                "安装确认", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            
            // DialogResult.Yes: 用户点击了"是"按钮
            if (confirmResult != DialogResult.Yes)
            {
                // 用户取消了安装，记录日志
                Logger.Info("用户取消了安装。");
                // 重新启用安装按钮
                btnInstall.Enabled = true;
                // 重置部署状态
                isDeploying = false;
                // return: 退出方法，不继续执行后续代码
                return;
            }

            // ========== 阶段 1.5：彻底清理旧版本 ==========
            // 更新状态标签文本
            lblStatus.Text = "正在彻底清理旧版本 Office 残留，请稍候...";
            // 设置状态标签为橙红色，表示正在进行重要操作
            lblStatus.ForeColor = Color.OrangeRed;
            // 记录日志：开始清理
            Logger.Info("开始彻底清理 Office 残留...");
            // ShowToast: 显示Toast提示消息
            // ToastType.Info: 信息类型提示
            ShowToast("开始清理", "正在清理旧版本 Office 残留...", ToastType.Info);
            // Task.Run: 在后台线程执行耗时操作
            // await: 等待任务完成，但不阻塞UI
            // ThoroughCleanupOffice(): 彻底清理Office的方法
            await Task.Run(() => ThoroughCleanupOffice());
            // 清理完成，更新状态
            lblStatus.Text = "清理完成，准备开始下载...";
            Logger.Info("Office 清理阶段完成。");
            // ToastType.Success: 成功类型提示
            ShowToast("清理完成", "旧版本 Office 已清理完成", ToastType.Success);
            
            // ========== 阶段2：下载安装程序 ==========
            // GOIConfig.SetupPath: 获取setup.exe的保存路径
            string setupPath = GOIConfig.SetupPath;
            
            // File.Exists: 检查文件是否已存在
            if (!File.Exists(setupPath))
            {
                // 文件不存在，需要下载
                lblStatus.Text = "正在从微软官网下载安装组件...";
                lblStatus.ForeColor = Color.Blue;
                Logger.Info("正在从微软下载 ODT setup.exe...");
                ShowToast("开始下载", "正在从微软官网下载安装组件...", ToastType.Info);
                
                // DownloadODT: 下载ODT安装程序的方法
                // await: 等待下载完成
                bool success = await DownloadODT(setupPath);
                // 检查下载是否成功
                if (!success)
                {
                    // 下载失败处理
                    lblStatus.Text = "下载失败，请检查网络连接。";
                    lblStatus.ForeColor = Color.Red;
                    Logger.Error("ODT 下载失败。");
                    // ToastType.Error: 错误类型提示
                    ShowToast("下载失败", "安装组件下载失败，请检查网络", ToastType.Error);
                    GlimMessageBox.Show("下载失败，请检查网络连接后重试。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    // 重新启用按钮并退出
                    btnInstall.Enabled = true;
                    return;
                }
                // 下载成功，记录日志
                Logger.Info("ODT 下载完成。路径: " + setupPath);
                ShowToast("下载完成", "安装组件下载完成", ToastType.Success);
            }

            // ========== 阶段3：生成配置文件 ==========
            lblStatus.Text = "正在生成配置文件...";
            // Path.Combine: 组合路径字符串
            // 生成configuration.xml的完整路径
            string xmlPath = Path.Combine(GOIConfig.RootPath, "configuration.xml");
            Logger.Info("正在生成 ODT 配置文件...");
            
            // try-catch: 异常处理结构
            // try: 尝试执行的代码块
            try
            {
                // GenerateXmlContent(): 生成XML配置内容的方法
                string xmlContent = GenerateXmlContent();
                // File.WriteAllText: 将内容写入文件（覆盖已存在的文件）
                File.WriteAllText(xmlPath, xmlContent);
                Logger.Info("配置文件已生成。内容:\n" + xmlContent);
            }
            // catch: 捕获异常
            catch (Exception ex)
            {
                // 异常处理：生成配置文件失败
                lblStatus.Text = "生成配置文件失败，请检查文件权限。";
                lblStatus.ForeColor = Color.Red;
                // ex: 异常对象，包含错误信息
                Logger.Error("生成配置文件失败。", ex);
                GlimMessageBox.Show("生成配置文件失败: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                btnInstall.Enabled = true;
                return;
            }

            // ========== 阶段4：安装Office ==========
            lblStatus.Text = "正在启动安装程序...";
            Logger.Info("启动 ODT 安装进程...");
            ShowToast("开始安装", "正在启动 Office 安装程序...", ToastType.Info);
            // 提示用户即将启动安装程序
            GlimMessageBox.Show("即将启动Office安装程序，请在弹出的窗口中完成安装。\n\n安装完成后请点击确定继续激活。", "安装阶段");
            
            try
            {
                // ProcessStartInfo: 进程启动信息类
                // 用于配置如何启动一个新进程
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    // FileName: 要启动的程序路径
                    FileName = setupPath,
                    // Arguments: 命令行参数
                    // /configure: ODT的配置命令
                    // string.Format: 格式化参数字符串
                    Arguments = string.Format("/configure \"{0}\"", xmlPath),
                    // UseShellExecute: 是否使用操作系统shell启动
                    // true: 使用shell启动，可以启动任何文件类型
                    UseShellExecute = true,
                    // Verb: 动作动词
                    // "runas": 以管理员权限运行
                    Verb = "runas",
                    // WorkingDirectory: 工作目录
                    WorkingDirectory = GOIConfig.RootPath
                };
                
                // Process.Start: 启动新进程
                // 返回Process对象，代表启动的进程
                Process proc = Process.Start(psi);
                lblStatus.Text = "安装程序运行中，请等待安装完成...";
                lblStatus.ForeColor = Color.Green;
                
                // 等待安装完成
                // Task.Run: 在后台线程等待，避免阻塞UI
                // WaitForExit: 阻塞当前线程直到进程退出
                await Task.Run(() => proc.WaitForExit());
                // ExitCode: 进程退出代码，0通常表示成功
                Logger.Info("ODT 安装进程已结束。退出码: " + proc.ExitCode);
                
                // ========== 阶段5：激活Office ==========
                lblStatus.Text = "正在准备激活...";
                Logger.Info("开始 Office 激活阶段...");
                ShowToast("开始激活", "正在激活 Office...", ToastType.Info);
                
                // RunOhookActivation: 运行激活方法
                bool activationSuccess = await RunOhookActivation();
                
                // 检查激活是否成功
                if (activationSuccess)
                {
                    // 激活成功
                    lblStatus.Text = "激活完成！Office已成功安装并激活。";
                    lblStatus.ForeColor = Color.Green;
                    Logger.Info("Office 激活成功。");
                    ShowToast("激活成功", "Office 已成功安装并激活！", ToastType.Success);
                    GlimMessageBox.Show("Office已成功安装并激活！\n\n您可以开始使用Office了。", "完成");
                }
                else
                {
                    // 激活失败
                    lblStatus.Text = "激活失败，请手动激活。";
                    lblStatus.ForeColor = Color.Red;
                    Logger.Warn("Office 自动激活可能未成功。");
                    // ToastType.Warning: 警告类型提示
                    ShowToast("激活失败", "Office 激活失败，请手动激活", ToastType.Warning);
                    GlimMessageBox.Show("自动激活失败，请检查激活工具是否存在，或手动运行激活程序。", "激活提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                
                // 重置部署状态
                isDeploying = false;
                // 重新启用安装按钮
                btnInstall.Enabled = true;
            }
            catch (Exception ex)
            {
                // 安装过程异常处理
                lblStatus.Text = "安装过程出错。";
                lblStatus.ForeColor = Color.Red;
                Logger.Error("安装过程发生异常。", ex);
                GlimMessageBox.Show("安装过程出错: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                isDeploying = false;
                btnInstall.Enabled = true;
            }
        }
        
        // ========== 获取安装类型文本方法 ==========
        // GetInstallTypeText: 根据当前选择的安装类型返回对应的文本
        // 返回值: string类型的Office版本名称
        private string GetInstallTypeText()
        {
            // switch: 多分支选择语句
            // 根据currentInstallType的值选择不同的分支
            switch (currentInstallType)
            {
                // case: 分支标签，匹配特定的值
                case InstallType.Office2024:
                    // return: 返回值并退出方法
                    return "Office 2024 LTSC";
                case InstallType.Office2021:
                    return "Office 2021 LTSC";
                case InstallType.Office2019:
                    return "Office 2019";
                case InstallType.Office2016:
                    return "Office 2016";
                case InstallType.Microsoft365Pro:
                    return "Microsoft 365 专业增强版";
                case InstallType.Microsoft365Home:
                    return "Microsoft 365 家庭个人版";
                // default: 默认分支，当所有case都不匹配时执行
                default:
                    return "Office";
            }
        }

        // ========== 彻底清理Office方法 ==========
        // ThoroughCleanupOffice: 深度清理系统中所有Office残留
        // 包括进程、文件、注册表等
        private void ThoroughCleanupOffice()
        {
            Logger.Info(">>> 开始执行深度清理 Office 任务 <<<");

            // ========== 1. 终止所有相关的进程 ==========
            // string[]: 字符串数组，存储需要终止的进程名称
            string[] processesToKill = { 
                // Office应用程序进程名
                "winword", "excel", "powerpnt", "outlook", "onenote", "publisher", 
                "infopath", "visio", "winproj", "msaccess", "lync", "groove", 
                // Office服务和辅助进程
                "teams", "officeclicktorun", "officeintegration", "setuphost",
                "officedebug", "msexcereport", "msosync", "msoia", "msoev",
                "splwow64", "office", "clview", "powerpnt", "mspub"
            };
            
            // foreach: 遍历数组中的每个元素
            // var: 隐式类型，编译器自动推断类型
            foreach (var procName in processesToKill)
            {
                try
                {
                    // Process.GetProcessesByName: 根据名称获取所有匹配的进程
                    // 返回Process数组
                    foreach (var proc in Process.GetProcessesByName(procName))
                    {
                        Logger.Info("终止进程: " + proc.ProcessName);
                        // proc.Kill: 强制终止进程
                        proc.Kill();
                        // WaitForExit(2000): 等待进程退出，最多等待2000毫秒
                        proc.WaitForExit(2000);
                    }
                }
                // 空catch: 忽略异常（某些进程可能无法终止）
                catch { }
            }

            // ========== 1.5 停止并清理 OfficeClickToRun 服务 ==========
            // OfficeClickToRun是Office的即点即用服务
            // 需要先停止服务，再删除服务
            try
            {
                Logger.Info("正在停止 OfficeClickToRun 服务...");
                // ProcessStartInfo: 配置进程启动参数
                ProcessStartInfo stopSvc = new ProcessStartInfo
                {
                    // sc.exe: Windows服务控制命令行工具
                    FileName = "sc.exe",
                    // stop ClickToRunSvc: 停止ClickToRun服务
                    Arguments = "stop ClickToRunSvc",
                    // CreateNoWindow: 不创建新窗口
                    CreateNoWindow = true,
                    // UseShellExecute: 不使用shell执行
                    UseShellExecute = false
                };
                // 启动停止服务的进程
                Process p1 = Process.Start(stopSvc);
                // WaitForExit(5000): 等待进程完成，最多等待5秒
                if (p1 != null) p1.WaitForExit(5000);

                Logger.Info("正在删除 ClickToRunSvc 服务...");
                // 删除服务的进程配置
                ProcessStartInfo delSvc = new ProcessStartInfo
                {
                    FileName = "sc.exe",
                    // delete ClickToRunSvc: 删除ClickToRun服务
                    Arguments = "delete ClickToRunSvc",
                    CreateNoWindow = true,
                    UseShellExecute = false
                };
                Process p2 = Process.Start(delSvc);
                if (p2 != null) p2.WaitForExit(5000);
            }
            // 捕获并记录异常
            catch (Exception ex) { Logger.Warn("停止/删除服务失败: " + ex.Message); }

            // ========== 2. 深度清理注册表 (HKCU 和 HKLM) ==========
            // 注册表路径数组：存储需要删除的Office相关注册表路径
            // @: 逐字字符串，不需要转义反斜杠
            string[] registryPaths = {
                // Office主键
                @"SOFTWARE\Microsoft\Office",
                // 即点即用相关
                @"SOFTWARE\Microsoft\Office\ClickToRun",
                // App-V虚拟化
                @"SOFTWARE\Microsoft\AppVisv",
                // 32位Office在64位系统上的注册表路径
                @"SOFTWARE\WOW6432Node\Microsoft\Office",
                // Office应用程序路径注册
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe",
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Excel.exe",
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Powerpnt.exe",
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Outlook.exe",
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Onenote.exe",
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Visio.exe",
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winproj.exe",
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Msaccess.exe",
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Mspub.exe",
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Lync.exe",
                // Office版本注册信息
                @"SOFTWARE\Microsoft\Office\16.0\Registration",  // Office 2016/2019/365
                @"SOFTWARE\Microsoft\Office\15.0\Registration",  // Office 2013
                @"SOFTWARE\Microsoft\Office\14.0\Registration",  // Office 2010
                // 映像文件执行选项（用于调试/拦截）
                @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\winword.exe",
                @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\excel.exe",
                @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\powerpnt.exe",
                // 安装程序数据
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\00006109C80000000000000000F01FEC",
                // 卸载信息
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Office16.PROPLUS",
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Office16.VISIOPRO",
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Office16.PROJECTPRO",
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Office16.OUTLOOK"
            };

            // 遍历所有注册表路径并删除
            foreach (var path in registryPaths)
            {
                // 删除当前用户(HKCU)下的注册表项
                DeleteRegistryKey(Microsoft.Win32.Registry.CurrentUser, path);
                // 删除本地机器(HKLM)下的注册表项
                DeleteRegistryKey(Microsoft.Win32.Registry.LocalMachine, path);
            }

            // ========== 3. 清理卸载残留 ==========
            // 卸载注册表路径数组
            string[] uninstallPaths = {
                // 64位程序的卸载路径
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                // 32位程序在64位系统上的卸载路径
                @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
            };

            // 遍历卸载路径
            foreach (var basePath in uninstallPaths)
            {
                try
                {
                    // using: 自动释放资源，确保注册表键被正确关闭
                    // OpenSubKey: 打开子键
                    // 第二个参数true: 表示需要写入权限
                    using (var rootKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(basePath, true))
                    {
                        // 检查键是否存在
                        if (rootKey != null)
                        {
                            // GetSubKeyNames: 获取所有子键名称
                            foreach (var subKeyName in rootKey.GetSubKeyNames())
                            {
                                try
                                {
                                    // 打开子键进行读取（不需要写入权限）
                                    using (var subKey = rootKey.OpenSubKey(subKeyName))
                                    {
                                        if (subKey != null)
                                        {
                                            // GetValue: 获取注册表值
                                            // as string: 类型转换，如果失败返回null
                                            // ??: 空合并运算符，如果左侧为null则使用右侧值
                                            string displayName = subKey.GetValue("DisplayName") as string ?? "";
                                            string publisher = subKey.GetValue("Publisher") as string ?? "";
                                            string uninstallString = subKey.GetValue("UninstallString") as string ?? "";

                                            // Contains: 检查字符串是否包含指定子串
                                            // StartsWith: 检查字符串是否以指定内容开头
                                            // 多条件判断：识别Office相关的卸载项
                                            if (displayName.Contains("Microsoft Office") || 
                                                displayName.Contains("Microsoft 365") || 
                                                displayName.Contains("Office 16") || 
                                                displayName.Contains("Office 15") ||
                                                subKeyName.StartsWith("Office1") || 
                                                subKeyName.Contains("0000-0000-0000000FF1CE") ||
                                                uninstallString.Contains("OfficeClickToRun") ||
                                                publisher.Contains("Microsoft Corporation"))
                                            {
                                                // 排除非 Office 的 Microsoft 产品
                                                // 使用!取反，排除不需要删除的项
                                                if (!displayName.Contains("Visual Studio") && 
                                                    !displayName.Contains("SQL Server") &&
                                                    !displayName.Contains("Edge") &&
                                                    !displayName.Contains("Windows"))
                                                {
                                                    Logger.Info("清理卸载项: " + displayName);
                                                    // DeleteSubKeyTree: 删除整个子键树
                                                    // 第二个参数false: 如果不存在不抛出异常
                                                    rootKey.DeleteSubKeyTree(subKeyName, false);
                                                }
                                            }
                                        }
                                    }
                                }
                                // 忽略单个子键处理异常
                                catch { }
                            }
                        }
                    }
                }
                // 忽略整个路径处理异常
                catch { }
            }

            // ========== 4. 清理计划任务 ==========
            // Windows计划任务可能包含Office相关的定时任务
            try
            {
                Logger.Info("正在清理 Office 相关计划任务...");
                // 任务关键词数组
                string[] taskKeywords = { "Office", "MicrosoftOffice", "OneNote", "Outlook" };
                foreach (var keyword in taskKeywords)
                {
                    // schtasks.exe: Windows计划任务命令行工具
                    ProcessStartInfo psi = new ProcessStartInfo
                    {
                        FileName = "schtasks.exe",
                        // /Delete: 删除任务
                        // /TN: 任务名称路径
                        // /F: 强制删除，不提示确认
                        Arguments = string.Format("/Delete /TN \"Microsoft\\Office\\{0}*\" /F", keyword),
                        CreateNoWindow = true,
                        UseShellExecute = false
                    };
                    Process p3 = Process.Start(psi);
                    if (p3 != null) p3.WaitForExit(2000);
                }
                
                // 尝试删除整个Office计划任务文件夹
                Process p4 = Process.Start(new ProcessStartInfo {
                    FileName = "schtasks.exe",
                    Arguments = "/Delete /TN \"Microsoft\\Office\" /F",
                    CreateNoWindow = true,
                    UseShellExecute = false
                });
                if (p4 != null) p4.WaitForExit(2000);
                
                Logger.Info("Office 计划任务清理尝试完成。");
            }
            catch (Exception ex) { Logger.Warn("清理计划任务时出现非致命错误: " + ex.Message); }

            // ========== 5. 清理残留文件 ==========
            // 需要清理的文件夹路径数组
            string[] foldersToClean = {
                // Program Files: 64位程序安装目录
                // Environment.SpecialFolder: 特殊文件夹枚举
                // GetFolderPath: 获取特殊文件夹路径
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Microsoft Office"),
                // Program Files (x86): 32位程序安装目录
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Microsoft Office"),
                // CommonApplicationData: 所有用户共享的应用数据目录 (C:\ProgramData)
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "Microsoft\\Office"),
                // LocalApplicationData: 当前用户的本地应用数据目录
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft\\Office"),
                // ApplicationData (Roaming): 当前用户的漫游应用数据目录
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft\\Office"),
                // Office共享组件目录
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Common Files\\microsoft shared\\OFFICE16"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Common Files\\microsoft shared\\OFFICE16"),
                // 即点即用目录
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Common Files\\Microsoft Shared\\ClickToRun"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "Microsoft\\ClickToRun"),
                // OneNote和Outlook数据目录
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft\\OneNote"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft\\Outlook")
            };

            // 遍历并删除所有残留目录
            foreach (var folder in foldersToClean)
            {
                try
                {
                    // Directory.Exists: 检查目录是否存在
                    if (Directory.Exists(folder))
                    {
                        Logger.Info("删除残留目录: " + folder);
                        // Directory.Delete: 删除目录
                        // 第二个参数true: 递归删除，包含所有子目录和文件
                        Directory.Delete(folder, true);
                    }
                }
                catch (Exception ex)
                {
                    Logger.Warn("无法删除目录 " + folder + ": " + ex.Message);
                }
            }

            Logger.Info("<<< 深度清理 Office 任务完成 >>>");
        }

        // ========== 删除注册表键方法 ==========
        // DeleteRegistryKey: 删除指定的注册表键
        // 参数root: 根键（如HKCU、HKLM）
        // 参数subKeyPath: 子键路径
        private void DeleteRegistryKey(Microsoft.Win32.RegistryKey root, string subKeyPath)
        {
            try
            {
                // DeleteSubKeyTree: 删除整个子键树（包括所有子键和值）
                // 第二个参数false: 如果键不存在，不抛出异常
                root.DeleteSubKeyTree(subKeyPath, false);
            }
            // 忽略所有异常（键可能不存在或无权限）
            catch { }
        }

        // ========== 删除目录内容方法 ==========
        // DeleteDirectoryContent: 删除目录内容
        // 参数path: 要清理的目录路径
        private void DeleteDirectoryContent(string path)
        {
            try
            {
                // 检查目录是否存在
                if (Directory.Exists(path))
                {
                    // 如果是 Office 核心目录，直接尝试删除整个文件夹
                    // Contains: 检查路径是否包含特定字符串
                    if (path.Contains("Microsoft Office") || path.Contains("OFFICE16"))
                    {
                        // 递归删除整个目录
                        Directory.Delete(path, true);
                    }
                    else
                    {
                        // 如果是临时目录，只清理内容，保留目录本身
                        // DirectoryInfo: 目录信息类，提供目录操作方法
                        DirectoryInfo di = new DirectoryInfo(path);
                        // GetFiles: 获取目录中所有文件
                        // FileInfo: 文件信息类
                        foreach (FileInfo file in di.GetFiles())
                        {
                            // 尝试删除每个文件，忽略异常
                            try { file.Delete(); } catch { }
                        }
                        // GetDirectories: 获取目录中所有子目录
                        foreach (DirectoryInfo dir in di.GetDirectories())
                        {
                            // 递归删除子目录
                            try { dir.Delete(true); } catch { }
                        }
                    }
                }
            }
            // 忽略所有异常
            catch { }
        }
        
        // ========== 显示隐藏功能窗口方法 ==========
        // ShowSecretWindow: 显示隐藏的功能窗口
        private void ShowSecretWindow()
        {
            // SecretWindow: 自定义的隐藏功能窗口类
            SecretWindow secretWindow = new SecretWindow();
            // ShowDialog: 以模态对话框方式显示窗口
            // 模态对话框：在关闭之前无法操作其他窗口
            secretWindow.ShowDialog();
        }

        // ========== 下载并安装教育软件方法 ==========
        // DownloadAndInstallSoftware: 下载并安装指定的教育软件
        // 参数softwareName: 软件名称
        // 参数downloadUrl: 下载链接
        // async void: 异步方法，不返回值
        private async void DownloadAndInstallSoftware(string softwareName, string downloadUrl)
        {
            try
            {
                Logger.Info("开始下载 " + softwareName + "...");
                
                // ========== 创建下载进度窗口 ==========
                // Form: Windows窗体基类
                Form progressForm = new Form
                {
                    // Text: 窗口标题
                    Text = "下载 " + softwareName,
                    // Size: 窗口大小
                    Size = new Size(450, 200),
                    // StartPosition: 窗口起始位置
                    // CenterScreen: 屏幕中央
                    StartPosition = FormStartPosition.CenterScreen,
                    // FormBorderStyle: 边框样式
                    // FixedDialog: 固定大小的对话框边框
                    FormBorderStyle = FormBorderStyle.FixedDialog,
                    // MaximizeBox: 是否显示最大化按钮
                    MaximizeBox = false,
                    // MinimizeBox: 是否显示最小化按钮
                    MinimizeBox = false,
                    // BackColor: 背景颜色
                    BackColor = Color.White
                };
                
                // ========== 加载图标 ==========
                try
                {
                    // Assembly: 程序集类
                    // GetExecutingAssembly: 获取当前执行的程序集
                    Assembly assembly = Assembly.GetExecutingAssembly();
                    // GetManifestResourceStream: 获取嵌入的资源流
                    // 嵌入资源是在编译时打包到exe中的文件
                    using (Stream stream = assembly.GetManifestResourceStream("OfficeInstaller.icon.ico"))
                    {
                        if (stream != null)
                        {
                            // Icon: 图标类，从流创建图标
                            progressForm.Icon = new Icon(stream);
                        }
                    }
                }
                // 忽略图标加载异常
                catch { }
                
                // ========== 状态标签 ==========
                Label lblStatus = new Label
                {
                    Text = "正在准备下载...",
                    // Location: 控件位置（相对于父容器）
                    Location = new Point(20, 30),
                    Size = new Size(400, 30),
                    // Font: 字体设置
                    // Microsoft YaHei UI: 微软雅黑字体
                    Font = new Font("Microsoft YaHei UI", 11),
                    // ForeColor: 前景色（文字颜色）
                    // FromArgb: 从RGB值创建颜色
                    ForeColor = Color.FromArgb(60, 60, 60)
                };
                // Controls.Add: 将控件添加到窗体
                progressForm.Controls.Add(lblStatus);
                
                // ========== 进度条 ==========
                // ProgressBar: 进度条控件
                ProgressBar progressBar = new ProgressBar
                {
                    Location = new Point(20, 80),
                    Size = new Size(400, 25),
                    // Style: 进度条样式
                    // Continuous: 连续样式，平滑显示
                    Style = ProgressBarStyle.Continuous,
                    // Minimum/Maximum: 进度范围
                    Minimum = 0,
                    Maximum = 100
                };
                progressForm.Controls.Add(progressBar);
                
                // ========== 进度百分比标签 ==========
                Label lblPercent = new Label
                {
                    Text = "0%",
                    Location = new Point(20, 115),
                    Size = new Size(400, 25),
                    Font = new Font("Microsoft YaHei UI", 10),
                    ForeColor = Color.FromArgb(100, 100, 100),
                    // TextAlign: 文本对齐方式
                    // MiddleCenter: 居中对齐
                    TextAlign = ContentAlignment.MiddleCenter
                };
                progressForm.Controls.Add(lblPercent);
                
                // ========== 显示进度窗口 ==========
                // Show: 非模态显示，允许继续执行后续代码
                progressForm.Show();
                
                // ========== 下载文件 ==========
                // 组合下载路径
                string downloadPath = Path.Combine(GOIConfig.DownloadPath, softwareName + "_Setup.exe");
                
                // WebClient: Web客户端类，用于下载文件
                // using: 确保资源被正确释放
                using (WebClient client = new WebClient())
                {
                    // DownloadProgressChanged: 下载进度变化事件
                    // +=: 订阅事件
                    // Lambda表达式: (sender, e) => { ... }
                    client.DownloadProgressChanged += (sender, e) =>
                    {
                        // Invoke: 在UI线程上执行委托
                        // 因为下载在后台线程，更新UI需要切换到UI线程
                        // MethodInvoker: 无参数委托
                        progressForm.Invoke((MethodInvoker)delegate
                        {
                            // ProgressPercentage: 下载进度百分比
                            progressBar.Value = e.ProgressPercentage;
                            lblPercent.Text = e.ProgressPercentage + "%";
                            lblStatus.Text = "正在下载 " + softwareName + "... " + e.ProgressPercentage + "%";
                        });
                    };
                    
                    // DownloadFileCompleted: 下载完成事件
                    client.DownloadFileCompleted += (sender, e) =>
                    {
                        progressForm.Invoke((MethodInvoker)delegate
                        {
                            // Error: 下载过程中的错误
                            if (e.Error != null)
                            {
                                progressForm.Close();
                                Logger.Error("下载 " + softwareName + " 失败", e.Error);
                                GlimMessageBox.Show("下载失败：" + e.Error.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            // Cancelled: 是否被取消
                            else if (e.Cancelled)
                            {
                                progressForm.Close();
                                Logger.Warn("下载 " + softwareName + " 已取消");
                                GlimMessageBox.Show("下载已取消。", "提示");
                            }
                            else
                            {
                                progressForm.Close();
                                Logger.Info("下载 " + softwareName + " 完成");
                                
                                // ========== 询问是否安装 ==========
                                DialogResult result = GlimMessageBox.Show(
                                    softwareName + " 下载完成！\n\n是否立即安装？",
                                    "下载完成",
                                    MessageBoxButtons.YesNo,
                                    MessageBoxIcon.Question);
                                
                                // 检查用户选择
                                if (result == DialogResult.Yes)
                                {
                                    // 调用安装方法
                                    InstallSoftware(softwareName, downloadPath);
                                }
                            }
                        });
                    };
                    
                    // ========== 开始下载 ==========
                    // DownloadFileTaskAsync: 异步下载文件
                    // Uri: 统一资源标识符类
                    // await: 等待下载完成
                    await client.DownloadFileTaskAsync(new Uri(downloadUrl), downloadPath);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("下载 " + softwareName + " 过程出错", ex);
                GlimMessageBox.Show("下载过程出错：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        // ========== 安装软件方法 ==========
        // InstallSoftware: 安装指定的软件
        // 参数softwareName: 软件名称
        // 参数installerPath: 安装程序路径
        private void InstallSoftware(string softwareName, string installerPath)
        {
            try
            {
                Logger.Info("开始安装 " + softwareName + "...");
                
                // 检查安装文件是否存在
                if (!File.Exists(installerPath))
                {
                    GlimMessageBox.Show("安装文件不存在：" + installerPath, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                
                // ========== 启动安装程序 ==========
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = installerPath,
                    // /S: 静默安装参数（适用于NSIS安装程序）
                    Arguments = "/S",
                    UseShellExecute = true,
                    // runas: 以管理员权限运行
                    Verb = "runas"
                };
                
                // 启动安装进程
                Process proc = Process.Start(psi);
                
                Logger.Info("启动 " + softwareName + " 安装程序");
                GlimMessageBox.Show(
                    softwareName + " 安装程序已启动。\n\n请按照安装向导完成安装。",
                    "安装启动",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Logger.Error("安装 " + softwareName + " 过程出错", ex);
                GlimMessageBox.Show("安装过程出错：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ========== 运行Ohook激活方法 ==========
        // RunOhookActivation: 使用Ohook方式激活Office
        // 返回值: Task<bool>，true表示激活成功
        private async Task<bool> RunOhookActivation()
        {
            try
            {
                // ========== 使用嵌入的Activator.cmd ==========
                // EmbeddedResource.ExtractActivator: 从嵌入资源提取激活脚本
                string activatorPath = EmbeddedResource.ExtractActivator();
                
                // 检查激活工具是否提取成功
                if (string.IsNullOrEmpty(activatorPath) || !File.Exists(activatorPath))
                {
                    Logger.Error("激活工具提取失败或路径不存在。");
                    return false;
                }
                
                lblStatus.Text = "正在执行 Ohook 激活，请稍候...";
                
                // 配置激活进程
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = activatorPath,
                    // /Ohook: 使用Ohook参数进行静默激活
                    Arguments = "/Ohook",
                    UseShellExecute = false,
                    // CreateNoWindow: 不创建新窗口
                    CreateNoWindow = true,
                    // WorkingDirectory: 工作目录
                    // Path.GetTempPath: 获取系统临时目录
                    WorkingDirectory = Path.GetTempPath()
                };
                
                // 启动激活进程
                Process proc = Process.Start(psi);
                // 异步等待进程完成
                await Task.Run(() => proc.WaitForExit());
                
                // ExitCode == 0 表示成功
                return proc.ExitCode == 0;
            }
            catch (Exception ex)
            {
                lblStatus.Text = "激活过程出错: " + ex.Message;
                GlimMessageBox.Show("激活过程出错: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        // ========== 单独激活按钮点击事件 ==========
        // BtnActivate_Click: 激活按钮的点击事件处理
        private async void BtnActivate_Click(object sender, EventArgs e)
        {
            // as: 安全类型转换，如果转换失败返回null
            ModernButton btn = sender as ModernButton;
            // 禁用按钮防止重复点击
            btn.Enabled = false;
            
            // ========== 第一次确认 ==========
            DialogResult result1 = GlimMessageBox.Show(
                "即将使用 Ohook 方式激活 Office。\n\n" +
                "此操作将：\n" +
                "1. 检查并安装必要的激活组件\n" +
                "2. 自动激活已安装的 Office 产品\n\n" +
                "是否继续？", 
                "激活确认 (1/2)", 
                MessageBoxButtons.YesNo, 
                MessageBoxIcon.Question);
                
            // 用户取消
            if (result1 != DialogResult.Yes)
            {
                btn.Enabled = true;
                return;
            }
            
            // ========== 第二次确认 - 免责声明 ==========
            DialogResult result2 = GlimMessageBox.Show(
                "重要提示：\n\n" +
                "• 本激活工具仅供学习和测试使用\n" +
                "• 请在24小时内删除激活后的软件\n" +
                "• 如需长期使用，请购买正版授权\n\n" +
                "点击 是 表示您已阅读并同意以上条款，\n" +
                "点击 否 取消激活操作。", 
                "激活确认 (2/2) - 免责声明", 
                MessageBoxButtons.YesNo, 
                MessageBoxIcon.Warning);
                
            // 用户不同意免责声明
            if (result2 != DialogResult.Yes)
            {
                btn.Enabled = true;
                return;
            }
            
            // 更新状态
            lblStatus.Text = "正在进行激活操作，请勿关闭本程序";
            // 执行激活
            bool success = await RunOhookActivation();
            
            // 处理激活结果
            if (success)
            {
                lblStatus.Text = "激活完成！Office 已成功激活。";
                lblStatus.ForeColor = Color.Green;
                GlimMessageBox.Show(
                    "Office 已成功激活！\n\n" +
                    "您可以打开 Word、Excel 等应用\n" +
                    "在\"账户\"中查看激活状态。", 
                    "激活成功");
            }
            else
            {
                // 激活失败处理
                lblStatus.Text = "激活失败，请检查是否已安装Office。";
                lblStatus.ForeColor = Color.Red;
                GlimMessageBox.Show(
                    "激活失败！\n\n" +
                    "可能的原因：\n" +
                    "1. 未安装 Office\n" +
                    "2. 激活工具运行出错\n" +
                    "3. 网络连接问题\n\n" +
                    "请确保已安装 Office 后重试。", 
                    "激活失败", 
                    MessageBoxButtons.OK, 
                    MessageBoxIcon.Error);
            }
            
            // 重新启用按钮
            btn.Enabled = true;
        }

        // ========== 下载ODT方法 ==========
        // DownloadODT: 下载Office部署工具(ODT)
        // 参数savePath: setup.exe的保存路径
        // 返回值: Task<bool>，true表示下载成功
        private async Task<bool> DownloadODT(string savePath)
        {
            try
            {
                string fileName;
                
                // ========== 根据安装类型选择显示名称 ==========
                switch (currentInstallType)
                {
                    case InstallType.Microsoft365Pro:
                        fileName = "Microsoft 365 专业增强版";
                        break;
                    case InstallType.Microsoft365Home:
                        fileName = "Microsoft 365 家庭个人版";
                        break;
                    case InstallType.Office2024:
                    default:
                        // default: 默认分支
                        fileName = "Office 2024";
                        break;
                }
                
                // ========== 统一使用 ODT (Office Deployment Tool) 下载 ==========
                // ODT下载链接（微软官方）
                string url = "https://c2rsetup.officeapps.live.com/c2r/officeDeploymentTool/officedeploymenttool.exe";
                
                using (WebClient client = new WebClient())
                {
                    // 订阅下载进度变化事件
                    client.DownloadProgressChanged += (sender, e) => {
                        // Invoke: 在UI线程更新标签
                        lblStatus.Invoke((MethodInvoker)delegate {
                            lblStatus.Text = string.Format("正在下载 {0}... {1}%", fileName, e.ProgressPercentage);
                        });
                    };
                    
                    // ========== 下载ODT自解压程序到 GOI/downloads ==========
                    string odtDownloadPath = Path.Combine(GOIConfig.DownloadPath, "officedeploymenttool.exe");
                    Logger.Info("开始下载 ODT 自解压程序: " + url);
                    // 异步下载文件
                    await client.DownloadFileTaskAsync(new Uri(url), odtDownloadPath);
                    
                    // ========== 静默解压ODT到 GOI 根目录 ==========
                    lblStatus.Invoke((MethodInvoker)delegate {
                        lblStatus.Text = "正在解压安装组件...";
                    });
                    Logger.Info("解压 ODT 到: " + GOIConfig.RootPath);
                    
                    // 配置解压进程
                    ProcessStartInfo psi = new ProcessStartInfo
                    {
                        FileName = odtDownloadPath,
                        // /quiet: 静默模式
                        // /extract: 解压到指定目录
                        Arguments = "/quiet /extract:\"" + GOIConfig.RootPath.TrimEnd('\\') + "\"",
                        UseShellExecute = false,
                        CreateNoWindow = true
                    };
                    
                    // 启动解压进程
                    Process proc = Process.Start(psi);
                    if (proc != null)
                    {
                        // 异步等待解压完成
                        await Task.Run(() => proc.WaitForExit());
                        Logger.Info("ODT 解压完成，退出码: " + proc.ExitCode);
                    }
                    
                    // ========== 清理下载的ODT安装程序 ==========
                    try { if (File.Exists(odtDownloadPath)) File.Delete(odtDownloadPath); } catch { }
                    
                    // ========== 检查setup.exe是否存在 ==========
                    if (!File.Exists(savePath))
                    {
                        Logger.Error("解压后未找到 setup.exe: " + savePath);
                    GlimMessageBox.Show("解压后未找到 setup.exe，请检查ODT下载是否成功。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                GlimMessageBox.Show("无法下载安装程序:\n" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        // ========== 生成XML配置内容方法 ==========
        // GenerateXmlContent: 生成ODT使用的configuration.xml内容
        // 返回值: XML配置字符串
        private string GenerateXmlContent()
        {
            // StringBuilder: 可变字符串构建器
            // 比直接使用string拼接更高效
            StringBuilder sb = new StringBuilder();
            
            // ========== 定义组件映射字典 - 严格遵循 ODT ID ==========
            // Dictionary<K,V>: 键值对集合
            // Key: 复选框显示名称，Value: ODT使用的组件ID
            var apps = new Dictionary<string, string> {
                {"Word", "Word"}, {"Excel", "Excel"}, {"PowerPoint", "PowerPoint"},
                {"Access", "Access"}, {"Outlook", "Outlook"}, {"OneNote", "OneNote"},
                {"Publisher", "Publisher"}, {"Lync", "Lync"}, {"OneDrive", "OneDrive"}, 
                {"Teams", "Teams"}, {"OneDrive (Groove)", "Groove"}
            };
            
            // AppendLine: 添加一行文本并换行
            sb.AppendLine("<Configuration>");
            
            // 版本架构：64位
            string edition = "64";
            // 更新通道
            string channel = "Current";
            // 产品ID
            string productId = "";
            
            // ========== 根据安装类型设置参数 ==========
            if (currentInstallType == InstallType.Office2024)
            {
                // PerpetualVL2024: Office 2024长期服务频道
                channel = "PerpetualVL2024";
                productId = "ProPlus2024Retail";
            }
            else if (currentInstallType == InstallType.Office2021)
            {
                channel = "PerpetualVL2021";
                productId = "ProPlus2021Retail";
            }
            else if (currentInstallType == InstallType.Office2019)
            {
                channel = "PerpetualVL2019";
                productId = "ProPlus2019Retail";
            }
            else if (currentInstallType == InstallType.Office2016)
            {
                channel = "PerpetualVL2016";
                productId = "ProPlusRetail";
            }
            else if (currentInstallType == InstallType.Microsoft365Pro)
            {
                // Microsoft 365专业增强版
                productId = "O365ProPlusRetail";
            }
            else if (currentInstallType == InstallType.Microsoft365Home)
            {
                // Microsoft 365家庭版
                productId = "O365HomePremRetail";
            }

            // ========== 生成XML配置内容 ==========
            // Add节点：添加产品配置
            // OfficeClientEdition: 架构版本（32或64位）
            // Channel: 更新通道
            sb.AppendLine(string.Format("  <Add OfficeClientEdition=\"{0}\" Channel=\"{1}\">", edition, channel));
            // Product节点：产品配置
            // ID: 产品标识符
            sb.AppendLine(string.Format("    <Product ID=\"{0}\">", productId));
            // Language节点：语言设置
            // zh-cn: 简体中文
            sb.AppendLine("      <Language ID=\"zh-cn\" />");

            // ========== 核心逻辑：用户没勾选的才 ExcludeApp (排除法) ==========
            // 遍历所有应用程序
            foreach(var app in apps)
            {
                // 如果 UI 字典中有该组件，且 CheckBox 没有被勾选，则添加排除标签
                // ContainsKey: 检查字典是否包含指定键
                // Checked: 复选框是否被选中
                if (productChecks.ContainsKey(app.Key) && !productChecks[app.Key].Checked)
                {
                    // ExcludeApp: 排除不需要安装的应用
                    sb.AppendLine(string.Format("      <ExcludeApp ID=\"{0}\" />", app.Value));
                }
            }
            sb.AppendLine("    </Product>");
            
            // ========== Visio 处理 - 使用零售版 ==========
            // 如果用户勾选了Visio，添加Visio产品配置
            if (productChecks.ContainsKey("Visio") && productChecks["Visio"].Checked)
            {
                string visioId = "VisioProRetail";
                // 根据Office版本选择对应的Visio版本
                if (currentInstallType == InstallType.Office2024) visioId = "VisioPro2024Retail";
                else if (currentInstallType == InstallType.Office2021) visioId = "VisioPro2021Retail";
                else if (currentInstallType == InstallType.Office2019) visioId = "VisioPro2019Retail";
                else if (currentInstallType == InstallType.Office2016) visioId = "VisioProRetail";
                sb.AppendLine(string.Format("    <Product ID=\"{0}\"><Language ID=\"zh-cn\" /></Product>", visioId));
            }
            
            // ========== Project 处理 - 使用零售版 ==========
            // 如果用户勾选了Project，添加Project产品配置
            if (productChecks.ContainsKey("Project") && productChecks["Project"].Checked)
            {
                string projectId = "ProjectProRetail";
                // 根据Office版本选择对应的Project版本
                if (currentInstallType == InstallType.Office2024) projectId = "ProjectPro2024Retail";
                else if (currentInstallType == InstallType.Office2021) projectId = "ProjectPro2021Retail";
                else if (currentInstallType == InstallType.Office2019) projectId = "ProjectPro2019Retail";
                else if (currentInstallType == InstallType.Office2016) projectId = "ProjectProRetail";
                sb.AppendLine(string.Format("    <Product ID=\"{0}\"><Language ID=\"zh-cn\" /></Product>", projectId));
            }

            // 关闭Add节点
            sb.AppendLine("  </Add>");
            // Display节点：显示设置
            // Level="Full": 显示完整安装界面
            // AcceptEULA="TRUE": 自动接受最终用户许可协议
            sb.AppendLine("  <Display Level=\"Full\" AcceptEULA=\"TRUE\" />");
            // Property节点：属性设置
            // SharedComputerLicensing: 共享计算机许可
            sb.AppendLine("  <Property Name=\"SharedComputerLicensing\" Value=\"0\" />");
            // FORCEAPPSHUTDOWN: 强制关闭Office应用
            sb.AppendLine("  <Property Name=\"FORCEAPPSHUTDOWN\" Value=\"TRUE\" />");
            // DeviceBasedLicensing: 基于设备的许可
            sb.AppendLine("  <Property Name=\"DeviceBasedLicensing\" Value=\"0\" />");
            // Updates节点：更新设置
            // Enabled="TRUE": 启用自动更新
            sb.AppendLine("  <Updates Enabled=\"TRUE\" />");
            // 关闭Configuration节点
            sb.AppendLine("</Configuration>");
            
            // ToString: 返回构建的字符串
            return sb.ToString();
        }

        // ========== 程序入口点 ==========
        // [STAThread]: 单线程单元特性
        // 表示应用程序使用单线程单元(COM)模型
        // Windows Forms应用程序必须使用此模型
        [STAThread]
        // Main: 程序入口方法
        static void Main()
        {
            // ========== 检查管理员权限 ==========
            if (!IsAdministrator())
            {
                // 以管理员权限重新启动
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    // Assembly.GetExecutingAssembly().Location: 获取当前程序路径
                    FileName = Assembly.GetExecutingAssembly().Location,
                    UseShellExecute = true,
                    // Verb = "runas": 以管理员身份运行
                    Verb = "runas"
                };

                try
                {
                    // 启动新的管理员进程
                    Process.Start(psi);
                }
                catch
                {
                    // 用户取消了UAC提示或无法提权
                    GlimMessageBox.Show("本程序需要管理员权限才能正常运行。", "权限不足", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                
                // 退出当前非管理员进程
                return;
            }

            // ========== 初始化 GOI 环境 ==========
            // Initialize: 初始化配置目录和文件
            GOIConfig.Initialize();
            Logger.Info("程序以管理员模式启动。");

            // ========== 设置TLS 1.2支持，确保HTTPS下载正常工作 ==========
            // SecurityProtocol: 安全协议类型
            // Tls12 | Tls11 | Tls: 使用位或运算组合多个协议
            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
            
            // ========== 设置DPI感知 - 支持不同分辨率 ==========
            // Environment.OSVersion.Version: 获取操作系统版本
            // Major >= 6: Windows Vista及更高版本
            if (Environment.OSVersion.Version.Major >= 6) 
            {
                // SetProcessDPIAware: 设置进程DPI感知
                SetProcessDPIAware();
                // 启用Per-Monitor DPI感知（Windows 8.1+）
                try
                {
                    // 2 = Per-monitor DPI aware
                    SetProcessDpiAwareness(2);
                }
                catch { }
            }
            
            // ========== 启动Windows Forms应用程序 ==========
            // EnableVisualStyles: 启用视觉样式（XP及更高版本）
            Application.EnableVisualStyles();
            // SetCompatibleTextRenderingDefault: 设置文本渲染默认值
            // false: 使用GDI+渲染文本（推荐）
            Application.SetCompatibleTextRenderingDefault(false);
            // Run: 启动应用程序主消息循环
            // new MainForm(): 创建主窗体实例
            Application.Run(new MainForm());
        }

        // ========== 检查管理员权限方法 ==========
        // IsAdministrator: 检查当前进程是否以管理员权限运行
        // 返回值: bool，true表示是管理员
        private static bool IsAdministrator()
        {
            // WindowsIdentity: Windows用户身份类
            // GetCurrent: 获取当前Windows用户身份
            var identity = System.Security.Principal.WindowsIdentity.GetCurrent();
            // WindowsPrincipal: Windows主体类，用于检查用户角色
            var principal = new System.Security.Principal.WindowsPrincipal(identity);
            // IsInRole: 检查用户是否属于指定角色
            // Administrator: 管理员角色
            return principal.IsInRole(System.Security.Principal.WindowsBuiltInRole.Administrator);
        }

        // ========== DLL导入 - DPI感知设置 ==========
        // DllImport: 导入外部DLL函数
        // user32.dll: Windows用户界面DLL
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        // extern: 表示方法在外部实现
        private static extern bool SetProcessDPIAware();
        
        // shcore.dll: Windows Shell核心DLL（Windows 8.1+）
        [System.Runtime.InteropServices.DllImport("shcore.dll")]
        private static extern int SetProcessDpiAwareness(int awareness);
    }

    // ========== 嵌入资源辅助类 ==========
    // EmbeddedResource: 处理嵌入在程序集中的资源文件
    // static: 静态类，不能实例化
    public static class EmbeddedResource {
        // ExtractActivator: 提取激活工具脚本
        // 返回值: 提取后的文件路径
        public static string ExtractActivator() {
            // 获取工具目录路径
            string toolsPath = GOIConfig.ToolsPath;
            // 组合激活脚本完整路径
            string activatorPath = Path.Combine(toolsPath, "Activator.cmd");
            
            try {
                // 如果工具目录不存在，创建它
                if (!Directory.Exists(toolsPath)) Directory.CreateDirectory(toolsPath);
                
                // 即使存在也重新提取，确保是最新版本
                // Assembly: 程序集类
                Assembly assembly = Assembly.GetExecutingAssembly();
                // 嵌入资源的完整名称
                // 格式: 命名空间.文件名
                string resourceName = "OfficeInstaller.Activator.cmd";
                
                // GetManifestResourceStream: 获取嵌入资源的流
                using (Stream stream = assembly.GetManifestResourceStream(resourceName)) {
                    // 检查资源是否存在
                    if (stream == null) {
                        Logger.Warn("未找到嵌入的激活资源: " + resourceName);
                        return "";
                    }
                    
                    // FileStream: 文件流类
                    // FileMode.Create: 创建新文件或覆盖现有文件
                    using (FileStream fileStream = new FileStream(activatorPath, FileMode.Create, FileAccess.Write)) {
                        // CopyTo: 将流内容复制到另一个流
                        stream.CopyTo(fileStream);
                    }
                }
                Logger.Info("激活工具已提取到: " + activatorPath);
            }
            catch (Exception ex) {
                Logger.Error("提取激活工具失败", ex);
                return "";
            }
            
            return activatorPath;
        }
    }

    // ========== 隐藏功能窗口类 ==========
    // SecretWindow: 隐藏功能窗口，连续点击banner 5次触发
    // 继承自Form类
    public class SecretWindow : Form
    {
        // glimBlue: Glim品牌蓝色
        private Color glimBlue = Color.FromArgb(0, 122, 204);
        
        // 构造函数：初始化窗口
        public SecretWindow()
        {
            this.Text = "Glim Office Installer - 隐藏功能";
            
            // ========== DPI感知设置 ==========
            // AutoScaleMode.Dpi: 根据DPI自动缩放
            this.AutoScaleMode = AutoScaleMode.Dpi;
            // AutoScaleDimensions: 自动缩放尺寸
            this.AutoScaleDimensions = new SizeF(96F, 96F);
            
            // 窗口基本属性设置
            this.Size = new Size(600, 550);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.BackColor = Color.White;
            
            // ========== 加载图标 ==========
            try
            {
                Assembly assembly = Assembly.GetExecutingAssembly();
                using (Stream stream = assembly.GetManifestResourceStream("OfficeInstaller.icon.ico"))
                {
                    if (stream != null)
                    {
                        this.Icon = new Icon(stream);
                    }
                }
            }
            catch { }
            
            // ========== 顶部装饰条 ==========
            Panel topBar = new Panel
            {
                Location = new Point(0, 0),
                Size = new Size(600, 5),
                BackColor = glimBlue
            };
            this.Controls.Add(topBar);
            
            // ========== 标题 ==========
            Label lblTitle = new Label
            {
                Text = "隐藏功能菜单",
                Font = EmbeddedFont.GetFont(20, FontStyle.Bold),
                ForeColor = glimBlue,
                Location = new Point(0, 25),
                Size = new Size(600, 45),
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleCenter,
                // Cursors.Hand: 鼠标悬停时显示手形光标
                Cursor = Cursors.Hand
            };
            
            // ========== 日志查看器触发逻辑 ==========
            // 闭包变量：记录点击次数
            int titleClickCount = 0;
            // Click事件：点击标题时触发
            lblTitle.Click += (s, e) => {
                titleClickCount++;
                // 连续点击5次显示日志查看器
                if (titleClickCount >= 5)
                {
                    titleClickCount = 0;
                    Logger.ShowLogViewer();
                }
            };
            this.Controls.Add(lblTitle);
            
            // ========== 提示文字 ==========
            Label lblHint = new Label
            {
                Text = "发现隐藏功能！请谨慎使用以下选项。",
                Font = EmbeddedFont.GetFont(11, FontStyle.Regular),
                ForeColor = Color.FromArgb(100, 100, 100),
                Location = new Point(0, 75),
                Size = new Size(600, 30),
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleCenter
            };
            this.Controls.Add(lblHint);
            
            // ========== 创建主界面风格的卡片按钮 ==========
            // Func<...>: 泛型委托类型，表示一个返回值的方法
            // 参数: title(标题), subtitle(副标题), icon(图标), yPos(Y坐标), clickAction(点击动作)
            // 返回值: Panel控件
            Func<string, string, string, int, Action, Panel> createCardButton = (title, subtitle, icon, yPos, clickAction) =>
            {
                Panel card = new Panel
                {
                    Location = new Point(100, yPos),
                    Size = new Size(400, 80),
                    BackColor = Color.White,
                    Cursor = Cursors.Hand,
                    // Tag: 控件的标签，可存储任意对象
                    Tag = clickAction
                };
                
                // 卡片圆角半径
                int cardRadius = 12;
                // 边框宽度
                int borderWidth = 2;
                // 鼠标悬停状态
                bool isHovered = false;
                // 选中状态
                bool isSelected = false;
                
                // ========== 绘制卡片 ==========
                // Paint事件：控件需要重绘时触发
                card.Paint += (sender, e) =>
                {
                    // Graphics: GDI+绘图对象
                    Graphics g = e.Graphics;
                    // SmoothingMode: 抗锯齿模式
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    
                    // 绘制区域矩形
                    Rectangle rect = new Rectangle(1, 1, card.Width - 3, card.Height - 3);
                    
                    // GraphicsPath: 图形路径，用于绘制复杂形状
                    using (GraphicsPath path = new GraphicsPath())
                    {
                        // AddArc: 添加圆弧
                        // 四个角分别添加90度圆弧
                        path.AddArc(rect.X, rect.Y, cardRadius * 2, cardRadius * 2, 180, 90);
                        path.AddArc(rect.Right - cardRadius * 2, rect.Y, cardRadius * 2, cardRadius * 2, 270, 90);
                        path.AddArc(rect.Right - cardRadius * 2, rect.Bottom - cardRadius * 2, cardRadius * 2, cardRadius * 2, 0, 90);
                        path.AddArc(rect.X, rect.Bottom - cardRadius * 2, cardRadius * 2, cardRadius * 2, 90, 90);
                        path.CloseFigure();
                        
                        // ========== 填充背景 ==========
                        using (SolidBrush brush = new SolidBrush(Color.White))
                        {
                            // FillPath: 填充路径
                            g.FillPath(brush, path);
                        }
                        
                        // ========== 绘制边框 ==========
                        // 三元运算符: 条件 ? 真值 : 假值
                        Color borderColor = isSelected ? glimBlue : (isHovered ? Color.FromArgb(150, 150, 150) : Color.FromArgb(220, 220, 220));
                        using (Pen pen = new Pen(borderColor, borderWidth))
                        {
                            // DrawPath: 绘制路径轮廓
                            g.DrawPath(pen, path);
                        }
                        
                        // ========== 绘制左边框强调色 ==========
                        if (isSelected || isHovered)
                        {
                            using (SolidBrush brush = new SolidBrush(glimBlue))
                            {
                                // FillRectangle: 填充矩形
                                g.FillRectangle(brush, 0, cardRadius, 4, card.Height - cardRadius * 2);
                            }
                        }
                    }
                };
                
                // ========== 图标标签 ==========
                Label lblIcon = new Label
                {
                    Text = icon,
                    // Segoe UI Emoji: Windows表情符号字体
                    Font = new Font("Segoe UI Emoji", 24),
                    Location = new Point(20, 10),
                    Size = new Size(50, 50),
                    AutoSize = false,
                    TextAlign = ContentAlignment.MiddleCenter,
                    BackColor = Color.Transparent
                };
                card.Controls.Add(lblIcon);
                
                // ========== 标题标签 ==========
                Label lblCardTitle = new Label
                {
                    Text = title,
                    Font = EmbeddedFont.GetFont(13, FontStyle.Bold),
                    ForeColor = Color.FromArgb(0, 90, 158),
                    Location = new Point(80, 10),
                    Size = new Size(310, 30),
                    AutoSize = false,
                    BackColor = Color.Transparent
                };
                card.Controls.Add(lblCardTitle);
                
                // ========== 副标题标签 ==========
                Label lblCardSubtitle = new Label
                {
                    Text = subtitle,
                    Font = EmbeddedFont.GetFont(9, FontStyle.Regular),
                    ForeColor = Color.FromArgb(100, 100, 100),
                    Location = new Point(80, 44),
                    Size = new Size(310, 26),
                    AutoSize = false,
                    BackColor = Color.Transparent
                };
                card.Controls.Add(lblCardSubtitle);
                
                // ========== 鼠标效果 ==========
                // MouseEnter: 鼠标进入控件区域
                card.MouseEnter += (sender, e) =>
                {
                    isHovered = true;
                    // Invalidate: 使控件无效，触发重绘
                    card.Invalidate();
                };
                // MouseLeave: 鼠标离开控件区域
                card.MouseLeave += (sender, e) =>
                {
                    isHovered = false;
                    card.Invalidate();
                };
                
                // ========== 子控件也绑定鼠标效果和点击事件传递 ==========
                // Control: 所有Windows控件的基类
                foreach (Control ctrl in card.Controls)
                {
                    ctrl.MouseEnter += (sender, e) =>
                    {
                        isHovered = true;
                        card.Invalidate();
                    };
                    ctrl.MouseLeave += (sender, e) =>
                    {
                        isHovered = false;
                        card.Invalidate();
                    };
                    // 点击事件传递给父Panel
                    ctrl.Click += (sender, e) =>
                    {
                        // 手动触发卡片点击
                        // is: 类型检查
                        if (card.Tag != null && card.Tag is Action)
                        {
                            // 强制类型转换并执行Action委托
                            ((Action)card.Tag)();
                        }
                    };
                }
                
                return card;
            };
            
            // ========== 创建功能卡片 ==========
            // 当前Y坐标
            int currentY = 120;
            // 卡片间距
            int cardSpacing = 90;
            
            // ========== 卡片1：激活管理 ==========
            Panel cardActivation = createCardButton("激活管理", "Windows/Office 激活工具集", "🔑", currentY, () => {
                // Lambda表达式作为点击回调
                ShowActivationManagementMenu();
            });
            // 订阅卡片点击事件
            cardActivation.Click += (s, e) => ((Action)cardActivation.Tag)();
            this.Controls.Add(cardActivation);
            currentY += cardSpacing;
            
            // ========== 卡片2：查看Windows激活信息 ==========
            Panel cardWinLicense = createCardButton("Windows 激活信息", "查看系统授权状态", "🪟", currentY, () => {
                ViewWindowsLicense();
            });
            cardWinLicense.Click += (s, e) => ((Action)cardWinLicense.Tag)();
            this.Controls.Add(cardWinLicense);
            currentY += cardSpacing;
            
            // ========== 卡片3：查看Office授权信息 ==========
            Panel cardOfficeLicense = createCardButton("Office 授权信息", "查看Office授权状态", "📋", currentY, () => {
                ViewOfficeLicense();
            });
            cardOfficeLicense.Click += (s, e) => ((Action)cardOfficeLicense.Tag)();
            this.Controls.Add(cardOfficeLicense);
            currentY += cardSpacing;
            
            // ========== 卡片4：更改Office授权人 ==========
            Panel cardChangeOwner = createCardButton("更改授权人名称", "修改Office显示的用户名", "✏️", currentY, () => {
                ChangeOfficeOwner();
            });
            cardChangeOwner.Click += (s, e) => ((Action)cardChangeOwner.Tag)();
            this.Controls.Add(cardChangeOwner);
        }
        
        // ========== 查看Office授权信息方法 ==========
        // ViewOfficeLicense: 显示Office产品的授权状态
        private void ViewOfficeLicense()
        {
            try
            {
                // ========== 使用cscript运行ospp.vbs查看授权信息 ==========
                // ospp.vbs: Office软件保护平台脚本
                // Environment.SpecialFolder.ProgramFiles: Program Files目录
                string officePath = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
                // 可能的Office安装路径数组
                string[] paths = new string[]
                {
                    // Office 2016/2019/365
                    Path.Combine(officePath, "Microsoft Office", "Office16"),
                    // Office 2013
                    Path.Combine(officePath, "Microsoft Office", "Office15"),
                    // Office 2010
                    Path.Combine(officePath, "Microsoft Office", "Office14"),
                    // 即点即用版本路径
                    Path.Combine(officePath, "Microsoft Office", "root", "Office16"),
                    Path.Combine(officePath, "Microsoft Office", "root", "Office15"),
                };
                
                // 查找ospp.vbs脚本
                string osppPath = null;
                foreach (var p in paths)
                {
                    string testPath = Path.Combine(p, "ospp.vbs");
                    if (File.Exists(testPath))
                    {
                        osppPath = testPath;
                        // break: 找到后退出循环
                        break;
                    }
                }
                
                // 如果找到ospp.vbs
                if (osppPath != null)
                {
                    // 配置进程启动信息
                    ProcessStartInfo psi = new ProcessStartInfo
                    {
                        // cscript.exe: Windows脚本宿主命令行版本
                        FileName = "cscript.exe",
                        // /dstatus: 显示授权状态
                        Arguments = "\"" + osppPath + "\" /dstatus",
                        UseShellExecute = false,
                        // RedirectStandardOutput: 重定向标准输出
                        RedirectStandardOutput = true,
                        CreateNoWindow = true
                    };
                    
                    // 启动进程并读取输出
                    using (Process proc = Process.Start(psi))
                    {
                        // StandardOutput: 标准输出流
                        // ReadToEnd: 读取所有输出
                        string output = proc.StandardOutput.ReadToEnd();
                        proc.WaitForExit();
                        
                        // ========== 显示结果 ==========
                        Form resultForm = new Form
                        {
                            Text = "Office授权信息",
                            Size = new Size(700, 500),
                            StartPosition = FormStartPosition.CenterScreen,
                            BackColor = Color.White
                        };
                        
                        // TextBox: 文本框控件
                        TextBox txtResult = new TextBox
                        {
                            // Multiline: 多行模式
                            Multiline = true,
                            // ScrollBars: 滚动条设置
                            ScrollBars = ScrollBars.Both,
                            // Dock: 停靠方式
                            // Fill: 填充整个父容器
                            Dock = DockStyle.Fill,
                            Text = output,
                            Font = EmbeddedFont.GetFont(10, FontStyle.Regular),
                            // ReadOnly: 只读模式
                            ReadOnly = true,
                            BackColor = Color.White
                        };
                        resultForm.Controls.Add(txtResult);
                        // ShowDialog: 模态显示
                        resultForm.ShowDialog();
                    }
                }
                else
                {
                    GlimMessageBox.Show("未找到Office授权管理工具。\n请确保已安装Office。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                GlimMessageBox.Show("查看授权信息失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        // ========== 查看Windows激活信息方法 ==========
        // ViewWindowsLicense: 显示Windows系统的激活状态
        private void ViewWindowsLicense()
        {
            try
            {
                // StringBuilder: 可变字符串构建器
                StringBuilder output = new StringBuilder();
                output.AppendLine("Windows 激活信息");
                output.AppendLine("================");
                output.AppendLine();
                
                // ========== 方法1: 使用PowerShell获取激活信息 ==========
                try
                {
                    // 配置PowerShell进程
                    ProcessStartInfo psi = new ProcessStartInfo
                    {
                        FileName = "powershell.exe",
                        // PowerShell命令：查询软件许可产品
                        // Get-WmiObject: 获取WMI对象
                        // SoftwareLicensingProduct: 软件许可产品类
                        // PartialProductKey is not null: 筛选有部分产品密钥的产品
                        Arguments = "-Command \"(Get-WmiObject -Query 'select * from SoftwareLicensingProduct where PartialProductKey is not null' | Select-Object -Property Name, ApplicationID, LicenseStatus, PartialProductKey | Format-List | Out-String)\"",
                        UseShellExecute = false,
                        RedirectStandardOutput = true,
                        // RedirectStandardError: 重定向标准错误流
                        RedirectStandardError = true,
                        CreateNoWindow = true,
                        // Environment.SystemDirectory: 系统目录（如C:\Windows\System32）
                        WorkingDirectory = Environment.SystemDirectory
                    };
                    
                    using (Process proc = Process.Start(psi))
                    {
                        // 读取标准输出
                        string result = proc.StandardOutput.ReadToEnd();
                        // 读取错误输出
                        string error = proc.StandardError.ReadToEnd();
                        proc.WaitForExit();
                        
                        // 检查结果是否有效
                        if (!string.IsNullOrEmpty(result) && result.Contains("Name"))
                        {
                            output.AppendLine("【方法1】PowerShell 查询结果:");
                            output.AppendLine(result);
                        }
                    }
                }
                catch (Exception ex1)
                {
                    output.AppendLine("【方法1】PowerShell 查询失败: " + ex1.Message);
                }
                
                output.AppendLine();
                output.AppendLine("================");
                output.AppendLine();
                
                // ========== 方法2: 使用systeminfo命令 ==========
                try
                {
                    ProcessStartInfo psi = new ProcessStartInfo
                    {
                        // systeminfo.exe: 显示系统信息
                        FileName = "systeminfo.exe",
                        Arguments = "",
                        UseShellExecute = false,
                        RedirectStandardOutput = true,
                        CreateNoWindow = true
                    };
                    
                    using (Process proc = Process.Start(psi))
                    {
                        string result = proc.StandardOutput.ReadToEnd();
                        proc.WaitForExit();
                        
                        // ========== 提取关键信息 ==========
                        // Split: 分割字符串
                        // StringSplitOptions.RemoveEmptyEntries: 移除空条目
                        string[] lines = result.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                        output.AppendLine("【方法2】系统信息:");
                        foreach (string line in lines)
                        {
                            // 筛选包含关键信息的行
                            if (line.Contains("OS 名称") || line.Contains("OS Name") ||
                                line.Contains("OS 版本") || line.Contains("OS Version") ||
                                line.Contains("产品 ID") || line.Contains("Product ID") ||
                                line.Contains("系统制造商") || line.Contains("System Manufacturer"))
                            {
                                // Trim: 去除首尾空白
                                output.AppendLine(line.Trim());
                            }
                        }
                    }
                }
                catch (Exception ex2)
                {
                    output.AppendLine("【方法2】系统信息查询失败: " + ex2.Message);
                }
                
                // ========== 显示结果 ==========
                Form resultForm = new Form
                {
                    Text = "Windows激活信息",
                    Size = new Size(700, 500),
                    StartPosition = FormStartPosition.CenterScreen,
                    BackColor = Color.White
                };
                
                TextBox txtResult = new TextBox
                {
                    Multiline = true,
                    ScrollBars = ScrollBars.Both,
                    Dock = DockStyle.Fill,
                    Text = output.ToString(),
                    Font = EmbeddedFont.GetFont(10, FontStyle.Regular),
                    ReadOnly = true,
                    BackColor = Color.White
                };
                resultForm.Controls.Add(txtResult);
                resultForm.ShowDialog();
            }
            catch (Exception ex)
            {
                GlimMessageBox.Show("查看Windows激活信息失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        // ========== 更改Office授权人方法 ==========
        // ChangeOfficeOwner: 修改Office显示的用户名和公司名
        private void ChangeOfficeOwner()
        {
            // ========== 创建输入对话框 ==========
            Form inputForm = new Form
            {
                Text = "更改授权人信息",
                Size = new Size(450, 320),
                StartPosition = FormStartPosition.CenterScreen,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                BackColor = Color.FromArgb(245, 245, 250)
            };
            
            // ========== 加载图标 ==========
            try
            {
                Assembly assembly = Assembly.GetExecutingAssembly();
                using (Stream stream = assembly.GetManifestResourceStream("OfficeInstaller.icon.ico"))
                {
                    if (stream != null)
                    {
                        inputForm.Icon = new Icon(stream);
                    }
                }
            }
            catch { }
            
            // ========== 标题 ==========
            Label lblTitle = new Label
            {
                Text = "更改 Office 用户信息",
                Location = new Point(0, 20),
                Size = new Size(450, 35),
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = EmbeddedFont.GetFont(14, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 122, 204)
            };
            inputForm.Controls.Add(lblTitle);
            
            // ========== 用户名标签 ==========
            Label lblName = new Label
            {
                Text = "用户名 (User Name):",
                Location = new Point(50, 70),
                Size = new Size(350, 25),
                Font = EmbeddedFont.GetFont(10, FontStyle.Regular),
                ForeColor = Color.FromArgb(60, 60, 60)
            };
            inputForm.Controls.Add(lblName);
            
            // ========== 用户名输入框 ==========
            TextBox txtName = new TextBox
            {
                Location = new Point(50, 95),
                Size = new Size(330, 30),
                Font = EmbeddedFont.GetFont(11, FontStyle.Regular),
                // BorderStyle: 边框样式
                // FixedSingle: 单线边框
                BorderStyle = BorderStyle.FixedSingle
            };
            inputForm.Controls.Add(txtName);
            
            // ========== 公司/组织标签 ==========
            Label lblOrg = new Label
            {
                Text = "公司/组织 (Organization):",
                Location = new Point(50, 140),
                Size = new Size(350, 25),
                Font = EmbeddedFont.GetFont(10, FontStyle.Regular),
                ForeColor = Color.FromArgb(60, 60, 60)
            };
            inputForm.Controls.Add(lblOrg);
            
            // ========== 公司输入框 ==========
            TextBox txtOrg = new TextBox
            {
                Location = new Point(50, 165),
                Size = new Size(330, 30),
                Font = EmbeddedFont.GetFont(11, FontStyle.Regular),
                BorderStyle = BorderStyle.FixedSingle
            };
            inputForm.Controls.Add(txtOrg);
            
            // ========== 获取当前值作为默认值 ==========
            try
            {
                // 打开注册表键读取当前用户信息
                using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\Common\UserInfo"))
                {
                    if (key != null)
                    {
                        // GetValue: 获取注册表值
                        txtName.Text = key.GetValue("UserName") as string ?? "";
                        txtOrg.Text = key.GetValue("Company") as string ?? "";
                    }
                }
            }
            catch { }
            
            // ========== 确定按钮 ==========
            Button btnOK = new Button();
            btnOK.Text = "保存更改";
            btnOK.Location = new Point(125, 215);
            btnOK.Size = new Size(200, 45);
            btnOK.Font = EmbeddedFont.GetFont(12, FontStyle.Bold);
            btnOK.ForeColor = Color.White;
            // FlatStyle: 按钮样式
            // Flat: 扁平样式
            btnOK.FlatStyle = FlatStyle.Flat;
            // FlatAppearance: 扁平外观设置
            btnOK.FlatAppearance.BorderSize = 0;
            btnOK.Cursor = Cursors.Hand;
            btnOK.BackColor = Color.Transparent;
            // DialogResult: 对话框结果
            // OK: 表示用户点击了确定
            btnOK.DialogResult = DialogResult.OK;
            
            // 按钮圆角半径
            int btnRadius = 12;
            // 自定义绘制按钮
            btnOK.Paint += (sender, e) =>
            {
                Graphics g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;
                
                Rectangle rect = new Rectangle(2, 2, btnOK.Width - 5, btnOK.Height - 5);
                
                using (GraphicsPath path = new GraphicsPath())
                {
                    // 绘制圆角矩形路径
                    path.AddArc(rect.X, rect.Y, btnRadius * 2, btnRadius * 2, 180, 90);
                    path.AddArc(rect.Right - btnRadius * 2, rect.Y, btnRadius * 2, btnRadius * 2, 270, 90);
                    path.AddArc(rect.Right - btnRadius * 2, rect.Bottom - btnRadius * 2, btnRadius * 2, btnRadius * 2, 0, 90);
                    path.AddArc(rect.X, rect.Bottom - btnRadius * 2, btnRadius * 2, btnRadius * 2, 90, 90);
                    path.CloseFigure();
                    
                    // ========== 橙色渐变填充 ==========
                    // LinearGradientBrush: 线性渐变画刷
                    // 从橙色渐变到红色
                    using (LinearGradientBrush brush = new LinearGradientBrush(
                        rect, Color.FromArgb(255, 112, 67), Color.FromArgb(232, 65, 37), LinearGradientMode.Horizontal))
                    {
                        g.FillPath(brush, path);
                    }
                    
                    // TextRenderer: 文本渲染器
                    // DrawText: 绘制文本
                    // TextFormatFlags: 文本格式标志
                    TextRenderer.DrawText(g, btnOK.Text, btnOK.Font, rect, Color.White,
                        TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                }
            };
            
            inputForm.Controls.Add(btnOK);
            // AcceptButton: 设置接受按钮（按Enter键触发）
            inputForm.AcceptButton = btnOK;
            
            // 显示对话框并检查结果
            if (inputForm.ShowDialog() == DialogResult.OK)
            {
                // 获取输入的用户名和公司名
                string newName = txtName.Text.Trim();
                string newOrg = txtOrg.Text.Trim();
                
                // 验证用户名不能为空
                if (string.IsNullOrEmpty(newName))
                {
                    GlimMessageBox.Show("用户名不能为空。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                try
                {
                    // 成功计数器
                    int successCount = 0;
                    StringBuilder log = new StringBuilder();
                    
                    // ========== 定义需要修改的注册表路径列表 ==========
                    // List<T>: 泛型列表
                    var regPaths = new List<string>
                    {
                        @"Software\Microsoft\Office\Common\UserInfo",
                        // Office 2016/2019/2021/365/2024
                        @"Software\Microsoft\Office\16.0\Common\UserInfo",
                        // Office 2013
                        @"Software\Microsoft\Office\15.0\Common\UserInfo",
                        // Office 2010
                        @"Software\Microsoft\Office\14.0\Common\UserInfo",
                        // Office 2007
                        @"Software\Microsoft\Office\12.0\Common\UserInfo",
                        // Office 2003
                        @"Software\Microsoft\Office\11.0\Common\UserInfo",
                        // Office 2024 特定路径
                        @"Software\Microsoft\Office\16.0\Common\General",
                        @"Software\Microsoft\Office\Common\General"
                    };

                    // ========== 1. 遍历修改 HKCU (当前用户配置) ==========
                    foreach (var path in regPaths)
                    {
                        try 
                        {
                            // CreateSubKey: 创建或打开子键
                            using (var key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(path))
                            {
                                if (key != null)
                                {
                                    // SetValue: 设置注册表值
                                    key.SetValue("UserName", newName);
                                    key.SetValue("Company", newOrg);
                                    
                                    // ========== 设置缩写 ==========
                                    string initials = "";
                                    if (newName.Length > 0)
                                    {
                                        // Split: 按空格分割名字
                                        string[] parts = newName.Split(new char[]{' '}, StringSplitOptions.RemoveEmptyEntries);
                                        if (parts.Length >= 2)
                                            // Substring: 截取子字符串
                                            // ToUpper: 转大写
                                            initials = (parts[0].Substring(0,1) + parts[1].Substring(0,1)).ToUpper();
                                        else if (newName.Length >= 2)
                                            initials = newName.Substring(0, 2).ToUpper();
                                        else
                                            initials = newName.ToUpper();
                                        key.SetValue("UserInitials", initials);
                                    }
                                    successCount++;
                                }
                            }
                        }
                        catch { }
                    }

                    // ========== 2. 修改 HKLM - Windows 注册信息 (需要管理员权限) ==========
                    try
                    {
                        // 标准路径
                        using (var key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion", true))
                        {
                            if (key != null)
                            {
                                key.SetValue("RegisteredOwner", newName);
                                key.SetValue("RegisteredOrganization", newOrg);
                                successCount++;
                            }
                        }
                        
                        // WOW6432Node (64位系统上的32位程序信息)
                        using (var key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"SOFTWARE\WOW6432Node\Microsoft\Windows NT\CurrentVersion", true))
                        {
                            if (key != null)
                            {
                                key.SetValue("RegisteredOwner", newName);
                                key.SetValue("RegisteredOrganization", newOrg);
                                successCount++;
                            }
                        }
                    }
                    catch { }

                    // ========== 3. 关键：修改 Office Click-To-Run 虚拟注册表 (针对 C2R/O365 版本) ==========
                    // 这是 "产品信息" 显示的关键位置
                    try
                    {
                        string c2rPath = @"SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Windows NT\CurrentVersion";
                        using (var key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(c2rPath, true))
                        {
                            if (key != null)
                            {
                                key.SetValue("RegisteredOwner", newName);
                                key.SetValue("RegisteredOrganization", newOrg);
                                successCount++;
                            }
                        }
                    }
                    catch { }

                    // ========== 4. 遍历所有 Office 版本的 Registration 注册表 (针对 MSI/批量许可版本) ==========
                    try
                    {
                        string[] rootPaths = { 
                            @"SOFTWARE\Microsoft\Office", 
                            @"SOFTWARE\WOW6432Node\Microsoft\Office" 
                        };

                        foreach (var rootPath in rootPaths)
                        {
                            using (var rootKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(rootPath))
                            {
                                if (rootKey == null) continue;

                                // 遍历版本号 (16.0, 15.0 等)
                                foreach (var verKeyName in rootKey.GetSubKeyNames())
                                {
                                    // 检查 Registration 子项
                                    string regBasePath = rootPath + "\\" + verKeyName + "\\Registration";
                                    using (var regKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(regBasePath))
                                    {
                                        if (regKey != null)
                                        {
                                            // 遍历 GUID
                                            foreach (var guid in regKey.GetSubKeyNames())
                                            {
                                                try
                                                {
                                                    string guidPath = regBasePath + "\\" + guid;
                                                    using (var guidKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(guidPath, true))
                                                    {
                                                        if (guidKey != null)
                                                        {
                                                            // 确保是有效的产品注册项
                                                            if (guidKey.GetValue("ProductName") != null || guidKey.GetValue("DigitalProductID") != null)
                                                            {
                                                                guidKey.SetValue("RegisteredOwner", newName);
                                                                guidKey.SetValue("RegisteredOrganization", newOrg);
                                                                successCount++;
                                                            }
                                                        }
                                                    }
                                                }
                                                catch { }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch { }
                    
                    // ========== 5. 尝试修改 Excel/Word 特定的 UserInfo (如果存在) ==========
                    // 应用程序数组
                    string[] apps = { "Word", "Excel", "PowerPoint" };
                    // Office版本数组
                    string[] versions = { "16.0", "15.0", "14.0" };
                    // 遍历所有版本
                    foreach (var ver in versions)
                    {
                        // 遍历所有应用程序
                        foreach (var app in apps)
                        {
                            try
                            {
                                // string.Format: 格式化字符串
                                string appPath = string.Format(@"Software\Microsoft\Office\{0}\{1}\UserInfo", ver, app);
                                using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(appPath, true))
                                {
                                    if (key != null)
                                    {
                                        key.SetValue("UserName", newName);
                                        key.SetValue("Company", newOrg);
                                    }
                                }
                            }
                            catch { }
                        }
                    }

                    // 显示结果
                    if (successCount > 0)
                    {
                        GlimMessageBox.Show(
                            "已强制更新所有本地授权信息！\n\n" +
                            "★ 重要提示 ★\n" +
                            "如果您当前已登录 Microsoft 账户（界面显示头像和邮箱）：\n" +
                            "Office 会优先显示云端账户名称，而非本地设置的名称。\n\n" +
                            "若要看到修改效果，请尝试：\n" +
                            "1. 在账户页面点击 [注销] 退出登录\n" +
                            "2. 重启 Office 应用程序\n" +
                            "3. 此时应显示您设置的本地名称", 
                            "修改完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        GlimMessageBox.Show("未能更新任何信息，请确保拥有管理员权限。", "更新失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                catch (Exception ex)
                {
                    GlimMessageBox.Show("发生错误：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        
        // ========== 激活Windows方法 ==========
        // ActivateWindows: 使用Activator.cmd激活Windows系统
        // 使用HWID方式获取数字权利激活
        private async Task ActivateWindows()
        {
            try
            {
                // ========== 显示确认弹窗 ==========
                // ShowConfirmDialog: 显示确认对话框
                if (!ShowConfirmDialog("激活确认", "即将使用 HWID 方式激活 Windows 系统。\n\n此操作将：\n1. 获取数字权利激活\n2. 永久激活您的 Windows 系统\n\n是否继续？", "开始激活"))
                    return;

                // 提取激活工具
                string activatorPath = EmbeddedResource.ExtractActivator();
                
                GlimMessageBox.Show("即将开始后台激活Windows系统，请稍候...", "激活Windows");
                
                // 配置激活进程
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = activatorPath,
                    // /Z-Windows: 使用HWID方式激活Windows
                    Arguments = "/Z-Windows",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WorkingDirectory = Path.GetTempPath()
                };
                
                // 启动并等待进程
                Process proc = Process.Start(psi);
                await Task.Run(() => proc.WaitForExit());
                
                // 检查退出代码
                if (proc.ExitCode == 0)
                {
                    GlimMessageBox.Show("Windows激活完成！", "激活成功");
                }
                else
                {
                    GlimMessageBox.Show("Windows激活可能未完成。", "激活提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                GlimMessageBox.Show("激活Windows过程出错：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        // ========== 显示激活管理菜单方法 ==========
        // ShowActivationManagementMenu: 显示包含所有激活功能的菜单
        private void ShowActivationManagementMenu()
        {
            // 创建菜单窗体
            Form menuForm = new Form
            {
                Text = "激活管理",
                AutoScaleMode = AutoScaleMode.Dpi,
                AutoScaleDimensions = new SizeF(96F, 96F),
                Size = new Size(600, 850),
                StartPosition = FormStartPosition.CenterScreen,
                FormBorderStyle = FormBorderStyle.FixedSingle,
                MaximizeBox = false,
                MinimizeBox = false,
                BackColor = Color.White
            };
            
            // ========== 加载图标 ==========
            try
            {
                Assembly assembly = Assembly.GetExecutingAssembly();
                using (Stream stream = assembly.GetManifestResourceStream("OfficeInstaller.icon.ico"))
                {
                    if (stream != null)
                    {
                        menuForm.Icon = new Icon(stream);
                    }
                }
            }
            catch { }
            
            // ========== 顶部装饰条 ==========
            Panel topBar = new Panel
            {
                Location = new Point(0, 0),
                Size = new Size(600, 5),
                BackColor = glimBlue
            };
            menuForm.Controls.Add(topBar);
            
            // ========== 标题 ==========
            Label lblTitle = new Label
            {
                Text = "激活管理",
                Font = EmbeddedFont.GetFont(20, FontStyle.Bold),
                ForeColor = glimBlue,
                Location = new Point(0, 25),
                Size = new Size(600, 45),
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleCenter
            };
            menuForm.Controls.Add(lblTitle);
            
            // ========== 提示文字 ==========
            Label lblHint = new Label
            {
                Text = "选择要执行的激活操作",
                Font = EmbeddedFont.GetFont(11, FontStyle.Regular),
                ForeColor = Color.FromArgb(100, 100, 100),
                Location = new Point(0, 75),
                Size = new Size(600, 30),
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleCenter
            };
            menuForm.Controls.Add(lblHint);
            
            // ========== 创建卡片按钮的辅助方法 ==========
            // Func委托：创建激活卡片
            Func<string, string, string, int, Action, Panel> createActivationCard = (title, subtitle, icon, yPos, clickAction) =>
            {
                Panel card = new Panel
                {
                    Location = new Point(100, yPos),
                    Size = new Size(400, 75),
                    BackColor = Color.White,
                    Cursor = Cursors.Hand,
                    Tag = clickAction
                };
                
                // 卡片圆角半径
                int cardRadius = 10;
                // 边框宽度
                int borderWidth = 2;
                // 鼠标悬停状态
                bool isHovered = false;
                
                // ========== 自定义绘制卡片 ==========
                card.Paint += (sender, e) =>
                {
                    Graphics g = e.Graphics;
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    
                    Rectangle rect = new Rectangle(1, 1, card.Width - 3, card.Height - 3);
                    
                    using (GraphicsPath path = new GraphicsPath())
                    {
                        // 绘制圆角矩形路径
                        path.AddArc(rect.X, rect.Y, cardRadius * 2, cardRadius * 2, 180, 90);
                        path.AddArc(rect.Right - cardRadius * 2, rect.Y, cardRadius * 2, cardRadius * 2, 270, 90);
                        path.AddArc(rect.Right - cardRadius * 2, rect.Bottom - cardRadius * 2, cardRadius * 2, cardRadius * 2, 0, 90);
                        path.AddArc(rect.X, rect.Bottom - cardRadius * 2, cardRadius * 2, cardRadius * 2, 90, 90);
                        path.CloseFigure();
                        
                        // 填充白色背景
                        using (SolidBrush brush = new SolidBrush(Color.White))
                        {
                            g.FillPath(brush, path);
                        }
                        
                        // 绘制边框
                        Color borderColor = isHovered ? glimBlue : Color.FromArgb(220, 220, 220);
                        using (Pen pen = new Pen(borderColor, borderWidth))
                        {
                            g.DrawPath(pen, path);
                        }
                        
                        // 悬停时绘制左侧强调条
                        if (isHovered)
                        {
                            using (SolidBrush brush = new SolidBrush(glimBlue))
                            {
                                g.FillRectangle(brush, 0, cardRadius, 4, card.Height - cardRadius * 2);
                            }
                        }
                    }
                };
                
                // ========== 图标标签 ==========
                Label lblIcon = new Label
                {
                    Text = icon,
                    Font = new Font("Segoe UI Emoji", 22),
                    Location = new Point(15, 8),
                    Size = new Size(50, 50),
                    AutoSize = false,
                    TextAlign = ContentAlignment.MiddleCenter,
                    BackColor = Color.Transparent
                };
                card.Controls.Add(lblIcon);
                
                // ========== 标题标签 ==========
                Label lblCardTitle = new Label
                {
                    Text = title,
                    Font = EmbeddedFont.GetFont(12, FontStyle.Bold),
                    ForeColor = Color.FromArgb(0, 90, 158),
                    Location = new Point(70, 10),
                    Size = new Size(320, 28),
                    AutoSize = false,
                    BackColor = Color.Transparent
                };
                card.Controls.Add(lblCardTitle);
                
                // ========== 副标题标签 ==========
                Label lblCardSubtitle = new Label
                {
                    Text = subtitle,
                    Font = EmbeddedFont.GetFont(9, FontStyle.Regular),
                    ForeColor = Color.FromArgb(100, 100, 100),
                    Location = new Point(70, 42),
                    Size = new Size(320, 20),
                    AutoSize = false,
                    BackColor = Color.Transparent
                };
                card.Controls.Add(lblCardSubtitle);
                
                // ========== 鼠标效果 ==========
                card.MouseEnter += (sender, e) =>
                {
                    isHovered = true;
                    card.Invalidate();
                };
                card.MouseLeave += (sender, e) =>
                {
                    isHovered = false;
                    card.Invalidate();
                };
                
                // 子控件事件传递
                foreach (Control ctrl in card.Controls)
                {
                    ctrl.MouseEnter += (sender, e) =>
                    {
                        isHovered = true;
                        card.Invalidate();
                    };
                    ctrl.MouseLeave += (sender, e) =>
                    {
                        isHovered = false;
                        card.Invalidate();
                    };
                    // 点击事件传递给父Panel
                    ctrl.Click += (sender, e) =>
                    {
                        // 触发卡片的Click事件
                        if (card.Tag != null && card.Tag is Action)
                        {
                            ((Action)card.Tag)();
                        }
                    };
                }
                
                return card;
            };
            
            // ========== 创建功能卡片 ==========
            int currentY = 120;
            int cardSpacing = 85;
            
            // ========== 卡片1：一键激活Windows ==========
            Panel cardWin = createActivationCard("一键激活 Windows", "使用 HWID 数字权利激活", "🪟", currentY, () => {
                menuForm.Close();
                ActivateWindows();
            });
            cardWin.Click += (s, e) => ((Action)cardWin.Tag)();
            menuForm.Controls.Add(cardWin);
            currentY += cardSpacing;
            
            // ========== 卡片2：一键激活Office ==========
            Panel cardOffice = createActivationCard("一键激活 Office", "使用 Ohook 永久激活", "📦", currentY, () => {
                menuForm.Close();
                ActivateOffice();
            });
            cardOffice.Click += (s, e) => ((Action)cardOffice.Tag)();
            menuForm.Controls.Add(cardOffice);
            currentY += cardSpacing;
            
            // ========== 卡片3：Office自定义KMS激活 ==========
            Panel cardKMS = createActivationCard("Office 自定义KMS激活", "输入自己的KMS服务器地址", "🔧", currentY, () => {
                menuForm.Close();
                ShowCustomKMSWindow();
            });
            cardKMS.Click += (s, e) => ((Action)cardKMS.Tag)();
            menuForm.Controls.Add(cardKMS);
            currentY += cardSpacing;
            
            // ========== 卡片4：Office TsForge激活 ==========
            Panel cardTsForge = createActivationCard("Office TsForge激活", "使用 TsForge 方式激活", "⚙️", currentY, () => {
                menuForm.Close();
                ActivateOfficeTsForge();
            });
            cardTsForge.Click += (s, e) => ((Action)cardTsForge.Tag)();
            menuForm.Controls.Add(cardTsForge);
            currentY += cardSpacing;
            
            // ========== 卡片5：Windows HWID激活 ==========
            Panel cardHWID = createActivationCard("Windows HWID激活", "数字权利永久激活", "🔐", currentY, () => {
                menuForm.Close();
                ActivateWindowsHWID();
            });
            cardHWID.Click += (s, e) => ((Action)cardHWID.Tag)();
            menuForm.Controls.Add(cardHWID);
            currentY += cardSpacing;
            
            // ========== 卡片6：Windows ESU激活 ==========
            Panel cardESU = createActivationCard("Windows ESU激活", "扩展安全更新激活", "🛡️", currentY, () => {
                menuForm.Close();
                ActivateESU();
            });
            cardESU.Click += (s, e) => ((Action)cardESU.Tag)();
            menuForm.Controls.Add(cardESU);
            currentY += cardSpacing;

            // ========== 卡片7：删除Office激活 ==========
            Panel cardRemoveOffice = createActivationCard("删除 Office 激活", "清除Office激活状态和密钥", "🗑️", currentY, () => {
                menuForm.Close();
                RemoveOfficeActivation();
            });
            cardRemoveOffice.Click += (s, e) => ((Action)cardRemoveOffice.Tag)();
            menuForm.Controls.Add(cardRemoveOffice);
            currentY += cardSpacing;

            // ========== 卡片8：删除Windows激活 ==========
            Panel cardRemoveWindows = createActivationCard("删除 Windows 激活", "清除Windows激活状态", "🗑️", currentY, () => {
                menuForm.Close();
                RemoveWindowsActivation();
            });
            cardRemoveWindows.Click += (s, e) => ((Action)cardRemoveWindows.Tag)();
            menuForm.Controls.Add(cardRemoveWindows);

            // 显示菜单窗体
            menuForm.ShowDialog();
        }

        // ========== 显示带LOGO的确认弹窗方法 ==========
        // ShowConfirmDialog: 显示带LOGO的确认对话框
        // 参数title: 对话框标题
        // 参数message: 消息内容
        // 参数buttonText: 确认按钮文本
        // 返回值: bool，true表示用户点击确认
        private bool ShowConfirmDialog(string title, string message, string buttonText = "确认")
        {
            // 创建确认对话框窗体
            Form confirmForm = new Form
            {
                Text = title,
                AutoScaleMode = AutoScaleMode.Dpi,
                AutoScaleDimensions = new SizeF(96F, 96F),
                Size = new Size(500, 280),
                StartPosition = FormStartPosition.CenterScreen,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                BackColor = Color.White
            };

            // ========== 加载图标 ==========
            try
            {
                Assembly assembly = Assembly.GetExecutingAssembly();
                using (Stream stream = assembly.GetManifestResourceStream("OfficeInstaller.icon.ico"))
                {
                    if (stream != null)
                    {
                        confirmForm.Icon = new Icon(stream);
                    }
                }
            }
            catch { }

            // ========== 顶部装饰条 ==========
            Panel topBar = new Panel
            {
                Location = new Point(0, 0),
                Size = new Size(500, 5),
                BackColor = glimBlue
            };
            confirmForm.Controls.Add(topBar);

            // ========== LOGO和标题区域 ==========
            Panel headerPanel = new Panel
            {
                Location = new Point(0, 20),
                Size = new Size(500, 60),
                BackColor = Color.White
            };
            confirmForm.Controls.Add(headerPanel);

            // ========== LOGO图标 ==========
            Label lblLogo = new Label
            {
                Text = "G",
                Font = new Font("Arial", 28, FontStyle.Bold),
                ForeColor = glimBlue,
                Location = new Point(30, 5),
                Size = new Size(50, 50),
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleCenter
            };
            headerPanel.Controls.Add(lblLogo);

            // ========== 软件名称 ==========
            Label lblAppName = new Label
            {
                Text = "Glim Office Installer",
                Font = EmbeddedFont.GetFont(14, FontStyle.Bold),
                ForeColor = glimBlue,
                Location = new Point(90, 10),
                Size = new Size(380, 25),
                AutoSize = false
            };
            headerPanel.Controls.Add(lblAppName);

            // ========== 版本信息 ==========
            Label lblVersion = new Label
            {
                Text = "Office 一键安装工具",
                Font = EmbeddedFont.GetFont(10, FontStyle.Regular),
                ForeColor = Color.FromArgb(100, 100, 100),
                Location = new Point(90, 35),
                Size = new Size(380, 20),
                AutoSize = false
            };
            headerPanel.Controls.Add(lblVersion);

            // ========== 分隔线 ==========
            Panel separator = new Panel
            {
                Location = new Point(30, 90),
                Size = new Size(440, 1),
                BackColor = Color.FromArgb(220, 220, 220)
            };
            confirmForm.Controls.Add(separator);

            // ========== 消息内容 ==========
            Label lblMessage = new Label
            {
                Text = message,
                Font = EmbeddedFont.GetFont(11, FontStyle.Regular),
                ForeColor = Color.FromArgb(60, 60, 60),
                Location = new Point(30, 110),
                Size = new Size(440, 80),
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleLeft
            };
            confirmForm.Controls.Add(lblMessage);

            // 结果变量（闭包使用）
            bool result = false;

            // ========== 确认按钮 ==========
            // ModernButton: 自定义现代风格按钮
            ModernButton btnConfirm = new ModernButton
            {
                Text = buttonText,
                Location = new Point(260, 200),
                Size = new Size(100, 35),
                BackColor = glimBlue
            };
            btnConfirm.Click += (s, e) =>
            {
                result = true;
                confirmForm.Close();
            };
            confirmForm.Controls.Add(btnConfirm);

            // ========== 取消按钮 ==========
            ModernButton btnCancel = new ModernButton
            {
                Text = "取消",
                Location = new Point(370, 200),
                Size = new Size(100, 35),
                BackColor = Color.FromArgb(150, 150, 150)
            };
            btnCancel.Click += (s, e) =>
            {
                result = false;
                confirmForm.Close();
            };
            confirmForm.Controls.Add(btnCancel);

            // 显示对话框并返回结果
            confirmForm.ShowDialog();
            return result;
        }

        // ========== 激活Office方法 ==========
        // ActivateOffice: 使用Activator.cmd的Ohook方式激活Office
        private async Task ActivateOffice()
        {
            try
            {
                // ========== 显示确认弹窗 ==========
                if (!ShowConfirmDialog("激活确认", "即将使用 Ohook 方式激活 Office。\n\n此操作将：\n1. 检查并安装必要的激活组件\n2. 自动激活已安装的 Office 产品\n\n是否继续？", "开始激活"))
                    return;

                // 提取激活工具
                string activatorPath = EmbeddedResource.ExtractActivator();
                
                GlimMessageBox.Show("即将开始后台激活Office，请稍候...", "激活Office");
                
                // 配置激活进程
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = activatorPath,
                    // /Ohook: 使用Ohook方式激活Office
                    Arguments = "/Ohook",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WorkingDirectory = Path.GetTempPath()
                };
                
                // 启动并等待进程
                Process proc = Process.Start(psi);
                await Task.Run(() => proc.WaitForExit());
                
                // 检查退出代码
                if (proc.ExitCode == 0)
                {
                    GlimMessageBox.Show("Office激活完成！", "激活成功");
                }
                else
                {
                    GlimMessageBox.Show("Office激活可能未完成。", "激活提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                GlimMessageBox.Show("激活Office过程出错：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        // ========== 激活Windows ESU方法 ==========
        // ActivateESU: 使用Activator.cmd激活Windows扩展安全更新
        // ESU: Extended Security Updates（扩展安全更新）
        private async Task ActivateESU()
        {
            try
            {
                // 提取激活工具
                string activatorPath = EmbeddedResource.ExtractActivator();
                
                // 检查激活工具是否存在
                if (string.IsNullOrEmpty(activatorPath) || !File.Exists(activatorPath))
                {
                    GlimMessageBox.Show("激活工具提取失败，请检查程序完整性。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                
                Logger.Info("开始激活 Windows ESU...");
                
                // 配置激活进程
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = activatorPath,
                    // /Z-ESU: 激活Windows扩展安全更新
                    Arguments = "/Z-ESU",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    // 重定向输出以捕获结果
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    WorkingDirectory = Path.GetTempPath()
                };
                
                using (Process proc = Process.Start(psi))
                {
                    // 异步读取输出
                    string output = await Task.Run(() => proc.StandardOutput.ReadToEnd());
                    string error = await Task.Run(() => proc.StandardError.ReadToEnd());
                    await Task.Run(() => proc.WaitForExit());
                    
                    // 记录日志
                    Logger.Info("ESU激活输出: " + output);
                    if (!string.IsNullOrEmpty(error))
                    {
                        Logger.Error("ESU激活错误: " + error, null);
                    }
                    
                    // 检查退出代码
                    if (proc.ExitCode == 0)
                    {
                        GlimMessageBox.Show("Windows ESU 激活完成！", "激活成功");
                        Logger.Info("Windows ESU 激活成功");
                    }
                    else
                    {
                        GlimMessageBox.Show("Windows ESU 激活可能未完成，请查看日志了解详情。", "激活提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        Logger.Warn("Windows ESU 激活返回非零退出码: " + proc.ExitCode);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("激活Windows ESU过程出错", ex);
                GlimMessageBox.Show("激活Windows ESU过程出错：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ========== 删除Office激活方法 ==========
        // RemoveOfficeActivation: 清除Office产品的激活状态和密钥
        private void RemoveOfficeActivation()
        {
            try
            {
                // 显示确认对话框
                DialogResult result = GlimMessageBox.Show(
                    "确定要删除Office激活状态吗？\n\n这将清除所有Office产品的激活信息和密钥。",
                    "确认删除",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);

                // 用户取消
                if (result != DialogResult.Yes)
                    return;

                // 提取激活工具
                string activatorPath = EmbeddedResource.ExtractActivator();

                // 配置进程
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = activatorPath,
                    // /RemoveOffice: 删除Office激活
                    Arguments = "/RemoveOffice",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WorkingDirectory = Path.GetTempPath()
                };

                // 启动并等待进程
                Process proc = Process.Start(psi);
                proc.WaitForExit();

                // 检查结果
                if (proc.ExitCode == 0)
                {
                    GlimMessageBox.Show("Office激活状态已清除！", "操作成功");
                    Logger.Info("Office激活状态已清除");
                }
                else
                {
                    GlimMessageBox.Show("清除Office激活状态可能未完成。", "操作提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    Logger.Warn("清除Office激活返回非零退出码: " + proc.ExitCode);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("清除Office激活过程出错", ex);
                GlimMessageBox.Show("清除Office激活过程出错：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ========== 删除Windows激活方法 ==========
        // RemoveWindowsActivation: 清除Windows的激活状态
        private void RemoveWindowsActivation()
        {
            try
            {
                // 显示确认对话框
                DialogResult result = GlimMessageBox.Show(
                    "确定要删除Windows激活状态吗？\n\n这将清除Windows的激活信息。",
                    "确认删除",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);

                // 用户取消
                if (result != DialogResult.Yes)
                    return;

                // 提取激活工具
                string activatorPath = EmbeddedResource.ExtractActivator();

                // 配置进程
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = activatorPath,
                    // /RemoveWindows: 删除Windows激活
                    Arguments = "/RemoveWindows",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WorkingDirectory = Path.GetTempPath()
                };

                // 启动并等待进程
                Process proc = Process.Start(psi);
                proc.WaitForExit();

                // 检查结果
                if (proc.ExitCode == 0)
                {
                    GlimMessageBox.Show("Windows激活状态已清除！", "操作成功");
                    Logger.Info("Windows激活状态已清除");
                }
                else
                {
                    GlimMessageBox.Show("清除Windows激活状态可能未完成。", "操作提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    Logger.Warn("清除Windows激活返回非零退出码: " + proc.ExitCode);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("清除Windows激活过程出错", ex);
                GlimMessageBox.Show("清除Windows激活过程出错：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ========== 显示自定义KMS激活窗口方法 ==========
        // ShowCustomKMSWindow: 显示自定义KMS服务器激活界面
        private void ShowCustomKMSWindow()
        {
            // 创建KMS激活窗体
            Form kmsForm = new Form
            {
                Text = "Office自定义KMS激活",
                Size = new Size(450, 280),
                StartPosition = FormStartPosition.CenterScreen,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                BackColor = Color.FromArgb(245, 245, 250)
            };
            
            // ========== 加载图标 ==========
            try
            {
                Assembly assembly = Assembly.GetExecutingAssembly();
                using (Stream stream = assembly.GetManifestResourceStream("OfficeInstaller.icon.ico"))
                {
                    if (stream != null)
                    {
                        kmsForm.Icon = new Icon(stream);
                    }
                }
            }
            catch { }
            
            // ========== 标题 ==========
            Label lblTitle = new Label
            {
                Text = "自定义KMS服务器激活",
                Font = EmbeddedFont.GetFont(14, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 122, 204),
                Location = new Point(0, 20),
                Size = new Size(450, 35),
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleCenter
            };
            kmsForm.Controls.Add(lblTitle);
            
            // ========== KMS服务器地址标签 ==========
            Label lblKMS = new Label
            {
                Text = "KMS服务器地址:",
                Location = new Point(50, 70),
                Size = new Size(350, 25),
                Font = EmbeddedFont.GetFont(10, FontStyle.Regular),
                ForeColor = Color.FromArgb(60, 60, 60)
            };
            kmsForm.Controls.Add(lblKMS);
            
            // ========== KMS服务器输入框 ==========
            TextBox txtKMS = new TextBox
            {
                Location = new Point(50, 95),
                Size = new Size(350, 30),
                Font = EmbeddedFont.GetFont(11, FontStyle.Regular),
                BorderStyle = BorderStyle.FixedSingle,
                // 默认示例地址
                Text = "kms.example.com"
            };
            kmsForm.Controls.Add(txtKMS);
            
            // ========== 提示标签 ==========
            Label lblHint = new Label
            {
                Text = "请输入您的KMS服务器地址，例如：kms.example.com",
                Location = new Point(50, 130),
                Size = new Size(350, 40),
                Font = EmbeddedFont.GetFont(9, FontStyle.Regular),
                ForeColor = Color.FromArgb(100, 100, 100)
            };
            kmsForm.Controls.Add(lblHint);
            
            // ========== 确定按钮 ==========
            Button btnOK = new Button();
            btnOK.Text = "开始激活";
            btnOK.Location = new Point(125, 180);
            btnOK.Size = new Size(200, 45);
            btnOK.Font = EmbeddedFont.GetFont(12, FontStyle.Bold);
            btnOK.ForeColor = Color.White;
            btnOK.FlatStyle = FlatStyle.Flat;
            btnOK.FlatAppearance.BorderSize = 0;
            btnOK.Cursor = Cursors.Hand;
            btnOK.BackColor = Color.Transparent;
            btnOK.DialogResult = DialogResult.OK;
            
            // 按钮圆角半径
            int btnRadius = 12;
            // 自定义绘制按钮
            btnOK.Paint += (sender, e) =>
            {
                Graphics g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;
                
                Rectangle rect = new Rectangle(2, 2, btnOK.Width - 5, btnOK.Height - 5);
                
                using (GraphicsPath path = new GraphicsPath())
                {
                    // 绘制圆角矩形
                    path.AddArc(rect.X, rect.Y, btnRadius * 2, btnRadius * 2, 180, 90);
                    path.AddArc(rect.Right - btnRadius * 2, rect.Y, btnRadius * 2, btnRadius * 2, 270, 90);
                    path.AddArc(rect.Right - btnRadius * 2, rect.Bottom - btnRadius * 2, btnRadius * 2, btnRadius * 2, 0, 90);
                    path.AddArc(rect.X, rect.Bottom - btnRadius * 2, btnRadius * 2, btnRadius * 2, 90, 90);
                    path.CloseFigure();
                    
                    // 紫色渐变填充
                    using (LinearGradientBrush brush = new LinearGradientBrush(
                        rect, Color.FromArgb(156, 39, 176), Color.FromArgb(123, 31, 162), LinearGradientMode.Horizontal))
                    {
                        g.FillPath(brush, path);
                    }
                    
                    // 绘制文本
                    TextRenderer.DrawText(g, btnOK.Text, btnOK.Font, rect, Color.White,
                        TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                }
            };
            
            kmsForm.Controls.Add(btnOK);
            kmsForm.AcceptButton = btnOK;
            
            // 显示对话框并处理结果
            if (kmsForm.ShowDialog() == DialogResult.OK)
            {
                // 获取输入的KMS服务器地址
                string kmsServer = txtKMS.Text.Trim();
                // 验证输入
                if (!string.IsNullOrEmpty(kmsServer) && kmsServer != "kms.example.com")
                {
                    // 调用KMS激活方法
                    ActivateOfficeWithKMS(kmsServer);
                }
                else
                {
                    GlimMessageBox.Show("请输入有效的KMS服务器地址。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }
        
        // ========== 使用自定义KMS服务器激活Office方法 ==========
        // ActivateOfficeWithKMS: 使用指定的KMS服务器激活Office
        // 参数kmsServer: KMS服务器地址
        private async void ActivateOfficeWithKMS(string kmsServer)
        {
            try
            {
                // 提取激活工具
                string activatorPath = EmbeddedResource.ExtractActivator();
                
                // 检查激活工具是否存在
                if (string.IsNullOrEmpty(activatorPath) || !File.Exists(activatorPath))
                {
                    GlimMessageBox.Show("激活工具提取失败，请检查程序完整性。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                
                Logger.Info("开始使用KMS服务器激活 Office: " + kmsServer);
                
                // 配置激活进程
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = activatorPath,
                    // /KMS: 使用KMS激活方式
                    // /KMS-Server: 指定KMS服务器地址
                    Arguments = "/KMS /KMS-Server:" + kmsServer,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    WorkingDirectory = Path.GetTempPath()
                };
                
                using (Process proc = Process.Start(psi))
                {
                    // 异步读取输出
                    string output = await Task.Run(() => proc.StandardOutput.ReadToEnd());
                    string error = await Task.Run(() => proc.StandardError.ReadToEnd());
                    await Task.Run(() => proc.WaitForExit());
                    
                    // 记录日志
                    Logger.Info("KMS激活输出: " + output);
                    if (!string.IsNullOrEmpty(error))
                    {
                        Logger.Error("KMS激活错误: " + error, null);
                    }
                    
                    // 检查结果
                    if (proc.ExitCode == 0)
                    {
                        GlimMessageBox.Show("Office KMS激活完成！\n\nKMS服务器: " + kmsServer, "激活成功");
                        Logger.Info("Office KMS激活成功: " + kmsServer);
                    }
                    else
                    {
                        GlimMessageBox.Show("Office KMS激活可能未完成，请查看日志了解详情。", "激活提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        Logger.Warn("Office KMS激活返回非零退出码: " + proc.ExitCode);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("KMS激活过程出错", ex);
                GlimMessageBox.Show("KMS激活过程出错：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        // ========== Office TsForge激活方法 ==========
        // ActivateOfficeTsForge: 使用TsForge方式激活Office
        private async void ActivateOfficeTsForge()
        {
            try
            {
                // 提取激活工具
                string activatorPath = EmbeddedResource.ExtractActivator();
                
                // 检查激活工具是否存在
                if (string.IsNullOrEmpty(activatorPath) || !File.Exists(activatorPath))
                {
                    GlimMessageBox.Show("激活工具提取失败，请检查程序完整性。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                
                Logger.Info("开始使用TsForge激活 Office...");
                
                // 配置激活进程
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = activatorPath,
                    // /TsForge: 使用TsForge激活方式
                    Arguments = "/TsForge",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    WorkingDirectory = Path.GetTempPath()
                };
                
                using (Process proc = Process.Start(psi))
                {
                    // 异步读取输出
                    string output = await Task.Run(() => proc.StandardOutput.ReadToEnd());
                    string error = await Task.Run(() => proc.StandardError.ReadToEnd());
                    await Task.Run(() => proc.WaitForExit());
                    
                    // 记录日志
                    Logger.Info("TsForge激活输出: " + output);
                    if (!string.IsNullOrEmpty(error))
                    {
                        Logger.Error("TsForge激活错误: " + error, null);
                    }
                    
                    // 检查结果
                    if (proc.ExitCode == 0)
                    {
                        GlimMessageBox.Show("Office TsForge激活完成！", "激活成功");
                        Logger.Info("Office TsForge激活成功");
                    }
                    else
                    {
                        GlimMessageBox.Show("Office TsForge激活可能未完成，请查看日志了解详情。", "激活提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        Logger.Warn("Office TsForge激活返回非零退出码: " + proc.ExitCode);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("TsForge激活过程出错", ex);
                GlimMessageBox.Show("TsForge激活过程出错：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        // ========== Windows HWID激活方法 ==========
        // ActivateWindowsHWID: 使用HWID方式激活Windows
        // HWID: Hardware ID，基于硬件ID的数字权利激活
        private async void ActivateWindowsHWID()
        {
            try
            {
                // 提取激活工具
                string activatorPath = EmbeddedResource.ExtractActivator();
                
                // 检查激活工具是否存在
                if (string.IsNullOrEmpty(activatorPath) || !File.Exists(activatorPath))
                {
                    GlimMessageBox.Show("激活工具提取失败，请检查程序完整性。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                
                Logger.Info("开始使用HWID激活 Windows...");
                
                // 配置激活进程
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = activatorPath,
                    // /HWID: 使用HWID数字权利激活
                    Arguments = "/HWID",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    WorkingDirectory = Path.GetTempPath()
                };
                
                using (Process proc = Process.Start(psi))
                {
                    // 异步读取输出
                    string output = await Task.Run(() => proc.StandardOutput.ReadToEnd());
                    string error = await Task.Run(() => proc.StandardError.ReadToEnd());
                    await Task.Run(() => proc.WaitForExit());
                    
                    // 记录日志
                    Logger.Info("HWID激活输出: " + output);
                    if (!string.IsNullOrEmpty(error))
                    {
                        Logger.Error("HWID激活错误: " + error, null);
                    }
                    
                    // 检查结果
                    if (proc.ExitCode == 0)
                    {
                        GlimMessageBox.Show("Windows HWID激活完成！\n\n这是数字权利激活，重装系统后自动激活。", "激活成功");
                        Logger.Info("Windows HWID激活成功");
                    }
                    else
                    {
                        GlimMessageBox.Show("Windows HWID激活可能未完成，请查看日志了解详情。", "激活提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        Logger.Warn("Windows HWID激活返回非零退出码: " + proc.ExitCode);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("HWID激活过程出错", ex);
                GlimMessageBox.Show("HWID激活过程出错：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    // ========== Toast通知类 ==========
    // ToastNotification: 右下角弹窗提示
    // 用于显示操作结果、警告等信息
    public class ToastNotification : Form
    {
        // Timer: 计时器控件
        // fadeTimer: 淡入淡出计时器
        private Timer fadeTimer;
        // closeTimer: 自动关闭计时器
        private Timer closeTimer;
        // 淡入淡出时间（毫秒）
        private int fadeDuration = 300;
        // 显示时间（毫秒）
        private int displayDuration = 5000;
        // 透明度变化步长
        private double opacityStep = 0.1;
        
        // 构造函数：创建Toast通知
        // 参数title: 标题
        // 参数message: 消息内容
        // 参数type: Toast类型（Info/Success/Warning/Error）
        public ToastNotification(string title, string message, ToastType type = ToastType.Info)
        {
            // 无边框窗体
            this.FormBorderStyle = FormBorderStyle.None;
            // 手动定位
            this.StartPosition = FormStartPosition.Manual;
            this.Size = new Size(350, 100);
            // 不在任务栏显示
            this.ShowInTaskbar = false;
            // 始终置顶
            this.TopMost = true;
            // 初始透明度为0（完全透明）
            this.Opacity = 0;
            
            // ========== 计算位置 - 右下角 ==========
            // Screen.PrimaryScreen: 主显示器
            Screen screen = Screen.PrimaryScreen;
            this.Location = new Point(
                // WorkingArea: 工作区域（排除任务栏）
                screen.WorkingArea.Width - this.Width - 20,
                screen.WorkingArea.Height - this.Height - 20
            );
            
            // ========== 设置背景色 ==========
            Color backColor = Color.FromArgb(50, 50, 50);
            Color borderColor = Color.FromArgb(0, 122, 204);
            
            // 根据类型设置边框颜色
            switch (type)
            {
                case ToastType.Success:
                    // 绿色：成功
                    borderColor = Color.FromArgb(76, 175, 80);
                    break;
                case ToastType.Warning:
                    // 橙色：警告
                    borderColor = Color.FromArgb(255, 152, 0);
                    break;
                case ToastType.Error:
                    // 红色：错误
                    borderColor = Color.FromArgb(244, 67, 54);
                    break;
            }
            
            this.BackColor = backColor;
            
            // ========== 绘制圆角和边框 ==========
            this.Paint += (s, e) => {
                Graphics g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;
                
                // 绘制圆角矩形
                using (GraphicsPath path = new GraphicsPath())
                {
                    int radius = 8;
                    Rectangle rect = new Rectangle(0, 0, this.Width - 1, this.Height - 1);
                    
                    // 四个角的圆弧
                    path.AddArc(rect.X, rect.Y, radius * 2, radius * 2, 180, 90);
                    path.AddArc(rect.Right - radius * 2, rect.Y, radius * 2, radius * 2, 270, 90);
                    path.AddArc(rect.Right - radius * 2, rect.Bottom - radius * 2, radius * 2, radius * 2, 0, 90);
                    path.AddArc(rect.X, rect.Bottom - radius * 2, radius * 2, radius * 2, 90, 90);
                    path.CloseFigure();
                    
                    // 填充背景
                    using (SolidBrush brush = new SolidBrush(backColor))
                    {
                        g.FillPath(brush, path);
                    }
                    
                    // 绘制左边框（彩色指示条）
                    using (SolidBrush brush = new SolidBrush(borderColor))
                    {
                        g.FillRectangle(brush, 0, 0, 5, this.Height);
                    }
                }
            };
            
            // ========== 标题标签 ==========
            Label lblTitle = new Label
            {
                Text = title,
                Font = new Font("Microsoft YaHei UI", 11, FontStyle.Bold),
                ForeColor = Color.White,
                Location = new Point(15, 12),
                Size = new Size(320, 25),
                AutoSize = false
            };
            this.Controls.Add(lblTitle);
            
            // ========== 消息标签 ==========
            Label lblMessage = new Label
            {
                Text = message,
                Font = new Font("Microsoft YaHei UI", 9),
                ForeColor = Color.FromArgb(200, 200, 200),
                Location = new Point(15, 40),
                Size = new Size(320, 50),
                AutoSize = false
            };
            this.Controls.Add(lblMessage);
            
            // ========== 点击关闭 ==========
            // 点击窗体任意位置关闭
            this.Click += (s, e) => this.Close();
            lblTitle.Click += (s, e) => this.Close();
            lblMessage.Click += (s, e) => this.Close();
            
            // ========== 淡入效果 ==========
            fadeTimer = new Timer();
            // Interval: 计时器间隔（毫秒）
            fadeTimer.Interval = 30;
            // Tick: 计时器触发事件
            fadeTimer.Tick += (s, e) => {
                // 逐渐增加透明度
                if (this.Opacity < 1)
                {
                    this.Opacity += opacityStep;
                }
                else
                {
                    // 淡入完成，停止计时器
                    fadeTimer.Stop();
                    // 启动关闭计时器
                    closeTimer = new Timer();
                    closeTimer.Interval = displayDuration;
                    closeTimer.Tick += (sender, args) => {
                        // 停止关闭计时器
                        closeTimer.Stop();
                        // 开始淡出
                        FadeOut();
                    };
                    closeTimer.Start();
                }
            };
        }
        
        // ========== 淡出效果方法 ==========
        // FadeOut: 逐渐降低透明度并关闭窗体
        private void FadeOut()
        {
            // 创建淡出计时器
            Timer fadeOutTimer = new Timer();
            fadeOutTimer.Interval = 30;
            fadeOutTimer.Tick += (s, e) => {
                // 逐渐降低透明度
                if (this.Opacity > 0)
                {
                    this.Opacity -= opacityStep;
                }
                else
                {
                    // 完全透明后停止计时器并关闭窗体
                    fadeOutTimer.Stop();
                    this.Close();
                }
            };
            fadeOutTimer.Start();
        }
        
        // ========== 显示通知方法 ==========
        // ShowToast: 显示Toast通知
        public void ShowToast()
        {
            // 显示窗体
            this.Show();
            // 启动淡入计时器
            fadeTimer.Start();
        }
    }
    
    // ========== Toast通知类型枚举 ==========
    // ToastType: 定义Toast通知的类型
    public enum ToastType
    {
        // Info: 信息提示（蓝色）
        Info,
        // Success: 成功提示（绿色）
        Success,
        // Warning: 警告提示（橙色）
        Warning,
        // Error: 错误提示（红色）
        Error
    }

    // ========== 版权页面类 ==========
    // CopyrightWindow: 显示软件版权信息的窗口
    // 公开测试版本，带水印
    public class CopyrightWindow : Form
    {
        // 构造函数：初始化版权窗口
        public CopyrightWindow()
        {
            this.Text = "关于 Glim Office Installer";
            this.Size = new Size(600, 550);
            // CenterParent: 在父窗口中央显示
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = Color.White;

            // ========== 加载图标 ==========
            try
            {
                Assembly assembly = Assembly.GetExecutingAssembly();
                using (Stream stream = assembly.GetManifestResourceStream("OfficeInstaller.icon.ico"))
                {
                    if (stream != null)
                    {
                        this.Icon = new Icon(stream);
                    }
                }
            }
            catch { }

            // ========== 标题标签 ==========
            Label lblTitle = new Label
            {
                Text = "Glim Office Installer",
                Font = new Font("Microsoft YaHei UI", 20, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(0, 25),
                Size = new Size(600, 45),
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleCenter
            };
            this.Controls.Add(lblTitle);

            // ========== 版本标签 ==========
            Label lblVersion = new Label
            {
                Text = "版本 1.0.0 (公开测试版)",
                Font = new Font("Microsoft YaHei UI", 11),
                ForeColor = Color.FromArgb(100, 100, 100),
                Location = new Point(0, 75),
                Size = new Size(600, 25),
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleCenter
            };
            this.Controls.Add(lblVersion);

            // ========== 分隔线 ==========
            Panel linePanel = new Panel
            {
                Location = new Point(50, 110),
                Size = new Size(500, 1),
                BackColor = Color.FromArgb(200, 200, 200)
            };
            this.Controls.Add(linePanel);

            // ========== 版权信息文本框 ==========
            // 使用一个TextBox，带滚动条
            TextBox txtContent = new TextBox
            {
                // \r\n: Windows换行符
                Text = "© 2025-2026 GlimStudio. All Rights Reserved.\r\n\r\n" +
                       "本软件由 GlimStudio 开发团队精心打造，旨在帮助用户快速、便捷地部署 Microsoft Office 办公套件。\r\n\r\n" +
                       "【软件功能】\r\n" +
                       "• 支持 Office 2024 LTSC 企业长期服务版\r\n" +
                       "• 支持 Microsoft 365 专业增强版\r\n" +
                       "• 一键下载、安装、激活全流程自动化\r\n" +
                       "• 深度清理旧版本 Office 残留\r\n" +
                       "• 智能组件选择，按需安装\r\n\r\n" +
                       "【免责声明】\r\n" +
                       "本软件仅供学习和研究使用，请确保您拥有合法的 Microsoft Office 授权。使用本软件所产生的任何直接或间接损失，开发者不承担任何责任。请遵守当地法律法规，支持正版软件。\r\n\r\n" +
                       "【开源组件】\r\n" +
                       "本软件使用了以下开源组件：\r\n" +
                       "• Microsoft Office Deployment Tool (ODT)\r\n" +
                       "• MAS (Microsoft Activation Scripts)\r\n\r\n" +
                       "【技术支持】\r\n" +
                       "如有问题或建议，请通过以下方式联系我们：\r\n" +
                       "• 官方网站：建设中\r\n" +
                       "• 反馈邮箱：wumiao.tech228@gmail.com (点击可复制)\r\n\r\n" +
                       "感谢您选择 Glim Office Installer！",
                Font = new Font("Microsoft YaHei UI", 10),
                ForeColor = Color.FromArgb(80, 80, 80),
                Location = new Point(50, 120),
                Size = new Size(500, 320),
                Multiline = true,
                // Vertical: 垂直滚动条
                ScrollBars = ScrollBars.Vertical,
                BorderStyle = BorderStyle.None,
                BackColor = Color.White,
                ReadOnly = true,
                // TabStop: 是否可以通过Tab键获取焦点
                TabStop = false,
                // WordWrap: 自动换行
                WordWrap = true
            };
            
            // ========== 点击邮箱时打开邮件客户端 ==========
            txtContent.Click += (s, e) => {
                // SelectionStart: 光标位置
                int pos = txtContent.SelectionStart;
                string text = txtContent.Text;
                // IndexOf: 查找子字符串位置
                int emailStart = text.IndexOf("wumiao.tech228@gmail.com");
                // 检查是否点击了邮箱地址
                if (emailStart >= 0 && pos >= emailStart && pos <= emailStart + "wumiao.tech228@gmail.com".Length)
                {
                    try
                    {
                        // mailto: 打开默认邮件客户端
                        Process.Start("mailto:wumiao.tech228@gmail.com?subject=Glim Office Installer 反馈");
                    }
                    catch (Exception ex)
                    {
                        GlimMessageBox.Show("无法打开邮件客户端：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            };
            
            // ========== 鼠标移动时改变光标 ==========
            txtContent.MouseMove += (s, e) => {
                // GetCharIndexFromPosition: 根据坐标获取字符索引
                int pos = txtContent.GetCharIndexFromPosition(e.Location);
                string text = txtContent.Text;
                int emailStart = text.IndexOf("wumiao.tech228@gmail.com");
                // 在邮箱地址上显示手形光标
                if (emailStart >= 0 && pos >= emailStart && pos <= emailStart + "wumiao.tech228@gmail.com".Length)
                {
                    txtContent.Cursor = Cursors.Hand;
                }
                else
                {
                    // 其他位置显示I形光标（文本选择）
                    txtContent.Cursor = Cursors.IBeam;
                }
            };
            
            this.Controls.Add(txtContent);

            // ========== 底部分隔线 ==========
            Panel linePanel2 = new Panel
            {
                Location = new Point(50, 455),
                Size = new Size(500, 1),
                BackColor = Color.FromArgb(200, 200, 200)
            };
            this.Controls.Add(linePanel2);

            // ========== 确定按钮 ==========
            Button btnOK = new Button
            {
                Text = "确定",
                Font = new Font("Microsoft YaHei UI", 10),
                Location = new Point(250, 475),
                Size = new Size(100, 35),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White
            };
            btnOK.FlatAppearance.BorderSize = 0;
            // 点击关闭窗口
            btnOK.Click += (s, e) => this.Close();
            this.Controls.Add(btnOK);

            // ========== 绘制水印 ==========
            // Paint事件：绘制"公开测试版"水印
            this.Paint += CopyrightWindow_Paint;
        }

        // ========== 版权窗口绘制事件 ==========
        // CopyrightWindow_Paint: 绘制半透明水印
        private void CopyrightWindow_Paint(object sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            // ========== 绘制半透明水印文字 ==========
            string watermark = "公开测试版";
            using (Font watermarkFont = new Font("Microsoft YaHei UI", 60, FontStyle.Bold))
            {
                // MeasureString: 测量字符串尺寸
                SizeF textSize = g.MeasureString(watermark, watermarkFont);

                // 设置半透明红色（Alpha=25，几乎透明）
                using (SolidBrush brush = new SolidBrush(Color.FromArgb(25, 255, 0, 0)))
                {
                    // ========== 坐标变换：旋转绘制 ==========
                    // TranslateTransform: 平移坐标系到中心
                    g.TranslateTransform(this.Width / 2, this.Height / 2);
                    // RotateTransform: 旋转坐标系-30度
                    g.RotateTransform(-30);
                    // DrawString: 绘制字符串
                    g.DrawString(watermark, watermarkFont, brush, -textSize.Width / 2, -textSize.Height / 2);
                    // ResetTransform: 重置变换
                    g.ResetTransform();
                }
            }
        }
    }

    // ========== 自定义消息弹窗类 ==========
    // GlimMessageBox: 自定义风格的消息对话框
    // 使用icon.ico图标，支持不同类型的消息图标
    public class GlimMessageBox : Form
    {
        // Glim品牌蓝色
        private static Color glimBlue = Color.FromArgb(0, 122, 204);
        // 对话框结果
        private DialogResult result = DialogResult.None;

        // ========== 静态Show方法（完整版） ==========
        // Show: 显示消息对话框
        // 参数message: 消息内容
        // 参数title: 标题
        // 参数buttons: 按钮类型
        // 参数icon: 图标类型
        // 返回值: DialogResult，用户点击的按钮
        public static DialogResult Show(string message, string title, MessageBoxButtons buttons, MessageBoxIcon icon)
        {
            // using: 确保对话框被正确释放
            using (GlimMessageBox box = new GlimMessageBox(message, title, buttons, icon))
            {
                // ShowDialog: 模态显示对话框
                box.ShowDialog();
                return box.result;
            }
        }

        // ========== 静态Show方法（简化版） ==========
        // Show: 显示简单的信息对话框
        public static void Show(string message, string title)
        {
            // 默认使用OK按钮和信息图标
            Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // ========== 私有构造函数 ==========
        // GlimMessageBox: 创建消息对话框实例
        private GlimMessageBox(string message, string title, MessageBoxButtons buttons, MessageBoxIcon icon)
        {
            this.Text = title;
            this.Size = new Size(450, 220);
            // CenterParent: 在父窗口中央显示
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = Color.White;
            this.AutoScaleMode = AutoScaleMode.Dpi;
            this.AutoScaleDimensions = new SizeF(96F, 96F);

            // ========== 加载图标 ==========
            try
            {
                Assembly assembly = Assembly.GetExecutingAssembly();
                using (Stream stream = assembly.GetManifestResourceStream("OfficeInstaller.icon.ico"))
                {
                    if (stream != null)
                    {
                        this.Icon = new Icon(stream);
                    }
                }
            }
            catch { }

            // ========== 顶部装饰条 ==========
            Panel topBar = new Panel
            {
                Location = new Point(0, 0),
                Size = new Size(450, 5),
                BackColor = glimBlue
            };
            this.Controls.Add(topBar);

            // ========== 图标区域 ==========
            // GetIconText: 根据图标类型获取对应的emoji
            string iconText = GetIconText(icon);
            Label lblIcon = new Label
            {
                Text = iconText,
                // Segoe UI Emoji: Windows表情符号字体
                Font = new Font("Segoe UI Emoji", 28),
                Location = new Point(20, 25),
                Size = new Size(60, 60),
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = Color.Transparent
            };
            this.Controls.Add(lblIcon);

            // ========== 消息内容 ==========
            Label lblMessage = new Label
            {
                Text = message,
                Font = new Font("Microsoft YaHei UI", 10),
                ForeColor = Color.FromArgb(60, 60, 60),
                Location = new Point(90, 30),
                Size = new Size(340, 100),
                AutoSize = false,
                BackColor = Color.Transparent
            };
            this.Controls.Add(lblMessage);

            // ========== 按钮区域 ==========
            int btnY = 150;
            // 根据按钮类型创建不同的按钮组合
            if (buttons == MessageBoxButtons.OK)
            {
                // 只有一个"确定"按钮
                Button btnOK = CreateButton("确定", 175, btnY);
                btnOK.Click += (s, e) => { result = DialogResult.OK; this.Close(); };
                this.Controls.Add(btnOK);
            }
            else if (buttons == MessageBoxButtons.YesNo)
            {
                // "是"和"否"按钮
                Button btnYes = CreateButton("是", 135, btnY);
                btnYes.Click += (s, e) => { result = DialogResult.Yes; this.Close(); };
                this.Controls.Add(btnYes);

                Button btnNo = CreateButton("否", 245, btnY);
                btnNo.Click += (s, e) => { result = DialogResult.No; this.Close(); };
                this.Controls.Add(btnNo);
            }
            else if (buttons == MessageBoxButtons.OKCancel)
            {
                // "确定"和"取消"按钮
                Button btnOK = CreateButton("确定", 135, btnY);
                btnOK.Click += (s, e) => { result = DialogResult.OK; this.Close(); };
                this.Controls.Add(btnOK);

                Button btnCancel = CreateButton("取消", 245, btnY);
                btnCancel.Click += (s, e) => { result = DialogResult.Cancel; this.Close(); };
                this.Controls.Add(btnCancel);
            }

            // ========== 设置默认按钮 ==========
            // OfType<T>: 筛选指定类型的控件
            // FirstOrDefault: 获取第一个元素
            this.AcceptButton = this.Controls.OfType<Button>().FirstOrDefault();
            // LastOrDefault: 获取最后一个元素
            this.CancelButton = this.Controls.OfType<Button>().LastOrDefault();
        }

        // ========== 创建按钮方法 ==========
        // CreateButton: 创建统一风格的按钮
        // 参数text: 按钮文本
        // 参数x: X坐标
        // 参数y: Y坐标
        // 返回值: Button控件
        private Button CreateButton(string text, int x, int y)
        {
            Button btn = new Button
            {
                Text = text,
                Font = new Font("Microsoft YaHei UI", 10),
                Location = new Point(x, y),
                Size = new Size(90, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = glimBlue,
                ForeColor = Color.White,
                Cursor = Cursors.Hand
            };
            btn.FlatAppearance.BorderSize = 0;
            return btn;
        }

        // ========== 获取图标文本方法 ==========
        // GetIconText: 根据消息图标类型返回对应的emoji
        // 参数icon: MessageBoxIcon枚举
        // 返回值: emoji字符串
        private string GetIconText(MessageBoxIcon icon)
        {
            switch (icon)
            {
                case MessageBoxIcon.Information:
                    // 信息图标
                    return "ℹ️";
                case MessageBoxIcon.Warning:
                    // 警告图标
                    return "⚠️";
                case MessageBoxIcon.Error:
                    // 错误图标
                    return "❌";
                case MessageBoxIcon.Question:
                    // 问题图标
                    return "❓";
                default:
                    // 默认信息图标
                    return "ℹ️";
            }
        }
    }
}
