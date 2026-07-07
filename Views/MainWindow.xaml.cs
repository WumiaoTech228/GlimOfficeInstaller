using System.Windows;
using GOI.Services;
using GOI.ViewModels;

namespace GOI.Views
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            var installService = new InstallService();
            var vm = new MainViewModel(installService);

            DataContext = vm;
        }

        protected override void OnSourceInitialized(System.EventArgs e)
        {
            base.OnSourceInitialized(e);

            // 精准计算操作系统窗口边框（Chrome）和标题栏的大小差异
            double widthDiff = this.ActualWidth - ((FrameworkElement)this.Content).ActualWidth;
            double heightDiff = this.ActualHeight - ((FrameworkElement)this.Content).ActualHeight;

            // 动态计算目标内容区域（Client Area）的大小，使其占屏幕面积的 1/7
            double screenWidth = SystemParameters.PrimaryScreenWidth;
            double screenHeight = SystemParameters.PrimaryScreenHeight;
            double targetArea = (screenWidth * screenHeight) / 7.0;
            double aspectRatio = 720.0 / 680.0;

            double targetClientWidth = System.Math.Sqrt(targetArea * aspectRatio);
            double targetClientHeight = targetClientWidth / aspectRatio;

            // 最终窗口大小 = 目标内容大小 + 边框/标题栏大小
            this.Width = targetClientWidth + widthDiff;
            this.Height = targetClientHeight + heightDiff;
        }
    }
}
