using System.Windows;
using GOI.Services;
using GOI.ViewModels;
using GOI.Models;

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

            // Dynamically translate the built-in Settings item
            Loaded += (s, e) =>
            {
                if (NavView.SettingsItem is iNKORE.UI.WPF.Modern.Controls.NavigationViewItem settingsItem)
                {
                    settingsItem.Content = vm.Loc.NavSettings;
                }
            };
            vm.PropertyChanged += (s, e) =>
            {
                if (string.IsNullOrEmpty(e.PropertyName) || e.PropertyName == "Loc")
                {
                    if (NavView.SettingsItem is iNKORE.UI.WPF.Modern.Controls.NavigationViewItem settingsItem)
                    {
                        settingsItem.Content = vm.Loc.NavSettings;
                    }
                }
            };
        }



        private void NavigationView_SelectionChanged(object sender, iNKORE.UI.WPF.Modern.Controls.NavigationViewSelectionChangedEventArgs e)
        {
            var vm = DataContext as MainViewModel;
            if (vm == null) return;

            if (e.IsSettingsSelected)
            {
                vm.CurrentProductType = ProductType.Settings;
            }
            else if (e.SelectedItem is iNKORE.UI.WPF.Modern.Controls.NavigationViewItem item)
            {
                switch (item.Tag?.ToString())
                {
                    case "MsOffice":
                        vm.CurrentProductType = ProductType.MsOffice;
                        break;
                    case "Wps":
                        vm.CurrentProductType = ProductType.Wps;
                        break;
                    case "Yozo":
                        vm.CurrentProductType = ProductType.Yozo;
                        break;
                    case "OnlyOffice":
                        vm.CurrentProductType = ProductType.OnlyOffice;
                        break;
                    case "LibreOffice":
                        vm.CurrentProductType = ProductType.LibreOffice;
                        break;
                }
            }
        }
    }
}
