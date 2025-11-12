using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using TESMEA_TMS.Configs;
using TESMEA_TMS.Localization;

namespace TESMEA_TMS.Views
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            System.Threading.Thread.CurrentThread.CurrentUICulture = WPFLocalizeExtension.Engine.LocalizeDictionary.Instance.Culture;
            //this.Title = Strings.TitleMainWindow;
            this.Title = "Hệ thống đo kiểm tự động đặc tính quạt công nghiệp";
        }

        private void MenuGroup_Click(object sender, RoutedEventArgs e)
        {
            var clickedItem = sender as MenuItem;
            var parentMenu = ItemsControl.ItemsControlFromItemContainer(clickedItem);
            foreach (var item in parentMenu.Items)
            {
                if (item is MenuItem menuItem && menuItem != clickedItem)
                {
                    menuItem.IsChecked = false;
                }
            }
            clickedItem.IsChecked = true;
        }

        // Logout sẽ call command để restart app
        protected override async void OnClosing(CancelEventArgs e)
        {
            if (Helpers.Validation.IsValidGuid(CurrentUser.Instance.SessionId))
            {
                var result = System.Windows.MessageBox.Show(
                    Strings.Common_ExitTitle,
                    Strings.Common_ExitMessage,
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);

                if (result != MessageBoxResult.Yes)
                {
                    e.Cancel = true;
                    return;
                }
                if (DataContext is ViewModels.MainViewModel vm)
                {
                    e.Cancel = true;
                    try
                    {
                        await vm.ExitApp();
                        base.OnClosing(e);
                    }
                    catch (Exception ex)
                    {
                        base.OnClosing(e);
                    }
                }
            }
            else
            {
                base.OnClosing(e);
            }
        }
    }
}
