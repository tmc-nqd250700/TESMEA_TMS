using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using TESMEA_TMS.ViewModels;

namespace TESMEA_TMS.Views
{
    /// <summary>
    /// Interaction logic for LoginWindow.xaml
    /// </summary>
    /// 
    public partial class LoginWindow : Window
    {
        private LoginViewModel? _vm;
        public LoginWindow()
        {
            InitializeComponent();
            Loaded += LoginWindow_Loaded;
        }

        private void LoginWindow_Loaded(object sender, RoutedEventArgs e)
        {
            _vm = DataContext as LoginViewModel;
        }

        private void PasswordBox_PasswordChanged(object sender, RoutedEventArgs e)
        {
            if (_vm != null)
            {
                _vm.Password = ((PasswordBox)sender).Password.Trim();
            }
                
        }

        private void PasswordBox_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                // Do not set e.Handled = true;
                // Let the event bubble up so the default button is triggered
            }
        }

        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_vm != null)
                _vm.ExecuteSelectLanguageCommand();
        }
    }
}
