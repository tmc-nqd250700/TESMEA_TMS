using System.Text.RegularExpressions;
using System.Windows.Controls;
using System.Windows.Input;
using TESMEA_TMS.ViewModels;

namespace TESMEA_TMS.Views
{
    /// <summary>
    /// Interaction logic for SettingDialog.xaml
    /// </summary>
    public partial class SettingDialog : System.Windows.Controls.UserControl
    {
        private SettingViewModel? _vm;
        public SettingDialog()
        {
            InitializeComponent();
            _vm = DataContext as SettingViewModel;
        }

        private static readonly Regex _regex = new Regex("[^0-9.,-]+");

        private void TextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = _regex.IsMatch(e.Text);
        }

        private void TextBox_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Space)
                e.Handled = true;
            if (e.Key == Key.V && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
                e.Handled = true;
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            var tb = sender as System.Windows.Controls.TextBox;
            if (tb != null && string.IsNullOrWhiteSpace(tb.Text))
            {
                tb.Text = "";
                tb.CaretIndex = tb.Text.Length;
            }
        }

        private void TextBox_TextChanged1(object sender, TextChangedEventArgs e)
        {
            var tb = sender as System.Windows.Controls.TextBox;
            if (tb != null && string.IsNullOrWhiteSpace(tb.Text))
            {
                tb.Text = "0";
                tb.CaretIndex = tb.Text.Length;
            }
        }

        private void TextBox_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            var ofd = new OpenFileDialog
            {
                //Filter = "Executable Files (*.exe)|*.exe",
                Title = "Chọn file thực thi Simetric/PLC"
            };
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                if (_vm != null)
                {
                    _vm.SimaticPath = ofd.FileName;
                }
            }
        }

        private void PasswordBox_PasswordChanged(object sender, System.Windows.RoutedEventArgs e)
        {
            if(_vm != null)
            {
                switch ((PasswordBox)sender)
                {
                    case { Name: "OldPasswordBox" }:
                        _vm.ChangePassword.CurrentPassword = ((PasswordBox)sender).Password;
                        break;
                    case { Name: "NewPasswordBox" }:
                        _vm.ChangePassword.NewPassword = ((PasswordBox)sender).Password;
                        break;
                    case { Name: "ConfirmPasswordBox" }:
                        _vm.ChangePassword.ConfirmNewPassword = ((PasswordBox)sender).Password;
                        break;
                    default:
                        break;
                }
            }
            return;
        }
    }
}
