using System.Text.RegularExpressions;
using System.Windows.Controls;
using System.Windows.Input;
using TESMEA_TMS.Configs;
using TESMEA_TMS.Helpers;
using TESMEA_TMS.ViewModels;

namespace TESMEA_TMS.Views
{
    /// <summary>
    /// Interaction logic for ProjectView.xaml
    /// </summary>
    public partial class ProjectView : System.Windows.Controls.UserControl
    {
        private ProjectViewModel? _vm;

        public ProjectView()
        {
            InitializeComponent();
            System.Threading.Thread.CurrentThread.CurrentUICulture = WPFLocalizeExtension.Engine.LocalizeDictionary.Instance.Culture;
            _vm = DataContext as ProjectViewModel;
            this.Loaded += ProjectView_Loaded;
        }

        private void ProjectView_Loaded(object sender, System.Windows.RoutedEventArgs e)
        {
            
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
                tb.Text = "";
                tb.CaretIndex = tb.Text.Length;
            }
        }

        private void TextBox_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            var dialog = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = "Chọn thư mục lưu dự án",
                ShowNewFolderButton = false,
                InitialDirectory = UserSetting.GetLocalAppPath()
            };
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (_vm != null)
                {
                    var selectedPath = dialog.SelectedPath;
                    var tenDuAn = _vm.ThamSo.TenDuAn;

                    if (!string.IsNullOrWhiteSpace(tenDuAn))
                    {
                        // Tạo thư mục mới với tên dự án
                        var projectFolder = System.IO.Path.Combine(selectedPath, tenDuAn);
                        if (!System.IO.Directory.Exists(projectFolder))
                        {
                            System.IO.Directory.CreateDirectory(projectFolder);
                        }
                        _vm.ThamSo.DuongDanLuuDuAn = projectFolder;
                    }
                    else
                    {
                        _vm.ThamSo.DuongDanLuuDuAn = selectedPath;
                    }
                    _vm.OnPropertyChanged(nameof(_vm.ThamSo));
                }
            }
            e.Handled = true;
        }

        private void btnPw_i_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            if (UserSetting.Instance.Language == "en")
                MessageBoxHelper.ShowInformation("Motor power is mentioned on the nameplate of the motor");
            else
                MessageBoxHelper.ShowInformation("Công suất động cơ là công suất in trên nhãn của động cơ");
        }

        private void btnSpeed_i_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            if (UserSetting.Instance.Language == "en")
                MessageBoxHelper.ShowInformation("Motor speed is mentioned on the nameplate of the motor");
            else
                MessageBoxHelper.ShowInformation("Tốc độ động cơ là tốc độ in trên nhãn của động cơ");
        }

        private void btnCosphi_i_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            if (UserSetting.Instance.Language == "en")
                MessageBoxHelper.ShowInformation("Cos (phi) is mentioned on the nameplate of the motor");
            else
                MessageBoxHelper.ShowInformation("Hệ số Cos (phi) là hệ số Cos (phi) in trên nhãn của động cơ");
        }

        private void btnEff_i_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            if (UserSetting.Instance.Language == "en")
                MessageBoxHelper.ShowInformation("Motor efficiency is mentioned on the nameplate of the motor");
            else
                MessageBoxHelper.ShowInformation("Hiệu suất động cơ là hiệu suất in trên nhãn của động cơ");
        }

        private void btnIdm_i_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            if (UserSetting.Instance.Language == "en")
                MessageBoxHelper.ShowInformation("Rated current is mentioned on the nameplate of the motor");
            else
                MessageBoxHelper.ShowInformation("Dòng điện định mức là dòng điện in trên nhãn của động cơ");
        }

        private void btnUdm_i_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            if (UserSetting.Instance.Language == "en")
                MessageBoxHelper.ShowInformation("Motor voltage is mentioned on the nameplate of the motor");
            else
                MessageBoxHelper.ShowInformation("Điện áp động cơ là điện áp in trên nhãn của động cơ");
        }

        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_vm != null)
            {
               _vm.OnLibraryChanged();
            }
        }
    }
}
