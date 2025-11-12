using System.IO;
using System.Windows;
using System.Windows.Input;
using TESMEA_TMS.Configs;
using TESMEA_TMS.Helpers;

namespace TESMEA_TMS.Views
{
    /// <summary>
    /// Interaction logic for ChooseDbDialog.xaml
    /// </summary>
    public partial class ChooseDbDialog : Window
    {
        public ChooseDbDialog()
        {
            InitializeComponent();
            txtDbPath.Text = UserSetting.Instance.DbPath;
        }

        private void TextBox_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            var dbPath = txtDbPath.Text.Trim();
            var ofd = new OpenFileDialog
            {
                Filter = "SQLite Database (*.db)|*.db",
            };
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (string.IsNullOrEmpty(dbPath))
                {
                    MessageBoxHelper.ShowWarning("Đường dẫn tới Simatic không được để trống");
                    return;
                }
                if (!File.Exists(dbPath))
                {
                    MessageBoxHelper.ShowWarning("Đường dẫn tới Simatic không tồn tại");
                    return;
                }
                var setting = Configs.UserSetting.Load();
                setting.DbPath = dbPath;
                setting.Save();
            }
        }
    }
}
