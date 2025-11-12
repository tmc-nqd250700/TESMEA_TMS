using System.IO;
using System.Windows;
using TESMEA_TMS.Helpers;

namespace TESMEA_TMS.Views
{
    /// <summary>
    /// Interaction logic for CalculationView.xaml
    /// </summary>
    public partial class CalculationView : System.Windows.Controls.UserControl
    {
        public CalculationView()
        {
            InitializeComponent();
        }

        private void txtPath_PreviewMouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            //if (radioExcel.IsChecked != true)
            //    return;

            //var ofd = new OpenFileDialog
            //{
            //    Filter = "Excel Files|*.xlsx;*.xls",
            //    Title = "Chọn dữ liệu thông số đo kiểm",
            //};

            //if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            //{
            //    txtPath.Text = ofd.FileName;
            //}
            //e.Handled = true;
        }

       

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "thongsodokiem_template.xlsx");
            if (!System.IO.File.Exists(templatePath))
            {
                MessageBoxHelper.ShowWarning("Không tìm thấy file template");
                return;
            }

            var ofd = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
                FileName = System.IO.Path.GetFileName(templatePath),
                Title = "Chọn nơi lưu file mẫu",
            };
            if (ofd.ShowDialog() == true)
            {
                try
                {
                    System.IO.File.Copy(templatePath, ofd.FileName, true);
                    MessageBoxHelper.ShowSuccess("Tải template thành công");
                }
                catch (Exception ex)
                {
                    MessageBoxHelper.ShowWarning("Không thể tải xuống template " + ex.Message);
                }
            }
        }
    }
}
