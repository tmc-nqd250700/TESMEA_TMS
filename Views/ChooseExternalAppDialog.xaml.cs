using System.Windows;
using System.Windows.Input;

namespace TESMEA_TMS.Views
{
    /// <summary>
    /// Interaction logic for ChooseExternalAppDialog.xaml
    /// </summary>
    public partial class ChooseExternalAppDialog : Window
    {
        private ViewModels.ChooseExternalAppViewModel? _vm;
        public ChooseExternalAppDialog()
        {
            InitializeComponent();
            _vm = DataContext as ViewModels.ChooseExternalAppViewModel;
        }

        private void TextBox_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            var ofd = new OpenFileDialog
            {
                //Filter = "Executable Files (*.exe)|*.exe",
                Title = "Chọn file dự án Simetric/PLC"
            };
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (_vm != null)
                {
                    _vm.SimaticPath = ofd.FileName;
                }
            }
        }

        private void txtWinccPath_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            var ofd = new OpenFileDialog
            {
                //Filter = "Executable Files (*.exe)|*.exe",
                Title = "Chọn file thực thi WinCC"
            };
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (_vm != null)
                {
                    _vm.WinCCPath = ofd.FileName;
                }
            }
        }
    }
}
