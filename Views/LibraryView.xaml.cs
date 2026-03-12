namespace TESMEA_TMS.Views
{
    /// <summary>
    /// Interaction logic for LibraryView.xaml
    /// </summary>
    public partial class LibraryView : System.Windows.Controls.UserControl
    {
        public LibraryView()
        {
            InitializeComponent();
        }

        private void btnToolbox_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            // lấy name của button được click rồi dùng common
            var button = sender as System.Windows.Controls.Button;
            if (button != null)
            {
                var name = button.Name;
                Helpers.Common.ShowMessageBoxHelper(name);
            }
        }
    }
}
