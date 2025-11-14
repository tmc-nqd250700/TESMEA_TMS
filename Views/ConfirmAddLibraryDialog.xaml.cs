using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using TESMEA_TMS.Helpers;

namespace TESMEA_TMS.Views
{
    /// <summary>
    /// Interaction logic for ConfirmAddLibraryDialog.xaml
    /// </summary>
    public partial class ConfirmAddLibraryDialog : Window, INotifyPropertyChanged
    {
        private string _inputText;


        public string InputText
        {
            get => _inputText;
            set
            {
                _inputText = value;
                OnPropertyChanged();
            }
        }

        public ConfirmAddLibraryDialog()
        {
            InitializeComponent();
            DataContext = this;
            Loaded += (s, e) => InputTextBox.Focus();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(InputText))
            {
                MessageBoxHelper.ShowWarning("Vui lòng nhập thông tin!");
                return;
            }

            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
