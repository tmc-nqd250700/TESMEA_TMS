using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using TESMEA_TMS.Helpers;

namespace TESMEA_TMS.Views
{
    /// <summary>
    /// Interaction logic for ConfirmSaveDialog.xaml
    /// </summary>
    public partial class ConfirmSaveDialog : Window, INotifyPropertyChanged
    {
        private string _dialogTitle = "Nhập thông tin";
        private string _message;
        private string _inputText;

        public string DialogTitle
        {
            get => _dialogTitle;
            set
            {
                _dialogTitle = value;
                OnPropertyChanged();
            }
        }

        public string Message
        {
            get => _message;
            set
            {
                _message = value;
                OnPropertyChanged();
            }
        }

        public string InputText
        {
            get => _inputText;
            set
            {
                _inputText = value;
                OnPropertyChanged();
            }
        }

        public ConfirmSaveDialog()
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
