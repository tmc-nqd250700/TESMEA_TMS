using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using TESMEA_TMS.Helpers;

namespace TESMEA_TMS.Views
{
    /// <summary>
    /// Interaction logic for ConfirmAddScenarioDialog.xaml
    /// </summary>
    public partial class ConfirmAddScenarioDialog : Window, INotifyPropertyChanged
    {
        private string _inputText;
        private float _standardDeviation = 0.28f;
        private float _timeRange = 20;
        public string InputText
        {
            get => _inputText;
            set
            {
                _inputText = value;
                OnPropertyChanged();
            }
        }

        public float StandardDeviation
        {
            get => _standardDeviation;
            set
            {
                _standardDeviation = value;
                OnPropertyChanged();
            }
        }

        public float TimeRange
        {
            get => _timeRange;
            set
            {
                _timeRange = value;
                OnPropertyChanged();
            }
        }

        public ConfirmAddScenarioDialog()
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
