using System.ComponentModel;

namespace TESMEA_TMS.ViewModels
{
    public class ProgressSplashViewModel : INotifyPropertyChanged
    {
        private string _message = "Đang xử lý, vui lòng chờ...";
        private bool _isIndeterminate = true;

        public string Message
        {
            get => _message;
            set
            {
                _message = value;
                OnPropertyChanged(nameof(Message));
            }
        }

        public bool IsIndeterminate
        {
            get => _isIndeterminate;
            set
            {
                _isIndeterminate = value;
                OnPropertyChanged(nameof(IsIndeterminate));
            }
        }

        private double _value;
        public double Value
        {
            get => _value;
            set
            {
                _value = value;
                OnPropertyChanged(nameof(Value));
            }
        }

        private double _maximum = 100;
        public double Maximum
        {
            get => _maximum;
            set
            {
                _maximum = value;
                OnPropertyChanged(nameof(Maximum));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
