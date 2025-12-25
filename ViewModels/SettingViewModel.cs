using System.IO;
using System.Windows;
using System.Windows.Input;
using TESMEA_TMS.Configs;
using TESMEA_TMS.DTOs;
using TESMEA_TMS.Helpers;
using TESMEA_TMS.Services;
using MessageBox = System.Windows.MessageBox;

namespace TESMEA_TMS.ViewModels
{
    public class SettingViewModel : ViewModelBase
    {
        private string _simaticPath;
        public string SimaticPath
        {
            get => _simaticPath;
            set
            {
                _simaticPath = value;
                OnPropertyChanged(nameof(SimaticPath));
            }
        }

        private int _timeout;
        public int Timeout
        {
            get => _timeout;
            set
            {
                _timeout = value;
                OnPropertyChanged(nameof(Timeout));
            }
        }

        private string _currentVerson, _lastUpdate;
        public string CurrentVersion
        {
            get => _currentVerson;
            set
            {
                if (_currentVerson != value)
                {
                    _currentVerson = value;
                    OnPropertyChanged(nameof(CurrentVersion));
                }
            }
        }

        public string LastUpdate
        {
            get => _lastUpdate;
            set
            {
                if (_lastUpdate != value)
                {
                    _lastUpdate = value;
                    OnPropertyChanged(nameof(LastUpdate));
                }
            }
        }

        private ChangePasswordDto _changePassword;
        public ChangePasswordDto ChangePassword
        {
            get => _changePassword;
            set
            {
                _changePassword = value;
                OnPropertyChanged(nameof(ChangePassword));
            }
        }
        // Dictionary để lưu giá trị ban đầu
        private Dictionary<string, object> _originalValues = new Dictionary<string, object>();

        // Property để check dirty state
        public bool IsDirty => HasChanges();
        // Commands
        public event Action? RequestClose;
        private UserSetting _userSetting { get; set; }
        public ICommand SaveCommand { get; }
        public ICommand CancelCommand { get; }

        private readonly IAuthenticationService _authenticationService;

        public SettingViewModel(IAuthenticationService authenticationService)
        {
            _authenticationService = authenticationService;
            ChangePassword = new ChangePasswordDto();
            SaveCommand = new ViewModelCommand(ExecuteSaveCommand);
            CancelCommand = new ViewModelCommand(ExecuteCancelCommand);
            _userSetting = UserSetting.Load();
            LoadUserSettings();
        }

        private void LoadUserSettings()
        {
            if(_userSetting == null)
            {
                _userSetting = UserSetting.Load();
            }
            SimaticPath = _userSetting.SimaticPath;
            Timeout = (int)Math.Round(_userSetting.TimeoutMilliseconds / 1000.0);
            ChangePassword = new ChangePasswordDto();


            _originalValues.Clear();
            _originalValues[nameof(SimaticPath)] = SimaticPath;
            _originalValues[nameof(Timeout)] = Timeout;
            _originalValues[nameof(ChangePassword)] = ChangePassword;
        }

        #region Dirty State Management
        private void SaveOriginalValues()
        {
            _originalValues.Clear();
            _originalValues[nameof(SimaticPath)] = SimaticPath;
            _originalValues[nameof(Timeout)] = Timeout;
            _originalValues[nameof(ChangePassword)] = ChangePassword;
        }

        private bool HasChanges()
        {
            if (_originalValues.Count == 0) return false;

            return !Equals(_originalValues[nameof(SimaticPath)], SimaticPath) ||
                   !Equals(_originalValues[nameof(Timeout)], Timeout) ||
                   !Equals(_originalValues[nameof(ChangePassword)], ChangePassword);
        }

        public bool CanClose()
        {
            if (!IsDirty)
            {
                LoadUserSettings();
                return true;
            }

            var result = MessageBox.Show(
                "Bạn có thay đổi chưa lưu. Bạn có muốn lưu không?",
                "Xác nhận",
                MessageBoxButton.YesNoCancel,
                MessageBoxImage.Question);

            switch (result)
            {
                case MessageBoxResult.Yes:
                    ExecuteSaveCommand(null);
                    return true;
                case MessageBoxResult.No:
                    return true;
                case MessageBoxResult.Cancel:
                    return false;
                default:
                    return false;
            }
        }

        #endregion

        private void ExecuteCancelCommand(object obj)
        {
            if (!IsDirty)
            {
                LoadUserSettings();
                RequestClose?.Invoke();
                return;
            }

            var result = MessageBox.Show(
                "Bạn có thay đổi chưa lưu. Bạn có muốn lưu không?",
                "Xác nhận",
                MessageBoxButton.YesNoCancel,
                MessageBoxImage.Question);

            switch (result)
            {
                case MessageBoxResult.Yes:
                    ExecuteSaveCommand(null);
                    break;
                case MessageBoxResult.No:
                    _originalValues.Clear();
                    RequestClose?.Invoke();
                    break;
                case MessageBoxResult.Cancel:
                default:
                    break;
            }
        }

        private async void ExecuteSaveCommand(object obj)
        {
            if (ChangePassword.NewPassword != ChangePassword.ConfirmNewPassword)
            {
                MessageBox.Show("Mật khẩu mới và mật khẩu xác nhận không khớp", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (string.IsNullOrEmpty(SimaticPath))
            {
                MessageBox.Show("Đường dẫn tới Simatic không được để trống", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (!File.Exists(SimaticPath))
            {
                MessageBox.Show("Đường dẫn tới Simatic không tồn tại", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (Timeout <= 0)
            {
                MessageBox.Show("Nhập timeout lớn hơn 0", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (Timeout > 360)
            {
                MessageBox.Show("Nhập timeout nhỏ hơn hoặc bằng 360 giây", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Warning);
            }

            // ChangePassword if have values
            bool changePasswordSuccess = false;
            if (!string.IsNullOrEmpty(ChangePassword.CurrentPassword) && !string.IsNullOrEmpty(ChangePassword.CurrentPassword) && !string.IsNullOrEmpty(ChangePassword.CurrentPassword))
            {
                if (!Validation.IsValidPassword(ChangePassword.CurrentPassword))
                {
                    MessageBox.Show("Mật khẩu hiện tại không hợp lệ", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                if (!Validation.IsValidPassword(ChangePassword.NewPassword))
                {
                    MessageBox.Show("Mật khẩu mới không hợp lệ", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                if (!Validation.IsValidPassword(ChangePassword.ConfirmNewPassword))
                {
                    MessageBox.Show("Mật khẩu xác nhận không hợp lệ", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                changePasswordSuccess = await _authenticationService.ChangePasswordAsync(ChangePassword);
                if (!changePasswordSuccess)
                {
                    MessageBox.Show("Có lỗi trong quá trình đổi mật khẩu", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
            }
            else
            {
                changePasswordSuccess = true;
            }

            // save new user setting
            if (changePasswordSuccess)
            {
                if (_userSetting == null)
                {
                    _userSetting = UserSetting.Load();
                }
                _userSetting.SimaticPath = SimaticPath;
                _userSetting.TimeoutMilliseconds = Timeout * 1000;
                _userSetting.Save();
                LoadUserSettings();
                RequestClose?.Invoke();
            }
        }
    }
}
