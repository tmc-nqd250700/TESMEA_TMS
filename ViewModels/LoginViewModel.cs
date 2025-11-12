using MaterialDesignThemes.Wpf;
using System.Windows.Input;
using TESMEA_TMS.Configs;
using TESMEA_TMS.Services;
using Application = System.Windows.Application;

namespace TESMEA_TMS.ViewModels
{
    public class LoginViewModel :ViewModelBase
    {
        //Fields
        private string _userName, _password;
        private string _errorMessage;

        //Properties
        public string UserName
        {
            get
            {
                return _userName;
            }

            set
            {
                _userName = value;
                OnPropertyChanged(nameof(UserName));
            }
        }

        public string Password
        {
            get
            {
                return _password;
            }

            set
            {
                _password = value;
                OnPropertyChanged(nameof(Password));
            }
        }

        public string ErrorMessage
        {
            get
            {
                return _errorMessage;
            }

            set
            {
                _errorMessage = value;
                OnPropertyChanged(nameof(ErrorMessage));
            }
        }

        private bool _isRememberMe;
        public bool IsRememberMe
        {
            get => _isRememberMe;
            set
            {
                if (_isRememberMe != value)
                {
                    _isRememberMe = value;
                    OnPropertyChanged(nameof(IsRememberMe));
                }
            }
        }
        public ICommand LoginCommand { get; }
        public ICommand RecoverPasswordCommand { get; }
        public ICommand ShowPasswordCommand { get; }
        public ICommand RememberPasswordCommand { get; }
        public ICommand ExitCommand { get; }

        // Repository 

        private readonly IAuthenticationService _authenticationService;
        private readonly IAppNavigationService _appNavigationService;
        public LoginViewModel(IAuthenticationService authenticationService, IAppNavigationService appNavigationService)
        {
            _authenticationService = authenticationService;
            _appNavigationService = appNavigationService;
            LoginCommand = new ViewModelCommand(CanExecuteLoginCommand, ExecuteLoginCommand);
            RecoverPasswordCommand = new ViewModelCommand(_ => ExecuteRecoverPassCommand("", ""));
            ExitCommand = new ViewModelCommand(_ => _appNavigationService.ExitApplication());

            var setting = UserSetting.Load();
            UserName = setting.LastUserName;
            IsRememberMe = !string.IsNullOrEmpty(UserName);
          
        }

        private bool CanExecuteLoginCommand(object obj)
        {
            bool validData;
            if (string.IsNullOrWhiteSpace(UserName) || UserName.Length < 3 ||
                Password == null || Password.Length < 3)
                validData = false;
            else
                validData = true;
            return validData;
        }

        private void ExecuteLoginCommand(object obj)
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                var splash = new Views.CustomControls.ProgressSplashContent
                {
                    DataContext = new ProgressSplashViewModel
                    {
                        Message = "Đang đăng nhập, vui lòng chờ...",
                        IsIndeterminate = true
                    }
                };
                DialogHost.Show(splash, "LoginDialogHost");
            });

            _ = Task.Run(async () =>
            {
                try
                {
                    var isSuccess = await _authenticationService.LoginAsync(UserName, Password);
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        if (DialogHost.IsDialogOpen("LoginDialogHost"))
                            DialogHost.Close("LoginDialogHost");

                        if (isSuccess)
                        {
                            if (IsRememberMe)
                            {
                                UserSetting.Instance.LastUserName = UserName;
                                UserSetting.Instance.Save();
                            }
                            _appNavigationService.CurrentUser = CurrentUser.Instance;
                            _appNavigationService.IsLoggedIn = true;
                            _appNavigationService.ShowMainWindow();
                        }
                        else
                        {
                            ErrorMessage = "Tên đăng nhập hoặc mật khẩu không chính xác";
                        }
                    });
                }
                catch (Exception ex)
                {
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        if (DialogHost.IsDialogOpen("LoginDialogHost"))
                            DialogHost.Close("LoginDialogHost");

                        ErrorMessage = $"Lỗi: {ex.Message}";
                    });
                }
            });
        }

        private void ExecuteRecoverPassCommand(string username, string email)
        {
            throw new NotImplementedException();
        }
    }
}
