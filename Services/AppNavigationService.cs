
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using TESMEA_TMS.Configs;
using TESMEA_TMS.ViewModels;
using TESMEA_TMS.Views;
using Application = System.Windows.Application;

namespace TESMEA_TMS.Services
{
    public interface IAppNavigationService : INotifyPropertyChanged
    {
        // Window Navigation
        void ShowLoginWindow();
        void ShowMainWindow();
        void ShowSettingDialog();
        //void ShowTrendDialog(int k);
        void RestartApplication();
        void ExitApplication();

        // Authentication & State
        CurrentUser CurrentUser { get; set; }
        bool IsLoggedIn { get; set; }
        Task Logout();
        // Events
        event EventHandler<WindowNavigationEventArgs> WindowNavigationRequested;
    }

    // AppNavigationService 
    public class AppNavigationService : IAppNavigationService
    {
        private bool _isLoggedIn;
        private CurrentUser _currentUser;
        private Window _loginWindow;
        private Window _mainWindow;
        private Window _chooseExternalAppWindow;

        public event EventHandler<WindowNavigationEventArgs> WindowNavigationRequested;
        public event PropertyChangedEventHandler PropertyChanged;

        public bool IsLoggedIn
        {
            get => _isLoggedIn;
            set
            {
                if (_isLoggedIn != value)
                {
                    _isLoggedIn = value;
                    OnPropertyChanged();
                }
            }
        }

        public CurrentUser CurrentUser
        {
            get => _currentUser;
            set
            {
                if (_currentUser != value)
                {
                    _currentUser = value;
                    OnPropertyChanged();
                }
            }
        }

        private readonly ViewModelLocator _viewModelLocator;

        public AppNavigationService(ViewModelLocator viewModelLocator)
        {
            _viewModelLocator = viewModelLocator;
        }

        // Window Navigation
        public void ShowLoginWindow()
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                if (_loginWindow == null)
                {
                    _loginWindow = new LoginWindow();
                }

                _loginWindow.Show();
                WindowNavigationRequested?.Invoke(this, new WindowNavigationEventArgs("Login"));

                _mainWindow?.Close();
                _mainWindow = null;
            });
        }

        public void ShowMainWindow()
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                // Check UserSetting and SimaticPath
                var userSetting = UserSetting.Instance;
                if (userSetting == null || string.IsNullOrEmpty(userSetting.SimaticPath))
                {
                    // Show ChooseExternalAppDialog if SimaticPath is not set
                    if (_chooseExternalAppWindow == null)
                    {
                        _chooseExternalAppWindow = new ChooseExternalAppDialog();
                    }
                    _chooseExternalAppWindow.Show();
                    WindowNavigationRequested?.Invoke(this, new WindowNavigationEventArgs("ChooseExternalApp"));

                    _mainWindow?.Close();
                    _mainWindow = null;
                    _loginWindow?.Close();
                    _loginWindow = null;
                    return;
                }


                IsLoggedIn = true;
                _mainWindow = new MainWindow();

                _mainWindow.Show();
                WindowNavigationRequested?.Invoke(this, new WindowNavigationEventArgs("Main"));

                _loginWindow?.Close();
                _loginWindow = null;
            });
        }

        public void ShowSettingDialog()
        {
            var dialog = new SettingDialog();
            if (dialog.DataContext is SettingViewModel vm)
            {
                vm.RequestClose += () =>
                {
                    var parentWindow = Window.GetWindow(dialog);
                    if (parentWindow != null)
                    {
                        parentWindow.Close();
                    }
                };
            }
            var mainWindow = Application.Current.Windows
                            .OfType<Window>()
                            .FirstOrDefault(w => w is TESMEA_TMS.Views.MainWindow);
            var window = new Window
            {
                Title = "Cài đặt",
                Content = dialog,
                SizeToContent = SizeToContent.WidthAndHeight,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                ResizeMode = ResizeMode.NoResize,
                WindowStyle = WindowStyle.ToolWindow,
                ShowInTaskbar = false,
            };
            if (mainWindow != null && mainWindow != window)
            {
                window.Owner = mainWindow;
            }
            window.Closing += (s, e) =>
            {
                if (dialog.DataContext is SettingViewModel vm)
                {
                    if (!vm.CanClose())
                    {
                        e.Cancel = true;
                    }
                }
            };
            window.ShowDialog();
        }

        //public void ShowTrendDialog(int k)
        //{
        //    var dialog = new TrendLineDialog(k);
        //    var mainWindow = Application.Current.Windows
        //                    .OfType<Window>()
        //                    .FirstOrDefault(w => w is TESMEA_TMS.Views.MainWindow);
        //    dialog.Title = $"Trend line k = {k}";
        //    dialog.WindowStartupLocation = WindowStartupLocation.CenterOwner;
        //    dialog.ResizeMode = ResizeMode.NoResize;
        //    dialog.WindowStyle = WindowStyle.ToolWindow;
        //    dialog.ShowInTaskbar = false;
        //    if (mainWindow != null && mainWindow != dialog)
        //    {
        //        dialog.Owner = mainWindow;
        //    }
        //    dialog.Closing += (s, e) =>
        //    {
        //        if (dialog.DataContext is SettingViewModel vm)
        //        {
        //            if (!vm.CanClose())
        //            {
        //                e.Cancel = true;
        //            }
        //        }
        //    };
        //    dialog.ShowDialog();
        //}

        public void RestartApplication()
        {
            System.Windows.Forms.Application.Restart();
            System.Windows.Application.Current.Shutdown();
        }

        public async void ExitApplication()
        {
            await ClearApplicationState();
            Application.Current.Shutdown();
        }

        // Authentication Implementation
        public async Task Logout()
        {
            IsLoggedIn = false;
            CurrentUser.Clear();
            RestartApplication();
        }

        private Task ClearApplicationState()
        {
            if (IsLoggedIn)
            {
                IsLoggedIn = false;
                CurrentUser.Clear();
                _viewModelLocator.Cleanup();
                _loginWindow?.Close();
            }
            return Task.CompletedTask;
        }

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    // Supporting Classes
    public class WindowNavigationEventArgs : EventArgs
    {
        public string WindowType { get; }

        public WindowNavigationEventArgs(string windowType)
        {
            WindowType = windowType;
        }
    }
}
