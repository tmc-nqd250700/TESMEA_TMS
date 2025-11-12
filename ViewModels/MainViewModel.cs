using MaterialDesignThemes.Wpf;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using TESMEA_TMS.Configs;
using TESMEA_TMS.Services;
using TESMEA_TMS.Views;
using Application = System.Windows.Application;

namespace TESMEA_TMS.ViewModels
{
    public class MainViewModel : ViewModelBase
    {
        private string _currentDateTime;
        public string CurrentDateTime
        {
            get => _currentDateTime;
            set
            {
                _currentDateTime = value;
                OnPropertyChanged(nameof(CurrentDateTime));
            }
        }

        private System.Windows.Controls.UserControl _currentView;
        public System.Windows.Controls.UserControl CurrentView
        {
            get => _currentView;
            set
            {
                if (_currentView != value)
                {
                    _currentView = value;
                    OnPropertyChanged(nameof(CurrentView));
                }
            }
        }


        public System.Windows.Controls.UserControl ProjectView { get; } = new ProjectView();
        //public System.Windows.Controls.UserControl LibraryView { get; } = new LibraryView();
        //public System.Windows.Controls.UserControl ScenarioView { get; } = new ScenarioView();
        //public System.Windows.Controls.UserControl MeasureView { get; } = new MeasureView();
        //public System.Windows.Controls.UserControl CalculationView { get; } = new CalculationView();


        public CurrentUser CurrentUser { get; }
        // services 
        private readonly IAppNavigationService _appNavigatioService;
        private readonly IExternalAppService _externalAppService;

        // Commands
        public ICommand ShowProjectViewCommand { get; }
        public ICommand ShowLibrabyViewCommand { get; }
        public ICommand ShowScenarioViewCommand { get; }
        public ICommand ShowMeasureViewCommand { get; }
        public ICommand ShowCalculationViewCommand { get; }
        public ICommand ShowSettingDialogCommand { get; }
        public ICommand LogoutCommand { get; }


        // simatic
        private DispatcherTimer _simaticMonitorTimer;

        private bool _isRestartingSimatic = false;
        public bool _isSimaticRunning
        {
            get
            {
                if (_externalAppService == null)
                    return false;
                return _externalAppService.IsAppRunning ? true : false;
            }
        }

        public MainViewModel(IAppNavigationService appNavigationService, IExternalAppService externalAppService)
        {
            _appNavigatioService = appNavigationService;
            _externalAppService = externalAppService;
            CurrentUser = CurrentUser.Instance;
            CurrentView = ProjectView;

            LogoutCommand = new ViewModelCommand(_ => _appNavigatioService.Logout());
            ShowProjectViewCommand = new ViewModelCommand(_ => CurrentView = ProjectView);
            ShowMeasureViewCommand = new ViewModelCommand(_ => CurrentView = new MeasureView());
            ShowCalculationViewCommand = new ViewModelCommand(_ => CurrentView = new CalculationView());
            ShowLibrabyViewCommand = new ViewModelCommand(_ => CurrentView = new LibraryView());
            ShowScenarioViewCommand = new ViewModelCommand(_ => CurrentView = new ScenarioView());
            ShowSettingDialogCommand = new ViewModelCommand(_ => ExecuteShowSettingCommand());
            Initialize();
        }

        private async void Initialize()
        {
            var timer = new DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(1)
            };
            timer.Tick += (s, e) =>
            {
                CurrentDateTime = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
            };
            timer.Start();
            await _externalAppService.StartAppAsync();
            _simaticMonitorTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(10)
            };
            _simaticMonitorTimer.Tick += SimaticMonitorTimer_Tick;
            _simaticMonitorTimer.Start();
        }


        private void SimaticMonitorTimer_Tick(object sender, EventArgs e)
        {
            if (!_isSimaticRunning && !_isRestartingSimatic)
            {
                // Simatic dừng và chưa restart
                _simaticMonitorTimer.Stop();
                ShowSimaticStoppedAlert();
            }
            else if (_isRestartingSimatic)
            {
                if (_isSimaticRunning)
                {
                    _isRestartingSimatic = false;
                    _simaticMonitorTimer.Start();
                    if (DialogHost.IsDialogOpen("MainDialogHost"))
                        DialogHost.Close("MainDialogHost");
                }
            }
        }

        private CancellationTokenSource? _restartCts;
        private async void ShowSimaticStoppedAlert()
        {
            // Cancel previous restart attempt
            _restartCts?.Cancel();
            _restartCts = new CancellationTokenSource();

            _isRestartingSimatic = true;
            var splashViewModel = new ProgressSplashViewModel
            {
                Message = "Ứng dụng Simatic đã dừng, đang khởi động lại...",
                IsIndeterminate = true
            };
            var splash = new Views.CustomControls.ProgressSplashContent { DataContext = splashViewModel };

            _ = Application.Current.Dispatcher.Invoke(
                 () => DialogHost.Show(splash, "MainDialogHost"),
                 System.Windows.Threading.DispatcherPriority.Send
             );

            try
            {
                await _externalAppService.StartAppAsync();
                if (_restartCts.Token.IsCancellationRequested) return;
                Application.Current.Dispatcher.Invoke(() =>
                {
                    if (DialogHost.IsDialogOpen("MainDialogHost"))
                        DialogHost.Close("MainDialogHost");
                });
            }
            catch (Exception ex)
            {
                _isRestartingSimatic = false;
                Application.Current.Dispatcher.Invoke(() =>
                {
                    if (DialogHost.IsDialogOpen("MainDialogHost"))
                        DialogHost.Close("MainDialogHost");
                });
            }
        }

        private void ExecuteShowSettingCommand()
        {
            _appNavigatioService.ShowSettingDialog();
        }

        public async Task ExitApp()
        {
            var splashViewModel = new ProgressSplashViewModel
            {
                Message = "Đang đóng ứng dụng...",
                IsIndeterminate = true
            };
            var splash = new Views.CustomControls.ProgressSplashContent { DataContext = splashViewModel };

            _ = Application.Current.Dispatcher.Invoke(
                 () => DialogHost.Show(splash, "MainDialogHost"),
                 System.Windows.Threading.DispatcherPriority.Send
             );
            _appNavigatioService.ExitApplication();
            Application.Current.Dispatcher.Invoke(() =>
            {
                if (DialogHost.IsDialogOpen("MainDialogHost"))
                    DialogHost.Close("MainDialogHost");
            });

        }
    }
}
