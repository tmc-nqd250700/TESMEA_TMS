using System.IO;
using System.Windows;
using System.Windows.Input;
using TESMEA_TMS.Services;
using MessageBox = System.Windows.MessageBox;

namespace TESMEA_TMS.ViewModels
{
    public class ChooseExternalAppViewModel : ViewModelBase
    {
        private string _winCCPath;
        public string WinCCPath
        {
            get => _winCCPath;
            set
            {
                _winCCPath = value;
                OnPropertyChanged(nameof(WinCCPath));
            }
        }

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

        public ICommand SaveCommand { get; }
        private readonly IAppNavigationService _appNavigationService;

        public ChooseExternalAppViewModel(IAppNavigationService appNavigationService)
        {
            _appNavigationService = appNavigationService;
            var setting = Configs.UserSetting.Load();
            SimaticPath = setting.SimaticPath;

            SaveCommand = new ViewModelCommand(CanExecuteCommand, ExecuteSaveCommand);
        }

        private bool CanExecuteCommand(object obj)
        {
            return !string.IsNullOrEmpty(SimaticPath) && !string.IsNullOrEmpty(WinCCPath);
        }

        private void ExecuteSaveCommand(object obj)
        {
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
            var setting = Configs.UserSetting.Load();
            setting.SimaticPath = SimaticPath;
            setting.WinccExePath = WinCCPath;
            setting.Save();
            
            // Showing main window after saving
            _appNavigationService.RestartApplication();
        }
    }
}
