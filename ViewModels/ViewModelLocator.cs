namespace TESMEA_TMS.ViewModels
{
    public class ViewModelLocator
    {
        private readonly Dictionary<Type, object> _viewModels;
        private static IServiceProvider _serviceProvider;

        public ViewModelLocator()
        {
            _viewModels = new Dictionary<Type, object>();
        }

        public static void Initialize(IServiceProvider serviceProvider)
        {
            _serviceProvider = serviceProvider;
        }

        // Generic method để lấy ViewModel
        public T GetViewModel<T>() where T : class
        {
            if (_serviceProvider == null)
                throw new InvalidOperationException("ServiceProvider chưa được khởi tạo. Hãy gọi ViewModelLocator.Initialize() trước.");

            var type = typeof(T);
            if (!_viewModels.ContainsKey(type))
            {
                var viewModel = _serviceProvider.GetService(typeof(T)) as T;
                if (viewModel != null)
                {
                    _viewModels[type] = viewModel;
                }
            }
            return _viewModels[type] as T;
        }

        //// Properties để truy cập trực tiếp các ViewModel chính
        public LoginViewModel LoginViewModel => GetViewModel<LoginViewModel>();
        public ChooseExternalAppViewModel ChooseExternalAppViewModel => GetViewModel<ChooseExternalAppViewModel>();
        public ProjectViewModel ProjectViewModel => GetViewModel<ProjectViewModel>();
        public MainViewModel MainViewModel => GetViewModel<MainViewModel>();
        public ProgressSplashViewModel ProgressSplashViewModel => GetViewModel<ProgressSplashViewModel>();
        public SettingViewModel SettingViewModel => GetViewModel<SettingViewModel>();
        public LibraryViewModel LibraryViewModel => GetViewModel<LibraryViewModel>();
        public MeasureViewModel MeasureViewModel => GetViewModel<MeasureViewModel>();
        public CalculationViewModel CalculationViewModel => GetViewModel<CalculationViewModel>();
        public ScenarioViewModel ScenarioViewModel => GetViewModel<ScenarioViewModel>();

        public void ResetViewModel<T>() where T : class
        {
            var type = typeof(T);
            if (_viewModels.ContainsKey(type))
            {
                _viewModels.Remove(type);
            }
        }

        // Cleanup resources
        public void Cleanup()
        {
            foreach (var viewModel in _viewModels.Values)
            {
                if (viewModel is IDisposable disposable)
                {
                    disposable.Dispose();
                }
            }
            _viewModels.Clear();
        }
    }
}
