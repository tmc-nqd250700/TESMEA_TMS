using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Windows;
using System.Windows.Input;
using TESMEA_TMS.Configs;
using TESMEA_TMS.DTOs;
using TESMEA_TMS.Helpers;
using TESMEA_TMS.Services;
using TESMEA_TMS.Views;
using Application = System.Windows.Application;

namespace TESMEA_TMS.ViewModels
{
    public class ScenarioViewModel : ViewModelBase
    {
        #region Properties

        private ObservableCollection<ScenarioDto> _scenarios;
        public ObservableCollection<ScenarioDto> Scenarios
        {
            get => _scenarios;
            set
            {
                _scenarios = value;
                OnPropertyChanged(nameof(Scenarios));
            }
        }

        private ScenarioDto _selectedScenario;
        public ScenarioDto SelectedScenario
        {
            get => _selectedScenario;
            set
            {
                _selectedScenario = value;
                OnPropertyChanged(nameof(SelectedScenario));
                OnPropertyChanged(nameof(IsScenarioSelected));
                OnPropertyChanged(nameof(CanEditParams));

                // KHÔNG tự động load params khi chỉ select
                // Chỉ load khi user click "Chọn" (ViewDetailCommand)
            }
        }

        private ObservableCollection<ScenarioParamDTO> _scenarioParams;
        public ObservableCollection<ScenarioParamDTO> ScenarioParams
        {
            get => _scenarioParams;
            set
            {
                _scenarioParams = value;
                OnPropertyChanged(nameof(ScenarioParams));
            }
        }

        // Properties cho UI binding
        public bool IsScenarioSelected => SelectedScenario != null;

        public bool CanEditParams => SelectedScenario != null &&
                                     !SelectedScenario.IsMarkedForDeletion;
        public bool HasUnsavedChanges =>
            Scenarios.Any(item => item.IsNew || item.IsEdited || item.IsMarkedForDeletion);

        #endregion

        #region Services & Fields

        private readonly IParameterService _parameterService;

        // Tracking để biết scenario nào đã có params thay đổi
        private HashSet<Guid> _scenariosWithChangedParams = new HashSet<Guid>();

        // Track scenario hiện đang xem chi tiết
        private Guid? _currentViewedScenarioId = null;

        #endregion

        #region Commands

        public ICommand NewCommand { get; }
        public ICommand SaveCommand { get; }
        public ICommand DeleteCommand { get; }
        public ICommand UndoDeleteCommand { get; }
        public ICommand ViewDetailCommand { get; }

        #endregion

        #region Constructor

        public ScenarioViewModel(IParameterService parameterService)
        {
            _parameterService = parameterService;

            Scenarios = new ObservableCollection<ScenarioDto>();
            ScenarioParams = new ObservableCollection<ScenarioParamDTO>();

            // Subscribe to collection changes để tự động đánh dấu edited
            ScenarioParams.CollectionChanged += ScenarioParams_CollectionChanged;

            // Initialize commands
            NewCommand = new ViewModelCommand(_ => true, ExecuteNewCommand);
            SaveCommand = new ViewModelCommand(CanExecuteSaveCommand, ExecuteSaveCommand);
            DeleteCommand = new ViewModelCommand(CanExecuteDeleteCommand, ExecuteDeleteCommand);
            UndoDeleteCommand = new ViewModelCommand(_ => true, ExecuteUndoDeleteCommand);
            ViewDetailCommand = new ViewModelCommand(_ => true, ExecuteViewDetailCommand);

            LoadScenarios();
        }

        #endregion

        #region Data Loading Methods

        public async void LoadScenarios()
        {
            try
            {
                var scenarios = await _parameterService.GetScenariosAsync();
                Scenarios.Clear();
                ScenarioParams.Clear();
                _scenariosWithChangedParams.Clear();

                if (scenarios == null) return;

                foreach (var scenario in scenarios)
                {
                    var dto = ScenarioDto.FromEntity(scenario, isNew: false);
                    Scenarios.Add(dto);
                }
            }
            catch (Exception ex)
            {
                MessageBoxHelper.ShowError($"Lỗi khi tải danh sách scenario: {ex.Message}");
            }
        }

        private async Task LoadScenarioParams(Guid scenarioId)
        {
            try
            {
                var paramList = await _parameterService.GetScenarioDetailAsync(scenarioId);
                ScenarioParams.Clear();

                if (paramList == null || !paramList.Any()) return;

                int stt = 1;
                foreach (var param in paramList.OrderBy(p => p.STT))
                {
                    var dto = ScenarioParamDTO.FromEntity(param, stt++, isNew: false);

                    // Subscribe to property changes để track editing
                    dto.PropertyChanged += (s, e) =>
                    {
                        if ((e.PropertyName == nameof(ScenarioParamDTO.S) ||
                             e.PropertyName == nameof(ScenarioParamDTO.CV)) &&
                            SelectedScenario != null)
                        {
                            MarkScenarioAsEdited();
                        }
                    };

                    ScenarioParams.Add(dto);
                }
                //if (ScenarioParams.Any())
                //{
                //    LoadScenarioParams(ScenarioParams.FirstOrDefault().ScenarioId);
                //}
            }
            catch (Exception ex)
            {
                MessageBoxHelper.ShowError($"Lỗi khi tải chi tiết scenario: {ex.Message}");
            }
        }

        #endregion

        #region Collection Change Handlers

        private void ScenarioParams_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            UpdateSTT();

            // Đánh dấu scenario là đã chỉnh sửa nếu thêm/xóa params
            if (SelectedScenario != null && !SelectedScenario.IsNew)
            {
                if (e.Action == NotifyCollectionChangedAction.Add ||
                    e.Action == NotifyCollectionChangedAction.Remove)
                {
                    MarkScenarioAsEdited();
                }
            }

            // Subscribe to new items
            if (e.NewItems != null)
            {
                foreach (ScenarioParamDTO item in e.NewItems)
                {
                    item.PropertyChanged += (s, ev) =>
                    {
                        if ((ev.PropertyName == nameof(ScenarioParamDTO.S) ||
                             ev.PropertyName == nameof(ScenarioParamDTO.CV)) &&
                            SelectedScenario != null && !SelectedScenario.IsNew)
                        {
                            MarkScenarioAsEdited();
                        }
                    };
                }
            }
        }

        private void UpdateSTT()
        {
            for (int i = 0; i < ScenarioParams.Count; i++)
            {
                ScenarioParams[i].STT = i + 1;
            }
        }

        private void MarkScenarioAsEdited()
        {
            if (SelectedScenario != null && !SelectedScenario.IsNew)
            {
                SelectedScenario.IsEdited = true;
                _scenariosWithChangedParams.Add(SelectedScenario.ScenarioId);
            }
        }

        #endregion

        #region Command Handlers

        private async void ExecuteViewDetailCommand(object parameter)
        {
            try
            {

                if (HasUnsavedChanges)
                {
                    MessageBoxHelper.ShowWarning("Còn thay đổi chưa lưu, vui lòng lưu để hoàn thành");
                    return;
                }

                if (parameter == null)
                    return;

                Guid scenarioId;
                if (parameter is Guid guid)
                {
                    scenarioId = guid;
                }
                else if (parameter is string str && Guid.TryParse(str, out var parsedGuid))
                {
                    scenarioId = parsedGuid;
                }
                else
                {
                    return;
                }

                var scenario = Scenarios.FirstOrDefault(s => s.ScenarioId == scenarioId);
                if (scenario == null || scenario.IsMarkedForDeletion)
                {
                    MessageBoxHelper.ShowWarning("Không thể xem chi tiết scenario này");
                    return;
                }

                SelectedScenario = scenario;
                _currentViewedScenarioId = scenarioId;

                // Load params tương ứng
                if (scenario.IsNew)
                {
                    ScenarioParams.Clear();
                }
                else
                {
                    await LoadScenarioParams(scenarioId);
                }
            }
            catch (Exception ex)
            {
                MessageBoxHelper.ShowError($"Lỗi khi xem chi tiết: {ex.Message}");
            }
        }

        private void ExecuteNewCommand(object obj)
        {
            try
            {
                var mainWindow = Application.Current.Windows
                          .OfType<Window>()
                          .FirstOrDefault(w => w is TESMEA_TMS.Views.MainWindow);


                var dialog = new ConfirmAddScenarioDialog();
                
                if (mainWindow != null && mainWindow != dialog)
                {
                    dialog.Owner = mainWindow;
                }


                if (dialog.ShowDialog() != true)
                {
                    return;
                }

                var scenarioName = dialog.InputText?.Trim();
                var standardDeviation = dialog.StandardDeviation;
                var timeRange = dialog.TimeRange;
                if(standardDeviation < 105)
                {
                    MessageBoxHelper.ShowWarning("Tỉ lệ giá trị max/min không được nhỏ hơn 1.05");
                }

                if (string.IsNullOrWhiteSpace(scenarioName))
                {
                    MessageBoxHelper.ShowWarning("Vui lòng nhập tên scenario");
                    return;
                }
                if (standardDeviation <= 0)
                {
                    MessageBoxHelper.ShowWarning("Độ lệch chuẩn phải lớn hơn 0");
                    return;
                }

                if (timeRange <= 0)
                {
                    MessageBoxHelper.ShowWarning("Khoảng thời gian phải lớn hơn 0");
                    return;
                }

                // Kiểm tra trùng tên
                if (Scenarios.Any(s => s.ScenarioName.Equals(scenarioName, StringComparison.OrdinalIgnoreCase)
                                       && !s.IsMarkedForDeletion))
                {
                    MessageBoxHelper.ShowWarning("Tên scenario đã tồn tại");
                    return;
                }

                // Tạo scenario mới
                var newScenario = new ScenarioDto
                {
                    ScenarioId = Guid.NewGuid(),
                    ScenarioName = scenarioName,
                    StandardDeviation = standardDeviation,
                    TimeRange = timeRange,
                    CreatedDate = DateTime.Now,
                    CreatedUser = CurrentUser.Instance.UserAccount.FullName,
                    IsNew = true,
                    IsEdited = false,
                    IsMarkedForDeletion = false
                };

                Scenarios.Add(newScenario);

                SelectedScenario = newScenario;
                _currentViewedScenarioId = newScenario.ScenarioId;
                ScenarioParams.Clear();
                ScenarioParams.Add(new ScenarioParamDTO
                {
                    ScenarioId = newScenario.ScenarioId,
                    STT = 1,
                    S = 0,
                    CV = 0,
                    IsNew = true
                });
                ScenarioParams.Add(new ScenarioParamDTO
                {
                    ScenarioId = newScenario.ScenarioId,
                    STT = 2,
                    S = 0,
                    CV = 0,
                    IsNew = true
                });

            }
            catch (Exception ex)
            {
                MessageBoxHelper.ShowError($"Lỗi khi tạo mới scenario: {ex.Message}");
            }
        }

        private bool CanExecuteSaveCommand(object parameter)
        {
            return Scenarios.Any(s => s.IsNew || s.IsEdited || s.IsMarkedForDeletion);
        }

        private async void ExecuteSaveCommand(object obj)
        {
            try
            {
                var result = MessageBoxHelper.ShowQuestion("Bạn có chắc chắn muốn lưu tất cả thay đổi?");

                if (!result)
                {
                    return;
                }
                if (!ScenarioParams.Any())
                {
                    MessageBoxHelper.ShowWarning("Không có thay đổi để lưu.");
                    return;
                }

                var scenarioToUpdate = new ScenarioUpdateDto();
                scenarioToUpdate.Scenario = SelectedScenario;
                scenarioToUpdate.Params = ScenarioParams.ToList();
                await _parameterService.UpdateScenarioAsync(scenarioToUpdate);
                _scenariosWithChangedParams.Clear();
                MessageBoxHelper.ShowSuccess("Lưu thành công");

                LoadScenarios();
                 // recall to reload current viewed param
                var locator = (ViewModelLocator)Application.Current.Resources["Locator"];
                locator.ProjectViewModel.LoadParam();
                if (SelectedScenario == null ||
                    !Scenarios.Any(s => s.ScenarioId == SelectedScenario.ScenarioId))
                {
                    ScenarioParams.Clear();
                    _currentViewedScenarioId = null;
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        private bool CanExecuteDeleteCommand(object parameter)
        {
            if (parameter is ScenarioDto scenario)
            {
                return !scenario.IsMarkedForDeletion;
            }

            return SelectedScenario != null && !SelectedScenario.IsMarkedForDeletion;
        }

        private async void ExecuteDeleteCommand(object parameter)
        {
            try
            {
                ScenarioDto scenarioToDelete = null;

                // Lấy scenario từ parameter
                if (parameter is ScenarioDto scenario)
                {
                    scenarioToDelete = scenario;
                }
                else if (SelectedScenario != null)
                {
                    scenarioToDelete = SelectedScenario;
                }

                if (scenarioToDelete == null) return;

                var result = MessageBoxHelper.ShowQuestion(
                    $"Bạn có chắc chắn muốn xóa scenario '{scenarioToDelete.ScenarioName}'?\n\n" +
                    "Lưu ý: Thao tác này có thể hoàn tác trước khi lưu");

                if (result)
                {
                    if (scenarioToDelete.IsNew)
                    {
                        // Xóa trực tiếp nếu là scenario mới chưa lưu
                        Scenarios.Remove(scenarioToDelete);

                        if (SelectedScenario == scenarioToDelete)
                        {
                            SelectedScenario = null;
                            ScenarioParams.Clear();
                            _currentViewedScenarioId = null;
                        }
                    }
                    else
                    {
                        // Đánh dấu để xóa (có thể hoàn tác)
                        scenarioToDelete.IsMarkedForDeletion = true;
                        await LoadScenarioParams(scenarioToDelete.ScenarioId);
                        _currentViewedScenarioId = null;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBoxHelper.ShowError($"Lỗi khi xóa scenario: {ex.Message}");
            }
        }

        private void ExecuteUndoDeleteCommand(object parameter)
        {
            try
            {
                if (parameter is Guid scenarioId)
                {
                    var scenario = Scenarios.FirstOrDefault(s => s.ScenarioId == scenarioId);

                    if (scenario != null && scenario.IsMarkedForDeletion)
                    {
                        scenario.IsMarkedForDeletion = false;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBoxHelper.ShowError($"Lỗi khi hoàn tác: {ex.Message}");
            }
        }

        #endregion
    }
}