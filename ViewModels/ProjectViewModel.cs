using System.Collections.ObjectModel;
using System.IO;
using System.Windows.Input;
using TESMEA_TMS.DTOs;
using TESMEA_TMS.Helpers;
using TESMEA_TMS.Models.Entities;
using TESMEA_TMS.Services;

namespace TESMEA_TMS.ViewModels
{
    public class ProjectViewModel : ViewModelBase
    {
        private List<ComboBoxInfo> _scenarioTypes, testTypes;

        public List<ComboBoxInfo> ScenarioTypes
        {
            get => _scenarioTypes;
            set
            {
                if (_scenarioTypes != value)
                {
                    _scenarioTypes = value;
                    OnPropertyChanged(nameof(ScenarioTypes));
                }
            }
        }

        public List<ComboBoxInfo> TestTypes
        {
            get => testTypes;
            set
            {
                if (testTypes != value)
                {
                    testTypes = value;
                    OnPropertyChanged(nameof(TestTypes));
                }
            }
        }

        private ThongTinDuAn _thongTinDuAn;
        public ThongTinDuAn ThongTinDuAn
        {
            get => _thongTinDuAn;
            set
            {
                if (_thongTinDuAn != value)
                {
                    _thongTinDuAn = value;
                    OnPropertyChanged(nameof(ThongTinDuAn));
                }
            }
        }

        private ThamSo _thamSo;
        public ThamSo ThamSo
        {
            get => _thamSo;
            set
            {
                if (_thamSo != value)
                {
                    _thamSo = value;
                    OnPropertyChanged(nameof(ThamSo));
                }
            }
        }

        private ThongTinChung _thongTinChung;
        public ThongTinChung ThongTinChung
        {
            get => _thongTinChung;
            set
            {
                if (_thongTinChung != value)
                {
                    _thongTinChung = value;
                    OnPropertyChanged(nameof(ThongTinChung));
                }
            }
        }

        private ThongTinMauThuNghiem _thongTinMauThuNghiem;
        public ThongTinMauThuNghiem ThongTinMauThuNghiem
        {
            get => _thongTinMauThuNghiem;
            set
            {
                if (_thongTinMauThuNghiem != value)
                {
                    _thongTinMauThuNghiem = value;
                    OnPropertyChanged(nameof(ThongTinMauThuNghiem));
                }
            }
        }

        private Library _currentParameter;
        public Library CurrentParameter
        {
            get => _currentParameter;
            set
            {
                _currentParameter = value;
                OnPropertyChanged(nameof(CurrentParameter));
            }
        }

        private BienTan _bienTan;
        public BienTan BienTan
        {
            get => _bienTan;
            set
            {
                _bienTan = value;
                OnPropertyChanged(nameof(BienTan));
            }
        }

        private CamBien _camBien;
        public CamBien CamBien
        {
            get => _camBien;
            set
            {
                _camBien = value;
                OnPropertyChanged(nameof(CamBien));
            }
        }

        private OngGio _ongGio;
        public OngGio OngGio
        {
            get => _ongGio;
            set
            {
                _ongGio = value;
                OnPropertyChanged(nameof(OngGio));
            }
        }

        private readonly IParameterService _parameterService;
        private readonly IFileService _fileService;

        // command
        public ICommand AddProjectCommand { get; }
        public ICommand AddCalculationCommand { get; }
        public ProjectViewModel(IParameterService parameterService, IFileService fileService)
        {
            _parameterService = parameterService;
            _fileService = fileService;

            CurrentParameter = new Library();
            BienTan = new BienTan();
            CamBien = new CamBien();
            OngGio = new OngGio();

            AddProjectCommand = new ViewModelCommand(CanExecuteCommand, ExecuteAddProjectCommand);
            AddCalculationCommand = new ViewModelCommand(CanExecuteCommand, ExecuteAddCalculationCommand);
            ThamSo = new ThamSo();
            ThongTinChung = new ThongTinChung();
            ThongTinMauThuNghiem = new ThongTinMauThuNghiem();
            ThongTinDuAn = new ThongTinDuAn(ThamSo, ThongTinChung, ThongTinMauThuNghiem);
            LoadParam();
        }

        public void ClearData()
        {
            ThamSo = new ThamSo();
            ThongTinChung = new ThongTinChung();
            ThongTinMauThuNghiem = new ThongTinMauThuNghiem();
            ThongTinDuAn = new ThongTinDuAn(ThamSo, ThongTinChung, ThongTinMauThuNghiem);
            BienTan = new BienTan();
            CamBien = new CamBien();
            OngGio = new OngGio();
        }

        public async void LoadParam()
        {
            var libraries = await _parameterService.GetLibrariesAsync();
            TestTypes = libraries
                    .Select(x => new ComboBoxInfo(x.LibId.ToString(), x.LibName))
                    .ToList();
            this.ThamSo.KieuKiemThu = TestTypes.FirstOrDefault().Value;
            (BienTan, CamBien, OngGio) = await _parameterService.GetLibraryByIdAsync(Guid.Parse(this.ThamSo.KieuKiemThu));

            var scenarios = await _parameterService.GetScenariosAsync();
            ScenarioTypes = scenarios
                            .Select(x => new ComboBoxInfo(x.ScenarioId.ToString(), x.ScenarioName))
                            .ToList();
            this.ThamSo.KichBan = ScenarioTypes.FirstOrDefault().Value;
        }

        private bool CanExecuteCommand(object parameter)
        {
            // Add logic to determine if the command can execute
            return true;
        }

        private async void ExecuteAddProjectCommand(object parameter)
        {
            try
            {
                var projectFolder = ThongTinDuAn.ThamSo.DuongDanLuuDuAn;
                if (string.IsNullOrEmpty(projectFolder) )
                {
                    MessageBoxHelper.ShowWarning("Vui lòng chọn đường dẫn lưu dự án");
                    return;
                }
                if(!Directory.Exists(projectFolder))
                {
                    MessageBoxHelper.ShowWarning("Thư mục lưu dự án không tồn tại");
                    return;
                }

                if (string.IsNullOrEmpty(ThongTinDuAn.ThamSo.KichBan))
                {
                    MessageBoxHelper.ShowWarning("Vui lòng chọn kịch bản kiểm thử");
                    return;
                }

                if (string.IsNullOrEmpty(ThongTinDuAn.ThamSo.KieuKiemThu))
                {
                    MessageBoxHelper.ShowWarning("Vui lòng chọn kiểu kiểm thử");
                    return;
                }

                ///// Tạo folder user trong folder Testdata, sau đó tạo file giống report có dạng {stt}_{timestamp}.xlsx trong folder user thuộc folder Testdata có dạng ẩn và bảo mật chỉ administrators mới có thể can thiệp
                //// Testdata
                //var testDataFolder = Path.Combine(exchangeFolder, "Testdata");
                //var userFolder = Path.Combine(testDataFolder, "User");
                //if (!Directory.Exists(testDataFolder))
                //{
                //    Directory.CreateDirectory(testDataFolder);
                //    Directory.CreateDirectory(userFolder);
                //    var dirInfo = new DirectoryInfo(userFolder);
                //    dirInfo.Attributes |= FileAttributes.Hidden;

                //    //// Set folder administrators only permissions
                //    //var dirSecurity = dirInfo.GetAccessControl();
                //    //dirSecurity.SetAccessRuleProtection(true, false); // Disable inheritance
                //    //var adminSid = new SecurityIdentifier(WellKnownSidType.BuiltinAdministratorsSid, null);
                //    //var accessRule = new FileSystemAccessRule(
                //    //    adminSid,
                //    //    FileSystemRights.FullControl,
                //    //    InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
                //    //    PropagationFlags.None,
                //    //    AccessControlType.Allow);
                //    //dirSecurity.SetAccessRule(accessRule);
                //    //dirInfo.SetAccessControl(dirSecurity);
                //}

                /////// Sau sẽ thay bằng logic chọn kiểu dữ liệu test từ màn hình
                ////int fileCount = Directory.GetFiles(userFolder, "*.xlsx").Length;
                ////// Create Excel file
                ////string fileName = $"{fileCount + 1}_{DateTime.Now:ddMMyyyy_HHmmss}.xlsx";
                ////_fileService.CreateExcelFile(Path.Combine(userFolder, fileName), "Testdata");

                //var ukFolder = Path.Combine(userFolder, "UK");
                //if (!Directory.Exists(ukFolder))
                //{
                //    Directory.CreateDirectory(ukFolder);
                //}


                var locator = ((App)System.Windows.Application.Current).Resources["Locator"] as ViewModelLocator;
               
                if (locator != null)
                {
                    locator.MeasureViewModel.ThongTinDuAn = this.ThongTinDuAn;
                    var scenarioParams = (await _parameterService.GetScenarioDetailAsync(Guid.Parse(ThongTinDuAn.ThamSo.KichBan))).OrderBy(x => x.STT);
                    locator.MeasureViewModel.MeasureRows.Clear();
                    var mearures = new List<Measure>();
                    foreach (var param in scenarioParams)
                    {
                        mearures.Add(new Measure
                        {
                            k = param.STT,
                            S = param.S,
                            CV = param.CV,
                            F = MeasureStatus.Pending
                        });
                    }
                    locator.MeasureViewModel.InitializePowerPlotModel();
                    locator.MeasureViewModel.InitializeEfficiencyPlotModel();
                    locator.MeasureViewModel.MeasureRows = new ObservableCollection<Measure>(mearures);
                    locator.MainViewModel.CurrentView = new TESMEA_TMS.Views.MeasureView();
                }
            }
            catch (Exception ex)
            {
                MessageBoxHelper.ShowError("Lỗi tạo dự án, vui lòng thử lại");
                return;
            }
        }

        private void ExecuteAddCalculationCommand(object parameter)
        {
            var locator = ((App)System.Windows.Application.Current).Resources["Locator"] as ViewModelLocator;
            if (locator != null)
            {
                locator.CalculationViewModel.ThongTinDuAn = this.ThongTinDuAn;
                locator.MainViewModel.CurrentView = new TESMEA_TMS.Views.CalculationView();
            }
        }

        public async void OnLibraryChanged()
        {
            var libId = ThongTinDuAn.ThamSo.KieuKiemThu;
            if (libId == null) return;
            (BienTan, CamBien, OngGio) = await _parameterService.GetLibraryByIdAsync(Guid.Parse(libId));
        }
    }
}
