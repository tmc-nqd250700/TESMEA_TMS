using OfficeOpenXml;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Input;
using TESMEA_TMS.DTOs;
using TESMEA_TMS.Helpers;
using TESMEA_TMS.Models.Entities;
using TESMEA_TMS.Services;
using TESMEA_TMS.Views;
using Application = System.Windows.Application;

namespace TESMEA_TMS.ViewModels
{
    public class LibraryViewModel : ViewModelBase
    {
        #region Properties

        private ObservableCollection<LibraryDto> _inputParameters;
        public ObservableCollection<LibraryDto> InputParameters
        {
            get => _inputParameters;
            set
            {
                _inputParameters = value;
                OnPropertyChanged(nameof(InputParameters));
            }
        }

        private LibraryDto _currentParameter;
        public LibraryDto CurrentParameter
        {
            get => _currentParameter;
            set
            {
                _currentParameter = value;
                OnPropertyChanged(nameof(CurrentParameter));
                OnPropertyChanged(nameof(IsParameterSelected));
            }
        }

        private BienTanDto _bienTan;
        public BienTanDto BienTan
        {
            get => _bienTan;
            set
            {
                _bienTan = value;
                OnPropertyChanged(nameof(BienTan));
            }
        }

        private CamBienDto _camBien;
        public CamBienDto CamBien
        {
            get => _camBien;
            set
            {
                _camBien = value;
                OnPropertyChanged(nameof(CamBien));
            }
        }

        private OngGioDto _ongGio;
        public OngGioDto OngGio
        {
            get => _ongGio;
            set
            {
                _ongGio = value;
                OnPropertyChanged(nameof(OngGio));
            }
        }

        public bool IsParameterSelected => CurrentParameter != null;
        public bool HasUnsavedChanges =>
            InputParameters.Any(item => item.IsNew || item.IsEdited || item.IsMarkedForDeletion);

        #endregion

        #region Services & Fields

        private readonly IParameterService _parameterService;

        // Track library hiện đang xem
        private Guid? _currentViewedParamId = null;

        #endregion

        #region Commands

        public ICommand ViewDetailCommand { get; }
        public ICommand NewCommand { get; }
        public ICommand OpenCommand { get; }
        public ICommand SaveCommand { get; }
        public ICommand DeleteCommand { get; }
        public ICommand UndoDeleteCommand { get; }

        #endregion

        #region Constructor

        public LibraryViewModel(IParameterService parameterService)
        {
            _parameterService = parameterService;

            InputParameters = new ObservableCollection<LibraryDto>();
            InitializeEmptyData();

            ViewDetailCommand = new ViewModelCommand(_ => true, ExecuteViewDetailCommand);
            NewCommand = new ViewModelCommand(_ => true, ExecuteNewCommand);
            OpenCommand = new ViewModelCommand(_ => true, ExecuteOpenCommand);
            SaveCommand = new ViewModelCommand(CanExecuteSaveCommand, ExecuteSaveCommand);
            DeleteCommand = new ViewModelCommand(_ => true, ExecuteDeleteCommand);
            UndoDeleteCommand = new ViewModelCommand(_ => true, ExecuteUndoDeleteCommand);

            LoadData();
        }

        #endregion

        #region Data Loading Methods

        public async void LoadData()
        {
            try
            {
                var inputParams = await _parameterService.GetLibrariesAsync();
                InputParameters.Clear();

                if (inputParams != null)
                {
                    foreach (var param in inputParams)
                    {
                        var dto = LibraryDto.FromEntity(param, isNew: false);
                        InputParameters.Add(dto);
                    }
                }

                InitializeEmptyData();
            }
            catch (Exception ex)
            {
                MessageBoxHelper.ShowError($"Lỗi khi tải danh sách library: {ex.Message}");
            }
        }

        private void InitializeEmptyData()
        {
            CurrentParameter = new LibraryDto
            {
                LibId = Guid.Empty,
                LibName = string.Empty,
                CreatedDate = DateTime.Now,
                CreatedUser = Environment.UserName
            };

            BienTan = new BienTanDto { LibId = Guid.Empty };
            CamBien = new CamBienDto { LibId = Guid.Empty };
            OngGio = new OngGioDto { LibId = Guid.Empty };

            // Subscribe to change events
            SubscribeToDataChanges();
        }

        private void SubscribeToDataChanges()
        {
            // Subscribe to entity changes để track editing
            if (BienTan != null)
            {
                BienTan.DataChanged -= OnDataChanged;
                BienTan.DataChanged += OnDataChanged;
            }

            if (CamBien != null)
            {
                CamBien.DataChanged -= OnDataChanged;
                CamBien.DataChanged += OnDataChanged;
            }

            if (OngGio != null)
            {
                OngGio.DataChanged -= OnDataChanged;
                OngGio.DataChanged += OnDataChanged;
            }
        }

        private void OnDataChanged(object sender, EventArgs e)
        {
            // Đánh dấu library là đã chỉnh sửa khi có thay đổi data
            if (CurrentParameter != null && !CurrentParameter.IsNew && _currentViewedParamId.HasValue)
            {
                CurrentParameter.IsEdited = true;
            }
        }

        #endregion

        #region Command Handlers

        private async void ExecuteViewDetailCommand(object parameter)
        {
            try
            {
                if (parameter == null)
                    return;

                Guid libId;
                if (parameter is Guid guid)
                {
                    libId = guid;
                }
                else if (parameter is string str && Guid.TryParse(str, out var parsedGuid))
                {
                    libId = parsedGuid;
                }
                else
                {
                    return;
                }

                var library = InputParameters.FirstOrDefault(p => p.LibId == libId);

                if (library == null || library.IsMarkedForDeletion)
                {
                    MessageBoxHelper.ShowWarning("Không thể xem chi tiết library này");
                    return;
                }

                // Set current parameter
                CurrentParameter = library;
                _currentViewedParamId = libId;

                if (library.IsNew)
                {
                    // Library mới - giữ data hiện tại
                    if (BienTan != null) BienTan.LibId = libId;
                    if (CamBien != null) CamBien.LibId = libId;
                    if (OngGio != null) OngGio.LibId = libId;
                }
                else
                {
                    // Library đã tồn tại - load từ DB
                    var detail = await _parameterService.GetLibraryByIdAsync(libId);

                    BienTan = BienTanDto.FromEntity(detail.Item1);
                    CamBien = CamBienDto.FromEntity(detail.Item2);
                    OngGio = OngGioDto.FromEntity(detail.Item3);

                    // Re-subscribe after loading new data
                    SubscribeToDataChanges();
                }
            }
            catch (Exception ex)
            {
                MessageBoxHelper.ShowError($"Lỗi khi xem chi tiết library: {ex.Message}");
            }
        }

        private void ExecuteNewCommand(object obj)
        {
            try
            {
                var mainWindow = Application.Current.Windows
                            .OfType<Window>()
                            .FirstOrDefault(w => w is TESMEA_TMS.Views.MainWindow);

                var dialog = new ConfirmAddLibraryDialog();
               
                if (mainWindow != null && mainWindow != dialog)
                {
                    dialog.Owner = mainWindow;
                }

                if (dialog.ShowDialog() != true)
                {
                    return;
                }

                var libraryName = dialog.InputText?.Trim();

                if (string.IsNullOrWhiteSpace(libraryName))
                {
                    MessageBoxHelper.ShowWarning("Vui lòng nhập tên library");
                    return;
                }

                // Kiểm tra trùng tên
                if (InputParameters.Any(s => s.LibName.Equals(libraryName, StringComparison.OrdinalIgnoreCase)
                                             && !s.IsMarkedForDeletion))
                {
                    MessageBoxHelper.ShowWarning("Tên library đã tồn tại");
                    return;
                }

                // Tạo library mới
                var newLibrary = new LibraryDto
                {
                    LibId = Guid.NewGuid(),
                    LibName = libraryName,
                    CreatedDate = DateTime.Now,
                    CreatedUser = Environment.UserName,
                    IsNew = true,
                    IsEdited = false,
                    IsMarkedForDeletion = false
                };

                InputParameters.Add(newLibrary);

                // Tự động chọn và khởi tạo data rỗng
                CurrentParameter = newLibrary;
                _currentViewedParamId = newLibrary.LibId;

                BienTan = new BienTanDto { LibId = newLibrary.LibId };
                CamBien = new CamBienDto { LibId = newLibrary.LibId };
                OngGio = new OngGioDto { LibId = newLibrary.LibId };

                SubscribeToDataChanges();
            }
            catch (Exception ex)
            {
                MessageBoxHelper.ShowError($"Lỗi khi tạo library mới: {ex.Message}");
            }
        }

        private void ExecuteOpenCommand(object obj)
        {
            try
            {
                var ofd = new OpenFileDialog
                {
                    Filter = "Excel Files|*.xlsx;*.xls",
                    Title = "Chọn file library"
                };

                if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    // Handle convert file to object
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    using (var package = new ExcelPackage(ofd.FileName))
                    {
                        ExcelWorksheet invSheet = package.Workbook.Worksheets.FirstOrDefault(x => x.Name.ToLower().Trim() == "inverter") ?? package.Workbook.Worksheets[0];
                        ExcelWorksheet sensorSheet = package.Workbook.Worksheets.FirstOrDefault(x => x.Name.ToLower().Trim() == "sensor") ?? package.Workbook.Worksheets[1];
                        ExcelWorksheet ductSheet = package.Workbook.Worksheets.FirstOrDefault(x => x.Name.ToLower().Trim() == "duct") ?? package.Workbook.Worksheets[2];

                        var bienTan = new BienTanDto
                        {
                            LibId = CurrentParameter?.LibId ?? Guid.Empty,
                            DienApVao = GetCellValue(invSheet, "B2"),
                            DongDienVao = GetCellValue(invSheet, "B3"),
                            TanSoVao = GetCellValue(invSheet, "B4"),
                            CongSuatVao = GetCellValue(invSheet, "B5"),
                            DienApRa = GetCellValue(invSheet, "B6"),
                            DongDienRa = GetCellValue(invSheet, "B7"),
                            TanSoRa = GetCellValue(invSheet, "B8"),
                            CongSuatTongRa = GetCellValue(invSheet, "B9"),
                            CongSuatHieuDungRa = GetCellValue(invSheet, "B10")
                        };
                      
                        var camBien = new CamBienDto
                        {
                            LibId = CurrentParameter?.LibId ?? Guid.Empty,
                            NhietDoMoiTruongMin = GetCellValue(sensorSheet, "B3"),
                            NhietDoMoiTruongMax = GetCellValue(sensorSheet, "C3"),
                            DoAmMoiTruongMin = GetCellValue(sensorSheet, "B4"),
                            DoAmMoiTruongMax = GetCellValue(sensorSheet, "C4"),
                            ApSuatKhiQuyenMin = GetCellValue(sensorSheet, "B5"),
                            ApSuatKhiQuyenMax = GetCellValue(sensorSheet, "C5"),
                            ChenhLechApSuatMin = GetCellValue(sensorSheet, "B6"),
                            ChenhLechApSuatMax = GetCellValue(sensorSheet, "C6"),
                            ApSuatTinhMin = GetCellValue(sensorSheet, "B7"),
                            ApSuatTinhMax = GetCellValue(sensorSheet, "C7"),
                            DoRungMin = GetCellValue(sensorSheet, "B8"),
                            DoRungMax = GetCellValue(sensorSheet, "C8"),
                            DoOnMin = GetCellValue(sensorSheet, "B9"),
                            DoOnMax = GetCellValue(sensorSheet, "C9"),
                            SoVongQuayMin = GetCellValue(sensorSheet, "B10"),
                            SoVongQuayMax = GetCellValue(sensorSheet, "C10"),
                            MomenMin = GetCellValue(sensorSheet, "B11"),
                            MomenMax = GetCellValue(sensorSheet, "C11"),
                            PhanHoiDongDienMin = GetCellValue(sensorSheet, "B12"),
                            PhanHoiDongDienMax = GetCellValue(sensorSheet, "C12"),
                            PhanHoiCongSuatMin = GetCellValue(sensorSheet, "B13"),
                            PhanHoiCongSuatMax = GetCellValue(sensorSheet, "C13"),
                            PhanHoiViTriVanMin = GetCellValue(sensorSheet, "B14"),
                            PhanHoiViTriVanMax = GetCellValue(sensorSheet, "C14")
                        };

                        var ongGio = new OngGioDto
                        {
                            LibId = CurrentParameter?.LibId ?? Guid.Empty,
                            //DuongKinhOngGio = GetCellValue(ductSheet, "B2"),
                            //ChieuDaiOngGioTruocQuat = GetCellValue(ductSheet, "B3"),
                            //ChieuDaiOngGioSauQuat = GetCellValue(ductSheet, "B4"),
                            //DuongKinhLoPhut = GetCellValue(ductSheet, "B5"),
                        };

                        BienTan = bienTan;
                        CamBien = camBien;
                        OngGio = ongGio;

                        SubscribeToDataChanges();

                        if (CurrentParameter != null && !CurrentParameter.IsNew)
                        {
                            CurrentParameter.IsEdited = true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBoxHelper.ShowError($"Lỗi khi mở file Excel: {ex.Message}");
            }
        }

        private float GetCellValue(ExcelWorksheet sheet, string cellAddress)
        {
            var text = sheet.Cells[cellAddress].Text;
            return string.IsNullOrWhiteSpace(text) ? 0 : sheet.Cells[cellAddress].GetValue<float>();
        }

        private bool CanExecuteSaveCommand(object parameter)
        {
            return InputParameters.Any(p => p.IsNew || p.IsEdited || p.IsMarkedForDeletion);
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

                // 1. Xóa các library đánh dấu xóa (đã tồn tại)
                var librariesToDelete = InputParameters
                    .Where(p => p.IsMarkedForDeletion && !p.IsNew)
                    .ToList();

                foreach (var library in librariesToDelete)
                {
                    await _parameterService.DeleteLibraryAsync(library.LibId);
                    InputParameters.Remove(library);
                }

                // 2. Xóa library mới bị đánh dấu xóa
                var newLibrariesToDelete = InputParameters
                    .Where(p => p.IsMarkedForDeletion && p.IsNew)
                    .ToList();

                foreach (var library in newLibrariesToDelete)
                {
                    InputParameters.Remove(library);
                }

                // 3. Lưu các library mới và đã chỉnh sửa
                var librariesToSave = InputParameters
                    .Where(p => (p.IsNew || p.IsEdited) && !p.IsMarkedForDeletion)
                    .ToList();

                foreach (var libraryDto in librariesToSave)
                {
                    var library = libraryDto.ToEntity();

                    // Lấy data tương ứng
                    BienTan bienTanEntity = null;
                    CamBien camBienEntity = null;
                    OngGio ongGioEntity = null;

                    if (_currentViewedParamId == libraryDto.LibId)
                    {
                        bienTanEntity = BienTan?.ToEntity();
                        camBienEntity = CamBien?.ToEntity();
                        ongGioEntity = OngGio?.ToEntity();
                    }
                    else
                    {
                        // Không đang xem - load từ DB
                        var detail = await _parameterService.GetLibraryByIdAsync(libraryDto.LibId);
                        bienTanEntity = detail.Item1;
                        camBienEntity = detail.Item2;
                        ongGioEntity = detail.Item3;
                    }

                    if (libraryDto.IsNew)
                    {
                        // Insert mới
                        await _parameterService.AddLibraryAsync(library, bienTanEntity, camBienEntity, ongGioEntity);
                    }
                    else if (libraryDto.IsEdited)
                    {
                        // Update
                        await _parameterService.UpdateLibraryAsync(library.LibId, bienTanEntity, camBienEntity, ongGioEntity);
                    }

                    // Reset flags
                    libraryDto.IsNew = false;
                    libraryDto.IsEdited = false;
                }

                MessageBoxHelper.ShowSuccess("Lưu thành công!");

                LoadData();
                // recall to reload current viewed param
                var locator = (ViewModelLocator)Application.Current.Resources["Locator"];
                locator.ProjectViewModel.LoadParam();

                if (CurrentParameter == null ||
                    !InputParameters.Any(p => p.LibId == CurrentParameter.LibId))
                {
                    InitializeEmptyData();
                    _currentViewedParamId = null;
                }
            }
            catch (Exception ex)
            {
                MessageBoxHelper.ShowError($"Lỗi khi lưu: {ex.Message}");
            }
        }

        private void ExecuteDeleteCommand(object parameter)
        {
            try
            {
                LibraryDto libraryToDelete = null;

                if (parameter is LibraryDto library)
                {
                    libraryToDelete = library;
                }
                else if (CurrentParameter != null)
                {
                    libraryToDelete = CurrentParameter;
                }

                if (libraryToDelete == null) return;

                var result = MessageBoxHelper.ShowQuestion(
                    $"Bạn có chắc chắn muốn xóa library '{libraryToDelete.LibName}'?\n\n" +
                    "Lưu ý: Thao tác này có thể hoàn tác trước khi lưu.");

                if (result)
                {
                    if (libraryToDelete.IsNew)
                    {
                        // Xóa trực tiếp
                        InputParameters.Remove(libraryToDelete);

                        if (CurrentParameter == libraryToDelete)
                        {
                            InitializeEmptyData();
                            _currentViewedParamId = null;
                        }
                    }
                    else
                    {
                        // Đánh dấu xóa
                        libraryToDelete.IsMarkedForDeletion = true;

                        if (_currentViewedParamId == libraryToDelete.LibId)
                        {
                            InitializeEmptyData();
                            _currentViewedParamId = null;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBoxHelper.ShowError($"Lỗi khi xóa library: {ex.Message}");
            }
        }

        private void ExecuteUndoDeleteCommand(object parameter)
        {
            try
            {
                if (parameter is Guid paramId)
                {
                    var library = InputParameters.FirstOrDefault(p => p.LibId == paramId);

                    if (library != null && library.IsMarkedForDeletion)
                    {
                        library.IsMarkedForDeletion = false;
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