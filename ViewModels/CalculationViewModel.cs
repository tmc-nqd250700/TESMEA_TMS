using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows.Input;
using TESMEA_TMS.Configs;
using TESMEA_TMS.DTOs;
using TESMEA_TMS.Helpers;
using TESMEA_TMS.Services;

namespace TESMEA_TMS.ViewModels
{
    public class CalculationViewModel : ViewModelBase
    {
        public ThongTinDuAn ThongTinDuAn { get; set; }

        private ThongSoDuongOngGio _ongGio;
        public ThongSoDuongOngGio OngGio
        {
            get => _ongGio;
            set
            {
                _ongGio = value;
                OnPropertyChanged(nameof(OngGio));
            }
        }
        private ThongSoCoBanCuaQuat _quat;
        public ThongSoCoBanCuaQuat Quat
        {
            get => _quat;
            set
            {
                _quat = value;
                OnPropertyChanged(nameof(Quat));
            }
        }
        private ObservableCollection<Measure> _danhSachThongSoDoKiem;
        public ObservableCollection<Measure> DanhSachThongSoDoKiem
        {
            get => _danhSachThongSoDoKiem;
            set
            {
                _danhSachThongSoDoKiem = value;
                OnPropertyChanged(nameof(DanhSachThongSoDoKiem));
            }
        }



        public ObservableCollection<ComboBoxInfo> ReportTemplates { get; set; }
        private ComboBoxInfo _selectedReportTemplate;
        public ComboBoxInfo SelectedReportTemplate
        {
            get => _selectedReportTemplate;
            set
            {
                _selectedReportTemplate = value;
                OnPropertyChanged(nameof(SelectedReportTemplate));
            }
        }
        private bool IsEn { get; set; } = UserSetting.Instance.Language == "en";
        public ICommand BrowseCommand { get; }
        public ICommand ExportResultCommand { get; }
        public ICommand ExportReportCommand { get; }
        public ICommand BackCommand { get; }
        public ICommand FinishCommand { get; }


        // services
        private readonly IFileService _fileService;
        public CalculationViewModel(IFileService fileService)
        {
            _fileService = fileService;


            OngGio = new ThongSoDuongOngGio();
            Quat = new ThongSoCoBanCuaQuat();
            DanhSachThongSoDoKiem = new ObservableCollection<Measure>();
            ReportTemplates = new ObservableCollection<ComboBoxInfo>();
            ReportTemplates.Add(new ComboBoxInfo("DESIGN", IsEn ? "Design condition" : "Điều kiện thiết kế"));
            ReportTemplates.Add(new ComboBoxInfo("NORMALIZED", IsEn ? "Normalized condition" : "Điều kiện tiêu chuẩn"));
            ReportTemplates.Add(new ComboBoxInfo("OPERATION", IsEn ? "Operation condition" : "Điều kiện hoạt động"));
            ReportTemplates.Add(new ComboBoxInfo("FULL", IsEn ? "Full" : "Tất cả"));
            SelectedReportTemplate = new ComboBoxInfo();
            SelectedReportTemplate.Value = ReportTemplates[0].Value;

            BrowseCommand = new ViewModelCommand(_ => ExecuteBrowseCommand());

            BackCommand = new ViewModelCommand(CanExecuteCommand, ExecuteBackCommand);
            FinishCommand = new ViewModelCommand(CanExecuteCommand, ExecuteFinishCommand);
            ExportResultCommand = new ViewModelCommand(CanExecuteCommand, ExecuteExportResultCommand);
            ExportReportCommand = new ViewModelCommand(CanExecuteCommand, ExecuteExportReportCommand);
        }

        private bool CanExecuteCommand(object obj)
        {
            return true;
        }

        private void ClearData()
        {
            OngGio = new ThongSoDuongOngGio();
            Quat = new ThongSoCoBanCuaQuat();
            DanhSachThongSoDoKiem.Clear();
        }

        private void ExecuteBackCommand(object obj)
        {
            var mainViewModel = ((App)System.Windows.Application.Current).Resources["Locator"] as ViewModelLocator;
            if (mainViewModel != null)
            {
                ClearData();
                mainViewModel.MainViewModel.CurrentView = new TESMEA_TMS.Views.ProjectView();
            }
        }

        private void ExecuteFinishCommand(object obj)
        {
            var isFinish = MessageBoxHelper.ShowQuestion("Bạn có chắc chắn muốn hoàn thành đo kiểm và quay lại trang chính không?");
            if (isFinish)
            {
                var mainViewModel = ((App)System.Windows.Application.Current).Resources["Locator"] as ViewModelLocator;
                if (mainViewModel != null)
                {
                    ClearData();
                    mainViewModel.ProjectViewModel.ClearData();
                    mainViewModel.MainViewModel.CurrentView = new TESMEA_TMS.Views.ProjectView();
                }
            }
        }

        private void ExecuteBrowseCommand()
        {
            var ofd = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xls",
                Title = "Chọn file thông số định mức"
            };
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                var importResult = _fileService.ImportCalculation(ofd.FileName);
                if (importResult != null)
                {
                    OngGio = importResult.ThongSoDuongOngGio;
                    Quat = importResult.ThongSoCoBanCuaQuat;
                    DanhSachThongSoDoKiem = new ObservableCollection<Measure>(importResult.DanhSachThongSoDoKiem);
                }
            }
        }

        private async void ExecuteExportResultCommand(object obj)
        {
            var timestamp = DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss");
            var sfd = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xls",
                Title = IsEn ? "Select where to save measurement calculation results" : "Chọn nơi lưu kết quả tính toán",
                FileName = $"Result_{timestamp}.xlsx",
                InitialDirectory = ThongTinDuAn.ThamSo.DuongDanLuuDuAn
            };
            if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                var tsdv = new ThongSoDauVao
                {
                    ThongSoDuongOngGio = OngGio,
                    ThongSoCoBanCuaQuat = Quat,
                    DanhSachThongSoDoKiem = DanhSachThongSoDoKiem.ToList()
                };
                await _fileService.ExportExcelTestResult(
                         outputPath: sfd.FileName,
                         option: SelectedReportTemplate?.Value ?? "DESIGN",
                         tsdv: tsdv,
                         project: ThongTinDuAn
                     );
                MessageBoxHelper.ShowSuccess(IsEn ? "Export result successfully" : "Kết quả tính toán đã được xuất thành công");
            }
        }

        private async void ExecuteExportReportCommand(object obj)
        {
            var timestamp = DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss");
            var sfd = new SaveFileDialog
            {
                Filter = "Word Files|*.docx",
                Title = IsEn ? "Select where to save measurement reports" : "Chọn nơi lưu báo cáo",
                FileName = $"Report_{timestamp}.docx",
                InitialDirectory = ThongTinDuAn.ThamSo.DuongDanLuuDuAn
            };
            if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                var tsdv = new ThongSoDauVao
                {
                    ThongSoDuongOngGio = OngGio,
                    ThongSoCoBanCuaQuat = Quat,
                    DanhSachThongSoDoKiem = DanhSachThongSoDoKiem.ToList()
                };
                await _fileService.ExportReportTestResult(
                         outputPath: sfd.FileName,
                         option: SelectedReportTemplate?.Value ?? "DESIGN",
                         tsdv: tsdv,
                         project: ThongTinDuAn
                     );
                MessageBoxHelper.ShowSuccess(IsEn ? "Export report successfully" : "Báo cáo đã được xuất thành công");
            }
        }
    }
}
