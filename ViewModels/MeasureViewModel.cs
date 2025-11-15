using MaterialDesignThemes.Wpf;
using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Legends;
using OxyPlot.Series;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Windows.Input;
using TESMEA_TMS.Configs;
using TESMEA_TMS.DTOs;
using TESMEA_TMS.Helpers;
using TESMEA_TMS.Models.Entities;
using TESMEA_TMS.Services;
using TESMEA_TMS.Views;
using Application = System.Windows.Application;
using MarkerType = OxyPlot.MarkerType;

namespace TESMEA_TMS.ViewModels
{

    public class MeasureViewModel : ViewModelBase
    {
        public ObservableCollection<Measure> MeasureRows { get; set; }
        public float StandardDeviation { get; set; }
        public float TimeRange { get; set; }


        private Measure _selectedMeasure;
        public Measure SelectedMeasure
        {
            get => _selectedMeasure;
            set
            {
                _selectedMeasure = value;
                OnPropertyChanged(nameof(SelectedMeasure));
                //if (TrendLineDialog != null && TrendLineDialog.IsLoaded && _selectedMeasure != null)
                //{
                //    TrendLineDialog.LoadTrendDataAndDraw(_selectedMeasure.k);
                //}
            }
        }

        private bool _isMeasuring;
        public bool CanMeasure => !_isMeasuring && _externalAppService.IsConnectedToSimatic;

        private bool _isConnected;
        private bool _isCompleted;
        private CamBien _camBien { get; set; } = new CamBien();


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
        public ThongTinDuAn ThongTinDuAn { get; set; } = new ThongTinDuAn();

        private ThongTinDoKiem _thongTinDoKiem;
        public ThongTinDoKiem ThongTinDoKiem
        {
            get => _thongTinDoKiem;
            set
            {
                _thongTinDoKiem = value;
                OnPropertyChanged(nameof(ThongTinDoKiem));
            }
        }


        #region Các thuộc tính hiển thị trên màn hình

        public string SimaticStatus
        {
            get
            {
                if (_externalAppService == null)
                    return "STOP";
                return _externalAppService.IsConnectedToSimatic ? "RUNNING" : "STOP";
            }
        }

        private ParameterShow _parameterShow;
        public ParameterShow ParameterShow
        {
            get => _parameterShow;
            set
            {
                _parameterShow = value;
                OnPropertyChanged(nameof(ParameterShow));
            }
        }


        #endregion

        public ObservableCollection<MeasureResponse> MeasureResponses { get; set; } = new ObservableCollection<MeasureResponse>();

        private PlotModel _powerPlotModel;
        public PlotModel PowerPlotModel
        {
            get => _powerPlotModel;
            set
            {
                _powerPlotModel = value;
                OnPropertyChanged(nameof(PowerPlotModel));
            }
        }

        private PlotModel _efficiencyPlotModel;
        public PlotModel EfficiencyPlotModel
        {
            get => _efficiencyPlotModel;
            set
            {
                _efficiencyPlotModel = value;
                OnPropertyChanged(nameof(EfficiencyPlotModel));
            }
        }



        public bool IsTorqueVisible { get; set; } = false;
        public bool IsTorque { get; set; } = false;
        private bool IsEn { get; set; } = UserSetting.Instance.Language == "en";

        private Views.TrendLineDialog TrendLineDialog;


        // services
        private readonly IExternalAppService _externalAppService;
        private readonly IParameterService _parameterService;
        private readonly IAppNavigationService _appNavigationService;

        // Commands
        public ICommand ConnectCommand { get; }
        public ICommand StartCommand { get; }
        public ICommand StopCommand { get; }
        public ICommand ResetCommand { get; }
        public ICommand TrendCommand { get; }

        public ICommand ExportResultCommand { get; }
        public ICommand ExportReportCommand { get; }
        public ICommand BackCommand { get; }
        public ICommand FinishCommand { get; }
        public MeasureViewModel(IExternalAppService externalAppService, IParameterService parameterService, IAppNavigationService appNavigationService)
        {
            _externalAppService = externalAppService;
            _parameterService = parameterService;
            _appNavigationService = appNavigationService;

            MeasureRows = new ObservableCollection<Measure>();
            ThongTinDoKiem = new ThongTinDoKiem();
            ParameterShow = new ParameterShow();
            InitializePowerPlotModel();
            InitializeEfficiencyPlotModel();

            _externalAppService.OnSimaticConnectionChanged += connected =>
            {
                Application.Current.Dispatcher.Invoke(() =>
                {
                    OnPropertyChanged(nameof(SimaticStatus));
                });
            };

            _externalAppService.OnSimaticResultReceived += result =>
            {
                Application.Current.Dispatcher.Invoke(() =>
                {
                    UpdateMeasureStatusFromSimaticResult(result);
                    SelectedMeasure = MeasureRows.FirstOrDefault(m => m.k == result.k);
                });
            };

            // event invoke khi mỗi điểm đo trên dải tần số hoàn thành
            _externalAppService.OnMeasurePointCompleted += OnMeasurePointCompletedHandler;
            // event invoke khi một dải tần số hoàn thành
            _externalAppService.OnMeasureRangeCompleted += OnMeasureRangeCompletedHandler;

            ReportTemplates = new ObservableCollection<ComboBoxInfo>();
            ReportTemplates.Add(new ComboBoxInfo("DESIGN", "Design condition"));
            ReportTemplates.Add(new ComboBoxInfo("NORMALIZED", "Normalized condition"));
            ReportTemplates.Add(new ComboBoxInfo("OPERATION", "Operation condition"));
            ReportTemplates.Add(new ComboBoxInfo("FULL", "Full"));
            SelectedReportTemplate = new ComboBoxInfo();
            SelectedReportTemplate.Value = ReportTemplates[0].Value;

            ConnectCommand = new ViewModelCommand(CanConnect, ExecuteConnectCommand);
            StartCommand = new ViewModelCommand(CanStart, ExecuteStartCommand);
            StopCommand = new ViewModelCommand(CanStop, ExecuteStopCommand);
            ResetCommand = new ViewModelCommand(CanReset, ExecuteResetCommand);
            BackCommand = new ViewModelCommand(CanExecuteCommand, ExecuteBackCommand);
            FinishCommand = new ViewModelCommand(CanExecuteCommand, ExecuteFinishCommand);
            TrendCommand = new ViewModelCommand(CanTrend, ExecuteTrendCommand);

            ExportResultCommand = new ViewModelCommand(CanExportCommand, ExecuteExportResultCommand);
            ExportReportCommand = new ViewModelCommand(CanExportCommand, ExecuteExportReportCommand);
        }


        private bool CanExecuteCommand(object obj)
        {
            return !_isMeasuring;
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

        public void InitializePowerPlotModel()
        {
            PowerPlotModel = new PlotModel
            {
                Title = IsEn ? "Air volume - Power curve" : "Đặc tuyến Công suất - Lưu lượng",
                Background = OxyColors.White,
                IsLegendVisible = true,
                PlotMargins = new OxyThickness(60, 10, 10, 60),
                Legends =
                    {
                        new Legend
                        {
                            LegendPlacement = LegendPlacement.Outside,
                            LegendPosition = LegendPosition.TopLeft,
                            LegendOrientation = LegendOrientation.Horizontal,
                            LegendMargin = 10
                        }
                    }
            };

            // Thêm trục X
            PowerPlotModel.Axes.Add(new LinearAxis
            {
                Position = AxisPosition.Bottom,
                Title = IsEn ? "Air volume (m3/h)" : "Lưu lượng (m3/h)",
                Minimum = 0,
                IsPanEnabled = false,
                IsZoomEnabled = false,
                MajorGridlineStyle = LineStyle.Solid,
                MinorGridlineStyle = LineStyle.Dot
            });

            PowerPlotModel.Axes.Add(new LinearAxis
            {
                Position = AxisPosition.Left,
                Title = IsEn ? "Power (kW)" : "Công suất (kW)",
                Minimum = 0,
                IsPanEnabled = false,
                IsZoomEnabled = false,
                MajorGridlineStyle = LineStyle.Solid,
                MinorGridlineStyle = LineStyle.Dot
            });

            // Thêm trục Y
            var scatterSeries = new ScatterSeries
            {
                Title = IsEn ? "Power" : "Công suất",
                MarkerType = MarkerType.Circle,
                MarkerFill = OxyColors.Red,
                MarkerStroke = OxyColors.Red,
                MarkerStrokeThickness = 1,
                MarkerSize = 4
            };
            PowerPlotModel.Series.Add(scatterSeries);
        }

        public void InitializeEfficiencyPlotModel()
        {
            EfficiencyPlotModel = new PlotModel
            {
                Title = IsEn ?
                    "Air volume - Pressure - Efficiency curve" :
                    "Đặc tuyến Áp suất - Hiệu suất - Lưu lượng",
                Background = OxyColors.White,
                PlotMargins = new OxyThickness(60, 10, 60, 60),
                Legends =
                    {
                        new Legend
                        {
                            LegendPlacement = LegendPlacement.Outside,
                            LegendPosition = LegendPosition.TopCenter,
                            LegendOrientation = LegendOrientation.Horizontal,
                            LegendMargin = 10
                        }
                    }
            };

            // Trục X (Air volume)
            EfficiencyPlotModel.Axes.Add(new LinearAxis
            {
                Position = AxisPosition.Bottom,
                Title = IsEn ? "Air volume (m3/h)" : "Lưu lượng (m3/h)",
                Minimum = 0,
                IsPanEnabled = false,
                IsZoomEnabled = false,
                MajorGridlineStyle = LineStyle.Solid,
                MinorGridlineStyle = LineStyle.Dot
            });

            // Trục Y1 (Pressure)
            EfficiencyPlotModel.Axes.Add(new LinearAxis
            {
                Position = AxisPosition.Left,
                Title = IsEn ? "Pressure (Pa)" : "Áp suất (Pa)",
                Minimum = 0,
                IsPanEnabled = false,
                IsZoomEnabled = false,
                MajorGridlineStyle = LineStyle.Solid,
                MinorGridlineStyle = LineStyle.Dot,
                Key = "PressureAxis"
            });

            // Trục Y2 (Efficiency)
            EfficiencyPlotModel.Axes.Add(new LinearAxis
            {
                Position = AxisPosition.Right,
                Title = IsEn ? "Efficiency (%)" : "Hiệu suất (%)",
                Minimum = 0,
                IsPanEnabled = false,
                IsZoomEnabled = false,
                MajorGridlineStyle = LineStyle.None,
                MinorGridlineStyle = LineStyle.None,
                TextColor = OxyColors.Blue,
                TitleColor = OxyColors.Blue,
                Key = "EfficiencyAxis"
            });

            // Thêm các series
            var staticPressureSeries = new ScatterSeries
            {
                Title = IsEn ? "Static Pressure" : "Áp suất tĩnh",
                MarkerType = MarkerType.Circle,
                MarkerFill = OxyColors.Red,
                MarkerStroke = OxyColors.Red,
                MarkerStrokeThickness = 1,
                MarkerSize = 4,
                YAxisKey = "PressureAxis"
            };

            var totalPressureSeries = new ScatterSeries
            {
                Title = IsEn ? "Total Pressure" : "Áp suất tổng",
                MarkerType = MarkerType.Diamond,
                MarkerFill = OxyColors.Blue,
                MarkerStroke = OxyColors.Blue,
                MarkerStrokeThickness = 1,
                MarkerSize = 4,
                YAxisKey = "PressureAxis"
            };

            var staticEfficiencySeries = new ScatterSeries
            {
                Title = IsEn ? "Static Efficiency" : "Hiệu suất tĩnh",
                MarkerType = MarkerType.Square,
                MarkerFill = OxyColors.Black,
                MarkerStroke = OxyColors.Black,
                MarkerStrokeThickness = 1,
                MarkerSize = 4,
                YAxisKey = "EfficiencyAxis"
            };

            var totalEfficiencySeries = new ScatterSeries
            {
                Title = IsEn ? "Total Efficiency" : "Hiệu suất tổng",
                MarkerType = MarkerType.Triangle,
                MarkerFill = OxyColors.DarkGreen,
                MarkerStroke = OxyColors.DarkGreen,
                MarkerStrokeThickness = 1,
                MarkerSize = 4,
                YAxisKey = "EfficiencyAxis"
            };

            EfficiencyPlotModel.Series.Add(staticPressureSeries);
            EfficiencyPlotModel.Series.Add(totalPressureSeries);
            EfficiencyPlotModel.Series.Add(staticEfficiencySeries);
            EfficiencyPlotModel.Series.Add(totalEfficiencySeries);
        }



        private void ExecuteFinishCommand(object obj)
        {
            var isFinish = MessageBoxHelper.ShowQuestion("Bạn có chắc chắn muốn hoàn thành đo kiểm và quay lại trang chính không?");
            if (isFinish)
            {
                var locator = ((App)System.Windows.Application.Current).Resources["Locator"] as ViewModelLocator;
                if (locator != null)
                {
                    ClearData();
                    locator.ProjectViewModel.ClearData();
                    locator.ProjectViewModel.ClearData();
                    locator.MainViewModel.CurrentView = new TESMEA_TMS.Views.ProjectView();
                }

                if (Directory.Exists(UserSetting.TOMFAN_folder))
                {
                    // Xóa tất cả file
                    foreach (var file in Directory.GetFiles(UserSetting.TOMFAN_folder))
                    {
                        try { File.WriteAllText(file, string.Empty); } catch { /* ignore */ }
                    }
                    // Xóa tất cả thư mục con
                    foreach (var dir in Directory.GetDirectories(UserSetting.TOMFAN_folder))
                    {
                        try { Directory.Delete(dir, true); } catch { }
                    }
                }
            }
        }


        private void ClearData()
        {
            MeasureRows.Clear();
            ThongTinDoKiem = new ThongTinDoKiem();
            ThongTinDuAn = new ThongTinDuAn();
        }
        private bool CanConnect(object obj) => !_isConnected;
        private bool CanStart(object obj) => _isConnected && !_isMeasuring && !_isCompleted;
        private bool CanStop(object obj) => _isMeasuring;
        private bool CanReset(object obj) => _isCompleted && !_isMeasuring;
        private bool CanTrend(object obj) => SelectedMeasure != null && (TrendLineDialog == null || !TrendLineDialog.IsLoaded);

        private async void ExecuteConnectCommand(object obj)
        {
            var splashViewModel = new ProgressSplashViewModel
            {
                Message = "Đang kiểm tra kết nối với Simatic...",
                IsIndeterminate = true
            };
            var splash = new Views.CustomControls.ProgressSplashContent { DataContext = splashViewModel };
            var dialogTask = DialogHost.Show(splash, "MainDialogHost");
            try
            {
                (var bienTan, _camBien, var ongGio) = await _parameterService.GetLibraryByIdAsync(Guid.Parse(ThongTinDuAn.ThamSo.KieuKiemThu));
                if (bienTan == null || _camBien == null || ongGio == null)
                {
                    throw new BusinessException("Không tìm thấy kiểu kiểm thử");
                }

                if (!MeasureRows.Any())
                {
                    throw new BusinessException("Không có dữ liệu đo kiểm cho kịch bản đã chọn");
                }
                SelectedMeasure = MeasureRows.First();
                var scenario = await _parameterService.GetScenarioAsync(ThongTinDuAn.ThamSo.KichBan);
                if (scenario == null)
                {
                    throw new BusinessException("Không tìm thấy kịch bản đo kiểm");
                }

                await _externalAppService.ConnectExchangeAsync(MeasureRows.ToList(), bienTan, _camBien, ongGio, ThongTinDuAn.ThongTinMauThuNghiem, scenario.StandardDeviation, scenario.TimeRange);
                _isConnected = true;
                _isCompleted = false;
                if (DialogHost.IsDialogOpen("MainDialogHost"))
                    DialogHost.Close("MainDialogHost");
                OnPropertyChanged(nameof(SimaticStatus));
            }
            catch (Exception ex)
            {
                if (DialogHost.IsDialogOpen("MainDialogHost"))
                    DialogHost.Close("MainDialogHost");
                throw;
            }
        }

        private async void ExecuteStartCommand(object obj)
        {
            _isMeasuring = true;
            _isCompleted = false;
            try
            {

                await _externalAppService.StartExchangeAsync();

            }
            catch (Exception ex)
            {
                _isMeasuring = false;
                MessageBoxHelper.ShowError(ex.Message);
            }
            finally
            {
                _isMeasuring = false;
                _isCompleted = true;
            }
        }

        private async void ExecuteStopCommand(object obj)
        {
            var splashViewModel = new ProgressSplashViewModel
            {
                Message = "Đang đóng kết nối với Simatic...",
                IsIndeterminate = true
            };
            var splash = new Views.CustomControls.ProgressSplashContent { DataContext = splashViewModel };

            _ = Application.Current.Dispatcher.Invoke(
                 () => DialogHost.Show(splash, "MainDialogHost"),
                 System.Windows.Threading.DispatcherPriority.Send
             );
            await Task.Delay(500);

            try
            {
                await _externalAppService.StopExchangeAsync();
                Application.Current.Dispatcher.Invoke(() =>
                {
                    DialogHost.Close("MainDialogHost");
                });
            }
            catch (Exception ex)
            {
                Application.Current.Dispatcher.Invoke(() =>
                {
                    DialogHost.Close("MainDialogHost");
                });
            }
            finally
            {
                _isMeasuring = false;
                _isCompleted = true;
                OnPropertyChanged(nameof(SimaticStatus));
                foreach (var measure in MeasureRows)
                {
                    if (measure.F != MeasureStatus.Completed)
                    {
                        measure.F = MeasureStatus.Error;
                    }
                }
                OnPropertyChanged(nameof(MeasureRows));
            }

        }

        private void ExecuteResetCommand(object obj)
        {
            foreach (var measure in MeasureRows)
            {
                measure.F = MeasureStatus.Pending;
            }
            OnPropertyChanged(nameof(MeasureRows));
            MeasureResponses.Clear();
            OnPropertyChanged(nameof(MeasureResponses));
            ClearPlots();
            ParameterShow = new ParameterShow();
            OnPropertyChanged(nameof(ParameterShow));
            if (Directory.Exists(UserSetting.TOMFAN_folder))
            {
                // Xóa tất cả file
                foreach (var file in Directory.GetFiles(UserSetting.TOMFAN_folder))
                {
                    try { File.WriteAllText(file, string.Empty); } catch { /* ignore */ }
                }
                foreach (var file in Directory.GetFiles(UserSetting.TOMFAN_folder))
                {
                    try { File.Delete(file); } catch { /* ignore */ }
                }
            }
            _externalAppService.StopExchangeAsync().Wait();
            _isCompleted = false;
            _isConnected = false;
        }

        private void ClearPlots()
        {
            // Xóa dữ liệu PowerPlotModel
            if (PowerPlotModel != null && PowerPlotModel.Series.Count > 0)
            {
                foreach (var series in PowerPlotModel.Series)
                {
                    if (series is ScatterSeries scatterSeries)
                        scatterSeries.Points.Clear();
                    if (series is LineSeries lineSeries)
                        lineSeries.Points.Clear();
                }
                PowerPlotModel.InvalidatePlot(true);
            }

            // Xóa dữ liệu EfficiencyPlotModel
            if (EfficiencyPlotModel != null && EfficiencyPlotModel.Series.Count > 0)
            {
                foreach (var series in EfficiencyPlotModel.Series)
                {
                    if (series is ScatterSeries scatterSeries)
                        scatterSeries.Points.Clear();
                    if (series is LineSeries lineSeries)
                        lineSeries.Points.Clear();
                }
                EfficiencyPlotModel.InvalidatePlot(true);
            }
        }

        private void ExecuteTrendCommand(object obj)
        {
            if (SelectedMeasure == null)
                return;

            if (TrendLineDialog == null || !TrendLineDialog.IsLoaded)
            {
                TrendLineDialog = new Views.TrendLineDialog(SelectedMeasure.k);
                TrendLineDialog.Closed += (s, e) => TrendLineDialog = null;
                TrendLineDialog.Show();
            }
            //else
            //{
            //    TrendLineDialog.LoadTrendDataAndDraw(SelectedMeasure.k);
            //    TrendLineDialog.Activate();
            //}


            //_appNavigationService.ShowTrendDialog(SelectedMeasure.k);
        }

        public void UpdateMeasureStatusFromSimaticResult(Measure result)
        {
            if (MeasureRows == null || result == null) return;

            var measure = MeasureRows.FirstOrDefault(m => m.k == result.k);
            if (measure != null)
            {
                measure.F = result.F;
                OnPropertyChanged(nameof(MeasureRows));
            }
        }

        private ScatterSeries _powerScatterLiveSeries;
        private ScatterSeries _psScatterLiveSeries;
        private ScatterSeries _ptScatterLiveSeries;
        private ScatterSeries _seScatterLiveSeries;
        private ScatterSeries _teScatterLiveSeries;
        private void OnMeasurePointCompletedHandler(MeasureResponse response, ParameterShow paramShow)
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                // Hiển thị kết quả lên màn hình
                ParameterShow = paramShow;
                OnPropertyChanged(nameof(ParameterShow));

                // Thêm vào danh sách kết quả
                MeasureResponses.Add(response);

                #region Power plot
                // Kiểm tra xem series ScatterPoints đầu tiên có tồn tại không.
                // Đây là series tạm thời dùng để vẽ các điểm đang đo.
                //var powerSeries = PowerPlotModel.Series.OfType<ScatterSeries>().FirstOrDefault();
                //if (powerSeries == null) return;
                //powerSeries.Points.Add(new ScatterPoint(response.Airflow, response.Power));
                
                if (_powerScatterLiveSeries == null)
                {
                    _powerScatterLiveSeries = PowerPlotModel.Series.OfType<ScatterSeries>().FirstOrDefault();
                }
                _powerScatterLiveSeries?.Points.Add(new ScatterPoint(response.Airflow, response.Power));


                // Cập nhật scale trục
                var xAxisPw = PowerPlotModel.Axes[0] as LinearAxis;
                if (xAxisPw != null && response.Airflow > xAxisPw.ActualMaximum - xAxisPw.MajorStep)
                {
                    xAxisPw.Maximum = response.Airflow + xAxisPw.MajorStep;
                }

                var yAxisPw = PowerPlotModel.Axes[1] as LinearAxis;
                if (yAxisPw != null && response.Power > yAxisPw.ActualMaximum - yAxisPw.MajorStep)
                {
                    yAxisPw.Maximum = response.Power + yAxisPw.MajorStep;
                }

                
                
                #endregion

                #region Eff Plot
                //Giả sử 4 series ScatterPoints đầu tiên đã được khởi tạo
                if (EfficiencyPlotModel?.Series == null || EfficiencyPlotModel.Series.Count < 4)
                    return;

                // Cần truy cập các ScatterSeries bằng Index hoặc Key/Title nếu bạn có nhiều series khác
                //var psPoint = EfficiencyPlotModel.Series[0] as ScatterSeries;
                //var ptPoint = EfficiencyPlotModel.Series[1] as ScatterSeries;
                //var sePoint = EfficiencyPlotModel.Series[2] as ScatterSeries;
                //var tePoint = EfficiencyPlotModel.Series[3] as ScatterSeries;

                //// thêm point cho eff plot
                //psPoint?.Points.Add(new ScatterPoint(response.Airflow, response.Ps));
                //ptPoint?.Points.Add(new ScatterPoint(response.Airflow, response.Pt));
                //sePoint?.Points.Add(new ScatterPoint(response.Airflow, response.SEff));
                //tePoint?.Points.Add(new ScatterPoint(response.Airflow, response.TEff));

                if (_psScatterLiveSeries == null)
                    _psScatterLiveSeries = EfficiencyPlotModel.Series[0] as ScatterSeries;
                if (_ptScatterLiveSeries == null)
                    _ptScatterLiveSeries = EfficiencyPlotModel.Series[1] as ScatterSeries;
                if (_seScatterLiveSeries == null)
                    _seScatterLiveSeries = EfficiencyPlotModel.Series[2] as ScatterSeries;
                if (_teScatterLiveSeries == null)
                    _teScatterLiveSeries = EfficiencyPlotModel.Series[3] as ScatterSeries;

                _psScatterLiveSeries?.Points.Add(new ScatterPoint(response.Airflow, response.Ps));
                _ptScatterLiveSeries?.Points.Add(new ScatterPoint(response.Airflow, response.Pt));
                _seScatterLiveSeries?.Points.Add(new ScatterPoint(response.Airflow, response.SEff));
                _teScatterLiveSeries?.Points.Add(new ScatterPoint(response.Airflow, response.TEff));

                // scale trục
                var xAxisEff = EfficiencyPlotModel.Axes[0] as LinearAxis;
                if (xAxisEff != null && response.Airflow > xAxisEff.ActualMaximum - xAxisEff.MajorStep)
                {

                    xAxisEff.Maximum = response.Airflow + xAxisEff.MajorStep;
                }

                var y1Axis = EfficiencyPlotModel.Axes.FirstOrDefault(a => a.Key == "PressureAxis") as LinearAxis;
                if (y1Axis != null && Math.Max(response.Ps, response.Pt) > y1Axis.ActualMaximum - y1Axis.MajorStep)
                {

                    y1Axis.Maximum = Math.Max(response.Ps, response.Pt) + y1Axis.MajorStep;
                }

                var y2Axis = EfficiencyPlotModel.Axes.FirstOrDefault(a => a.Key == "EfficiencyAxis") as LinearAxis;
                if (y2Axis != null && Math.Max(response.SEff, response.TEff) > y2Axis.ActualMaximum - y2Axis.MajorStep)
                {
                    y2Axis.Maximum = Math.Max(response.SEff, response.TEff) + y2Axis.MajorStep;
                }

                #endregion

                PowerPlotModel.InvalidatePlot(true);
                EfficiencyPlotModel.InvalidatePlot(true);
            });
        }
        private void OnMeasureRangeCompletedHandler(MeasureFittingFC fitting, Measure rangeMeasure)
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                // Log toàn bộ giá trị fitting
                Debug.WriteLine("==== LOG FITTING ====");
                for (int i = 0; i < fitting.FlowPoint_ft.Length; i++)
                {
                    Debug.WriteLine($"i={i} | FlowPoint_ft={fitting.FlowPoint_ft[i]} | PrPoint_ft={fitting.PrPoint_ft[i]} | PsPoint_ft={fitting.PsPoint_ft[i]} | PtPoint_ft={fitting.PtPoint_ft[i]} | EsPoint_ft={fitting.EsPoint_ft[i]} | EtPoint_ft={fitting.EtPoint_ft[i]} | PrtPoint_ft={fitting.PrtPoint_ft[i]} | EstPoint_ft={fitting.EstPoint_ft[i]} | EttPoint_ft={fitting.EttPoint_ft[i]} | Ope_FlowPoint={fitting.Ope_FlowPoint[i]} | Ope_PsPoint={fitting.Ope_PsPoint[i]} | Ope_PtPoint={fitting.Ope_PtPoint[i]} | Ope_EsPoint={fitting.Ope_EsPoint[i]} | Ope_EtPoint={fitting.Ope_EtPoint[i]} | Ope_PrPoint={fitting.Ope_PrPoint[i]} | Ope_EstPoint={fitting.Ope_EstPoint[i]} | Ope_EttPoint={fitting.Ope_EttPoint[i]} | PrtPoint={fitting.PrtPoint[i]}");
                }
                Debug.WriteLine("==== END LOG FITTING ====");
                //if (_powerScatterLiveSeries != null) PowerPlotModel.Series.Remove(_powerScatterLiveSeries);
                //if (_psScatterLiveSeries != null) EfficiencyPlotModel.Series.Remove(_psScatterLiveSeries);
                //if (_ptScatterLiveSeries != null) EfficiencyPlotModel.Series.Remove(_ptScatterLiveSeries);
                //if (_seScatterLiveSeries != null) EfficiencyPlotModel.Series.Remove(_seScatterLiveSeries);
                //if (_teScatterLiveSeries != null) EfficiencyPlotModel.Series.Remove(_teScatterLiveSeries);
                //_powerScatterLiveSeries = _psScatterLiveSeries = _ptScatterLiveSeries = _seScatterLiveSeries = _teScatterLiveSeries = null;

                #region powerPlot
                PowerPlotModel.Title = IsEn ? "Air volume - Power curve" : "Đặc tuyến Công suất - Lưu lượng";
                PowerPlotModel.Background = OxyColors.White;
                // Cập nhật trục X
                var xAxisPw = PowerPlotModel.Axes.FirstOrDefault(a => a.Position == AxisPosition.Bottom) as LinearAxis;
                if (xAxisPw == null)
                {
                    xAxisPw = new LinearAxis
                    {
                        Position = AxisPosition.Bottom,
                        Title = IsEn ? "Air volume (m3/h)" : "Lưu lượng (m3/h)",
                        Minimum = 0,
                        Maximum = fitting.FlowPoint_ft.Max() + xAxisPw.MajorStep,
                        MajorGridlineStyle = LineStyle.Solid,
                        MinorGridlineStyle = LineStyle.Dot
                    };
                    PowerPlotModel.Axes.Add(xAxisPw);
                }
                else
                {
                    xAxisPw.Title = IsEn ? "Air volume (m3/h)" : "Lưu lượng (m3/h)";
                    xAxisPw.Minimum = 0;
                    xAxisPw.Maximum = fitting.FlowPoint_ft.Max() + xAxisPw.MajorStep;
                }

                // Cập nhật trục Y
                var yAxisPw = PowerPlotModel.Axes.FirstOrDefault(a => a.Position == AxisPosition.Left) as LinearAxis;
                if (yAxisPw == null)
                {
                    yAxisPw = new LinearAxis
                    {
                        Position = AxisPosition.Left,
                        Title = IsEn ? "Power (kW)" : "Công suất (kW)",
                        Minimum = 0,
                        Maximum = fitting.PrtPoint_ft.Max() + yAxisPw.MajorStep,
                        MajorGridlineStyle = LineStyle.Solid,
                        MinorGridlineStyle = LineStyle.Dot
                    };
                    PowerPlotModel.Axes.Add(yAxisPw);
                }
                else
                {
                    yAxisPw.Title = IsEn ? "Power (kW)" : "Công suất (kW)";
                    yAxisPw.Minimum = 0;
                    yAxisPw.Maximum = fitting.PrtPoint_ft.Max() + yAxisPw.MajorStep;
                }

                ScatterSeries pwSeries = new ScatterSeries
                {
                    MarkerType = MarkerType.Circle,
                    MarkerFill = OxyColors.Red,
                    MarkerStroke = OxyColors.Red,
                    MarkerSize = 4,
                    Title = "", 
                };

                LineSeries pwLine = new LineSeries
                {
                    Color = OxyColors.Red,
                    StrokeThickness = 2,
                    Title = "",
                    LineLegendPosition = LineLegendPosition.None
                };

                if (!IsTorque)
                {
                    // Dữ liệu thường: dùng PrPoint_ft
                    for (int i = 0; i < fitting.FlowPoint_ft.Length; i++)
                    {
                        pwSeries.Points.Add(new ScatterPoint(fitting.Ope_FlowPoint[i], fitting.Ope_PrPoint[i]));
                        pwLine.Points.Add(new DataPoint(fitting.FlowPoint_ft[i], fitting.PrPoint_ft[i]));
                    }
                    
                }
                else
                {
                    // Dữ liệu momen xoắn: dùng PrtPoint_ft
                    for (int i = 0; i < fitting.FlowPoint_ft.Length; i++)
                    {
                        pwSeries.Points.Add(new ScatterPoint(fitting.Ope_FlowPoint[i], fitting.PrtPoint[i]));
                        pwLine.Points.Add(new DataPoint(fitting.FlowPoint_ft[i], fitting.PrtPoint_ft[i]));
                    }
                }

                //PowerPlotModel.Series.Add(pwSeries);
                PowerPlotModel.Series.Add(pwLine);

                // Lấy giá trị S của dải hiện tại
                var measure = rangeMeasure != null ? rangeMeasure.S : 0;
                float sValue = _camBien.PhanHoiTanSoMin + (_camBien.PhanHoiTanSoMax - _camBien.PhanHoiTanSoMin) * measure / 100;
                // Tính điểm giữa của line
                int midIndex = fitting.FlowPoint_ft.Length / 2;
                double midX = fitting.FlowPoint_ft[midIndex];
                double midY = !IsTorque ? fitting.PrPoint_ft[midIndex] : fitting.PrtPoint_ft[midIndex];
                double offsetY = (yAxisPw.ActualMaximum - yAxisPw.ActualMinimum) * 0.05;
                double textY = midY + offsetY;
                var annotation = new OxyPlot.Annotations.TextAnnotation
                {
                    Text = $"Tần số {sValue} Hz",
                    TextPosition = new OxyPlot.DataPoint(midX, textY),
                    TextColor = OxyColors.Black,
                    Stroke = OxyColors.White,
                    FontSize = 11,
                    FontWeight = OxyPlot.FontWeights.Bold,
                    Background = OxyColors.White,
                    Padding = new OxyThickness(2),
                    TextHorizontalAlignment = OxyPlot.HorizontalAlignment.Center,
                    TextVerticalAlignment = OxyPlot.VerticalAlignment.Middle
                };

                //PowerPlotModel.Annotations.Add(annotation);


                #endregion

                #region efficiency Plot
                // Khởi tạo tiêu đề và nhãn trục
                EfficiencyPlotModel.Title = IsEn ? "Air volume - Pressure - Efficiency curve" : "Đặc tuyến Áp suất - Hiệu suất - Lưu lượng";
                EfficiencyPlotModel.Background = OxyColors.White;

                // Trục X
                var xAxisEff = EfficiencyPlotModel.Axes.FirstOrDefault(a => a.Position == AxisPosition.Bottom) as LinearAxis;
                if (xAxisEff == null)
                {
                    xAxisEff = new LinearAxis
                    {
                        Position = AxisPosition.Bottom,
                        Title = IsEn ? "Air volume (m3/h)" : "Lưu lượng (m3/h)",
                        Minimum = 0,
                        Maximum = fitting.FlowPoint_ft.Max() + xAxisEff.MajorStep,
                        MajorGridlineStyle = LineStyle.Solid,
                        MinorGridlineStyle = LineStyle.Dot
                    };
                    EfficiencyPlotModel.Axes.Add(xAxisEff);
                }
                else
                {
                    xAxisEff.Title = IsEn ? "Air volume (m3/h)" : "Lưu lượng (m3/h)";
                    xAxisEff.Minimum = 0;
                    xAxisEff.Maximum = fitting.FlowPoint_ft.Max() + xAxisEff.MajorStep;
                }

                // Trục Y1 (Pressure)
                var yAxisEff = EfficiencyPlotModel.Axes.FirstOrDefault(a => a.Position == AxisPosition.Left) as LinearAxis;
                if (yAxisEff == null)
                {
                    yAxisEff = new LinearAxis
                    {
                        Position = AxisPosition.Left,
                        Title = IsEn ? "Pressure (Pa)" : "Áp suất (Pa)",
                        //Minimum = 0,
                        Minimum = Math.Min(fitting.PsPoint_ft.Min(), fitting.PtPoint_ft.Min()),
                        Maximum = Math.Max(fitting.PsPoint_ft.Max(), fitting.PtPoint_ft.Max()) + yAxisEff.MajorStep,
                        MajorGridlineStyle = LineStyle.Solid,
                        MinorGridlineStyle = LineStyle.Dot,
                        Key = "PressureAxis"
                    };
                    EfficiencyPlotModel.Axes.Add(yAxisEff);
                }
                else
                {
                    yAxisEff.Title = IsEn ? "Pressure (Pa)" : "Áp suất (Pa)";
                    //yAxisEff.Minimum = 0;
                    yAxisEff.Minimum = Math.Min(fitting.PsPoint_ft.Min(), fitting.PtPoint_ft.Min());
                    yAxisEff.Maximum = Math.Max(fitting.PsPoint_ft.Max(), fitting.PtPoint_ft.Max()) + yAxisEff.MajorStep;
                }

                // Trục Y2 (Efficiency)
                var y2Axis = EfficiencyPlotModel.Axes.FirstOrDefault(a => a.Position == AxisPosition.Right) as LinearAxis;
                if (y2Axis == null)
                {
                    y2Axis = new LinearAxis
                    {
                        Position = AxisPosition.Right,
                        Title = IsEn ? "Efficiency (%)" : "Hiệu suất (%)",
                        Minimum = 0,
                        Maximum= 100,
                        MajorGridlineStyle = LineStyle.None,
                        MinorGridlineStyle = LineStyle.None,
                        TextColor = OxyColors.Blue,
                        TitleColor = OxyColors.Blue,
                        Key = "EfficiencyAxis"
                    };
                    EfficiencyPlotModel.Axes.Add(y2Axis);
                }
                else
                {
                    y2Axis.Title = IsEn ? "Efficiency (%)" : "Hiệu suất (%)";
                    y2Axis.Minimum = 0;
                    y2Axis.Maximum = 100;
                }

                var staticPressureScatter = new ScatterSeries
                {
                    MarkerType = MarkerType.Circle,
                    MarkerFill = OxyColors.Red,
                    MarkerStroke = OxyColors.Red,
                    MarkerSize = 4,
                    Title = "",
                    YAxisKey = "PressureAxis"
                };
                var totalPressureScatter = new ScatterSeries
                {
                    MarkerType = MarkerType.Diamond,
                    MarkerFill = OxyColors.Blue,
                    MarkerStroke = OxyColors.Blue,
                    MarkerSize = 4,
                    Title = "",
                    YAxisKey = "PressureAxis"
                };

                var staticEfficiencyScatter = new ScatterSeries
                {
                    MarkerType = MarkerType.Square,
                    MarkerFill = OxyColors.Black,
                    MarkerStroke = OxyColors.Black,
                    MarkerSize = 4,
                    Title = "",
                    YAxisKey = "EfficiencyAxis"
                };
                var totalEfficiencyScatter = new ScatterSeries
                {
                    MarkerType = MarkerType.Triangle,
                    MarkerFill = OxyColors.DarkGreen,
                    MarkerStroke = OxyColors.DarkGreen,
                    MarkerSize = 4,
                    Title = "",
                    YAxisKey = "EfficiencyAxis"
                };

                var staticPressureLine = new LineSeries
                {
                    Color = OxyColors.Red,
                    StrokeThickness = 2,
                    Title = "",
                    YAxisKey = "PressureAxis",
                    LineLegendPosition = LineLegendPosition.None
                };
                var totalPressureLine = new LineSeries
                {
                    Color = OxyColors.Blue,
                    StrokeThickness = 2,
                    Title = "",
                    YAxisKey = "PressureAxis",
                    LineStyle = LineStyle.Dash,
                    LineLegendPosition = LineLegendPosition.None
                };

                var staticEfficiencyLine = new LineSeries
                {
                    Color = OxyColors.Black,
                    StrokeThickness = 2,
                    Title = "",
                    YAxisKey = "EfficiencyAxis",
                    LineStyle = LineStyle.DashDot,
                    LineLegendPosition = LineLegendPosition.None
                };
                var totalEfficiencyLine = new LineSeries
                {
                    Color = OxyColors.DarkGreen,
                    StrokeThickness = 2,
                    Title = "",
                    YAxisKey = "EfficiencyAxis",
                    LineStyle = LineStyle.Dot,
                    LineLegendPosition = LineLegendPosition.None
                };


                for (int i = 0; i < fitting.FlowPoint_ft.Length; i++)
                {
                    staticPressureScatter.Points.Add(new ScatterPoint(fitting.Ope_FlowPoint[i], fitting.Ope_PsPoint[i])); // Ps
                    totalPressureScatter.Points.Add(new ScatterPoint(fitting.Ope_FlowPoint[i], fitting.Ope_PtPoint[i])); // Pt

                    staticPressureLine.Points.Add(new DataPoint(fitting.FlowPoint_ft[i], fitting.PsPoint_ft[i])); // Ps Fitting
                    totalPressureLine.Points.Add(new DataPoint(fitting.FlowPoint_ft[i], fitting.PtPoint_ft[i])); // Pt Fitting
                   
                    
                    if (!IsTorque)
                    {
                        staticEfficiencyScatter.Points.Add(new ScatterPoint(fitting.Ope_FlowPoint[i], fitting.Ope_EsPoint[i])); // Es
                        totalEfficiencyScatter.Points.Add(new ScatterPoint(fitting.Ope_FlowPoint[i], fitting.Ope_EtPoint[i])); // Et

                        staticEfficiencyLine.Points.Add(new DataPoint(fitting.FlowPoint_ft[i], fitting.EsPoint_ft[i])); // Es Fitting
                        totalEfficiencyLine.Points.Add(new DataPoint(fitting.FlowPoint_ft[i], fitting.EtPoint_ft[i])); // Et Fitting
                    }
                    else
                    {
                        staticEfficiencyScatter.Points.Add(new ScatterPoint(fitting.Ope_FlowPoint[i], fitting.Ope_EstPoint[i])); // Est
                        totalEfficiencyScatter.Points.Add(new ScatterPoint(fitting.Ope_FlowPoint[i], fitting.Ope_EttPoint[i])); // Ett

                        staticEfficiencyLine.Points.Add(new DataPoint(fitting.FlowPoint_ft[i], fitting.EstPoint_ft[i])); // Est Fitting
                        totalEfficiencyLine.Points.Add(new DataPoint(fitting.FlowPoint_ft[i], fitting.EttPoint_ft[i])); // Ett Fitting
                    }
                }

                // Thêm series mới vào model
                //EfficiencyPlotModel.Series.Add(staticPressureScatter);
                //EfficiencyPlotModel.Series.Add(totalPressureScatter);
                //EfficiencyPlotModel.Series.Add(staticEfficiencyScatter);
                //EfficiencyPlotModel.Series.Add(totalEfficiencyScatter);


                EfficiencyPlotModel.Series.Add(staticPressureLine);
                EfficiencyPlotModel.Series.Add(totalPressureLine);
                EfficiencyPlotModel.Series.Add(staticEfficiencyLine);
                EfficiencyPlotModel.Series.Add(totalEfficiencyLine);
                EfficiencyPlotModel.InvalidatePlot(true);
                #endregion

                PowerPlotModel.InvalidatePlot(true);
                EfficiencyPlotModel.InvalidatePlot(true);
            });
        }


        #region xuất kết quả và báo cáo
        private bool CanExportCommand(object obj)
        {
            return true;
        }


        private void ExecuteExportResultCommand(object obj)
        {

        }

        private async void ExecuteExportReportCommand(object obj)
        {
            // Implement export report logic here
        }

        #endregion
    }
}
