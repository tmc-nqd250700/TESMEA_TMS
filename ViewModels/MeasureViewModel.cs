using MaterialDesignThemes.Wpf;
using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Legends;
using OxyPlot.Series;
using System.Collections.ObjectModel;
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
            }
        }

        private bool _isMeasuring;
        public bool CanMeasure => !_isMeasuring && _externalAppService.IsConnectedToSimatic;

        private bool _isConnected;
        private bool _isConnectedRow1;
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
        private readonly IFileService _fileService;

        // Commands
        public ICommand ConnectCommand { get; }
        public ICommand ConnectCommand2 { get; }
        public ICommand StartCommand { get; }
        public ICommand StopCommand { get; }
        public ICommand ResetCommand { get; }
        public ICommand TrendCommand { get; }

        public ICommand ExportResultCommand { get; }
        public ICommand ExportReportCommand { get; }
        public ICommand BackCommand { get; }
        public ICommand FinishCommand { get; }
        public MeasureViewModel(IExternalAppService externalAppService, IParameterService parameterService, IFileService fileService)
        {
            _externalAppService = externalAppService;
            _parameterService = parameterService;
            _fileService = fileService;

            MeasureRows = new ObservableCollection<Measure>();
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
            // event invoke khi hoàn thành toàn bộ đo kiểm
            _externalAppService.OnSimaticExchangeCompleted += OnExchangeCompletedHandler;

            ReportTemplates = new ObservableCollection<ComboBoxInfo>();
            ReportTemplates.Add(new ComboBoxInfo("DESIGN", IsEn ? "Design condition" : "Điều kiện thiết kế"));
            ReportTemplates.Add(new ComboBoxInfo("NORMALIZED", IsEn ? "Normalized condition" : "Điều kiện tiêu chuẩn"));
            ReportTemplates.Add(new ComboBoxInfo("OPERATION", IsEn ? "Operation condition" : "Điều kiện hoạt động"));
            ReportTemplates.Add(new ComboBoxInfo("FULL", IsEn ? "Full" : "Tất cả"));
            SelectedReportTemplate = new ComboBoxInfo();
            SelectedReportTemplate.Value = ReportTemplates[0].Value;

            ConnectCommand = new ViewModelCommand(CanConnect, ExecuteConnectCommand);
            ConnectCommand2 = new ViewModelCommand(CanConnect2, ExecuteConnectCommand2);

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
                Key = "PowerXAxis",
                Minimum = 0,
                IsPanEnabled = false,
                IsZoomEnabled = false,
                MajorGridlineStyle = LineStyle.Solid,
                MinorGridlineStyle = LineStyle.Dot,
            });

            // Thêm trục Y
            PowerPlotModel.Axes.Add(new LinearAxis
            {
                Position = AxisPosition.Left,
                Title = IsEn ? "Power (kW)" : "Công suất (kW)",
                Key = "PowerYAxis",
                Minimum = 0,
                IsPanEnabled = false,
                IsZoomEnabled = false,
                MajorGridlineStyle = LineStyle.Solid,
                MinorGridlineStyle = LineStyle.Dot,
            });

           
            var scatterSeries = new ScatterSeries
            {
                Title = IsEn ? "Power" : "Công suất",
                MarkerType = MarkerType.Circle,
                MarkerFill = OxyColors.Red,
                MarkerStroke = OxyColors.Red,
                MarkerStrokeThickness = 1,
                MarkerSize = 2
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
                MinorGridlineStyle = LineStyle.Dot,
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
                Key = "PressureAxis",
            });

            // Trục Y2 (Efficiency)
            EfficiencyPlotModel.Axes.Add(new LinearAxis
            {
                Position = AxisPosition.Right,
                Title = IsEn ? "Efficiency (%)" : "Hiệu suất (%)",
                Minimum = 0,
                Maximum = 100,
                MajorStep = 10,
                IsPanEnabled = false,
                IsZoomEnabled = false,
                MajorGridlineStyle = LineStyle.None,
                MinorGridlineStyle = LineStyle.None,
                TextColor = OxyColors.Blue,
                TitleColor = OxyColors.Blue,
                Key = "EfficiencyAxis",
            });

            // Thêm các series
            var staticPressureSeries = new ScatterSeries
            {
                Title = IsEn ? "Static Pressure" : "Áp suất tĩnh",
                MarkerType = MarkerType.Circle,
                MarkerFill = OxyColors.Red,
                MarkerStroke = OxyColors.Red,
                MarkerStrokeThickness = 1,
                MarkerSize = 2,
                YAxisKey = "PressureAxis",
            };

            var totalPressureSeries = new ScatterSeries
            {
                Title = IsEn ? "Total Pressure" : "Áp suất tổng",
                MarkerType = MarkerType.Diamond,
                MarkerFill = OxyColors.Blue,
                MarkerStroke = OxyColors.Blue,
                MarkerStrokeThickness = 1,
                MarkerSize = 2,
                YAxisKey = "PressureAxis"
            };

            var staticEfficiencySeries = new ScatterSeries
            {
                Title = IsEn ? "Static Efficiency" : "Hiệu suất tĩnh",
                MarkerType = MarkerType.Square,
                MarkerFill = OxyColors.Black,
                MarkerStroke = OxyColors.Black,
                MarkerStrokeThickness = 1,
                MarkerSize = 2,
                YAxisKey = "EfficiencyAxis"
            };

            var totalEfficiencySeries = new ScatterSeries
            {
                Title = IsEn ? "Total Efficiency" : "Hiệu suất tổng",
                MarkerType = MarkerType.Triangle,
                MarkerFill = OxyColors.DarkGreen,
                MarkerStroke = OxyColors.DarkGreen,
                MarkerStrokeThickness = 1,
                MarkerSize = 2,
                YAxisKey = "EfficiencyAxis"
            };

            EfficiencyPlotModel.Series.Add(staticPressureSeries);
            EfficiencyPlotModel.Series.Add(totalPressureSeries);
            EfficiencyPlotModel.Series.Add(staticEfficiencySeries);
            EfficiencyPlotModel.Series.Add(totalEfficiencySeries);
        }

        private async void ExecuteFinishCommand(object obj)
        {
            var isFinish = MessageBoxHelper.ShowQuestion("Bạn có chắc chắn muốn hoàn thành đo kiểm và quay lại trang chính không?");
            if (isFinish)
            {
                await _externalAppService.StopExchangeAsync();
                var locator = ((App)System.Windows.Application.Current).Resources["Locator"] as ViewModelLocator;
                if (locator != null)
                {
                    ClearData();
                    locator.ProjectViewModel.ClearData();
                    locator.ProjectViewModel.ClearData();
                    locator.MainViewModel.CurrentView = new TESMEA_TMS.Views.ProjectView();
                }

                var exchangeFolder = UserSetting.TOMFAN_folder;
                if (Directory.Exists(exchangeFolder))
                {
                    var files = Directory.GetFiles(exchangeFolder, "*.csv");
                    if (files.Length == 0)
                    {
                        using (var writer2 = new StreamWriter(Path.Combine(exchangeFolder, "1_T_OUT.csv")))
                        {
                        }

                        using (var writer2 = new StreamWriter(Path.Combine(exchangeFolder, "2_S_IN.csv")))
                        {
                        }
                    }
                    else
                    {
                        foreach (var file in Directory.GetFiles(exchangeFolder))
                        {
                            try { File.WriteAllText(file, string.Empty); } catch { }
                        }

                    }

                    // trend folder
                    if (!Directory.Exists(Path.Combine(exchangeFolder, "Trend")))
                    {
                        Directory.CreateDirectory(Path.Combine(exchangeFolder, "Trend"));
                    }
                    else
                    {
                        // delete all files in trend folder
                        foreach (var file in Directory.GetFiles(Path.Combine(exchangeFolder, "Trend")))
                        {
                            try { File.Delete(file); } catch { }
                        }
                    }
                }
            }
        }


        private void ClearData()
        {
            MeasureRows.Clear();
            ThongTinDuAn = new ThongTinDuAn();
        }
        private bool CanConnect(object obj) => !_isConnected && !_isConnectedRow1;
        private bool CanConnect2(object obj) => !_isConnected && _isConnectedRow1;
        private bool CanStart(object obj) => _isConnected && _isConnectedRow1 && !_isMeasuring && !_isCompleted;
        //private bool CanStop(object obj) => _isMeasuring;
        private bool CanStop(object obj) => true;
        private bool CanReset(object obj) => _isCompleted && !_isMeasuring;
        private bool CanTrend(object obj) => SelectedMeasure != null && (TrendLineDialog == null || !TrendLineDialog.IsLoaded);

        private async void ExecuteConnectCommand2(object obj)
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
                if (!MeasureRows.Where(x => x.k == 2).Any())
                {
                    throw new BusinessException("Không có dữ liệu cho dòng kết nối thứ 2");
                }
                SelectedMeasure = MeasureRows.Where(x => x.k == 2).First();
                var scenario = await _parameterService.GetScenarioAsync(ThongTinDuAn.ThamSo.KichBan);
                if (scenario == null)
                {
                    throw new BusinessException("Không tìm thấy kịch bản đo kiểm");
                }

                if (await _externalAppService.ConnectExchangeAsync(SelectedMeasure, scenario.StandardDeviation))
                {
                    _isConnected = true;
                    _isCompleted = false;
                    if (DialogHost.IsDialogOpen("MainDialogHost"))
                        DialogHost.Close("MainDialogHost");
                    OnPropertyChanged(nameof(SimaticStatus));
                }

            }
            catch (Exception ex)
            {
                if (DialogHost.IsDialogOpen("MainDialogHost"))
                    DialogHost.Close("MainDialogHost");
                throw;
            }
        }

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
                (var bienTan, _camBien, var ongGio) = await _parameterService.GetLibraryByIdAsync(Guid.Parse(ThongTinDuAn.ThamSo.ThongSo));
                if (bienTan == null || _camBien == null || ongGio == null)
                {
                    throw new BusinessException("Không tìm thấy thông số kiểm thử");
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
                _isConnectedRow1 = true;
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
                _isCompleted = true;
                MessageBoxHelper.ShowError(ex.Message);
            }
        }

        private async void ExecuteStopCommand(object obj)
        {
            var splashViewModel = new ProgressSplashViewModel
            {
                Message = "Dừng khẩn cấp kết nối với Simatic...",
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
            DataProcess.Initialize(MeasureRows.Count);
            MeasureResponses.Clear();
            OnPropertyChanged(nameof(MeasureResponses));
            ParameterShow = new ParameterShow();
            OnPropertyChanged(nameof(ParameterShow));
            var exchangeFolder = UserSetting.TOMFAN_folder;
            if (Directory.Exists(exchangeFolder))
            {
                var files = Directory.GetFiles(exchangeFolder, "*.csv");
                if (files.Length == 0)
                {
                    using (var writer2 = new StreamWriter(Path.Combine(exchangeFolder, "1_T_OUT.csv")))
                    {
                    }

                    using (var writer2 = new StreamWriter(Path.Combine(exchangeFolder, "2_S_IN.csv")))
                    {
                    }
                }
                else
                {
                    foreach (var file in Directory.GetFiles(exchangeFolder))
                    {
                        try { File.WriteAllText(file, string.Empty); } catch { }
                    }

                }

                // trend folder
                if (!Directory.Exists(Path.Combine(exchangeFolder, "Trend")))
                {
                    Directory.CreateDirectory(Path.Combine(exchangeFolder, "Trend"));
                }
                else
                {
                    // delete all files in trend folder
                    foreach (var file in Directory.GetFiles(Path.Combine(exchangeFolder, "Trend")))
                    {
                        try { File.Delete(file); } catch { }
                    }
                }
            }
            ClearPlots();
            _isMeasuring = false;
            _isConnectedRow1 = false;
            _isCompleted = false;
            _isConnected = false;
            OnPropertyChanged(nameof(CanMeasure));
            OnPropertyChanged(nameof(SimaticStatus));
        }

        private void ClearPlots()
        {
            if (PowerPlotModel != null)
            {
                PowerPlotModel.Series.Clear();
                InitializePowerPlotModel();
                PowerPlotModel.InvalidatePlot(true);
            }

            if (EfficiencyPlotModel != null)
            {
                EfficiencyPlotModel.Series.Clear();
                InitializeEfficiencyPlotModel();
                EfficiencyPlotModel.InvalidatePlot(true);
            }
        }

        private void ExecuteTrendCommand(object obj)
        {
            if (SelectedMeasure == null)
                return;

            if (TrendLineDialog == null || !TrendLineDialog.IsLoaded)
            {
                TrendLineDialog = new Views.TrendLineDialog(SelectedMeasure.k, _camBien);
                TrendLineDialog.Closed += (s, e) => TrendLineDialog = null;
                TrendLineDialog.Show();
            }
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

        //private ScatterSeries _powerScatterLiveSeries;
        //private ScatterSeries _psScatterLiveSeries;
        //private ScatterSeries _ptScatterLiveSeries;
        //private ScatterSeries _seScatterLiveSeries;
        //private ScatterSeries _teScatterLiveSeries;


        private void OnMeasurePointCompletedHandler(MeasureResponse response, ParameterShow paramShow)
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                ParameterShow = paramShow;
                OnPropertyChanged(nameof(ParameterShow));

                // Thêm vào danh sách kết quả
                MeasureResponses.Add(response);

               // Power Plot Update
                var powerSeries = PowerPlotModel.Series.OfType<ScatterSeries>().FirstOrDefault();
                if (powerSeries != null)
                {
                    powerSeries.Points.Add(new ScatterPoint(response.Airflow, response.Power));

                    //var xAxis = PowerPlotModel.Axes.FirstOrDefault(a => a.Position == AxisPosition.Bottom) as LinearAxis;
                    //var yAxis = PowerPlotModel.Axes.FirstOrDefault(a => a.Position == AxisPosition.Left) as LinearAxis;

                    //// Cập nhật trục X
                    //if (xAxis != null)
                    //{
                    //    var currentMax = xAxis.Maximum;
                    //    if (response.Airflow > currentMax)
                    //    {
                    //        xAxis.Maximum = Common.RoundUpToNearest(response.Airflow * 1.1f);
                    //        _externalAppService.WriteTomfanLog($"PowerPlot: Cập nhật X axis từ {currentMax} lên {xAxis.Maximum}");
                    //    }
                    //}

                    //// Cập nhật trục Y
                    //if (yAxis != null)
                    //{
                    //    var currentMax = yAxis.Maximum;
                    //    if (response.Power > currentMax)
                    //    {
                    //        yAxis.Maximum = Common.RoundUpToNearest(response.Power * 1.1f);
                    //        _externalAppService.WriteTomfanLog($"PowerPlot: Cập nhật Y axis từ {currentMax} lên {yAxis.Maximum}");
                    //    }
                    //}

                    PowerPlotModel.InvalidatePlot(true);
                }

                // Efficiency Plot Update
                if (EfficiencyPlotModel.Series.Count >= 4)
                {
                    var psPoint = EfficiencyPlotModel.Series[0] as ScatterSeries;
                    var ptPoint = EfficiencyPlotModel.Series[1] as ScatterSeries;
                    var sePoint = EfficiencyPlotModel.Series[2] as ScatterSeries;
                    var tePoint = EfficiencyPlotModel.Series[3] as ScatterSeries;

                    psPoint?.Points.Add(new ScatterPoint(response.Airflow, response.Ps));
                    ptPoint?.Points.Add(new ScatterPoint(response.Airflow, response.Pt));
                    sePoint?.Points.Add(new ScatterPoint(response.Airflow, response.SEff));
                    tePoint?.Points.Add(new ScatterPoint(response.Airflow, response.TEff));

                    // Cập nhật trục cho Efficiency Plot
                    //var xAxisEff = EfficiencyPlotModel.Axes.FirstOrDefault(a => a.Position == AxisPosition.Bottom) as LinearAxis;
                    //var yPressAxis = EfficiencyPlotModel.Axes.FirstOrDefault(a => a.Key == "PressureAxis") as LinearAxis;

                    //// Cập nhật trục X
                    //if (xAxisEff != null)
                    //{
                    //    var currentMax = xAxisEff.Maximum;
                    //    if (response.Airflow > currentMax)
                    //    {
                    //        xAxisEff.Maximum = Common.RoundUpToNearest(response.Airflow * 1.1f);
                    //    }
                    //}

                    //// Cập nhật trục Y (Pressure)
                    //if (yPressAxis != null)
                    //{
                    //    float maxPress = Math.Max(response.Ps, response.Pt);
                    //    var currentMax = yPressAxis.Maximum;
                    //    if (maxPress > currentMax)
                    //    {
                    //        yPressAxis.Maximum = Common.RoundUpToNearest(maxPress * 1.1f);
                    //    }
                    //}
                    EfficiencyPlotModel.InvalidatePlot(true);
                }
            });
        }

        private void OnMeasureRangeCompletedHandler(MeasureFittingFC fitting, Measure rangeMeasure)
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                _externalAppService.WriteTomfanLog("==== LOG FITTING ====");

                // Power Plot - Thêm Line mới
                // Tạo LineSeries cho dải vừa hoàn thành
                var pwLine = new LineSeries
                {
                    Color = OxyColors.Red,
                    StrokeThickness = 2,
                    LineLegendPosition = LineLegendPosition.None
                };

                for (int i = 0; i < fitting.FlowPoint_ft.Length; i++)
                {
                    double x = fitting.Ope_FlowPoint[i];
                    double x2 = fitting.FlowPoint_ft[i];
                    double y = fitting.Ope_PrPoint[i];
                    double y2 = fitting.PrPoint_ft[i];
                    double y3 = fitting.PrtPoint[i];
                    double y4 = fitting.PrtPoint_ft[i];
                    // fitting co torque
                    //pwLine.Points.Add(new DataPoint(x2, y4));
                    // fitting ps
                    pwLine.Points.Add(new DataPoint(x2, y2));
                }

                PowerPlotModel.Series.Add(pwLine);

              
                var list = new List<DataPoint>();
                var list2 = new List<DataPoint>();
                var list3 = new List<DataPoint>();
                var list4 = new List<DataPoint>();
                var list5 = new List<DataPoint>();
                var list6 = new List<DataPoint>();
                var list7 = new List<DataPoint>();
                var list8 = new List<DataPoint>();
                // T
                var list9 = new List<DataPoint>();
                var list10 = new List<DataPoint>();
                var list11 = new List<DataPoint>();
                var list12 = new List<DataPoint>();
                for (int i = 0; i < fitting.FlowPoint_ft.Length; i++)
                {
                    double x = fitting.Ope_FlowPoint[i];
                    double x2 = fitting.FlowPoint_ft[i];
                    double y = fitting.Ope_PsPoint[i];
                    double y2 = fitting.Ope_PtPoint[i];
                    double y3 = fitting.Ope_EsPoint[i];
                    double y4 = fitting.Ope_EtPoint[i];
                    double y5 = fitting.PsPoint_ft[i];
                    double y6 = fitting.PtPoint_ft[i];
                    double y7 = fitting.EsPoint_ft[i];
                    double y8 = fitting.EtPoint_ft[i];

                    // T
                    double y9 = fitting.Ope_EstPoint[i];
                    double y10 = fitting.Ope_EttPoint[i];
                    double y11 = fitting.EstPoint_ft[i];
                    double y12 = fitting.EttPoint_ft[i];

                    list.Add(new DataPoint(x, y));
                    list2.Add(new DataPoint(x, y2));
                    list3.Add(new DataPoint(x, y3));
                    list4.Add(new DataPoint(x, y4));
                    list5.Add(new DataPoint(x2, y5));
                    list6.Add(new DataPoint(x2, y6));
                    list7.Add(new DataPoint(x2, y7));
                    list8.Add(new DataPoint(x2, y8));
                    // T
                    list9.Add(new DataPoint(x, y9));
                    list10.Add(new DataPoint(x, y10));
                    list11.Add(new DataPoint(x2, y11));
                    list12.Add(new DataPoint(x2, y12));
                }
                List<ScatterPoint> ToScatterPoints(List<DataPoint> dataPoints)
                {
                    return dataPoints.Select(dp => new ScatterPoint(dp.X, dp.Y)).ToList();
                }
                if (EfficiencyPlotModel.Series.Count >= 4)
                {
                    var scatter1 = EfficiencyPlotModel.Series[0] as ScatterSeries;
                    var scatter2 = EfficiencyPlotModel.Series[1] as ScatterSeries;
                    var scatter3 = EfficiencyPlotModel.Series[2] as ScatterSeries;
                    var scatter4 = EfficiencyPlotModel.Series[3] as ScatterSeries;

                    scatter1?.Points.Clear();
                    scatter2?.Points.Clear();
                    scatter3?.Points.Clear();
                    scatter4?.Points.Clear();

                    if (scatter1 != null)
                    {
                        scatter1.Points.AddRange(ToScatterPoints(list));
                    }
                    if (scatter2 != null)
                    {
                        scatter2.Points.AddRange(ToScatterPoints(list2));
                    }
                    if (scatter3 != null && scatter4 != null)
                    {
                        if (!IsTorque)
                        {
                            scatter3.Points.AddRange(ToScatterPoints(list3));
                            scatter4.Points.AddRange(ToScatterPoints(list4));
                        }
                        else
                        {
                            scatter3.Points.AddRange(ToScatterPoints(list9));
                            scatter4.Points.AddRange(ToScatterPoints(list10));
                        }
                    }
                }


                var staticPressureLine = new LineSeries { Color = OxyColors.Red, StrokeThickness = 2, YAxisKey = "PressureAxis" };
                var totalPressureLine = new LineSeries { Color = OxyColors.Blue, StrokeThickness = 2, YAxisKey = "PressureAxis", LineStyle = LineStyle.Dash };
                var staticEfficiencyLine = new LineSeries { Color = OxyColors.Black, StrokeThickness = 2, YAxisKey = "EfficiencyAxis", LineStyle = LineStyle.DashDot };
                var totalEfficiencyLine = new LineSeries { Color = OxyColors.DarkGreen, StrokeThickness = 2, YAxisKey = "EfficiencyAxis", LineStyle = LineStyle.Dot };

                staticPressureLine.Points.AddRange(list5);
                totalPressureLine.Points.AddRange(list6);


                if (!IsTorque)
                {
                    staticEfficiencyLine.Points.AddRange(list7);
                    totalEfficiencyLine.Points.AddRange(list8);
                }
                else
                {
                    staticEfficiencyLine.Points.AddRange(list11);
                    totalEfficiencyLine.Points.AddRange(list12);
                }

                EfficiencyPlotModel.Series.Add(staticPressureLine);
                EfficiencyPlotModel.Series.Add(totalPressureLine);
                EfficiencyPlotModel.Series.Add(staticEfficiencyLine);
                EfficiencyPlotModel.Series.Add(totalEfficiencyLine);

                PowerPlotModel.InvalidatePlot(true);
                EfficiencyPlotModel.InvalidatePlot(true);

                _externalAppService.WriteTomfanLog("==== END LOG FITTING ====");
            });
        }



        private void OnExchangeCompletedHandler(List<Measure> measures)
        {
            _isMeasuring = false;
            _isCompleted = true;
        }


        #region xuất kết quả và báo cáo
        private bool CanExportCommand(object obj)
        {
            return true;
        }

        private async Task<ThongSoDauVao> ConvertData()
        {
            try
            {
                // convert measure sang thông số đầu vào cho phần tính toán và xuất báo cáo
                ThongSoDauVao tsdv = new ThongSoDauVao();
                ThongSoDuongOngGio tsOngGio = new ThongSoDuongOngGio();
                ThongSoCoBanCuaQuat tsQuat = new ThongSoCoBanCuaQuat();
                List<Measure> dstsDoKiem = new List<Measure>();

                (var bienTan, _camBien, var ongGio) = await _parameterService.GetLibraryByIdAsync(Guid.Parse(ThongTinDuAn.ThamSo.ThongSo));
                if (bienTan == null || _camBien == null || ongGio == null)
                {
                    throw new BusinessException("Không tìm thấy thông số kiểm thử");
                }

                // thông số đường ống gió
                tsOngGio.DuongKinhOngD5 = ongGio.DuongKinhOngD5;
                tsOngGio.ChieuDaiOngGioTonThatL = ongGio.ChieuDaiConQuat;
                tsOngGio.DuongKinhOngGioD3 = ongGio.DuongKinhMiengQuat;
                tsOngGio.TietDienOngD5 = 3.14 * Math.Pow(tsOngGio.DuongKinhOngD5 / 1000, 2) / 4;
                tsOngGio.HeSoMaSatOngK = ongGio.HeSoMaSat;
                tsOngGio.TietDienOngGioD3 = 3.14 * Math.Pow(tsOngGio.DuongKinhOngGioD3 / 1000, 2) / 4;

                // thông số cơ bản của quạt
                var mauThuNghiem = ThongTinDuAn.ThongTinMauThuNghiem;
                tsQuat.SoVongQuayCuaQuatNLT = mauThuNghiem.TocDoThietKeCuaQuat;
                tsQuat.CongSuatDongCo = mauThuNghiem.CongSuatDongCo;
                tsQuat.HeSoDongCo = mauThuNghiem.HeSoCongSuatDongCo;
                tsQuat.Tanso = mauThuNghiem.TanSoDongCoTheoThietKe;
                tsQuat.HieuSuatDongCo = mauThuNghiem.HieuSuatDongCo;
                tsQuat.DoNhotKhongKhi = 0; // chuwa cos
                tsQuat.ApSuatKhiQuyen = 110110; // chuw cos
                tsQuat.NhietDoLamViec = mauThuNghiem.NhietDoThietKeLamViec;
                if (!MeasureRows.Any())
                {
                    throw new BusinessException("Chưa có kết quả đo kiểm");
                }
                //int stt = 1;
                //foreach (var item in MeasureRows)
                //{
                //    Measure tsDokiem = new Measure();
                //    tsDokiem.KiemTraSo = stt;
                //    tsDokiem.NhietDoBauKho = item.NhietDoMoiTruong_sen;
                //    tsDokiem.DoAmTuongDoi = item.DoAm_sen;
                //    tsDokiem.SoVongQuayNTT = item.SoVongQuay_sen;
                //    tsDokiem.ChenhLechApSuat = item.ChenhLechApSuat_sen;
                //    tsDokiem.ApSuatTinh = item.ApSuatTinh_sen;
                //    tsDokiem.DongLamViec = item.DongDien_fb;
                //    tsDokiem.DienAp = item.DienAp_fb;
                //    tsDokiem.TanSo = item.TanSo_fb;
                //    dstsDoKiem.Add(tsDokiem);
                //}
                tsdv.ThongSoDuongOngGio = tsOngGio;
                tsdv.ThongSoCoBanCuaQuat = tsQuat;
                tsdv.DanhSachThongSoDoKiem = MeasureRows.ToList();
                return tsdv;
            }
            catch(BusinessException ex)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private async void ExecuteExportResultCommand(object obj)
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
                var splashViewModel = new ProgressSplashViewModel
                {
                    Message = "Đang tạo file kết quả đo kiểm, vui lòng chờ...",
                    IsIndeterminate = true
                };
                var splash = new Views.CustomControls.ProgressSplashContent { DataContext = splashViewModel };
                var dialogTask = DialogHost.Show(splash, "MainDialogHost");
                try
                {
                    await _fileService.ExportReportTestResult(
                       outputPath: sfd.FileName,
                       option: SelectedReportTemplate?.Value ?? "DESIGN",
                       tsdv: await ConvertData(),
                       project: ThongTinDuAn
                   );
                    if (DialogHost.IsDialogOpen("MainDialogHost"))
                        DialogHost.Close("MainDialogHost");

                }
                catch (Exception ex)
                {
                    if (DialogHost.IsDialogOpen("MainDialogHost"))
                        DialogHost.Close("MainDialogHost");
                    throw;
                }
                MessageBoxHelper.ShowSuccess(IsEn ? "Export report successfully" : "Báo cáo đã được xuất thành công");
            }
        }

        private async void ExecuteExportReportCommand(object obj)
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
                var splashViewModel = new ProgressSplashViewModel
                {
                    Message = "Đang tạo báo cáo đo kiểm, vui lòng chờ...",
                    IsIndeterminate = true
                };
                var splash = new Views.CustomControls.ProgressSplashContent { DataContext = splashViewModel };
                var dialogTask = DialogHost.Show(splash, "MainDialogHost");
                try
                {
                    await _fileService.ExportExcelTestResult(
                        outputPath: sfd.FileName,
                        option: SelectedReportTemplate?.Value ?? "DESIGN",
                        tsdv: await ConvertData(),
                        project: ThongTinDuAn
                    );

                    if (DialogHost.IsDialogOpen("MainDialogHost"))
                        DialogHost.Close("MainDialogHost");

                }
                catch (Exception ex)
                {
                    if (DialogHost.IsDialogOpen("MainDialogHost"))
                        DialogHost.Close("MainDialogHost");
                    throw;
                }
                MessageBoxHelper.ShowSuccess(IsEn ? "Export result successfully" : "Kết quả tính toán đã được xuất thành công");
            }
        }

        #endregion
    }
}
