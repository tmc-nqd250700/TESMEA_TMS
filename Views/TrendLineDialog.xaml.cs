using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Series;
using System.IO;
using System.Windows;
using TESMEA_TMS.DTOs;
using TESMEA_TMS.Configs;
using System.Diagnostics;
using System.Collections.ObjectModel;
using System.ComponentModel;
using TESMEA_TMS.Models.Entities;
using System.Globalization;

namespace TESMEA_TMS.Views
{
    public partial class TrendLineDialog : Window, INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        public PlotModel TrendPlotModel { get; set; }
        public ObservableCollection<ComboBoxInfo> PVTypes { get; set; }
        private string _selectedPV;
        public string SelectedPV
        {
            get => _selectedPV;
            set
            {
                if (_selectedPV != value)
                {
                    _selectedPV = value;
                    OnPropertyChanged(nameof(SelectedPV));
                    LoadTrendDataAndDraw(_currentK, _selectedPV);
                }
            }
        }
        private int _currentK;
        private CamBien _sensor;

        public TrendLineDialog(int k, CamBien sensor)
        {
            InitializeComponent();
            DataContext = this;
            _currentK = k;
            _sensor = sensor;
            TrendPlotModel = new PlotModel
            {
                Title = $"Trend line k = {k}",
                Background = OxyColors.White,
               
            };
            TrendPlotModel.Axes.Add(new LinearAxis
            {
                Position = AxisPosition.Bottom,
                Title = "Thời gian (millisecond)",
                Minimum = 0,
                MajorGridlineStyle = LineStyle.Solid,
                MinorGridlineStyle = LineStyle.Dot,
                IsPanEnabled = false,
                IsZoomEnabled = false,
            });
            TrendPlotModel.Axes.Add(new LinearAxis
            {
                Position = AxisPosition.Left,
                Title = "Giá trị",
                Minimum = 0,
                MajorGridlineStyle = LineStyle.Solid,
                MinorGridlineStyle = LineStyle.Dot,
                IsPanEnabled = false,
                IsZoomEnabled = false,
            });

            PVTypes = new ObservableCollection<ComboBoxInfo>
            {
                new ComboBoxInfo("NhietDoMoiTruong_sen", "Nhiệt độ môi trường"),
                new ComboBoxInfo("DoAm_sen", "Độ ẩm"),
                new ComboBoxInfo("ApSuatkhiQuyen_sen", "Áp suất khí quyển"),
                new ComboBoxInfo("ChenhLechApSuat_sen", "Chênh lệch áp suất"),
                new ComboBoxInfo("ApSuatTinh_sen", "Áp suất tĩnh"),
                new ComboBoxInfo("DoRung_sen", "Độ rung"),
                new ComboBoxInfo("DoOn_sen", "Độ ồn"),
                new ComboBoxInfo("SoVongQuay_sen", "Tốc độ quay"),
                new ComboBoxInfo("Momen_sen", "Mômen"),
                new ComboBoxInfo("DongDien_fb", "Dòng điện"),
                new ComboBoxInfo("CongSuat_fb", "Công suất"),
                new ComboBoxInfo("ViTriVan_fb", "Vị trí van")
            };
            SelectedPV = PVTypes[0].Value;

            LoadTrendDataAndDraw(_currentK, SelectedPV);
        }

        private float CalcSimatic(float minValue, float maxValue, float percent)
        {
            return minValue + (maxValue - minValue) * percent / 100f;
        }
        public void LoadTrendDataAndDraw(int k, string pvType)
        {
            var fileFormat = "csv"; 
            var trendFolder = Path.Combine(UserSetting.TOMFAN_folder, "Trend");
            if (!Directory.Exists(trendFolder)) return;

            var files = Directory.GetFiles(trendFolder, $"{k}.{fileFormat}");
            if (files.Length == 0) return;

            var filePath = files[0];
            var trendList = new List<TrendTime>();

            using (var reader = new StreamReader(filePath))
            {
                string? line;
                int row = 0;
                while ((line = reader.ReadLine()) != null)
                {
                    row++;
                    if (row < 0) continue;
                    var values = line.Split(' ');
                    if (values.Length < 13) continue;
                    var trend = new TrendTime
                    {
                        Index = row,
                        Time = float.TryParse(values[0], NumberStyles.Float, CultureInfo.InvariantCulture, out var time) ? time : 0,
                        NhietDoMoiTruong_sen = CalcSimatic(_sensor.NhietDoMoiTruongMin, _sensor.NhietDoMoiTruongMax, float.TryParse(values[1], NumberStyles.Float, CultureInfo.InvariantCulture, out var pv1) ? pv1 : 0),
                        DoAm_sen = CalcSimatic(_sensor.DoAmMoiTruongMin, _sensor.DoAmMoiTruongMax, float.TryParse(values[2], NumberStyles.Float, CultureInfo.InvariantCulture, out var pv2) ? pv2 : 0),
                        ApSuatkhiQuyen_sen = CalcSimatic(_sensor.ApSuatKhiQuyenMin, _sensor.ApSuatKhiQuyenMax, float.TryParse(values[3], NumberStyles.Float, CultureInfo.InvariantCulture, out var pv3) ? pv3 : 0),
                        ChenhLechApSuat_sen = CalcSimatic(_sensor.ChenhLechApSuatMin, _sensor.ChenhLechApSuatMax, float.TryParse(values[4], NumberStyles.Float, CultureInfo.InvariantCulture, out var pv4) ? pv4 : 0),
                        ApSuatTinh_sen = CalcSimatic(_sensor.ApSuatTinhMin, _sensor.ApSuatTinhMax, float.TryParse(values[5], NumberStyles.Float, CultureInfo.InvariantCulture, out var pv5) ? pv5 : 0),
                        DoRung_sen = CalcSimatic(_sensor.DoRungMin, _sensor.DoRungMax, float.TryParse(values[6], NumberStyles.Float, CultureInfo.InvariantCulture, out var pv6) ? pv6 : 0),
                        DoOn_sen = CalcSimatic(_sensor.DoOnMin, _sensor.DoOnMax, float.TryParse(values[7], NumberStyles.Float, CultureInfo.InvariantCulture, out var pv7) ? pv7 : 0),
                        SoVongQuay_sen = CalcSimatic(_sensor.SoVongQuayMin, _sensor.SoVongQuayMax, float.TryParse(values[8], NumberStyles.Float, CultureInfo.InvariantCulture, out var pv8) ? pv8 : 0),
                        Momen_sen = CalcSimatic(_sensor.MomenMin, _sensor.MomenMax, float.TryParse(values[9], NumberStyles.Float, CultureInfo.InvariantCulture, out var pv9) ? pv9 : 0),
                        DongDien_fb = CalcSimatic(_sensor.PhanHoiDongDienMin, _sensor.PhanHoiDongDienMax, float.TryParse(values[10], NumberStyles.Float, CultureInfo.InvariantCulture, out var fb1) ? fb1 : 0),
                        CongSuat_fb = CalcSimatic(_sensor.PhanHoiCongSuatMin, _sensor.PhanHoiCongSuatMax, float.TryParse(values[11], NumberStyles.Float, CultureInfo.InvariantCulture, out var fb2) ? fb2 : 0),
                        ViTriVan_fb = CalcSimatic(_sensor.PhanHoiViTriVanMin, _sensor.PhanHoiViTriVanMax, float.TryParse(values[12], NumberStyles.Float, CultureInfo.InvariantCulture, out var fb3) ? fb3 : 0)
                    };
                    trendList.Add(trend);
                }
            }

            if (trendList.Count == 0)
            {
                return;
            }

            var parametersToDisplay = new List<string>()
            {
                "NhietDoMoiTruong_sen",
                "DoAm_sen",
                "ApSuatkhiQuyen_sen",
                "ChenhLechApSuat_sen",
                "ApSuatTinh_sen",
                "DoRung_sen",
                "DoOn_sen",
                "SoVongQuay_sen",
                "Momen_sen",
                "DongDien_fb",
                "CongSuat_fb",
                "ViTriVan_fb"
            };

            var thongSoTrend = CalculateTrendStatistics(trendList, parametersToDisplay);

            TrendPlotModel.Title = $"Trend line k = {k} - {GetDisplayName(pvType)}";
            this.Title = TrendPlotModel.Title;
            TrendPlotModel.Series.Clear();

            var color = OxyColors.Red;
            var lineSeries = CreateLineSeries(trendList, pvType, color);
            if (lineSeries != null)
            {
                TrendPlotModel.Series.Add(lineSeries);
            }
            if (trendList.Count > 0)
            {
                var xAxis = TrendPlotModel.Axes.FirstOrDefault(a => a.Position == AxisPosition.Bottom) as LinearAxis;
                if (xAxis != null)
                {
                    xAxis.Minimum = trendList.Min(t => t.Time);
                    xAxis.Maximum = trendList.Max(t => t.Time);
                    xAxis.Title = "Thời gian phản hồi";
                }

                var yAxis = TrendPlotModel.Axes.FirstOrDefault(a => a.Position == AxisPosition.Left) as LinearAxis;
                if (yAxis != null)
                {
                    var values = trendList.Select(t => GetPropertyValue(t, pvType)).ToList();
                    yAxis.Minimum = Math.Min(0, values.Min() - 5);
                    yAxis.Maximum = values.Max() + 5;
                    yAxis.Title = GetDisplayName(pvType);
                }
            }
            TrendPlotModel.InvalidatePlot(true);
        }
        private Dictionary<string, (float Max, float Min, float Average)> CalculateTrendStatistics(List<TrendTime> trendList, List<string> parameters)
        {
            var result = new Dictionary<string, (float Max, float Min, float Average)>();

            foreach (var param in parameters)
            {
                var values = trendList.Select(t => GetPropertyValue(t, param)).ToList();
                if (values.Count == 0) continue;

                float max = values.Max();
                float min = values.Min();
                float average = (min != 0 && (max / min) < 1.03f) ? values.Average() : -1f;

                result[param] = (max, min, average);
            }

            return result;
        }



        private LineSeries CreateLineSeries(List<TrendTime> trendList, string parameterName, OxyColor color)
        {
            try
            {
                var lineSeries = new LineSeries
                {
                    Title = GetDisplayName(parameterName),
                    Color = color,
                    StrokeThickness = 2,
                    MarkerType = MarkerType.None
                };

                var points = new List<DataPoint>();
                foreach (var trend in trendList)
                {
                    var value = GetPropertyValue(trend, parameterName);
                    points.Add(new DataPoint(trend.Time, value));
                }

                lineSeries.Points.AddRange(points);
                return lineSeries;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Lỗi tạo series cho {parameterName}: {ex.Message}");
                return null;
            }
        }

        private float GetPropertyValue(TrendTime trend, string propertyName)
        {
            try
            {
                var property = typeof(TrendTime).GetProperty(propertyName);
                return property != null ? (float)property.GetValue(trend) : 0;
            }
            catch
            {
                return 0;
            }
        }

        private string GetDisplayName(string propertyName)
        {
            return propertyName switch
            {
                "NhietDoMoiTruong_sen" => "Nhiệt độ môi trường (°C)",
                "DoAm_sen" => "Độ ẩm (%)",
                "ApSuatkhiQuyen_sen" => "Áp suất khí quyển (Pa)",
                "ChenhLechApSuat_sen" => "Chênh lệch áp suất",
                "ApSuatTinh_sen" => "Áp suất tĩnh (Pa)",
                "DoRung_sen" => "Độ rung (mm/s)",
                "DoOn_sen" => "Độ ồn (dB)",
                "SoVongQuay_sen" => "Tốc độ quay (RPM)",
                "Momen_sen" => "Mômen (Nm)",
                "DongDien_fb" => "Dòng điện (A)",
                "CongSuat_fb" => "Công suất (kW)",
                "ViTriVan_fb" => "Vị trí van",
                _ => propertyName
            };
        }
    }
}