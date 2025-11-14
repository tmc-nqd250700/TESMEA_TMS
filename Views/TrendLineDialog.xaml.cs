using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Series;
using OfficeOpenXml;
using System.IO;
using System.Windows;
using TESMEA_TMS.DTOs;
using TESMEA_TMS.Configs;
using System.Diagnostics;
using System.Collections.ObjectModel;
using System.ComponentModel;
using LicenseContext = OfficeOpenXml.LicenseContext;

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

        public TrendLineDialog(int k)
        {
            InitializeComponent();
            DataContext = this;
            _currentK = k;
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
                new ComboBoxInfo("ApSuatTinh_sen", "Áp suất tính"),
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

        //public void LoadTrendDataAndDraw(int k, string pvType)
        //{
        // 
        public void LoadTrendDataAndDraw(int k, string pvType)
        {
            var fileFormat = "csv"; // Temporary setting to "csv"
            var trendFolder = Path.Combine(UserSetting.TOMFAN_folder);
            if (!Directory.Exists(trendFolder)) return;

            // Dynamically select files based on the file format
            var files = Directory.GetFiles(trendFolder, $"{k}.*.{fileFormat}");
            if (files.Length == 0) return;

            var filePath = files[0];
            var trendList = new List<TrendTime>();

            if (fileFormat == "xlsx")
            {
                // Handle Excel files
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                    if (worksheet != null)
                    {
                        for (int row = 3; row <= worksheet.Dimension.End.Row; row++)
                        {
                            var trend = new TrendTime
                            {
                                Index = int.TryParse(worksheet.Cells[row, 1].Text, out var idx) ? idx : 0,
                                Time = float.TryParse(worksheet.Cells[row, 2].Text, out var time) ? time : 0,
                                NhietDoMoiTruong_sen = float.TryParse(worksheet.Cells[row, 3].Text, out var pv1) ? pv1 : 0,
                                DoAm_sen = float.TryParse(worksheet.Cells[row, 4].Text, out var pv2) ? pv2 : 0,
                                ApSuatkhiQuyen_sen = float.TryParse(worksheet.Cells[row, 5].Text, out var pv3) ? pv3 : 0,
                                ChenhLechApSuat_sen = float.TryParse(worksheet.Cells[row, 6].Text, out var pv4) ? pv4 : 0,
                                ApSuatTinh_sen = float.TryParse(worksheet.Cells[row, 7].Text, out var pv5) ? pv5 : 0,
                                DoRung_sen = float.TryParse(worksheet.Cells[row, 8].Text, out var pv6) ? pv6 : 0,
                                DoOn_sen = float.TryParse(worksheet.Cells[row, 9].Text, out var pv7) ? pv7 : 0,
                                SoVongQuay_sen = float.TryParse(worksheet.Cells[row, 10].Text, out var pv8) ? pv8 : 0,
                                Momen_sen = float.TryParse(worksheet.Cells[row, 11].Text, out var pv9) ? pv9 : 0,
                                DongDien_fb = float.TryParse(worksheet.Cells[row, 12].Text, out var fb1) ? fb1 : 0,
                                CongSuat_fb = float.TryParse(worksheet.Cells[row, 13].Text, out var fb2) ? fb2 : 0,
                                ViTriVan_fb = float.TryParse(worksheet.Cells[row, 14].Text, out var fb3) ? fb3 : 0,
                            };
                            trendList.Add(trend);
                        }
                    }
                }
            }
            else if (fileFormat == "csv")
            {
                // Handle CSV files
                using (var reader = new StreamReader(filePath))
                {
                    string? line;
                    int row = 0;
                    while ((line = reader.ReadLine()) != null)
                    {
                        row++;
                        if (row < 3) continue; // Skip header rows

                        var values = line.Split(',');
                        if (values.Length < 14) continue; // Ensure sufficient columns

                        var trend = new TrendTime
                        {
                            Index = int.TryParse(values[0], out var idx) ? idx : 0,
                            Time = float.TryParse(values[1], out var time) ? time : 0,
                            NhietDoMoiTruong_sen = float.TryParse(values[2], out var pv1) ? pv1 : 0,
                            DoAm_sen = float.TryParse(values[3], out var pv2) ? pv2 : 0,
                            ApSuatkhiQuyen_sen = float.TryParse(values[4], out var pv3) ? pv3 : 0,
                            ChenhLechApSuat_sen = float.TryParse(values[5], out var pv4) ? pv4 : 0,
                            ApSuatTinh_sen = float.TryParse(values[6], out var pv5) ? pv5 : 0,
                            DoRung_sen = float.TryParse(values[7], out var pv6) ? pv6 : 0,
                            DoOn_sen = float.TryParse(values[8], out var pv7) ? pv7 : 0,
                            SoVongQuay_sen = float.TryParse(values[9], out var pv8) ? pv8 : 0,
                            Momen_sen = float.TryParse(values[10], out var pv9) ? pv9 : 0,
                            DongDien_fb = float.TryParse(values[11], out var fb1) ? fb1 : 0,
                            CongSuat_fb = float.TryParse(values[12], out var fb2) ? fb2 : 0,
                            ViTriVan_fb = float.TryParse(values[13], out var fb3) ? fb3 : 0,
                        };
                        trendList.Add(trend);
                    }
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
                "ApSuatkhiQuyen_sen" => "Áp suất khí quyển (hPa)",
                "ChenhLechApSuat_sen" => "Chênh lệch áp suất",
                "ApSuatTinh_sen" => "Áp suất tính (hPa)",
                "DoRung_sen" => "Độ rung (mm/s)",
                "DoOn_sen" => "Độ ồn (dB)",
                "SoVongQuay_sen" => "Tốc độ quay (RPM)",
                "Momen_sen" => "Mômen (N.m)",
                "DongDien_fb" => "Dòng điện (A)",
                "CongSuat_fb" => "Công suất (kW)",
                "ViTriVan_fb" => "Vị trí van (%)",
                _ => propertyName
            };
        }
    }
}