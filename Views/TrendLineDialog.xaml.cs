using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Series;
using OfficeOpenXml;
using System.IO;
using System.Windows;
using TESMEA_TMS.DTOs;
using TESMEA_TMS.Configs;
using System.Diagnostics;

namespace TESMEA_TMS.Views
{
    public partial class TrendLineDialog : Window
    {
        public PlotModel TrendPlotModel { get; set; }

        public TrendLineDialog(int k)
        {
            InitializeComponent();
            DataContext = this;
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
                Title = "Nhiệt độ môi trường (°C)",
                Minimum = 0,
                MajorGridlineStyle = LineStyle.Solid,
                MinorGridlineStyle = LineStyle.Dot,
                IsPanEnabled = false,
                IsZoomEnabled = false,
            });

            LoadTrendDataAndDraw(k);
        }

        public void LoadTrendDataAndDraw(int k)
        {
            var trendFolder = Path.Combine(UserSetting.TOMFAN_folder);
            if (!Directory.Exists(trendFolder)) return;

            // Chỉ lấy file có k đúng với lựa chọn
            var files = Directory.GetFiles(trendFolder, $"{k}.*.xlsx");
            if (files.Length == 0) return;

            var filePath = files[0];
            var trendList = new List<TrendTime>();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                if (worksheet != null)
                {
                    for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
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

            if (trendList.Count == 0)
            {
                return;
            }

            TrendPlotModel.Title = $"Trend line k = {k}";
            this.Title = $"Trend line k = {k}";
            TrendPlotModel.Series.Clear();


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

            var colors = new[]
            {
                OxyColors.Red, OxyColors.Blue, OxyColors.Green,
                OxyColors.Orange, OxyColors.Purple, OxyColors.Brown
            };
            int colorIndex = 0;
            foreach (var param in parametersToDisplay)
            {
                var lineSeries = CreateLineSeries(trendList, param, colors[colorIndex % colors.Length]);
                if (lineSeries != null)
                {
                    TrendPlotModel.Series.Add(lineSeries);
                }
                colorIndex++;
            }

            //// Vẽ dữ liệu lên biểu đồ
            //var lineSeries = new LineSeries
            //{
            //    Title = "Nhiệt độ môi trường",
            //    Color = OxyColors.Red,
            //    StrokeThickness = 2
            //};

            //foreach (var trend in trendList)
            //{
            //    lineSeries.Points.Add(new DataPoint(trend.Time, trend.NhietDoMoiTruong_sen));
            //}

            //TrendPlotModel.Series.Add(lineSeries);
            TrendPlotModel.InvalidatePlot(true);
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

                // Sử dụng AddRange để tăng hiệu suất
                var points = new List<DataPoint>();
                foreach (var trend in trendList)
                {
                    var value = GetPropertyValue(trend, parameterName);
                    var timeInSeconds = trend.Time / 1000f; // Chuyển ms sang giây
                    points.Add(new DataPoint(timeInSeconds, value));
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