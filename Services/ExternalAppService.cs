using OfficeOpenXml;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;
using TESMEA_TMS.Configs;
using TESMEA_TMS.DTOs;
using TESMEA_TMS.Helpers;
using TESMEA_TMS.Models.Entities;

namespace TESMEA_TMS.Services
{
    public interface IExternalAppService
    {
        Task StartAppAsync();
        Task StopAppAsync();
        Task<bool> ConnectExchangeAsync(List<Measure> measures, BienTan inv, CamBien sensor, OngGio duct, ThongTinMauThuNghiem input, string kieuKiemThu, float maxmin, float timeRange);
        //Task<bool> ConnectExchangeAsync(float timeRange);
        Task StartExchangeAsync();
        Task StopExchangeAsync();
        bool IsAppRunning { get; }
        bool IsConnectedToSimatic { get; }
        event Action<Measure> OnSimaticResultReceived;
        event Action<bool> OnSimaticConnectionChanged;
        event Action<List<Measure>> OnSimaticExchangeCompleted;
        event Action<MeasureResponse, ParameterShow> OnMeasurePointCompleted;
        event Action<MeasureFittingFC, Measure> OnMeasureRangeCompleted;
    }

    public class ExternalAppService : IExternalAppService
    {
        private Process? _process;
        private CancellationTokenSource? _cts;

        private int _currentIndex = 0;
        private List<Measure> _measures = new List<Measure>();
        private BienTan _inv = new BienTan();
        private CamBien _sensor = new CamBien();
        private OngGio _duct = new OngGio();
        private ThongTinMauThuNghiem _input = new ThongTinMauThuNghiem();

        private List<Measure> _simaticResults = new List<Measure>();
        private FileSystemWatcher _watcher;
        private string _exchangeFolder;
        private string _trendFolder;
        private bool _isComma = true;

        private bool _isEStop = false;
        private float[] avgs = new float[13];
        private string _lastReadTargetLine = string.Empty;
        public ExternalAppService()
        {
            _exchangeFolder = UserSetting.TOMFAN_folder;
            _trendFolder = Path.Combine(_exchangeFolder, "Trend");
        }

        // mở wincc
        public async Task StartAppAsync()
        {
            string winccExePath = UserSetting.Instance.WinccExePath;
            string simaticProjectPath = UserSetting.Instance.SimaticPath;


            // Kiểm tra File Thực thi WinCC
            if (string.IsNullOrEmpty(winccExePath) || !File.Exists(winccExePath))
            {
                MessageBoxHelper.ShowWarning("Đường dẫn file thực thi WinCC (.exe) không tồn tại");
                return;
            }

            // Kiểm tra File Dự án (Để chắc chắn có thể mở được)
            if (string.IsNullOrEmpty(simaticProjectPath) || !File.Exists(simaticProjectPath))
            {
                MessageBoxHelper.ShowWarning("Đường dẫn file dự án Simatic/WinCC không tồn tại");
                return;
            }
            foreach (var proc in Process.GetProcessesByName(Path.GetFileNameWithoutExtension(simaticProjectPath)))
            {
                try
                {
                    string fileName = string.Empty;
                    try
                    {
                        fileName = proc.MainModule.FileName;
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Không thể truy cập MainModule: {ex.Message}");
                        continue;
                    }

                    if (fileName.Equals(simaticProjectPath, StringComparison.OrdinalIgnoreCase) && !proc.HasExited)
                    {
                        proc.Kill(true);
                        proc.WaitForExit(2000);
                    }
                }
                catch (System.ComponentModel.Win32Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Win32Exception: {ex.Message}");
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Exception: {ex.Message}");
                }
            }
            _cts = new CancellationTokenSource();
            try
            {
                _process = new Process
                {
                    //StartInfo = new ProcessStartInfo(winccExePath)
                    //{
                    //    Arguments = $"\"{simaticProjectPath}\"",
                    //    UseShellExecute = false,
                    //    CreateNoWindow = false,
                    //    WindowStyle = ProcessWindowStyle.Normal
                    //}

                    StartInfo = new ProcessStartInfo(simaticProjectPath)
                    {
                        UseShellExecute = false,
                        CreateNoWindow = false,
                        WindowStyle = ProcessWindowStyle.Normal
                    }
                };
                _process.StartInfo.UseShellExecute = false;
                // đăng ký event để monitor process
                _process.Exited += (s, e) =>
                {
                    System.Diagnostics.Debug.WriteLine($"Simatic process exited unexpectedly at {DateTime.Now}");
                };
                _process.EnableRaisingEvents = true;
                _process.Start();
                //var started = await Task.Run(() =>
                //{
                //    return _process.WaitForInputIdle(UserSetting.Instance.TimeoutMilliseconds);
                //}, _cts.Token);

                //if (!started)
                //{
                //    // timeout exception
                //    await StopAppAsync();
                //    MessageBoxHelper.ShowError("Hết thời gian chờ khởi động phần mềm Simatic");
                //}
            }
            catch (Win32Exception ex)
            {
                await StopAppAsync();
                MessageBoxHelper.ShowError("Lỗi hệ thống khi khởi động phần mềm Simatic");
            }
            catch (UnauthorizedAccessException ex)
            {
                await StopAppAsync();
                MessageBoxHelper.ShowError("Không đủ quyền để khởi động phần mềm Simatic");
            }
            catch (InvalidOperationException ex)
            {
                await StopAppAsync();
                MessageBoxHelper.ShowError("Thao tác không hợp lệ với process Simatic");
            }
            catch (OperationCanceledException ex)
            {
                await StopAppAsync();
                MessageBoxHelper.ShowError("Thao tác khởi động phần mềm Simatic bị hủy");
            }
            catch (AggregateException ex)
            {
                await StopAppAsync();
                MessageBoxHelper.ShowError("Có lỗi bất đồng bộ khi khởi động phần mềm Simatic");
            }
            catch (NotSupportedException ex)
            {
                await StopAppAsync();
                MessageBoxHelper.ShowError("Đường dẫn hoặc thao tác không được hỗ trợ");
            }
            catch (Exception ex)
            {
                await StopAppAsync();
                MessageBoxHelper.ShowError($"Lỗi khởi động Simatic: {ex.Message}");
            }
        }

        // stop wincc
        public async Task StopAppAsync()
        {
            try
            {
                foreach (var proc in Process.GetProcessesByName(Path.GetFileNameWithoutExtension(UserSetting.Instance.SimaticPath)))
                {
                    try
                    {
                        string fileName = string.Empty;
                        try
                        {
                            fileName = proc.MainModule.FileName;
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Không thể truy cập MainModule: {ex.Message}");
                            continue;
                        }

                        if (fileName.Equals(UserSetting.Instance.SimaticPath, StringComparison.OrdinalIgnoreCase) && !proc.HasExited)
                        {
                            try
                            {
                                proc.Kill(true);
                                proc.WaitForExit(2000);
                            }
                            catch (System.ComponentModel.Win32Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"Win32Exception khi kill tiến trình: {ex.Message}");
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"Exception khi kill tiến trình: {ex.Message}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Exception ngoài: {ex.Message}");
                    }
                }


                if (_process != null && !_process.HasExited)
                {
                    await Task.Run(() =>
                    {
                        _process.Kill(true);
                        _process.Dispose();
                        _process = null;
                    });
                }
                if (_cts != null)
                {
                    _cts.Cancel();
                    _cts.Dispose();
                    _cts = null;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in StopAppAsync: {ex.Message}");
            }
        }

        // kiểm tra trạng thái wincc
        public bool IsAppRunning
        {
            get
            {
                try
                {
                    if (_process == null)
                    {
                        System.Diagnostics.Debug.WriteLine("Process Simatic is null");
                        return false;
                    }

                    var hasExited = _process.HasExited;
                    if (hasExited)
                    {
                        try
                        {
                            System.Diagnostics.Debug.WriteLine($"Process Simatic has exited. ExitCode: {_process.ExitCode}, ExitTime: {_process.ExitTime}");
                        }
                        catch
                        {
                            System.Diagnostics.Debug.WriteLine("Process Simatic has exited but cannot get exit details");
                        }
                    }

                    return !hasExited;
                }
                catch (InvalidOperationException ex)
                {
                    System.Diagnostics.Debug.WriteLine($"InvalidOperationException in IsAppRunning Simatic: {ex.Message}");
                    return false;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Exception in IsAppRunning Simatic: {ex.Message}");
                    return false;
                }
            }
        }

        public bool IsConnectedToSimatic { get; set; } = false;
        private TaskCompletionSource<bool>? _connectCompletionSource;

        // kết nối dòng 1
        public async Task<bool> ConnectExchangeAsync(List<Measure> measures, BienTan inv, CamBien sensor, OngGio duct, ThongTinMauThuNghiem input, string kieuKiemThu, float maxmin, float timeRange)
        {
            try
            {
                WriteTomfanLog("========== BẮT ĐẦU QUY TRÌNH KẾT NỐI (CONNECT) ==========");
                avgs = new float[13];
                _measures = measures;
                _inv = inv;
                _sensor = sensor;
                _duct = duct;
                _input = input;
                _isEStop = false;
                DataProcess.Initialize(_measures.Count, _inv, _sensor, _duct, _input, kieuKiemThu);
                if (!Directory.Exists(_exchangeFolder))
                {
                    WriteTomfanLog($"Thư mục trao đổi không tồn tại: {_exchangeFolder}");
                    throw new BusinessException("Thư mục trao đổi dữ liệu với Simatic không tồn tại");
                }

                if (!Directory.Exists(Path.Combine(_exchangeFolder, "ZERO")))
                {
                    Directory.CreateDirectory(Path.Combine(_exchangeFolder, "ZERO"));
                }

                _simaticResults.Clear();
                _lastReadTargetLine = string.Empty;

                var m = _measures[0];
                WriteTomfanLog($"Connect - Ghi file và chờ WinCC phản hồi..");
                // tao file 0.csv để zero-span
                string zeroSpanPath = Path.Combine(_exchangeFolder, "ZERO", "0.csv");
                using (var fs = File.Create(zeroSpanPath)) { }
                using (var fs = File.Create(Path.Combine(_exchangeFolder, "History", "History.csv"))) { }

                await WriteDataToFilesAsync(m, maxmin);
                var result = await WaitForResultAsync(1, isConnection: true);
                if (result == null || Math.Abs(result.S - m.S) > 0.01)
                {
                    string errorMsg = result == null ? "Timeout" : "Dữ liệu không khớp";
                    WriteTomfanLog($"Kết nối thất bại");
                    throw new Exception($"Không thể kết nối, lỗi: {errorMsg}");
                }

                //try
                //{
                //    using (var fs = new FileStream(zeroSpanPath, FileMode.Open, FileAccess.Read, FileShare.Read))
                //    using (var sr = new StreamReader(fs))
                //    {
                //        var allLines = await File.ReadAllLinesAsync(zeroSpanPath);
                //        if (allLines.Length == 0)
                //            throw new BusinessException("Không có dữ liệu từ file 0.csv");

                //        float[] sums = new float[13];
                //        int count = 0;
                //        foreach (var l in allLines)
                //        {
                //            var vals = l.Split(' ');
                //            if (vals.Length < 13)
                //            {
                //                WriteTomfanLog($"Dòng {count + 1} không đủ 12 tín hiệu cảm biến");
                //                continue;
                //            }
                //            for (int i = 0; i < 13; i++)
                //            {
                //                if (float.TryParse(vals[i], NumberStyles.Float, CultureInfo.InvariantCulture, out var v))
                //                    sums[i] += v;
                //            }
                //            count++;
                //        }

                //        if (count == 0)
                //            throw new BusinessException("Không có dữ liệu hợp lệ trong file 0.csv");


                //        avgs = sums.Select(x => x / count).ToArray();
                //        if (avgs.Any(x => float.IsNaN(x) || float.IsInfinity(x)))
                //            throw new BusinessException("Lỗi hiệu chỉnh cảm biến");

                //        using (var package = new ExcelPackage(new FileInfo(Path.Combine(_exchangeFolder, "MeasurementSummary.xlsx"))))
                //        {
                //            var ws2 = package.Workbook.Worksheets.FirstOrDefault(x => x.Name == "ZeroSpan");
                //            if (ws2 == null)
                //            {
                //                ws2 = package.Workbook.Worksheets.Add("ZeroSpan");
                //            }
                //            ws2.Cells[1, 1].Value = "T môi trường (%)";
                //            ws2.Cells[1, 2].Value = "Độ ẩm (%)";
                //            ws2.Cells[1, 3].Value = "Vị trí van (%)";
                //            ws2.Cells[1, 4].Value = "Momen (%)";
                //            ws2.Cells[1, 5].Value = "Hồng ngoại (%)";
                //            ws2.Cells[1, 6].Value = "Độ rung (%)";
                //            ws2.Cells[1, 7].Value = "Số vòng quay (%)";
                //            ws2.Cells[1, 8].Value = "Dòng diện - AM (%)";
                //            ws2.Cells[1, 9].Value = "Áp suất tĩnh - Chênh áp 2 (%)";
                //            ws2.Cells[1, 10].Value = "Công suất (%)";
                //            ws2.Cells[1, 11].Value = "Chênh áp - Chênh áp 1 (%)";
                //            ws2.Cells[1, 12].Value = "Áp suất khí quyển (%)";
                //            using (var range = ws2.Cells[1, 1, 1, 12])
                //            {
                //                range.Style.Font.Bold = true;
                //                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                //                range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                //            }
                //            ws2.Cells[2, 1].Value = avgs[1];
                //            ws2.Cells[2, 2].Value = avgs[2];
                //            ws2.Cells[2, 3].Value = avgs[3];
                //            ws2.Cells[2, 4].Value = avgs[4];
                //            ws2.Cells[2, 5].Value = avgs[5];
                //            ws2.Cells[2, 6].Value = avgs[6];
                //            ws2.Cells[2, 7].Value = avgs[7];
                //            ws2.Cells[2, 8].Value = avgs[8];
                //            ws2.Cells[2, 9].Value = avgs[9];
                //            ws2.Cells[2, 10].Value = avgs[10];
                //            ws2.Cells[2, 11].Value = avgs[11];
                //            ws2.Cells[2, 12].Value = avgs[12];

                //            ws2.Cells.AutoFitColumns();
                //            package.Save();
                //        }
                //    }
                //}
                //catch (Exception ex)
                //{
                //    WriteTomfanLog($"Lỗi khi ZeroSpan từ file 0.csv: {ex.Message}");
                //    throw;
                //}
                m.F = MeasureStatus.Completed;
                _currentIndex = m.k;
                return await ConnectExchangeAsync(maxmin);
            }
            catch (Exception ex)
            {
                WriteTomfanLog($"Lỗi trong ConnectExchangeAsync: {ex.Message}");
                throw;
            }
        }

        // kết nối dòng 2
        public async Task<bool> ConnectExchangeAsync(float maxmin)
        {
            try
            {
                //var m = new Measure
                //{
                //    k = 0,
                //    S = 0,
                //    CV = 0
                //};
                var m = _measures[1];
                await WriteDataToFilesAsync(m, maxmin);

                var result = await WaitForResultAsync(2, isConnection: true);
                if (result == null || Math.Abs(result.S - m.S) > 0.01)
                {
                    string errorMsg = result == null ? "Timeout" : "Dữ liệu không khớp";
                    WriteTomfanLog($"Kết nối thất bại");
                    throw new Exception($"Không thể kết nối, lỗi: {errorMsg}");
                }

                m.F = MeasureStatus.Completed;
                WriteTomfanLog($"Connect thành công");

                IsConnectedToSimatic = true;
                OnSimaticConnectionChanged?.Invoke(true);
                _currentIndex = m.k;

                WriteTomfanLog("Đã thiết lập kết nối với Simatic thành công");
                return true;
            }
            catch (Exception ex)
            {
                WriteTomfanLog($"Lỗi trong ConnectExchangeAsync: {ex.Message}");
                throw;
            }
        }

        // đo kiểm từ dòng 3
        public async Task StartExchangeAsync()
        {
            try
            {
                if (!IsConnectedToSimatic)
                {
                    WriteTomfanLog("StartExchange bị từ chối: Chưa kết nối Simatic");
                    MessageBoxHelper.ShowWarning("Chưa kết nối với Simatic");
                    return;
                }

                WriteTomfanLog("========== Bắt đầu đo kiểm ==========");

                if (!Directory.Exists(_trendFolder))
                    Directory.CreateDirectory(_trendFolder);

                List<Measure> currentRange = new List<Measure>();
                double currentS = _measures[_currentIndex].S;
                int startIndex = 2;
                int timeoutMs = UserSetting.Instance.TimeoutMilliseconds;

                for (int i = _currentIndex; i < _measures.Count; i++)
                {
                    _currentIndex = _measures[i].k;
                    var m = _measures[i];

                    // check chuyển tần số thì fitting
                    if (_measures[i].S != currentS)
                    {
                        var fitting = DataProcess.FittingFC(currentRange.Count, startIndex);
                        WriteTomfanLog($"Hoàn tất dải đo với tần số {currentRange.First().S} từ CV={currentRange.First().CV}% đến {currentRange.Last().CV}%");
                        OnMeasureRangeCompleted?.Invoke(fitting, currentRange.LastOrDefault());
                        currentRange.Clear();
                        currentS = _measures[i].S;
                        startIndex = i;
                        WriteTomfanLog("Chuyển sang dải đo tiếp theo");
                    }

                    WriteTomfanLog($"Đang xử lý điểm đo k={m.k}/{_measures.Count}");
                    using (var fs = File.Create(Path.Combine(_trendFolder, $"{m.k}.csv"))) { }
                    await WriteDataToFilesAsync(m);
                    if (_currentIndex > 3)
                    {
                        // delay 15s den khi ghi dong tiep theo
                        WriteTomfanLog("Delay 15s sau đó chờ kết quả dòng tiếp theo");
                        await Task.Delay(1000);
                        WriteTomfanLog("Delay xong, tiếp tục lắng nghe dòng tiếp theo");
                    }
                    // Chờ kết quả xử lý thực tế (isConnection = false để tính toán sensor)
                    // Luôn lắng nghe dòng 3
                    var result = await WaitForResultAsync(m.k, isConnection: false);

                    if (result != null)
                    {
                        if(result.CongSuat_fb > _input.CongSuatDongCo)
                        {
                            WriteTomfanLog("======= QUÁ TẢI ======= \nKết quả công suất từ Simatic vượt quá công suất định mức của động cơ, dừng đo kiểm và gửi lệnh E-Stop =======");
                            if(MessageBoxHelper.ShowQuestion("Quá tải công suất của động cơ, có muốn tiếp tục đo kiểm không?", "Cảnh báo quá tải"))
                                await StopExchangeAsync();
                        }
                        WriteTomfanLog($"Đã nhận kết quả k={m.k}");
                        m = result;
                        m.F = MeasureStatus.Completed;
                        _simaticResults.Add(result);

                        var measurePoint = DataProcess.OnePointMeasure(result);
                        OnMeasurePointCompleted?.Invoke(measurePoint, DataProcess.ParaShow(result));
                        OnSimaticResultReceived?.Invoke(m);

                        // thêm kết quả hiện tại vào range
                        if (!currentRange.Any(m => m.k == _measures[i].k))
                        {
                            currentRange.Add(_measures[i]);
                        }

                        // check là điểm đo cuối cùng thì thực hiện fitting FC và vẽ line cho chart
                        if (i == _measures.Count - 1)
                        {
                            var fitting = DataProcess.FittingFC(currentRange.Count, startIndex);
                            WriteTomfanLog($"Hoàn tất dải đo cuối cùng với tần số {currentRange.First().S} từ CV={currentRange.First().CV}% đến {currentRange.Last().CV}%");
                            OnMeasureRangeCompleted?.Invoke(fitting, currentRange.LastOrDefault());
                            currentRange.Clear();
                        }
                        else
                        {
                            // delay 15s den khi ghi dong tiep theo
                            WriteTomfanLog("Delay 15s trước khi ghi dòng tiếp theo");
                            await Task.Delay(1000);
                            WriteTomfanLog("Delay xong, tiếp tục ghi dữ liệu dòng tiếp theo");
                        }
                        WriteTomfanLog($"Hoàn tất điểm đo k={m.k}");
                    }
                    else
                    {
                        WriteTomfanLog($"Điểm đo k={m.k} thất bại");
                        m.F = MeasureStatus.Error;
                        OnSimaticResultReceived?.Invoke(m);
                    }

                }

                WriteTomfanLog("========== HOÀN TẤT TOÀN BỘ KỊCH BẢN ĐO KIỂM, DỪNG ĐO KIỂM, GỬI LỆNH 96 TỚI SIMATIC ==========");
                await StopExchangeAsync();
                OnSimaticExchangeCompleted?.Invoke(_simaticResults);
            }
            catch (Exception ex)
            {
                WriteTomfanLog($"Lỗi trong StartExchangeAsync: {ex.Message} {ex.StackTrace}");
                throw;
            }
        }

        // estop
        public async Task StopExchangeAsync()
        {
            try
            {
                WriteTomfanLog("--- LỆNH DỪNG KHẨN CẤP (E-STOP) ---");
                int nextIndex = _currentIndex + 1;

                // Tạo object eStop cho dòng mới
                var eStopMeasure = new Measure
                {
                    k = nextIndex,
                    S = 0,
                    CV = 0
                };
                await WriteDataToFilesAsync(eStopMeasure, 0, true);

                IsConnectedToSimatic = false;
                _isEStop = true;
                WriteTomfanLog($"Đã chèn lệnh E-Stop (96) vào dòng mới k={nextIndex}");

                try
                {
                    string timeStamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    string folderCopy = new DirectoryInfo(_exchangeFolder).Name;
                    string destFolder = Path.Combine(UserSetting.GetLocalAppPath(), timeStamp, folderCopy);

                    Directory.CreateDirectory(destFolder);

                    foreach (string dirPath in Directory.GetDirectories(_exchangeFolder, "*", SearchOption.AllDirectories))
                    {
                        Directory.CreateDirectory(dirPath.Replace(_exchangeFolder, destFolder));
                    }

                    foreach (string filePath in Directory.GetFiles(_exchangeFolder, "*", SearchOption.AllDirectories))
                    {
                        string destFilePath = filePath.Replace(_exchangeFolder, destFolder);
                        File.Copy(filePath, destFilePath, true);
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
            catch (Exception ex)
            {
                WriteTomfanLog($"Lỗi trong StopExchangeAsync: {ex.Message}");
            }
        }

        // tính giá trị trả về từ %
        private float CalcSimatic(float minValue, float maxValue, float percent, float sensorIdx, int indexK)
        {
            if (percent <= 0)
            {
                WriteTomfanLog("Không hội tụ được giá trị từ PLC, thực hiện hội tụ từ trendline");
                /// Vì file 2_S_IN.csv phía PLC trả về có 3 giá trị đầu là
                /// 0. Mã lệnh điều khiển: 100 - yêu cầu tính hội tụ, 96 - lệnh dừng khẩn cấp (E-Stop)
                /// 1. Giá trị S (% tần số) 
                /// 2. Giá trị CV (% van điều khiển)
                /// => ngoại trừ tín hiệu tần số phản hồi, các tín hiệu khác lấy từ 3 => -2 để fit với index bên trend
                percent = CalculateConvergingByTrend(sensorIdx == 1 ? sensorIdx : sensorIdx - 2, indexK);
            }

            if(percent > 118)
            {
                WriteTomfanLog($"Không kết nối được tới tín hiệu cảm biến {indexK}");
                return 0;
            }

            return minValue + (maxValue - minValue) * percent / 100f;
        }

        //private float CalcSimatic(float minValue, float maxValue, float percent, float zeroSpan)
        //{
        //    var newMax = maxValue * (100 - zeroSpan) / 100;
        //    return minValue + (newMax - minValue) * percent / 100f;
        //}

        // retry nếu file bị khóa
        private async Task<bool> ExecuteWithRetryAsync(Func<Task> action, int retries = 20, int delay = 200)
        {
            for (int i = 0; i < retries; i++)
            {
                try
                {
                    await action();
                    return true;
                }
                catch (IOException ex)
                {
                    WriteTomfanLog($"retry {i + 1}/{retries}: File đang bị khóa bởi WinCC: {ex.Message}");
                    if (i == retries - 1)
                    {
                        WriteTomfanLog("Đã thử lại tối đa nhưng vẫn không thể truy cập file");
                        throw;
                    }
                    await Task.Delay(delay);
                }
            }
            return false;
        }

        private async Task WriteDataToFilesAsync(Measure m, float col4Value = 0, bool eStop = false)
        {
            try
            {

                // force stop khi đã ấn estop, tránh tình trạng bên plc trả về kết quả dẫn đến việc tiếp tục gửi lệnh sau khi đã stop ở đây
                if (_isEStop) return;

                //string xlsxPath = Path.Combine(_exchangeFolder, "1_T_OUT.xlsx");
                string csvPath = Path.Combine(_exchangeFolder, "1_T_OUT.csv");
                char sep = _isComma ? ' ' : ';';
                WriteTomfanLog($">>> Ghi dữ liệu dòng k={m.k} (S={m.S}, CV={m.CV})");

                // Ghi CSV trực tiếp
                int kValueToPrint = eStop ? 96 : 100;
                int rowIdx = m.k > 0 ? m.k : 1;
                await ExecuteWithRetryAsync(async () =>
                {
                    List<string> lines = new List<string>();
                    if (File.Exists(csvPath))
                    {
                        lines = (await File.ReadAllLinesAsync(csvPath)).ToList();
                    }
                    else
                    {
                        throw new BusinessException("File 1_T_OUT.csv không tồn tại");
                    }

                    // Tạo nội dung dòng mới
                    string newLine = col4Value == 0
                        ? $"{kValueToPrint}{sep}{m.S}{sep}{m.CV}"
                        : $"{m.k}{sep}{m.S}{sep}{m.CV}{sep}{col4Value}";

                    WriteTomfanLog($"Dữ liệu dòng mới: {newLine}");
                    while (lines.Count < rowIdx)
                    {
                        lines.Add("");
                    }
                    lines[rowIdx - 1] = newLine;

                    // Ghi đè lại toàn bộ file trực tiếp
                    using (var fs = new FileStream(csvPath, FileMode.Truncate, FileAccess.Write, FileShare.None))
                    using (var sw = new StreamWriter(fs, Encoding.UTF8))
                    {
                        foreach (var line in lines)
                            await sw.WriteLineAsync(line);
                        await sw.FlushAsync();
                        fs.Flush(true);
                    }
                    WriteTomfanLog($"Step: CSV row {rowIdx} ghi thành công");
                });
            }
            catch (Exception ex)
            {
                WriteTomfanLog($"Error WriteDataToFilesAsync: {ex}");
            }
        }

        // tính hội tụ của tín hiệu cảm biến để lấy giá trị chốt, tránh trường hợp tín hiệu không ổn định, kết quả là %
        public float CalculateConvergingByTrend(float sensor, int index)
        {
            try
            {
                string trendFilePath = Path.Combine(_trendFolder, $"{index}.csv");
                if (!File.Exists(trendFilePath))
                {
                    WriteTomfanLog($"File trend không tồn tại: {trendFilePath}");
                    return 0;
                }

                var trendTimes = new List<TrendTime>();

                using (var reader = new StreamReader(trendFilePath))
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
                            // độ ồn và momen chưa có giá trị
                            Index = row,
                            Time = float.TryParse(values[0], NumberStyles.Float, CultureInfo.InvariantCulture, out var time) ? time : 0,
                            NhietDoMoiTruong_sen = float.TryParse(values[1], NumberStyles.Float, CultureInfo.InvariantCulture, out var v1) ? v1 : 0,
                            DoAm_sen = float.TryParse(values[2], NumberStyles.Float, CultureInfo.InvariantCulture, out var v2) ? v2 : 0,
                            ViTriVan_fb = float.TryParse(values[3], NumberStyles.Float, CultureInfo.InvariantCulture, out var v3) ? v3 : 0,
                            Momen_sen = float.TryParse(values[4], NumberStyles.Float, CultureInfo.InvariantCulture, out var v4) ? v4 : 0,
                            NhietDoGoi_sen = float.TryParse(values[5], NumberStyles.Float, CultureInfo.InvariantCulture, out var v5) ? v5 : 0,
                            DoRung_sen = float.TryParse(values[6], NumberStyles.Float, CultureInfo.InvariantCulture, out var v6) ? v6 : 0,
                            SoVongQuay_sen = float.TryParse(values[7], NumberStyles.Float, CultureInfo.InvariantCulture, out var v7) ? v7 : 0,
                            DongDien_fb = float.TryParse(values[8], NumberStyles.Float, CultureInfo.InvariantCulture, out var v8) ? v8 : 0,
                            ApSuatTinh_sen = float.TryParse(values[9], NumberStyles.Float, CultureInfo.InvariantCulture, out var v9) ? v9 : 0,
                            CongSuat_fb = float.TryParse(values[10], NumberStyles.Float, CultureInfo.InvariantCulture, out var v10) ? v10 : 0,
                            ChenhLechApSuat_sen = float.TryParse(values[11], NumberStyles.Float, CultureInfo.InvariantCulture, out var v11) ? v11 : 0,
                            ApSuatkhiQuyen_sen = float.TryParse(values[12], NumberStyles.Float, CultureInfo.InvariantCulture, out var v12) ? v12 : 0,
                        };
                        trendTimes.Add(trend);
                    }
                }

                if (trendTimes.Count <= 20)
                {
                    WriteTomfanLog($"Trendtime không đủ 20 giá trị để chốt dữ liệu. Tổng {trendTimes.Count}");
                    return 0;
                }


                float GetContinuousAverage(Func<TrendTime, float> selector, string property, float percent = 5)
                {
                    WriteTomfanLog($"=== Thực hiện tính trung bình cho tín hiệu cảm biến {property}");
                    int windowSize = 20;
                    float maxPercent = 100;
                    while (percent <= maxPercent)
                    {
                        float lower = 1f - percent / 100f;
                        float upper = 1f + percent / 100f;
                        WriteTomfanLog($"Tính giá trị trung bình với tiêu chuẩn {percent}%");

                        for (int start = 0; start <= trendTimes.Count - windowSize; start++)
                        {
                            bool isValid = true;
                            for (int i = start; i < start + windowSize - 1; i++)
                            {
                                float v1 = selector(trendTimes[i]);
                                float v2 = selector(trendTimes[i + 1]);
                                if (v1 == 0) continue; // bỏ qua giá trị 0
                                double ratio = v2 / v1;
                                if (ratio < lower || ratio > upper)
                                {
                                    isValid = false;
                                    break;
                                }
                            }
                            if (isValid)
                            {
                                var res = trendTimes.Skip(start).Take(windowSize).Average(selector);
                                WriteTomfanLog($"===Tìm thấy dải hợp lệ từ index {start} đến {start + windowSize - 1} với tiêu chuẩn {percent}, giá trị chốt: {res}");
                                return res;
                            }
                        }
                        percent++;
                    }
                    WriteTomfanLog($"Không tìm thấy dải hợp lệ sau khi tăng tiêu chuẩn đến giới hạn {maxPercent}");
                    return 0;
                }

                float val = 0;
                switch (sensor)
                {
                    case 1:
                        val = GetContinuousAverage(x => x.NhietDoMoiTruong_sen, "Nhiệt độ môi trường");
                        break;
                    case 2:
                        val = GetContinuousAverage(x => x.DoAm_sen, "Độ ẩm");
                        break;
                    case 3:
                        val = GetContinuousAverage(x => x.ViTriVan_fb, "Vị trí van");
                        break;
                    case 4:
                        val = GetContinuousAverage(x => x.Momen_sen, "Momen");
                        break;
                    case 5:
                        val = GetContinuousAverage(x => x.NhietDoGoi_sen, "Nhiệt độ hồng ngoại");
                        break;
                    case 6:
                        val = GetContinuousAverage(x => x.DoRung_sen, "Độ rung");
                        break;
                    case 7:
                        val = GetContinuousAverage(x => x.SoVongQuay_sen, "Số vòng quay");
                        break;
                    case 8:
                        val = GetContinuousAverage(x => x.DongDien_fb, "Dòng điện");
                        break;
                    case 9:
                        val = GetContinuousAverage(x => x.ApSuatTinh_sen, "Áp suất tĩnh");
                        break;
                    case 10:
                        val = GetContinuousAverage(x => x.CongSuat_fb, "Công suất");
                        break;
                    case 11:
                        val = GetContinuousAverage(x => x.ChenhLechApSuat_sen, "Chênh lệch áp suất");
                        break;
                    case 12:
                        val = GetContinuousAverage(x => x.ApSuatkhiQuyen_sen, "Áp suất khí quyển");
                        break;
                }

                return val;
            }
            catch (Exception ex)
            {
                throw;
            }
        }


        // chờ kết quả từ file 2.csv
        private async Task<Measure?> WaitForResultAsync(int expectedK, bool isConnection)
        {
            string path2 = Path.Combine(_exchangeFolder, "2_S_IN.csv");
            var sw = Stopwatch.StartNew();
            char sep = isConnection ? ' ' : ' ';
            WriteTomfanLog($"--- Bắt đầu chờ kết quả từ WinCC cho k={expectedK} ---");
            while (sw.ElapsedMilliseconds < UserSetting.Instance.TimeoutMilliseconds)
            {
                try
                {
                    if (File.Exists(path2))
                    {
                        using (var fs = new FileStream(path2, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                        using (var sr = new StreamReader(fs))
                        {
                            string[] lines = await File.ReadAllLinesAsync(path2);
                            //int targetIndex = isConnection ? expectedK - 1 : expectedK - 1;
                            int targetIndex = expectedK - 1;
                            if (lines.Length > 0)
                            {
                                var isNewLine = false;
                                var targetLine = lines[targetIndex];

                                // Kiểm tra nếu dòng có dữ liệu
                                // phía plc có thể gặp lỗi row luôn bị xóa khi chuyển sang điểm đo mới -> kết quả luôn nằm ở row này -> tạo thêm check với row3 để đảm bảo có dữ liệu trả về
                                if (!string.IsNullOrWhiteSpace(targetLine))
                                {
                                    isNewLine = true;
                                }
                                else if (lines.Length > 2)
                                {
                                    targetLine = lines[2];
                                    if (!string.IsNullOrEmpty(targetLine))
                                        isNewLine = true;
                                }

                                if (isNewLine)
                                {
                                    if (_lastReadTargetLine == targetLine) continue;
                                    _lastReadTargetLine = targetLine;
                                    var parts = targetLine.Split(sep);
                                    if (parts.Length < 3) continue;

                                    // do cái này đang check với 100
                                    //if (int.TryParse(parts[0], out int k) && k == expectedK)
                                    //{
                                    WriteTomfanLog($"Tìm thấy dòng k={expectedK} sau {sw.ElapsedMilliseconds}ms");
                                    WriteTomfanLog($"Nội dung dòng: {string.Join(", ", parts)}");

                                    var m = new Measure
                                    {
                                        k = expectedK,
                                        S = float.Parse(parts[1], CultureInfo.InvariantCulture),
                                        CV = float.Parse(parts[2], CultureInfo.InvariantCulture)
                                    };


                                    // 15 part gồm 3 parts đầu là hiển thị tín hiệu (100) - %S - %CV
                                    // 12 parts còn lại tương ứng với tín hiệu trả về của 12 cảm biến
                                    if (!isConnection && parts.Length > 10)
                                    {
                                        // tần số tính từ %S
                                        m.TanSo_fb = _sensor.IsImportPhanHoiTanSo
                                                    ? _sensor.PhanHoiTanSoValue
                                                    : CalcSimatic(
                                                        minValue: _sensor.PhanHoiTanSoMin,
                                                        maxValue: _sensor.PhanHoiTanSoMax,
                                                        percent: float.Parse(parts[1], CultureInfo.InvariantCulture),
                                                        sensorIdx: 1,
                                                        indexK: m.k);

                                        // 1. nhiệt độ môi trường
                                        m.NhietDoMoiTruong_sen = _sensor.IsImportNhietDoMoiTruong
                                            ? _sensor.NhietDoMoiTruongValue
                                            : CalcSimatic(
                                                minValue: _sensor.NhietDoMoiTruongMin,
                                                maxValue: _sensor.NhietDoMoiTruongMax,
                                                percent: float.Parse(parts[3], CultureInfo.InvariantCulture) - avgs[1],
                                                sensorIdx: 3,
                                                indexK: m.k);

                                        // 2. độ ẩm
                                        m.DoAm_sen = _sensor.IsImportDoAmMoiTruong
                                            ? _sensor.DoAmMoiTruongValue
                                            : CalcSimatic(
                                                minValue: _sensor.DoAmMoiTruongMin,
                                                maxValue: _sensor.DoAmMoiTruongMax,
                                                percent: float.Parse(parts[4], CultureInfo.InvariantCulture) - avgs[2],
                                                sensorIdx: 4,
                                                indexK: m.k);

                                        // 3. phản hồi vị trí van
                                        m.ViTriVan_fb = _sensor.IsImportPhanHoiViTriVan
                                            ? _sensor.PhanHoiViTriVanValue
                                            : CalcSimatic(
                                                minValue: _sensor.PhanHoiViTriVanMin,
                                                maxValue: _sensor.PhanHoiViTriVanMax,
                                                percent: float.Parse(parts[5], CultureInfo.InvariantCulture) - avgs[3],
                                                sensorIdx: 5,
                                                indexK: m.k);

                                        // 4. Momen
                                        m.Momen_sen = _sensor.IsImportMomen
                                            ? _sensor.MomenValue
                                            : CalcSimatic(
                                                minValue: _sensor.MomenMin,
                                                maxValue: _sensor.MomenMax,
                                                percent: float.Parse(parts[6], CultureInfo.InvariantCulture) - avgs[4],
                                                sensorIdx: 6,
                                                indexK: m.k);

                                        // 5. nhiệt độ hồng ngoại
                                        m.NhietDoGoi = _sensor.IsImportNhietDoGoiTruc
                                            ? _sensor.NhietDoGoiTrucValue
                                            : CalcSimatic(
                                                minValue: _sensor.NhietDoGoiTrucMin,
                                                maxValue: _sensor.NhietDoGoiTrucMax,
                                                percent: float.Parse(parts[7], CultureInfo.InvariantCulture) - avgs[5],
                                                sensorIdx: 7,
                                                indexK: m.k);

                                        // 6. độ rung
                                        m.DoRung_sen = _sensor.IsImportDoRung
                                            ? _sensor.DoRungValue
                                            : CalcSimatic(
                                                minValue: _sensor.DoRungMin,
                                                maxValue: _sensor.DoRungMax,
                                                percent: float.Parse(parts[8], CultureInfo.InvariantCulture) - avgs[6],
                                                sensorIdx: 8,
                                                indexK: m.k);

                                        // 7. số vòng quay
                                        // nếu là thông số nhập tay, lấy số vòng quay value (là số định mức của động cơ tương đương tần số cao nhất đê tính ra số vòng quay ở tần số hiện tại) 
                                        m.SoVongQuay_sen = _sensor.IsImportSoVongQuay
                                            ? (_sensor.SoVongQuayValue * m.S / _input.TanSoDongCoTheoThietKe)
                                            : CalcSimatic(
                                                minValue: _sensor.SoVongQuayMin,
                                                maxValue: _sensor.SoVongQuayMax,
                                                percent: float.Parse(parts[9], CultureInfo.InvariantCulture) - avgs[7],
                                                sensorIdx: 9,
                                                indexK: m.k);

                                        // 8. phản hồi dòng điện
                                        // phía plc chỉ trả về 50% nên nhân với hệ số 2
                                        m.DongDien_fb = _sensor.IsImportPhanHoiDongDien
                                            ? _sensor.PhanHoiDongDienValue
                                            : CalcSimatic(
                                                minValue: _sensor.PhanHoiDongDienMin,
                                                maxValue: _sensor.PhanHoiDongDienMax,
                                                //percent: (float.Parse(parts[10], CultureInfo.InvariantCulture) - avgs[8]) * 2,
                                                percent: (float.Parse(parts[10], CultureInfo.InvariantCulture) - avgs[8]),
                                                sensorIdx: 10,
                                                indexK: m.k);

                                        // 9. áp suất tĩnh
                                        m.ApSuatTinh_sen = _sensor.IsImportApSuatTinh
                                            ? _sensor.ApSuatTinhValue
                                            : CalcSimatic(
                                                minValue: _sensor.ApSuatTinhMin,
                                                maxValue: _sensor.ApSuatTinhMax,
                                                percent: float.Parse(parts[11], CultureInfo.InvariantCulture) - avgs[9],
                                                sensorIdx: 11,
                                                indexK: m.k);

                                        // 10. phản hồi công suất
                                        m.CongSuat_fb = _sensor.IsImportPhanHoiCongSuat
                                            ? _sensor.PhanHoiCongSuatValue
                                            : CalcSimatic(
                                                minValue: _sensor.PhanHoiCongSuatMin,
                                                maxValue: _sensor.PhanHoiCongSuatMax,
                                                percent: float.Parse(parts[12], CultureInfo.InvariantCulture) - avgs[10],
                                                sensorIdx: 12,
                                                indexK: m.k);

                                        // 11. chênh lệch áp suất
                                        m.ChenhLechApSuat_sen = _sensor.IsImportChenhLechApSuat
                                            ? _sensor.ChenhLechApSuatValue
                                            : CalcSimatic(
                                                minValue: _sensor.ChenhLechApSuatMin,
                                                maxValue: _sensor.ChenhLechApSuatMax,
                                                percent: float.Parse(parts[13], CultureInfo.InvariantCulture) - avgs[11],
                                                sensorIdx: 13,
                                                indexK: m.k);

                                        // 12. áp suất khí quyển
                                        m.ApSuatkhiQuyen_sen = _sensor.IsImportApSuatKhiQuyen
                                            ? _sensor.ApSuatKhiQuyenValue
                                            : CalcSimatic(
                                                minValue: _sensor.ApSuatKhiQuyenMin,
                                                maxValue: _sensor.ApSuatKhiQuyenMax,
                                                percent: float.Parse(parts[14], CultureInfo.InvariantCulture) - avgs[12],
                                                sensorIdx: 14,
                                                indexK: m.k);

                                        // điện áp luôn lấy theo giá trị nhập vào
                                        m.DienAp_fb = _sensor.IsImportPhanHoiDienAp
                                            ? _sensor.PhanHoiDienApValue
                                            : _sensor.PhanHoiDienApValue;

                                        m.DoOn_sen = _sensor.IsImportDoOn
                                            ? _sensor.DoOnValue
                                            : _sensor.DoOnValue;

                                        WriteTomfanLog("Hoàn thành tính toán từ file 2_S_IN.csv");
                                        WriteTomfanLog($"Nhiệt độ môi trường: {m.NhietDoMoiTruong_sen}");
                                        WriteTomfanLog($"Độ ẩm: {m.DoAm_sen}");
                                        WriteTomfanLog($"Áp suất khí quyển: {m.ApSuatkhiQuyen_sen}");
                                        WriteTomfanLog($"Chênh lệch áp suất: {m.ChenhLechApSuat_sen}");
                                        WriteTomfanLog($"Áp suất tĩnh: {m.ApSuatTinh_sen}");
                                        WriteTomfanLog($"Độ rung: {m.DoRung_sen}");
                                        WriteTomfanLog($"Độ ồn: {m.DoOn_sen}");
                                        WriteTomfanLog($"Số vòng quay: {m.SoVongQuay_sen}");
                                        WriteTomfanLog($"Momen: {m.Momen_sen}");
                                        WriteTomfanLog($"Dòng điện phản hồi: {m.DongDien_fb}");
                                        WriteTomfanLog($"Công suất phản hồi: {m.CongSuat_fb}");
                                        WriteTomfanLog($"Vị trí van phản hồi: {m.ViTriVan_fb}");
                                        WriteTomfanLog($"Tần số phản hồi: {m.TanSo_fb}");
                                        WriteTomfanLog($"Nhiệt độ gối trục: {m.NhietDoGoi}");
                                    }

                                    if (expectedK >= 2)
                                    {
                                        try
                                        {
                                            string tempPath = Path.Combine(_exchangeFolder, "MeasurementSummary.xlsx");
                                            FileInfo fileInfo = new FileInfo(tempPath);

                                            using (var package = new ExcelPackage(fileInfo))
                                            {
                                                var ws = package.Workbook.Worksheets.FirstOrDefault();
                                                if (ws == null)
                                                {
                                                    ws = package.Workbook.Worksheets.Add("MeasureData");
                                                }

                                                if (expectedK == 2)
                                                {
                                                    ws.Cells[1, 1].Value = "k";
                                                    ws.Cells[1, 2].Value = "Tần số (%)";
                                                    ws.Cells[1, 3].Value = "Góc mở van (%)";

                                                    ws.Cells[1, 4].Value = "Nhiệt độ MT (%)";
                                                    ws.Cells[1, 5].Value = "Độ ẩm (%)";
                                                    ws.Cells[1, 6].Value = "Vị trí van (%)";
                                                    ws.Cells[1, 7].Value = "Momen (%)";
                                                    ws.Cells[1, 8].Value = "Hồng ngoại (%)";
                                                    ws.Cells[1, 9].Value = "Độ rung (%)";
                                                    ws.Cells[1, 10].Value = "Số vòng quay (%)";
                                                    ws.Cells[1, 11].Value = "Dòng điện (%)";
                                                    ws.Cells[1, 12].Value = "Áp suất tĩnh (%)";
                                                    ws.Cells[1, 13].Value = "Công suất (%)";
                                                    ws.Cells[1, 14].Value = "Chênh áp (%)";
                                                    ws.Cells[1, 15].Value = "Áp suất khí quyển (%)";

                                                    ws.Cells[1, 16].Value = "Nhiệt độ MT (oC)";
                                                    ws.Cells[1, 17].Value = "Độ ẩm (%)";
                                                    ws.Cells[1, 18].Value = "Vị trí van (%)";
                                                    ws.Cells[1, 19].Value = "Momen";
                                                    ws.Cells[1, 20].Value = "Nhiệt độ gối (oC)";
                                                    ws.Cells[1, 21].Value = "Độ rung";
                                                    ws.Cells[1, 22].Value = "Số vòng quay (RPM)";
                                                    ws.Cells[1, 23].Value = "Dòng điện (A)";
                                                    ws.Cells[1, 24].Value = "Áp suất tĩnh (Pa)";
                                                    ws.Cells[1, 25].Value = "Công suất (kW)";
                                                    ws.Cells[1, 26].Value = "Chênh lệch áp (Pa)";
                                                    ws.Cells[1, 27].Value = "Áp suất khí quyển (Pa)";

                                                    ws.Cells[1, 28].Value = "Tần số (Hz)";
                                                    ws.Cells[1, 29].Value = "Điện áp (V)";
                                                    ws.Cells[1, 30].Value = "Độ ồn (dB)";
                                                    ws.Cells.AutoFitColumns();

                                                    // Style cho header
                                                    using (var range = ws.Cells[1, 1, 1, 30])
                                                    {
                                                        range.Style.Font.Bold = true;
                                                        range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                                        range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                                                        range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                                    }

                                                    var ws1 = package.Workbook.Worksheets.FirstOrDefault(ws => ws.Name == "Sensor");
                                                    if (ws1 == null)
                                                    {
                                                        ws1 = package.Workbook.Worksheets.Add("Sensor");
                                                    }

                                                    // Header cho worksheet cấu hình
                                                    ws1.Cells[1, 1].Value = "Tín hiệu";
                                                    ws1.Cells[1, 2].Value = "Nhập tay (Import)";
                                                    ws1.Cells[1, 3].Value = "Giá trị (dùng khi import)";
                                                    ws1.Cells[1, 4].Value = "Min";
                                                    ws1.Cells[1, 5].Value = "Max";

                                                    // Style header
                                                    using (var range = ws1.Cells[1, 1, 1, 5])
                                                    {
                                                        range.Style.Font.Bold = true;
                                                        range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                                        range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                                                        range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                                    }

                                                    // Danh sách cấu hình cảm biến theo thứ tự
                                                    int row = 2;

                                                    // 1. Nhiệt độ môi trường
                                                    ws1.Cells[row, 1].Value = "Nhiệt độ môi trường";
                                                    ws1.Cells[row, 2].Value = _sensor.IsImportNhietDoMoiTruong ? "TRUE" : "FALSE";
                                                    ws1.Cells[row, 3].Value = _sensor.NhietDoMoiTruongValue;
                                                    ws1.Cells[row, 4].Value = _sensor.NhietDoMoiTruongMin;
                                                    ws1.Cells[row, 5].Value = _sensor.NhietDoMoiTruongMax;
                                                    row++;

                                                    // 2. Độ ẩm môi trường
                                                    ws1.Cells[row, 1].Value = "Độ ẩm môi trường";
                                                    ws1.Cells[row, 2].Value = _sensor.IsImportDoAmMoiTruong ? "TRUE" : "FALSE";
                                                    ws1.Cells[row, 3].Value = _sensor.DoAmMoiTruongValue;
                                                    ws1.Cells[row, 4].Value = _sensor.DoAmMoiTruongMin;
                                                    ws1.Cells[row, 5].Value = _sensor.DoAmMoiTruongMax;
                                                    row++;

                                                    // 3. Phản hồi vị trí van
                                                    ws1.Cells[row, 1].Value = "Phản hồi vị trí van";
                                                    ws1.Cells[row, 2].Value = _sensor.IsImportPhanHoiViTriVan ? "TRUE" : "FALSE";
                                                    ws1.Cells[row, 3].Value = _sensor.PhanHoiViTriVanValue;
                                                    ws1.Cells[row, 4].Value = _sensor.PhanHoiViTriVanMin;
                                                    ws1.Cells[row, 5].Value = _sensor.PhanHoiViTriVanMax;
                                                    row++;

                                                    // 4. Momen
                                                    ws1.Cells[row, 1].Value = "Momen";
                                                    ws1.Cells[row, 2].Value = _sensor.IsImportMomen ? "TRUE" : "FALSE";
                                                    ws1.Cells[row, 3].Value = _sensor.MomenValue;
                                                    ws1.Cells[row, 4].Value = _sensor.MomenMin;
                                                    ws1.Cells[row, 5].Value = _sensor.MomenMax;
                                                    row++;



                                                    // 5. Nhiệt độ gối trục
                                                    ws1.Cells[row, 1].Value = "Nhiệt độ gối trục";
                                                    ws1.Cells[row, 2].Value = _sensor.IsImportNhietDoGoiTruc ? "TRUE" : "FALSE";
                                                    ws1.Cells[row, 3].Value = _sensor.NhietDoGoiTrucValue;
                                                    ws1.Cells[row, 4].Value = _sensor.NhietDoGoiTrucMin;
                                                    ws1.Cells[row, 5].Value = _sensor.NhietDoGoiTrucMax;
                                                    row++;

                                                    // 6. Độ rung
                                                    ws1.Cells[row, 1].Value = "Độ rung";
                                                    ws1.Cells[row, 2].Value = _sensor.IsImportDoRung ? "TRUE" : "FALSE";
                                                    ws1.Cells[row, 3].Value = _sensor.DoRungValue;
                                                    ws1.Cells[row, 4].Value = _sensor.DoRungMin;
                                                    ws1.Cells[row, 5].Value = _sensor.DoRungMax;
                                                    row++;

                                                    // 7. Số vòng quay
                                                    ws1.Cells[row, 1].Value = "Số vòng quay";
                                                    ws1.Cells[row, 2].Value = _sensor.IsImportSoVongQuay ? "TRUE" : "FALSE";
                                                    ws1.Cells[row, 3].Value = _sensor.SoVongQuayValue;
                                                    ws1.Cells[row, 4].Value = _sensor.SoVongQuayMin;
                                                    ws1.Cells[row, 5].Value = _sensor.SoVongQuayMax;
                                                    row++;

                                                    // 8. Phản hồi dòng điện
                                                    ws1.Cells[row, 1].Value = "Phản hồi dòng điện";
                                                    ws1.Cells[row, 2].Value = _sensor.IsImportPhanHoiDongDien ? "TRUE" : "FALSE";
                                                    ws1.Cells[row, 3].Value = _sensor.PhanHoiDongDienValue;
                                                    ws1.Cells[row, 4].Value = _sensor.PhanHoiDongDienMin;
                                                    ws1.Cells[row, 5].Value = _sensor.PhanHoiDongDienMax;
                                                    row++;

                                                    // 9. Chênh lệch áp suất
                                                    ws1.Cells[row, 1].Value = "Chênh lệch áp suất";
                                                    ws1.Cells[row, 2].Value = _sensor.IsImportChenhLechApSuat ? "TRUE" : "FALSE";
                                                    ws1.Cells[row, 3].Value = _sensor.ChenhLechApSuatValue;
                                                    ws1.Cells[row, 4].Value = _sensor.ChenhLechApSuatMin;
                                                    ws1.Cells[row, 5].Value = _sensor.ChenhLechApSuatMax;
                                                    row++;

                                                    // 10. Phản hồi công suất
                                                    ws1.Cells[row, 1].Value = "Phản hồi công suất";
                                                    ws1.Cells[row, 2].Value = _sensor.IsImportPhanHoiCongSuat ? "TRUE" : "FALSE";
                                                    ws1.Cells[row, 3].Value = _sensor.PhanHoiCongSuatValue;
                                                    ws1.Cells[row, 4].Value = _sensor.PhanHoiCongSuatMin;
                                                    ws1.Cells[row, 5].Value = _sensor.PhanHoiCongSuatMax;
                                                    row++;

                                                    // 11. Áp suất tĩnh
                                                    ws1.Cells[row, 1].Value = "Áp suất tĩnh";
                                                    ws1.Cells[row, 2].Value = _sensor.IsImportApSuatTinh ? "TRUE" : "FALSE";
                                                    ws1.Cells[row, 3].Value = _sensor.ApSuatTinhValue;
                                                    ws1.Cells[row, 4].Value = _sensor.ApSuatTinhMin;
                                                    ws1.Cells[row, 5].Value = _sensor.ApSuatTinhMax;
                                                    row++;

                                                    // 12. Áp suất khí quyển
                                                    ws1.Cells[row, 1].Value = "Áp suất khí quyển";
                                                    ws1.Cells[row, 2].Value = _sensor.IsImportApSuatKhiQuyen ? "TRUE" : "FALSE";
                                                    ws1.Cells[row, 3].Value = _sensor.ApSuatKhiQuyenValue;
                                                    ws1.Cells[row, 4].Value = _sensor.ApSuatKhiQuyenMin;
                                                    ws1.Cells[row, 5].Value = _sensor.ApSuatKhiQuyenMax;




                                                    // 13. Phản hồi tần số
                                                    ws1.Cells[row, 1].Value = "Phản hồi tần số";
                                                    ws1.Cells[row, 2].Value = _sensor.IsImportPhanHoiTanSo ? "TRUE" : "FALSE";
                                                    ws1.Cells[row, 3].Value = _sensor.PhanHoiTanSoValue;
                                                    ws1.Cells[row, 4].Value = _sensor.PhanHoiTanSoMin;
                                                    ws1.Cells[row, 5].Value = _sensor.PhanHoiTanSoMax;
                                                    row++;

                                                    // 14. Phản hồi điện áp
                                                    ws1.Cells[row, 1].Value = "Phản hồi điện áp";
                                                    ws1.Cells[row, 2].Value = _sensor.IsImportPhanHoiDienAp ? "TRUE" : "FALSE";
                                                    ws1.Cells[row, 3].Value = _sensor.PhanHoiDienApValue;
                                                    ws1.Cells[row, 4].Value = _sensor.PhanHoiDienApMin;
                                                    ws1.Cells[row, 5].Value = _sensor.PhanHoiDienApMax;
                                                    row++;

                                                    // 15. Độ ồn
                                                    ws1.Cells[row, 1].Value = "Độ ồn";
                                                    ws1.Cells[row, 2].Value = _sensor.IsImportDoOn ? "TRUE" : "FALSE";
                                                    ws1.Cells[row, 3].Value = _sensor.DoOnValue;
                                                    ws1.Cells[row, 4].Value = _sensor.DoOnMin;
                                                    ws1.Cells[row, 5].Value = _sensor.DoOnMax;
                                                    row++;

                                                    ws1.Cells.AutoFitColumns();

                                                }

                                                if (!isConnection && parts.Length > 10)
                                                {
                                                    int dataRow = expectedK - 1;

                                                    ws.Cells[dataRow, 1].Value = m.k;
                                                    ws.Cells[dataRow, 2].Value = m.S;
                                                    ws.Cells[dataRow, 3].Value = m.CV;

                                                    if (parts.Length > 3)
                                                    {
                                                        for (int i = 3; i < Math.Min(parts.Length, 15); i++)
                                                        {
                                                            if (float.TryParse(parts[i], NumberStyles.Float, CultureInfo.InvariantCulture, out float value))
                                                            {
                                                                ws.Cells[dataRow, i + 1].Value = value;
                                                            }
                                                        }
                                                    }

                                                    ws.Cells[dataRow, 16].Value = m.NhietDoMoiTruong_sen;
                                                    ws.Cells[dataRow, 17].Value = m.DoAm_sen;
                                                    ws.Cells[dataRow, 18].Value = m.ViTriVan_fb;
                                                    ws.Cells[dataRow, 19].Value = m.Momen_sen;
                                                    ws.Cells[dataRow, 20].Value = m.NhietDoGoi;
                                                    ws.Cells[dataRow, 21].Value = m.DoRung_sen;
                                                    ws.Cells[dataRow, 22].Value = m.SoVongQuay_sen;
                                                    ws.Cells[dataRow, 23].Value = m.DongDien_fb;
                                                    ws.Cells[dataRow, 24].Value = m.ApSuatTinh_sen;
                                                    ws.Cells[dataRow, 25].Value = m.CongSuat_fb;
                                                    ws.Cells[dataRow, 26].Value = m.ChenhLechApSuat_sen;
                                                    ws.Cells[dataRow, 27].Value = m.ApSuatkhiQuyen_sen;

                                                    ws.Cells[dataRow, 28].Value = m.TanSo_fb;
                                                    ws.Cells[dataRow, 29].Value = m.DienAp_fb;
                                                    ws.Cells[dataRow, 30].Value = m.DoOn_sen;

                                                }
                                                ws.Cells.AutoFitColumns();
                                                package.Save();

                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                        }
                                    }

                                    return m;
                                }

                            }
                        }
                    }
                }
                catch (IOException)
                {
                }
                await Task.Delay(200);
            }

            WriteTomfanLog($"Không nhận được phản hồi cho k={expectedK} sau {UserSetting.Instance.TimeoutMilliseconds}ms ----- TIMEOUT");
            return null;
        }


        public void WriteTomfanLog(string message)
        {
            try
            {
                string logPath = Path.Combine(_exchangeFolder, "tomfan_log.txt");
                string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} : {message}{Environment.NewLine}";

                // Lock để đảm bảo an toàn đa luồng
                lock (this)
                {
                    File.AppendAllText(logPath, logEntry);
                }
            }
            catch
            {
                // Tránh treo app vì lỗi ghi log
            }
        }

        public event Action<bool> OnSimaticConnectionChanged;
        public event Action<Measure> OnSimaticResultReceived;
        public event Action<List<Measure>> OnSimaticExchangeCompleted;
        public event Action<MeasureResponse, ParameterShow> OnMeasurePointCompleted;
        public event Action<MeasureFittingFC, Measure> OnMeasureRangeCompleted;
    }
}
