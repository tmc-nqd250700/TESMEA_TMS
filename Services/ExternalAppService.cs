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
        Task<bool> ConnectExchangeAsync(List<Measure> measures, BienTan inv, CamBien sensor, OngGio duct, ThongTinMauThuNghiem input, float maxmin, float timeRange);
        Task<bool> ConnectExchangeAsync(Measure measure, float timeRange);
        Task StartExchangeAsync();
        Task StopExchangeAsync();
        void WriteTomfanLog(string message);
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
        public async Task<bool> ConnectExchangeAsync(List<Measure> measures, BienTan inv, CamBien sensor, OngGio duct, ThongTinMauThuNghiem input, float maxmin, float timeRange)
        {
            try
            {
                WriteTomfanLog("========== BẮT ĐẦU QUY TRÌNH KẾT NỐI (CONNECT) ==========");
                _measures = measures;
                _inv = inv;
                _sensor = sensor;
                _duct = duct;
                _input = input;

                if (!Directory.Exists(_exchangeFolder))
                {
                    WriteTomfanLog($"[FATAL] Thư mục trao đổi không tồn tại: {_exchangeFolder}");
                    throw new BusinessException("Thư mục trao đổi dữ liệu với Simatic không tồn tại");
                }

                _simaticResults.Clear();
                WriteTomfanLog($"Step: Kiểm tra kết nối qua dòng đầu tiên.");
                var m = measures.First();
                WriteTomfanLog($"Connect - Thử dòng k={m.k}: Ghi file và chờ WinCC phản hồi...");
                await WriteDataToFilesAsync(m, timeRange);

                var result = await WaitForResultAsync(m.k, isConnection: true);
                if (result == null || Math.Abs(result.S - m.S) > 0.01)
                {
                    string errorMsg = result == null ? "Timeout" : $"Dữ liệu không khớp (S_gửi={m.S}, S_nhận={result.S})";
                    WriteTomfanLog($"Kết nối thất bại tại dòng k={m.k}: {errorMsg}");
                    throw new Exception($"Không thể kết nối. Dòng {m.k} lỗi: {errorMsg}");
                }

                m.F = MeasureStatus.Completed;
                WriteTomfanLog($"Connect - Dòng k={m.k} THÀNH CÔNG");
                _currentIndex = m.k;
                return true;
            }
            catch (Exception ex)
            {
                WriteTomfanLog($"Lỗi trong ConnectExchangeAsync: {ex.Message}");
                throw;
            }
        }

        // kết nối dòng 2
        public async Task<bool> ConnectExchangeAsync(Measure measure, float maxmin)
        {
            try
            {
                WriteTomfanLog("========== BẮT ĐẦU QUY TRÌNH KẾT NỐI ==========");
                if (!Directory.Exists(_exchangeFolder))
                {
                    WriteTomfanLog($"Thư mục trao đổi không tồn tại: {_exchangeFolder}");
                    throw new BusinessException("Thư mục trao đổi dữ liệu với Simatic không tồn tại");
                }

                WriteTomfanLog($"Connect - Thử dòng k={measure.k}: Ghi file và chờ WinCC phản hồi...");
                await WriteDataToFilesAsync(measure, maxmin);

                var result = await WaitForResultAsync(measure.k, isConnection: true);
                if (result == null || Math.Abs(result.S - measure.S) > 0.01)
                {
                    string errorMsg = result == null ? "Timeout (Không có phản hồi)" : $"Dữ liệu không khớp (S_gửi={measure.S}, S_nhận={result.S})";
                    WriteTomfanLog($"Kết nối thất bại tại dòng k={measure.k}: {errorMsg}");
                    throw new Exception($"Không thể kết nối. Dòng {measure.k} lỗi: {errorMsg}");
                }

                measure.F = MeasureStatus.Completed;
                WriteTomfanLog($"Connect - Dòng k={measure.k} THÀNH CÔNG. Nghỉ 5s chờ WinCC sẵn sàng...");
                DataProcess.Initialize(_measures.Count);
                IsConnectedToSimatic = true;
                OnSimaticConnectionChanged?.Invoke(true);
                _currentIndex = measure.k;
                WriteTomfanLog("Đã thiết lập kết nối với Simatic thành công.");
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
                    WriteTomfanLog("StartExchange bị từ chối: Chưa kết nối Simatic.");
                    MessageBoxHelper.ShowWarning("Chưa kết nối với Simatic.");
                    return;
                }

                WriteTomfanLog("========== BẮT ĐẦU TRAO ĐỔI DỮ LIỆU HÀNG LOẠT ==========");

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
                    // tạo file trend.csv theo k
                    using (var fs = File.Create(Path.Combine(_trendFolder, $"{m.k}.csv"))) { }
                    await WriteDataToFilesAsync(m);
                    if (_currentIndex > 3)
                    {
                        // delay 30s den khi ghi dong tiep theo
                        WriteTomfanLog("Delay 15s sau đó chờ kết quả dòng tiếp theo");
                        await Task.Delay(15000);
                        WriteTomfanLog("Delay xong, tiếp tục lắng nghe dòng tiếp theo");
                    }
                    // Chờ kết quả xử lý thực tế (isConnection = false để tính toán sensor)
                    // Luôn lắng nghe dòng 3
                    var result = await WaitForResultAsync(m.k, isConnection: false);

                    if (result != null)
                    {
                        WriteTomfanLog($"Đã nhận kết quả k={m.k}");
                        m.F = MeasureStatus.Completed;
                        _simaticResults.Add(result);

                        var measurePoint = DataProcess.OnePointMeasure(result, _inv, _sensor, _duct, _input);
                        OnMeasurePointCompleted?.Invoke(measurePoint, DataProcess.ParaShow(result, _inv, _sensor, _duct, _input));
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
                        WriteTomfanLog($"Hoàn tất điểm đo k={m.k}");
                    }
                    else
                    {
                        WriteTomfanLog($"Điểm đo k={m.k} thất bại");
                        m.F = MeasureStatus.Error;
                        OnSimaticResultReceived?.Invoke(m);
                    }
                    // delay 15s den khi ghi dong tiep theo
                    WriteTomfanLog("Delay 15s trước khi ghi dòng tiếp theo");
                    await Task.Delay(15000);
                    WriteTomfanLog("Delay xong, tiếp tục ghi dữ liệu dòng tiếp theo");
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
                WriteTomfanLog($"Đã chèn lệnh E-Stop (96) vào dòng mới k={nextIndex}");
            }
            catch (Exception ex)
            {
                WriteTomfanLog($"Lỗi trong StopExchangeAsync: {ex.Message}");
            }
        }

        // tính giá trị trả về từ %
        private float CalcSimatic(float minValue, float maxValue, float percent)
        {
            return minValue + (maxValue - minValue) * percent / 100f;
        }

        // retry n lần nếu file bị khóa
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
                        WriteTomfanLog("Đã thử lại tối đa nhưng vẫn không thể truy cập file.");
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
                        throw new BusinessException("File 1_T_OUT.csv không tồn tại.");
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
                    WriteTomfanLog($"Step: CSV row {rowIdx} ghi thành công.");
                });
            }
            catch (Exception ex)
            {
                WriteTomfanLog($"Error WriteDataToFilesAsync: {ex}");
            }
        }

        // lấy dữ liệu từ trendline file để tính toán
        public Measure CalculateTrendData(float S, float CV, int index)
        {
            try
            {
                Measure m = new Measure
                {
                    k = index,
                    S = S,
                    CV = CV
                };
                string trendFilePath = Path.Combine(_trendFolder, $"{index}.csv");
                if (!File.Exists(trendFilePath))
                {
                    throw new FileNotFoundException($"File trend không tồn tại: {trendFilePath}");
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
                            ViTriVan_fb = float.TryParse(values[3], NumberStyles.Float, CultureInfo.InvariantCulture, out var v12) ? v12 : 0,
                            DoOn_sen = float.TryParse(values[4], NumberStyles.Float, CultureInfo.InvariantCulture, out var v7) ? v7 : 0,
                            Momen_sen = float.TryParse(values[4], NumberStyles.Float, CultureInfo.InvariantCulture, out var v9) ? v9 : 0,
                            DoRung_sen = float.TryParse(values[6], NumberStyles.Float, CultureInfo.InvariantCulture, out var v6) ? v6 : 0,
                            SoVongQuay_sen = float.TryParse(values[7], NumberStyles.Float, CultureInfo.InvariantCulture, out var v8) ? v8 : 0,
                            CongSuat_fb = float.TryParse(values[8], NumberStyles.Float, CultureInfo.InvariantCulture, out var v11) ? v11 : 0,
                            ChenhLechApSuat_sen = float.TryParse(values[9], NumberStyles.Float, CultureInfo.InvariantCulture, out var v4) ? v4 : 0,
                            DongDien_fb = float.TryParse(values[10], NumberStyles.Float, CultureInfo.InvariantCulture, out var v10) ? v10 : 0,
                            ApSuatTinh_sen = float.TryParse(values[11], NumberStyles.Float, CultureInfo.InvariantCulture, out var v5) ? v5 : 0,
                            ApSuatkhiQuyen_sen = float.TryParse(values[12], NumberStyles.Float, CultureInfo.InvariantCulture, out var v3) ? v3 : 0,
                        };
                        trendTimes.Add(trend);
                    }
                }

                if (trendTimes.Count <= 20)
                {
                    WriteTomfanLog($"Trendtime không đủ 20 giá trị để chốt dữ liệu. Tổng {trendTimes.Count}");
                    return null;
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

                // tần số tính từ %S
                m.TanSo_fb = _sensor.IsImportPhanHoiTanSo
                            ? _sensor.PhanHoiTanSoValue
                            : CalcSimatic(_sensor.PhanHoiTanSoMin, _sensor.PhanHoiTanSoMax, S);

                // 1. nhiệt độ môi trường
                m.NhietDoMoiTruong_sen = _sensor.IsImportNhietDoMoiTruong
                    ? _sensor.NhietDoMoiTruongValue
                    : CalcSimatic(_sensor.NhietDoMoiTruongMin, _sensor.NhietDoMoiTruongMax, GetContinuousAverage(x=>x.NhietDoMoiTruong_sen, "Nhiệt độ môi trường"));

                // 2. độ ẩm
                m.DoAm_sen = _sensor.IsImportDoAmMoiTruong
                    ? _sensor.DoAmMoiTruongValue
                    : CalcSimatic(_sensor.DoAmMoiTruongMin, _sensor.DoAmMoiTruongMax, GetContinuousAverage(x => x.DoAm_sen, "Độ ẩm"));

                // 3. phản hồi vị trí van
                m.ViTriVan_fb = _sensor.IsImportPhanHoiViTriVan
                    ? _sensor.PhanHoiViTriVanValue
                    : CalcSimatic(_sensor.PhanHoiViTriVanMin, _sensor.PhanHoiViTriVanMax, GetContinuousAverage(x => x.ViTriVan_fb, "Vị trí van"));

                // 4. chưa có
                m.DoOn_sen = _sensor.IsImportDoOn
                    ? _sensor.DoOnValue
                    : CalcSimatic(_sensor.DoOnMin, _sensor.DoOnMax, GetContinuousAverage(x => x.ViTriVan_fb, "Độ ồn"));

                m.Momen_sen = _sensor.IsImportMomen
                    ? _sensor.MomenValue
                    : CalcSimatic(_sensor.MomenMin, _sensor.MomenMax, GetContinuousAverage(x => x.ViTriVan_fb, "momen"));

                // điện áp luôn lấy theo giá trị nhập vào

                m.DienAp_fb = _sensor.IsImportPhanHoiDienAp
                    ? _sensor.PhanHoiDienApValue
                    : _sensor.PhanHoiDienApValue;
                //CalcSimatic(_sensor.PhanHoiDienApMin, _sensor.PhanHoiDienApMax, float.Parse(parts[6], CultureInfo.InvariantCulture));
                // 5. nhiệt độ hồng ngoại

                // 6. độ rung
                m.DoRung_sen = _sensor.IsImportDoRung
                    ? _sensor.DoRungValue
                    : CalcSimatic(_sensor.DoRungMin, _sensor.DoRungMax, GetContinuousAverage(x => x.DoRung_sen, "Độ rung"));

                //7. số vòng quay
                m.SoVongQuay_sen = _sensor.IsImportSoVongQuay
                    ? _sensor.SoVongQuayValue
                    : CalcSimatic(_sensor.SoVongQuayMin, _sensor.SoVongQuayMax, GetContinuousAverage(x => x.SoVongQuay_sen, "Số vòng quay"));

                // 8. phản hồi công suất
                m.CongSuat_fb = _sensor.IsImportPhanHoiCongSuat
                  ? _sensor.PhanHoiCongSuatValue
                  : CalcSimatic(_sensor.PhanHoiCongSuatMin, _sensor.PhanHoiCongSuatMax, GetContinuousAverage(x => x.CongSuat_fb, "Công suất"));

                // 9. chênh lệch áp suất
                m.ChenhLechApSuat_sen = _sensor.IsImportChenhLechApSuat
                   ? _sensor.ChenhLechApSuatValue
                   : CalcSimatic(_sensor.ChenhLechApSuatMin, _sensor.ChenhLechApSuatMax, GetContinuousAverage(x => x.ChenhLechApSuat_sen, "Chênh lệch áp suất"));

                // 10. phản hồi dòng điện
                m.DongDien_fb = _sensor.IsImportPhanHoiDongDien
                  ? _sensor.PhanHoiDongDienValue
                  : CalcSimatic(_sensor.PhanHoiDongDienMin, _sensor.PhanHoiDongDienMax, GetContinuousAverage(x => x.DongDien_fb, "Dòng điện"));
                //: _sensor.PhanHoiDienApValue;


                // 11. áp suất tĩnh
                m.ApSuatTinh_sen = _sensor.IsImportApSuatTinh
                    ? _sensor.ApSuatTinhValue
                    : CalcSimatic(_sensor.ApSuatTinhMin, _sensor.ApSuatTinhMax, GetContinuousAverage(x => x.ApSuatTinh_sen, "Áp suất tĩnh"));

                // 12. áp suất khí quyển
                m.ApSuatkhiQuyen_sen = _sensor.IsImportApSuatKhiQuyen
                    ? _sensor.ApSuatKhiQuyenValue
                    : CalcSimatic(_sensor.ApSuatKhiQuyenMin, _sensor.ApSuatKhiQuyenMax, GetContinuousAverage(x => x.ApSuatkhiQuyen_sen, "Áp suất khí quyển"));

                WriteTomfanLog("Hoàn thành tính toán kết quả từ trendline");
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
                return m;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        // chờ kết quả từ file 2.csv
        private async Task<Measure?> WaitForResultAsync(int expectedK, bool isConnection)
        {
            string path2 = Path.Combine(_exchangeFolder, "2_S_IN.csv");
            var sw = Stopwatch.StartNew();
            char sep = _isComma ? ' ' : ';';
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
                            int targetIndex = isConnection ? expectedK - 1 : 2;
                            if (lines.Length > targetIndex)
                            {
                                string targetLine = lines[targetIndex];

                                // Kiểm tra nếu dòng có dữ liệu
                                if (!string.IsNullOrWhiteSpace(targetLine))
                                {
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
                                        // check nếu parts có giá trị -1 -> lỗi chốt dữ liệu từ simatic thì sẽ truy cập trực tiếp vào file trend lấy dữ liệu và thực hiện chốt ở đây
                                        bool isInvalid = parts.Any(p =>
                                        {
                                            if (float.TryParse(p, NumberStyles.Float, CultureInfo.InvariantCulture, out var val))
                                                return val <= 0;
                                            return false;
                                        });
                                        if (!isInvalid)
                                        {
                                            // 6,12, 13, 14 chưa có

                                            // tần số tính từ %S
                                            m.TanSo_fb = _sensor.IsImportPhanHoiTanSo
                                                        ? _sensor.PhanHoiTanSoValue
                                                        : CalcSimatic(_sensor.PhanHoiTanSoMin, _sensor.PhanHoiTanSoMax, float.Parse(parts[1], CultureInfo.InvariantCulture));
                                            // 1. nhiệt độ môi trường
                                            m.NhietDoMoiTruong_sen = _sensor.IsImportNhietDoMoiTruong
                                                ? _sensor.NhietDoMoiTruongValue
                                                : CalcSimatic(_sensor.NhietDoMoiTruongMin, _sensor.NhietDoMoiTruongMax, float.Parse(parts[3], CultureInfo.InvariantCulture));

                                            // 2. độ ẩm
                                            m.DoAm_sen = _sensor.IsImportDoAmMoiTruong
                                                ? _sensor.DoAmMoiTruongValue
                                                : CalcSimatic(_sensor.DoAmMoiTruongMin, _sensor.DoAmMoiTruongMax, float.Parse(parts[4], CultureInfo.InvariantCulture));

                                            // 3. phản hồi vị trí van
                                            m.ViTriVan_fb = _sensor.IsImportPhanHoiViTriVan
                                                ? _sensor.PhanHoiViTriVanValue
                                                : CalcSimatic(_sensor.PhanHoiViTriVanMin, _sensor.PhanHoiViTriVanMax, float.Parse(parts[5], CultureInfo.InvariantCulture));

                                            // 4. chưa có
                                            m.DoOn_sen = _sensor.IsImportDoOn
                                                ? _sensor.DoOnValue
                                                : CalcSimatic(_sensor.DoOnMin, _sensor.DoOnMax, float.Parse(parts[6], CultureInfo.InvariantCulture));

                                            m.Momen_sen = _sensor.IsImportMomen
                                                ? _sensor.MomenValue
                                                : CalcSimatic(_sensor.MomenMin, _sensor.MomenMax, float.Parse(parts[6], CultureInfo.InvariantCulture));

                                            // điện áp luôn lấy theo giá trị nhập vào

                                            m.DienAp_fb = _sensor.IsImportPhanHoiDienAp
                                                ? _sensor.PhanHoiDienApValue
                                                : _sensor.PhanHoiDienApValue;
                                            //CalcSimatic(_sensor.PhanHoiDienApMin, _sensor.PhanHoiDienApMax, float.Parse(parts[6], CultureInfo.InvariantCulture));
                                            // 5. nhiệt độ hồng ngoại

                                            // 6. độ rung
                                            m.DoRung_sen = _sensor.IsImportDoRung
                                                ? _sensor.DoRungValue
                                                : CalcSimatic(_sensor.DoRungMin, _sensor.DoRungMax, float.Parse(parts[8], CultureInfo.InvariantCulture));

                                            //7. số vòng quay
                                            m.SoVongQuay_sen = _sensor.IsImportSoVongQuay
                                                ? _sensor.SoVongQuayValue
                                                : CalcSimatic(_sensor.SoVongQuayMin, _sensor.SoVongQuayMax, float.Parse(parts[9], CultureInfo.InvariantCulture));

                                            // 8. phản hồi công suất
                                            m.CongSuat_fb = _sensor.IsImportPhanHoiCongSuat
                                              ? _sensor.PhanHoiCongSuatValue
                                              : CalcSimatic(_sensor.PhanHoiCongSuatMin, _sensor.PhanHoiCongSuatMax, float.Parse(parts[10], CultureInfo.InvariantCulture));

                                            // 9. chênh lệch áp suất
                                            m.ChenhLechApSuat_sen = _sensor.IsImportChenhLechApSuat
                                               ? _sensor.ChenhLechApSuatValue
                                               : CalcSimatic(_sensor.ChenhLechApSuatMin, _sensor.ChenhLechApSuatMax, float.Parse(parts[11], CultureInfo.InvariantCulture));

                                            // 10. phản hồi dòng điện
                                            m.DongDien_fb = _sensor.IsImportPhanHoiDongDien
                                              ? _sensor.PhanHoiDongDienValue
                                              : CalcSimatic(_sensor.PhanHoiDongDienMin, _sensor.PhanHoiDongDienMax, float.Parse(parts[12], CultureInfo.InvariantCulture));
                                            //: _sensor.PhanHoiDienApValue;

                                            // 11. áp suất tĩnh
                                            m.ApSuatTinh_sen = _sensor.IsImportApSuatTinh
                                                ? _sensor.ApSuatTinhValue
                                                : CalcSimatic(_sensor.ApSuatTinhMin, _sensor.ApSuatTinhMax, float.Parse(parts[13], CultureInfo.InvariantCulture));

                                            // 12. áp suất khí quyển
                                            m.ApSuatkhiQuyen_sen = _sensor.IsImportApSuatKhiQuyen
                                                ? _sensor.ApSuatKhiQuyenValue
                                                : CalcSimatic(_sensor.ApSuatKhiQuyenMin, _sensor.ApSuatKhiQuyenMax, float.Parse(parts[14], CultureInfo.InvariantCulture));

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
                                        }
                                        else
                                        {
                                            WriteTomfanLog("Dữ liệu không hợp lệ, thực hiện tính thủ công từ trendline");
                                            m = CalculateTrendData(m.S, m.CV, m.k);
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
