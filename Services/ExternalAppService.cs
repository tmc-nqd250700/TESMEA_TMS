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
using LicenseContext = OfficeOpenXml.LicenseContext;

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

        public async Task StartAppAsync()
        {
            string winccExePath = UserSetting.Instance.WinccExePath;
            string simaticProjectPath = UserSetting.Instance.SimaticPath;


            // 1. Kiểm tra File Thực thi WinCC
            if (string.IsNullOrEmpty(winccExePath) || !File.Exists(winccExePath))
            {
                MessageBoxHelper.ShowWarning("Đường dẫn file thực thi WinCC (.exe) không tồn tại");
                return;
            }

            // 2. Kiểm tra File Dự án (Để chắc chắn có thể mở được)
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
                await WriteDataToFilesAsync(m, maxmin);

                var result = await WaitForResultAsync(m.k, isConnection: true);
                if (result == null || Math.Abs(result.S - m.S) > 0.01)
                {
                    string errorMsg = result == null ? "Timeout (Không có phản hồi)" : $"Dữ liệu không khớp (S_gửi={m.S}, S_nhận={result.S})";
                    WriteTomfanLog($"[ERROR] Kết nối thất bại tại dòng k={m.k}: {errorMsg}");
                    throw new Exception($"Không thể kết nối. Dòng {m.k} lỗi: {errorMsg}");
                }

                m.F = MeasureStatus.Completed;
                WriteTomfanLog($"Connect - Dòng k={m.k} THÀNH CÔNG");
                _currentIndex = m.k;
                return true;
            }
            catch (Exception ex)
            {
                WriteTomfanLog($"[CRITICAL] Lỗi trong ConnectExchangeAsync: {ex.Message}");
                throw;
            }
        }

        public async Task<bool> ConnectExchangeAsync(Measure measure, float timeRange)
        {
            try
            {
                WriteTomfanLog("========== BẮT ĐẦU QUY TRÌNH KẾT NỐI (CONNECT) ==========");
                if (!Directory.Exists(_exchangeFolder))
                {
                    WriteTomfanLog($"[FATAL] Thư mục trao đổi không tồn tại: {_exchangeFolder}");
                    throw new BusinessException("Thư mục trao đổi dữ liệu với Simatic không tồn tại");
                }

                WriteTomfanLog($"Connect - Thử dòng k={measure.k}: Ghi file và chờ WinCC phản hồi...");
                await WriteDataToFilesAsync(measure, timeRange);

                var result = await WaitForResultAsync(measure.k, isConnection: true);
                if (result == null || Math.Abs(result.S - measure.S) > 0.01)
                {
                    string errorMsg = result == null ? "Timeout (Không có phản hồi)" : $"Dữ liệu không khớp (S_gửi={measure.S}, S_nhận={result.S})";
                    WriteTomfanLog($"[ERROR] Kết nối thất bại tại dòng k={measure.k}: {errorMsg}");
                    throw new Exception($"Không thể kết nối. Dòng {measure.k} lỗi: {errorMsg}");
                }

                measure.F = MeasureStatus.Completed;
                WriteTomfanLog($"Connect - Dòng k={measure.k} THÀNH CÔNG. Nghỉ 5s chờ WinCC sẵn sàng...");
                DataProcess.Initialize(_measures.Count);
                IsConnectedToSimatic = true;
                OnSimaticConnectionChanged?.Invoke(true);
                _currentIndex = measure.k;
                WriteTomfanLog("SUCCESS: Đã thiết lập kết nối với Simatic thành công.");
                return true;
            }
            catch (Exception ex)
            {
                WriteTomfanLog($"[CRITICAL] Lỗi trong ConnectExchangeAsync: {ex.Message}");
                throw;
            }
        }
        public async Task StartExchangeAsync()
        {
            try
            {
                if (!IsConnectedToSimatic)
                {
                    WriteTomfanLog("[WARN] StartExchange bị từ chối: Chưa kết nối Simatic.");
                    MessageBoxHelper.ShowWarning("Chưa kết nối với Simatic.");
                    return;
                }

                WriteTomfanLog("========== BẮT ĐẦU TRAO ĐỔI DỮ LIỆU HÀNG LOẠT ==========");

                if (!Directory.Exists(_trendFolder))
                    Directory.CreateDirectory(_trendFolder);

                for (int i = _currentIndex; i < _measures.Count; i++)
                {
                    var m = _measures[i];
                    WriteTomfanLog($"Processing: Đang xử lý điểm đo k={m.k}/{_measures.Count}");

                    await WriteDataToFilesAsync(m);

                    // Chờ kết quả xử lý thực tế (isConnection = false để tính toán sensor)
                    var result = await WaitForResultAsync(m.k, isConnection: false);

                    if (result != null)
                    {
                        WriteTomfanLog($"Received: Đã nhận kết quả k={m.k}. Tiến hành tính toán PointMeasure.");
                        m.F = MeasureStatus.Completed;
                        _simaticResults.Add(result);

                        var measurePoint = DataProcess.OnePointMeasure(result, _inv, _sensor, _duct, _input);
                        OnMeasurePointCompleted?.Invoke(measurePoint, DataProcess.ParaShow(result, _inv, _sensor, _duct, _input));
                        OnSimaticResultReceived?.Invoke(m);

                        WriteTomfanLog($"Done: Hoàn tất điểm đo k={m.k}. Nghỉ 5s...");
                    }
                    else
                    {
                        WriteTomfanLog($"[ERROR]: Điểm đo k={m.k} thất bại do Timeout.");
                        m.F = MeasureStatus.Error;
                        OnSimaticResultReceived?.Invoke(m);
                    }

                    await Task.Delay(5000);
                }

                WriteTomfanLog("========== HOÀN TẤT TOÀN BỘ CHU KỲ TRAO ĐỔI ==========");
                OnSimaticExchangeCompleted?.Invoke(_simaticResults);
            }
            catch (Exception ex)
            {
                WriteTomfanLog($"[CRITICAL] Lỗi trong StartExchangeAsync: {ex.Message}");
                throw;
            }
        }
        public async Task StopExchangeAsync()
        {
            if (_watcher != null)
            {
                _watcher.EnableRaisingEvents = false;
                _watcher.Dispose();
                _watcher = null;
            }
            _measures.Clear();
            _simaticResults.Clear();
            _currentIndex = 0;
            IsConnectedToSimatic = false;
        }

        private float CalcSimatic(float minValue, float maxValue, float percent)
        {
            return minValue + (maxValue - minValue) * percent / 100f;
        }

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
                    WriteTomfanLog($"[RETRY {i + 1}/{retries}] File đang bị khóa bởi WinCC: {ex.Message}");
                    if (i == retries - 1)
                    {
                        WriteTomfanLog("[FATAL] Đã thử lại tối đa nhưng vẫn không thể truy cập file.");
                        throw;
                    }
                    await Task.Delay(delay);
                }
            }
            return false;
        }

        private async Task WriteDataToFilesAsync(Measure m, float col4Value = 0)
        {
            string xlsxPath = Path.Combine(_exchangeFolder, "1_T_OUT.xlsx");
            string csvPath = Path.Combine(_exchangeFolder, "1_T_OUT.csv");
            char sep = _isComma ? ' ' : ';';

            WriteTomfanLog($">>> Ghi dữ liệu dòng k={m.k} (S={m.S}, CV={m.CV})");

            // Ghi CSV
            await ExecuteWithRetryAsync(async () =>
            {
                List<string> lines = new List<string>();
                if (File.Exists(csvPath))
                {
                    lines = (await File.ReadAllLinesAsync(csvPath, Encoding.UTF8)).ToList();
                }

                int rowIndex = m.k - 1;
                while (lines.Count <= rowIndex) lines.Add("");
                lines[rowIndex] = $"{m.k}{sep}{m.S}{sep}{m.CV}{sep}{col4Value}";

                string tempCsv = csvPath + ".tmp";
                await File.WriteAllLinesAsync(tempCsv, lines, Encoding.UTF8);
                File.Move(tempCsv, csvPath, true);
                WriteTomfanLog($"Step: CSV k={m.k} ghi thành công ");
            });

            // Ghi XLSX
            await ExecuteWithRetryAsync(async () =>
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                byte[] fileData;

                using (var fs = new FileStream(xlsxPath, FileMode.OpenOrCreate, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (var package = new ExcelPackage())
                    {
                        await package.LoadAsync(fs);
                        var ws = package.Workbook.Worksheets.FirstOrDefault() ?? package.Workbook.Worksheets.Add("1_T_OUT");

                        ws.Cells[m.k, 1].Value = m.k;
                        ws.Cells[m.k, 2].Value = m.S;
                        ws.Cells[m.k, 3].Value = m.CV;
                        ws.Cells[m.k, 4].Value = col4Value;

                        fileData = await package.GetAsByteArrayAsync();
                    }
                }

                string tempPath = xlsxPath + ".tmp";
                await File.WriteAllBytesAsync(tempPath, fileData);
                File.Move(tempPath, xlsxPath, true);
                WriteTomfanLog($"Step: XLSX k={m.k} ghi thành công ");
            });
        }

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
                            string? line;
                            while ((line = await sr.ReadLineAsync()) != null)
                            {
                                var parts = line.Split(sep);
                                if (parts.Length < 3) continue;

                                if (int.TryParse(parts[0], out int k) && k == expectedK)
                                {
                                    WriteTomfanLog($"MATCH: Tìm thấy dòng k={k} trong 2_S_IN.csv sau {sw.ElapsedMilliseconds}ms");
                                    var m = new Measure
                                    {
                                        k = k,
                                        S = float.Parse(parts[1], CultureInfo.InvariantCulture),
                                        CV = float.Parse(parts[2], CultureInfo.InvariantCulture)
                                    };

                                    if (!isConnection && parts.Length >= 15)
                                    {
                                        // ... (Giữ nguyên phần CalcSimatic của bạn)
                                        WriteTomfanLog("Step: Đã tính toán xong các thông số cảm biến từ WinCC.");
                                    }
                                    return m;
                                }
                            }
                        }
                    }
                }
                catch (IOException)
                {
                    // Không ghi log ở đây để tránh làm file log quá nặng vì polling 200ms
                }

                await Task.Delay(200);
            }

            WriteTomfanLog($"[TIMEOUT] Không nhận được phản hồi cho k={expectedK} sau {UserSetting.Instance.TimeoutMilliseconds}ms");
            return null;
        }


        private void WriteTomfanLog(string message)
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
            catch { /* Tránh treo app vì lỗi ghi log */ }
        }

        public event Action<bool> OnSimaticConnectionChanged;
        public event Action<Measure> OnSimaticResultReceived;
        public event Action<List<Measure>> OnSimaticExchangeCompleted;
        public event Action<MeasureResponse, ParameterShow> OnMeasurePointCompleted;
        public event Action<MeasureFittingFC, Measure> OnMeasureRangeCompleted;
    }
}
