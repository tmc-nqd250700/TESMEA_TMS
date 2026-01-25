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
                await WriteDataToFilesAsync(m, maxmin);

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
                WriteTomfanLog($"[CRITICAL] Lỗi trong ConnectExchangeAsync: {ex.Message}");
                throw;
            }
        }

        // kết nối dòng 2
        public async Task<bool> ConnectExchangeAsync(Measure measure, float timeRange)
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
                await WriteDataToFilesAsync(measure, timeRange);

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
                //int sIndex = 0;
                //int cvIndex = 0;
                double currentS = _measures[_currentIndex].S;
                int startIndex = 2; // track index của điểm bắt đầu trong dải đo
                int timeoutMs = UserSetting.Instance.TimeoutMilliseconds;
                
                for (int i = _currentIndex; i < _measures.Count; i++)
                {
                    _currentIndex = _measures[i].k;
                    var m = _measures[i];
                    WriteTomfanLog($"Đang xử lý điểm đo k={m.k}/{_measures.Count}");

                    // tạo file trend.csv theo k
                    using (var fs = File.Create(Path.Combine(_trendFolder, $"{m.k}.csv"))) { }
                    await WriteDataToFilesAsync(m);

                    // Chờ kết quả xử lý thực tế (isConnection = false để tính toán sensor)
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

                        //OnSimaticResultReceived?.Invoke(_measures[i]);
                        //var paramShow = DataProcess.ParaShow(result, _inv, _sensor, _duct, _input);
                        //OnMeasurePointCompleted?.Invoke(measurePoint, paramShow);


                        // check nếu chuyển % tần số hoặc là điểm đo cuối cùng thì thực hiện fitting FC và vẽ line cho chart
                        if (_measures[i].S != currentS || i == _measures.Count - 1)
                        {
                            var fitting = DataProcess.FittingFC(currentRange.Count, startIndex);
                            WriteTomfanLog($"Hoàn tất dải đo từ k={currentRange.First().k} đến k={currentRange.Last().k}");
                            OnMeasureRangeCompleted?.Invoke(fitting, currentRange.LastOrDefault());
                            currentRange.Clear();
                            currentS = _measures[i].S;
                            startIndex = i;
                            //sIndex++;
                            //cvIndex = 0;
                        }
                        
                        //string fileName = $"{sIndex}{cvIndex}.csv";
                        //string filePath = Path.Combine(_trendFolder, fileName);
                        ////using (var fs = File.Create(filePath)) { }
                        //cvIndex++;

                        WriteTomfanLog($"Hoàn tất điểm đo k={m.k}");
                    }
                    else
                    {
                        WriteTomfanLog($"Điểm đo k={m.k} thất bại do Timeout");
                        m.F = MeasureStatus.Error;
                        OnSimaticResultReceived?.Invoke(m);
                    }
                }

                // HOÀN THÀNH KỊCH BẢN ĐO KIỂM, DỪNG TRAO ĐỔI
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
            //if (_watcher != null)
            //{
            //    _watcher.EnableRaisingEvents = false;
            //    _watcher.Dispose();
            //    _watcher = null;
            //}
            //_measures.Clear();
            //_simaticResults.Clear();
            //_currentIndex = 0;
            //IsConnectedToSimatic = false;

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

                // Gửi lệnh dừng: tham số eStop = true để ép số 96 vào cột k
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

        // ghi dữ liệu vào file 1.csv, 1.xlsx
        //private async Task WriteDataToFilesAsync(Measure m, float col4Value = 0, bool eStop =false)
        //{
        //    try
        //    {
        //        string xlsxPath = Path.Combine(_exchangeFolder, "1_T_OUT.xlsx");
        //        string csvPath = Path.Combine(_exchangeFolder, "1_T_OUT.csv");
        //        char sep = _isComma ? ' ' : ';';
        //        string tempCsvPath = csvPath + ".tmp";
        //        string tempXlsxPath = xlsxPath + ".tmp";
        //        WriteTomfanLog($">>> Ghi dữ liệu dòng k={m.k} (S={m.S}, CV={m.CV})");

        //        // Ghi CSV
        //        int kValueToPrint = eStop ? 96 : 100;
        //        int rowIdx = m.k > 0 ? m.k : 1;
        //        await ExecuteWithRetryAsync(async () =>
        //        {
        //            List<string> lines = new List<string>();
        //            if (File.Exists(csvPath))
        //            {
        //                lines = (await File.ReadAllLinesAsync(csvPath)).ToList();
        //            }
        //            else
        //            {
        //                throw new BusinessException("File 1_T_OUT.csv không tồn tại.");
        //            }

        //            // Tạo nội dung dòng mới
        //            string newLine = col4Value == 0
        //                ? $"{kValueToPrint}{sep}{m.S}{sep}{m.CV}"
        //                : $"{m.k}{sep}{m.S}{sep}{m.CV}{sep}{col4Value}";

        //            WriteTomfanLog($"Dữ liệu dòng mới: {newLine}");
        //            while (lines.Count < rowIdx)
        //            {
        //                lines.Add("");
        //            }
        //            lines[rowIdx - 1] = newLine;
        //            await File.WriteAllLinesAsync(tempCsvPath, lines, Encoding.UTF8);
        //            File.Move(tempCsvPath, csvPath, true);
        //            WriteTomfanLog($"Step: CSV row {rowIdx} ghi thành công.");
        //        });

        //        // Ghi XLSX
        //        await ExecuteWithRetryAsync(async () =>
        //        {
        //            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        //            byte[] fileData;

        //            using (var package = new ExcelPackage())
        //            {
        //                if (File.Exists(xlsxPath))
        //                {
        //                    using (var fsRead = new FileStream(xlsxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        //                    {
        //                        await package.LoadAsync(fsRead);
        //                    }
        //                }
        //                else
        //                {
        //                    throw new BusinessException("File 1_T_OUT.xlsx không tồn tại.");
        //                }

        //                var ws = package.Workbook.Worksheets.FirstOrDefault() ?? package.Workbook.Worksheets.Add("1_T_OUT");

        //                // make sure index dòng >= 1
        //                int rowIdx = m.k > 0 ? m.k : 1;
        //                // insert m.k khi connect
        //                ws.Cells[rowIdx, 1].Value = kValueToPrint;
        //                if (col4Value != 0)
        //                {
        //                    // override khi la do kiem
        //                    ws.Cells[rowIdx, 1].Value = m.k;
        //                    ws.Cells[rowIdx, 4].Value = col4Value;
        //                }
        //                ws.Cells[rowIdx, 2].Value = m.S;
        //                ws.Cells[rowIdx, 3].Value = m.CV;
        //                fileData = await package.GetAsByteArrayAsync();
        //            }

        //            using (var fsWrite = new FileStream(tempXlsxPath, FileMode.Create, FileAccess.Write, FileShare.None))
        //            {
        //                await fsWrite.WriteAsync(fileData, 0, fileData.Length);
        //                await fsWrite.FlushAsync();
        //                fsWrite.Flush(true);
        //            }
        //            File.Move(tempXlsxPath, xlsxPath, true);
        //        });
        //    }
        //    catch (Exception ex)
        //    {
        //        WriteTomfanLog($"Error WriteDataToFilesAsync: {ex}");
        //    }
        //}

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

                //// Ghi XLSX trực tiếp
                //await ExecuteWithRetryAsync(async () =>
                //{
                //    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                //    byte[] fileData;

                //    using (var package = new ExcelPackage())
                //    {
                //        if (File.Exists(xlsxPath))
                //        {
                //            using (var fsRead = new FileStream(xlsxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                //            {
                //                await package.LoadAsync(fsRead);
                //            }
                //        }
                //        else
                //        {
                //            throw new BusinessException("File 1_T_OUT.xlsx không tồn tại.");
                //        }

                //        var ws = package.Workbook.Worksheets.FirstOrDefault() ?? package.Workbook.Worksheets.Add("1_T_OUT");

                //        int rowIdx = m.k > 0 ? m.k : 1;
                //        ws.Cells[rowIdx, 1].Value = kValueToPrint;
                //        if (col4Value != 0)
                //        {
                //            ws.Cells[rowIdx, 1].Value = m.k;
                //            ws.Cells[rowIdx, 4].Value = col4Value;
                //        }
                //        ws.Cells[rowIdx, 2].Value = m.S;
                //        ws.Cells[rowIdx, 3].Value = m.CV;
                //        fileData = await package.GetAsByteArrayAsync();
                //    }

                //    // Ghi đè lại toàn bộ file trực tiếp
                //    using (var fsWrite = new FileStream(xlsxPath, FileMode.Create, FileAccess.Write, FileShare.None))
                //    {
                //        await fsWrite.WriteAsync(fileData, 0, fileData.Length);
                //        await fsWrite.FlushAsync();
                //        fsWrite.Flush(true);
                //    }
                //    WriteTomfanLog($"Step: XLSX row {m.k} ghi thành công.");
                //});
            }
            catch (Exception ex)
            {
                WriteTomfanLog($"Error WriteDataToFilesAsync: {ex}");
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
                            int targetIndex = expectedK - 1;
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
                                    if (!isConnection && parts.Length == 15)
                                    {
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
                                            : CalcSimatic(_sensor.DoRungMin, _sensor.DoRungMax, float.Parse(parts[9], CultureInfo.InvariantCulture));

                                        //7. số vòng quay
                                        m.SoVongQuay_sen = _sensor.IsImportSoVongQuay
                                            ? _sensor.SoVongQuayValue
                                            : CalcSimatic(_sensor.SoVongQuayMin, _sensor.SoVongQuayMax, float.Parse(parts[10], CultureInfo.InvariantCulture));

                                        // 8. phản hồi công suất
                                        m.CongSuat_fb = _sensor.IsImportPhanHoiCongSuat
                                          ? _sensor.PhanHoiCongSuatValue
                                          : CalcSimatic(_sensor.PhanHoiCongSuatMin, _sensor.PhanHoiCongSuatMax, float.Parse(parts[11], CultureInfo.InvariantCulture));

                                        // 9. chênh lệch áp suất
                                        m.ChenhLechApSuat_sen = _sensor.IsImportChenhLechApSuat
                                           ? _sensor.ChenhLechApSuatValue
                                           : CalcSimatic(_sensor.ChenhLechApSuatMin, _sensor.ChenhLechApSuatMax, float.Parse(parts[12], CultureInfo.InvariantCulture));

                                        // 10. phản hồi dòng điện
                                        m.DongDien_fb = _sensor.IsImportPhanHoiDongDien
                                          ? _sensor.PhanHoiDongDienValue
                                          : CalcSimatic(_sensor.PhanHoiDongDienMin, _sensor.PhanHoiDongDienMax, float.Parse(parts[13], CultureInfo.InvariantCulture));
                                        //: _sensor.PhanHoiDienApValue;


                                        // 11. áp suất tĩnh
                                        m.ApSuatTinh_sen = _sensor.IsImportApSuatTinh
                                            ? _sensor.ApSuatTinhValue
                                            : CalcSimatic(_sensor.ApSuatTinhMin, _sensor.ApSuatTinhMax, float.Parse(parts[14], CultureInfo.InvariantCulture));

                                        // 12. áp suất khí quyển
                                        m.ApSuatkhiQuyen_sen = _sensor.IsImportApSuatKhiQuyen
                                            ? _sensor.ApSuatKhiQuyenValue
                                            : CalcSimatic(_sensor.ApSuatKhiQuyenMin, _sensor.ApSuatKhiQuyenMax, float.Parse(parts[15], CultureInfo.InvariantCulture));
                                    }

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
                            }
                        }
                    }
                }
                catch (IOException)
                {
                }

                await Task.Delay(200);
            }

            WriteTomfanLog($"Không nhận được phản hồi cho k={expectedK} sau {UserSetting.Instance.TimeoutMilliseconds}ms");
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
