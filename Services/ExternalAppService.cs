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
        Task ConnectExchangeAsync(List<Measure> measures, BienTan inv, CamBien sensor, OngGio duct, ThongTinMauThuNghiem input, float maxmin, float timeRange);
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

        public async Task ConnectExchangeAsync(List<Measure> measures, BienTan inv, CamBien sensor, OngGio duct, ThongTinMauThuNghiem input, float maxmin, float timeRange)
        {
            try
            {
                _measures = measures;
                _inv = inv;
                _sensor = sensor;
                _duct = duct;
                _input = input;

                if (!Directory.Exists(_exchangeFolder)) throw new BusinessException("Thư mục trao đổi dữ liệu với Simatic không tồn tại");

                _simaticResults.Clear();
                int connectedRows = 0;
                int timeoutMs = UserSetting.Instance.TimeoutMilliseconds;
                // lấy 2 row đầu để check kết nối
                int connectRows = Math.Min(2, _measures.Count);

                //var mConnects = measures.Take(2).ToArray();
                //if(await WriteDataConnectionToFilesAsync(mConnects, maxmin, timeRange))
                //{
                //    await StartAppAsync();
                //}    
                //// Chờ phản hồi từ file 2_S_IN.csv
                //var resultMeasures = await WaitForConnectResultAsync();
                //if (resultMeasures == null || resultMeasures.Length != connectRows)
                //    throw new Exception("Không thể kết nối. Không đủ dữ liệu phản hồi từ Simatic");

                //for (int i = 0; i < connectRows; i++)
                //{
                //    var m = _measures[i];
                //    var result = resultMeasures[i];
                //    if (Math.Abs(result.S - m.S) > 0.01)
                //        throw new Exception($"Không thể kết nối. Dòng {m.k} không khớp dữ liệu");

                //    m.F = MeasureStatus.Completed;
                //    connectedRows++;
                //}

                for (int i = 0; i < connectRows; i++)
                {
                    var m = _measures[i];
                    float col4 = (i == 0) ? maxmin : timeRange;
                    await WriteDataToFilesAsync(m, col4);
                    // start tu row dau tien
                    if(i == 0)
                    {
                        await StartAppAsync();
                    }
                    var result = await WaitForResultAsync(m.k, isConnection: true);
                    if (result == null || Math.Abs(result.S - m.S) > 0.01)
                        throw new Exception($"Không thể kết nối. Dòng {m.k} không khớp dữ liệu.");

                    m.F = MeasureStatus.Completed;
                    connectedRows++;
                }

                if (connectedRows == connectRows)
                {
                    DataProcess.Initialize(_measures.Count);
                    IsConnectedToSimatic = true;
                    OnSimaticConnectionChanged?.Invoke(true);
                    _currentIndex = connectedRows;
                }

            }
            catch (BusinessException)
            {
                throw;
            }
            catch (Exception ex)
            {
#if DEBUG
                throw new BusinessException("Lỗi khi kết nối tới Simatic: " + ex.Message);
#else
                return;
#endif
            }
            finally
            {
                if (!IsConnectedToSimatic && _connectCompletionSource?.Task.IsCompleted == false)
                {
                    OnSimaticConnectionChanged?.Invoke(false);
                }
            }
        }

        public async Task StartExchangeAsync()
        {
            try
            {
                if (!IsConnectedToSimatic)
                {
                    MessageBoxHelper.ShowWarning("Chưa kết nối với Simatic. Vui lòng thực hiện kết nối trước khi trao đổi dữ liệu.");
                    return;
                }
                // tạo thư mục lưu trend nếu chưa có
                if (!Directory.Exists(_trendFolder))
                    Directory.CreateDirectory(_trendFolder);

                for (int i = _currentIndex; i < _measures.Count; i++)
                {
                    var m = _measures[i];
                    await WriteDataToFilesAsync(m);
                    var result = await WaitForResultAsync(m.k, isConnection: false);

                    if (result != null)
                    {
                        m.F = MeasureStatus.Completed;
                        _simaticResults.Add(result);

                        var measurePoint = DataProcess.OnePointMeasure(result, _inv, _sensor, _duct, _input);
                        OnMeasurePointCompleted?.Invoke(measurePoint, DataProcess.ParaShow(result, _inv, _sensor, _duct, _input));

                        OnSimaticResultReceived?.Invoke(m);
                    }
                    else
                    {
                        m.F = MeasureStatus.Error;
                        OnSimaticResultReceived?.Invoke(m);
                    }

                    await Task.Delay(500);
                }
                OnSimaticExchangeCompleted?.Invoke(_simaticResults);
            }
            catch (Exception ex)
            {
#if DEBUG
                throw;

#else
                return;

#endif
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

        private async Task<bool> ExecuteWithRetryAsync(Func<Task> action, int retries = 5, int delay = 200)
        {
            for (int i = 0; i < retries; i++)
            {
                try
                {
                    await action();
                    return true;
                }
                catch (IOException)
                {
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

            // ghi vào file csv
            await ExecuteWithRetryAsync(async () =>
            {
                using (var fs = new FileStream(csvPath, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
                using (var sw = new StreamWriter(fs, Encoding.UTF8))
                {
                    await sw.WriteLineAsync($"{m.k}{sep}{m.S}{sep}{m.CV}{sep}{col4Value}");
                }
            });

            await ExecuteWithRetryAsync(async () =>
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                byte[] bin;
                using (var package = new ExcelPackage())
                {
                    var ws = package.Workbook.Worksheets.Add("1_T_OUT");
                    ws.Cells[1, 1].Value = m.k;
                    ws.Cells[1, 2].Value = m.S;
                    ws.Cells[1, 3].Value = m.CV;
                    ws.Cells[1, 4].Value = col4Value;
                    bin = await package.GetAsByteArrayAsync();
                }

                using (var fs = new FileStream(xlsxPath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
                {
                    await fs.WriteAsync(bin, 0, bin.Length);
                }
            });
        }

        private async Task<Measure?> WaitForResultAsync(int expectedK, bool isConnection)
        {
            string path2 = Path.Combine(_exchangeFolder, "2_S_IN.csv");
            var sw = Stopwatch.StartNew();
            char sep = _isComma ? ' ' : ';';

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
                                    var m = new Measure
                                    {
                                        k = k,
                                        S = float.Parse(parts[1], CultureInfo.InvariantCulture),
                                        CV = float.Parse(parts[2], CultureInfo.InvariantCulture)
                                    };


                                    // tính toán dữ liệu từ dữ liệu chốt trung bình
                                    if (!isConnection && parts.Length >= 15)
                                    {
                                        m.NhietDoMoiTruong_sen = CalcSimatic(_sensor.NhietDoMoiTruongMin, _sensor.NhietDoMoiTruongMax, float.Parse(parts[3], CultureInfo.InvariantCulture));
                                        m.DoAm_sen = CalcSimatic(_sensor.DoAmMoiTruongMin, _sensor.DoAmMoiTruongMax, float.Parse(parts[4], CultureInfo.InvariantCulture));
                                        m.ApSuatkhiQuyen_sen = CalcSimatic(_sensor.ApSuatKhiQuyenMin, _sensor.ApSuatKhiQuyenMax, float.Parse(parts[5], CultureInfo.InvariantCulture));
                                        m.ChenhLechApSuat_sen = CalcSimatic(_sensor.ChenhLechApSuatMin, _sensor.ChenhLechApSuatMax, float.Parse(parts[6], CultureInfo.InvariantCulture));
                                        m.ApSuatTinh_sen = CalcSimatic(_sensor.ApSuatTinhMin, _sensor.ApSuatTinhMax, float.Parse(parts[7], CultureInfo.InvariantCulture));
                                        m.DoRung_sen = CalcSimatic(_sensor.DoRungMin, _sensor.DoRungMax, float.Parse(parts[8], CultureInfo.InvariantCulture));
                                        m.DoOn_sen = CalcSimatic(_sensor.DoOnMin, _sensor.DoOnMax, float.Parse(parts[9], CultureInfo.InvariantCulture));
                                        m.SoVongQuay_sen = CalcSimatic(_sensor.SoVongQuayMin, _sensor.SoVongQuayMax, float.Parse(parts[10], CultureInfo.InvariantCulture));
                                        m.Momen_sen = CalcSimatic(_sensor.MomenMin, _sensor.MomenMax, float.Parse(parts[11], CultureInfo.InvariantCulture));
                                        m.DongDien_fb = CalcSimatic(_sensor.PhanHoiDongDienMin, _sensor.PhanHoiDongDienMax, float.Parse(parts[12], CultureInfo.InvariantCulture));
                                        m.CongSuat_fb = CalcSimatic(_sensor.PhanHoiCongSuatMin, _sensor.PhanHoiCongSuatMax, float.Parse(parts[13], CultureInfo.InvariantCulture));
                                        m.ViTriVan_fb = CalcSimatic(_sensor.PhanHoiViTriVanMin, _sensor.PhanHoiViTriVanMax, float.Parse(parts[14], CultureInfo.InvariantCulture));
                                        m.TanSo_fb = CalcSimatic(_sensor.PhanHoiTanSoMin, _sensor.PhanHoiTanSoMax, float.Parse(parts[1], CultureInfo.InvariantCulture));
                                    }
                                    return m;
                                }
                            }
                        }
                    }
                }
                catch (IOException) { /* File đang bị WinCC lock, đợi vòng lặp sau */ }

                await Task.Delay(200); // Polling interval
            }
            return null; // Timeout
        }

        public event Action<bool> OnSimaticConnectionChanged;
        public event Action<Measure> OnSimaticResultReceived;
        public event Action<List<Measure>> OnSimaticExchangeCompleted;
        public event Action<MeasureResponse, ParameterShow> OnMeasurePointCompleted;
        public event Action<MeasureFittingFC, Measure> OnMeasureRangeCompleted;
    }
}
