using OfficeOpenXml;
using System.Diagnostics;
using System.IO;
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
        Task ConnectExchangeAsync(List<Measure> measures, BienTan inv, CamBien sensor, OngGio duct, ThongTinMauThuNghiem input, float maxmin, float timeRange, string fileFormat = "xlsx");
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
        private string _exchangeFilePath;
        private string _trendlineFolder;
        private string _fileFormat = "xlsx";
        public ExternalAppService()
        {
            _exchangeFolder = UserSetting.TOMFAN_folder;
            _exchangeFilePath = Path.Combine(_exchangeFolder, "1_T_OUT.xlsx");
            _trendlineFolder = Path.Combine(_exchangeFolder, "Testdata");
        }

        public async Task StartAppAsync()
        {
            string simaticPath = UserSetting.Instance.SimaticPath;

            if (string.IsNullOrEmpty(simaticPath) || !File.Exists(simaticPath))
            {
                MessageBoxHelper.ShowWarning("Đường dẫn Simatic không hợp lệ");
                return;
            }
            //foreach (var proc in Process.GetProcessesByName(Path.GetFileNameWithoutExtension(simaticPath)))
            //{
            //    try
            //    {
            //        if (proc.MainModule.FileName.Equals(simaticPath, StringComparison.OrdinalIgnoreCase))
            //        {
            //            proc.Kill(true);
            //            proc.WaitForExit(2000);
            //        }
            //    }
            //    catch
            //    {

            //    }
            //}
            foreach (var proc in Process.GetProcessesByName(Path.GetFileNameWithoutExtension(simaticPath)))
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

                    if (fileName.Equals(simaticPath, StringComparison.OrdinalIgnoreCase) && !proc.HasExited)
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
                    StartInfo = new ProcessStartInfo(simaticPath)
                    {
                        UseShellExecute = false,
                        CreateNoWindow = false,
                        WindowStyle = ProcessWindowStyle.Normal
                    }
                };
                _process.StartInfo.UseShellExecute = false;
                // Đăng ký event để monitor process
                _process.Exited += (s, e) =>
                {
                    System.Diagnostics.Debug.WriteLine($"Simatic process exited unexpectedly at {DateTime.Now}");
                };
                _process.EnableRaisingEvents = true;
                _process.Start();
                var started = await Task.Run(() =>
                {
                    return _process.WaitForInputIdle(UserSetting.Instance.TimeoutMilliseconds);
                }, _cts.Token);

                if (!started)
                {
                    // timeout exception
                    await StopAppAsync();
                    MessageBoxHelper.ShowError("Hết thời gian chờ khởi động phần mềm Simatic");
                }
            }
            //catch (Win32Exception ex)
            //{
            //    await StopAppAsync();
            //    MessageBoxHelper.ShowError("Lỗi hệ thống khi khởi động phần mềm Simatic");
            //}
            //catch (UnauthorizedAccessException ex)
            //{
            //    await StopAppAsync();
            //    MessageBoxHelper.ShowError("Không đủ quyền để khởi động phần mềm Simatic");
            //}
            //catch (InvalidOperationException ex)
            //{
            //    await StopAppAsync();
            //    MessageBoxHelper.ShowError("Thao tác không hợp lệ với process Simatic");
            //}
            //catch (OperationCanceledException ex)
            //{
            //    await StopAppAsync();
            //    MessageBoxHelper.ShowError("Thao tác khởi động phần mềm Simatic bị hủy");
            //}
            //catch (AggregateException ex)
            //{
            //    await StopAppAsync();
            //    MessageBoxHelper.ShowError("Có lỗi bất đồng bộ khi khởi động phần mềm Simatic");
            //}
            //catch (NotSupportedException ex)
            //{
            //    await StopAppAsync();
            //    MessageBoxHelper.ShowError("Đường dẫn hoặc thao tác không được hỗ trợ");
            //}
            //catch (Exception ex)
            //{
            //    await StopAppAsync();
            //    MessageBoxHelper.ShowError("Lỗi khi khởi động phần mềm Simatic");
            //}
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

        public async Task ConnectExchangeAsync(List<Measure> measures, BienTan inv, CamBien sensor, OngGio duct, ThongTinMauThuNghiem input, float maxmin, float timeRange, string fileFormat)
        {
            try
            {

                _measures = measures;
                _inv = inv;
                _sensor = sensor;
                _duct = duct;
                _input = input;
                _fileFormat = fileFormat;

                _currentIndex = 0;
                _simaticResults.Clear();
                if (!Directory.Exists(_exchangeFolder))
                {
                    Directory.CreateDirectory(_exchangeFolder);
                }

                if (_fileFormat == "csv")
                {
                    using (var writer = new StreamWriter(_exchangeFilePath))
                    {
                        for (int i = 0; i < Math.Min(2, _measures.Count); i++)
                        {
                            var m = _measures[i];
                            var row = string.Join(",", new object[]
                            {
                                m.k,
                                m.S,
                                m.CV
                            });
                            writer.WriteLine(row);
                        }
                    }
                }
                else
                {
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                    using (var package = new ExcelPackage(new FileInfo(_exchangeFilePath)))
                    {
                        var ws = package.Workbook.Worksheets["1_T_OUT"] ?? package.Workbook.Worksheets.Add("1_T_OUT");
                        //ws.Cells[1, 1].Value = "k";
                        //ws.Cells[1, 2].Value = "S";
                        //ws.Cells[1, 3].Value = "CV";

                        //// thông số dải cảm biến
                        //ws.Cells[1, 4].Value = "Cảm biến nhiệt độ môi trường (oC)";
                        //ws.Cells[1, 5].Value = "Cảm biến độ ẩm môi trường (%)";
                        //ws.Cells[1, 6].Value = "Cảm biến áp suất khí quyển (Pa)";
                        //ws.Cells[1, 7].Value = "Cảm biến chênh lệch áp suất (Pa)";
                        //ws.Cells[1, 8].Value = "Cảm biến áp suất tĩnh (Pa)";
                        //ws.Cells[1, 9].Value = "Cảm biến độ rung";
                        //ws.Cells[1, 10].Value = "Cảm biến độ ồn";
                        //ws.Cells[1, 11].Value = "Số vòng quay (Tốc độ thực của guồng cánh)";
                        //ws.Cells[1, 12].Value = "Cảm biến momen xoắn";
                        //ws.Cells[1, 13].Value = "Phản hồi dòng điện (AO1_INV)";
                        //ws.Cells[1, 14].Value = "Phản hồi công suất (AO2_INV)";
                        //ws.Cells[1, 15].Value = "Phản hồi vị trí van (CV)";
                        //ws.Cells[1, 16].Value = "Tần số";
                        //ws.Cells[1, 17].Value = "Nhiệt độ gối trục";

                        for (int i = 0; i < Math.Min(2, _measures.Count); i++)
                        {
                            // từ row 2 trong excel
                            var m = _measures[i];
                            ws.Cells[i + 1, 1].Value = m.k;
                            ws.Cells[i + 1, 2].Value = m.S;
                            ws.Cells[i + 1, 3].Value = m.CV;
                        }

                        ws.Cells[1, 4].Value = maxmin;
                        ws.Cells[2, 4].Value = timeRange;

                        ws.Cells[ws.Dimension.Address].AutoFitColumns();
                        package.Save();
                    }
                }

                //ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                //using (var package = new ExcelPackage(new FileInfo(_exchangeFilePath)))
                //{
                //    var ws = package.Workbook.Worksheets["1_T_OUT"] ?? package.Workbook.Worksheets.Add("1_T_OUT");
                //    ws.Cells[1, 1].Value = "k";
                //    ws.Cells[1, 2].Value = "S";
                //    ws.Cells[1, 3].Value = "CV";

                //    // thông số dải cảm biến
                //    ws.Cells[1, 4].Value = "Cảm biến nhiệt độ môi trường (oC)";
                //    ws.Cells[1, 5].Value = "Cảm biến độ ẩm môi trường (%)";
                //    ws.Cells[1, 6].Value = "Cảm biến áp suất khí quyển (Pa)";
                //    ws.Cells[1, 7].Value = "Cảm biến chênh lệch áp suất (Pa)";
                //    ws.Cells[1, 8].Value = "Cảm biến áp suất tĩnh (Pa)";
                //    ws.Cells[1, 9].Value = "Cảm biến độ rung";
                //    ws.Cells[1, 10].Value = "Cảm biến độ ồn";
                //    ws.Cells[1, 11].Value = "Số vòng quay (Tốc độ thực của guồng cánh)";
                //    ws.Cells[1, 12].Value = "Cảm biến momen xoắn";
                //    ws.Cells[1, 13].Value = "Phản hồi dòng điện (AO1_INV)";
                //    ws.Cells[1, 14].Value = "Phản hồi công suất (AO2_INV)";
                //    ws.Cells[1, 15].Value = "Phản hồi vị trí van (CV)";
                //    //ws.Cells[1, 16].Value = "Tần số";
                //    //ws.Cells[1, 17].Value = "Nhiệt độ gối trục";

                //    // fill thông tin
                //    for (int i = 0; i < Math.Min(2, _measures.Count); i++)
                //    {
                //        // từ row 2 trong excel
                //        var m = _measures[i];
                //        ws.Cells[i + 2, 1].Value = m.k;
                //        ws.Cells[i + 2, 2].Value = m.S;
                //        ws.Cells[i + 2, 3].Value = m.CV;

                //        // Các cột cảm biến: dòng đầu lấy min, dòng thứ hai lấy max
                //        ws.Cells[i + 2, 4].Value = i == 0 ? sensor.NhietDoMoiTruongMin : sensor.NhietDoMoiTruongMax;
                //        ws.Cells[i + 2, 5].Value = i == 0 ? sensor.DoAmMoiTruongMin : sensor.DoAmMoiTruongMax;
                //        ws.Cells[i + 2, 6].Value = i == 0 ? sensor.ApSuatKhiQuyenMin : sensor.ApSuatKhiQuyenMax;
                //        ws.Cells[i + 2, 7].Value = i == 0 ? sensor.ChenhLechApSuatMin : sensor.ChenhLechApSuatMax;
                //        ws.Cells[i + 2, 8].Value = i == 0 ? sensor.ApSuatTinhMin : sensor.ApSuatTinhMax;
                //        ws.Cells[i + 2, 9].Value = i == 0 ? sensor.DoRungMin : sensor.DoRungMax;
                //        ws.Cells[i + 2, 10].Value = i == 0 ? sensor.DoOnMin : sensor.DoOnMax;
                //        ws.Cells[i + 2, 11].Value = i == 0 ? sensor.SoVongQuayMin : sensor.SoVongQuayMax;
                //        ws.Cells[i + 2, 12].Value = i == 0 ? sensor.MomenMin : sensor.MomenMax;
                //        ws.Cells[i + 2, 13].Value = i == 0 ? sensor.PhanHoiDongDienMin : sensor.PhanHoiDongDienMax;
                //        ws.Cells[i + 2, 14].Value = i == 0 ? sensor.PhanHoiCongSuatMin : sensor.PhanHoiCongSuatMax;
                //        ws.Cells[i + 2, 15].Value = i == 0 ? sensor.PhanHoiViTriVanMin : sensor.PhanHoiViTriVanMax;
                //        //ws.Cells[i + 2, 16].Value = i == 0 ? sensor.TanSoMin : sensor.TanSoMax;
                //        //ws.Cells[i + 2, 17].Value = i == 0 ? sensor.NhietDoGoiTrucMin : sensor.NhietDoGoiTrucMax;
                //    }

                //    ws.Cells[ws.Dimension.Address].AutoFitColumns();
                //    package.Save();
                //}

                _connectCompletionSource = new TaskCompletionSource<bool>();
                int timeoutMs = UserSetting.Instance.TimeoutMilliseconds;
                int receivedCount = 0;

                _watcher = new FileSystemWatcher(_exchangeFolder, $"2_S_IN.{_fileFormat}")
                {
                    NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.Size
                };
                _watcher.Changed += async (s, e) =>
                {
                    if (_watcher == null) return;
                    var results = await ReadSimaticResult(e.FullPath, _sensor);
                    foreach (var m in results)
                    {
                        if (_measures.Take(2).Any(x => x.k == m.k))
                        {
                            var measure = _measures.FirstOrDefault(x => x.k == m.k);
                            if (measure != null && measure.F != MeasureStatus.Completed)
                            {
                                measure.F = MeasureStatus.Completed;
                                //OnSimaticResultReceived?.Invoke(measure);
                                receivedCount++;
                                _currentIndex++;
                            }
                        }
                    }
                    // stop watcher khi đã nhận đủ 2 dòng
                    if (receivedCount == 2)
                    {
                        if (_watcher != null)
                        {
                            _watcher.EnableRaisingEvents = false;
                            _watcher.Dispose();
                            _watcher = null;
                        }

                        // khởi tạo mảng dữ liệu cho quá trình đo
                        DataProcess.Initialize(_measures.Count);
                        IsConnectedToSimatic = true;
                        OnSimaticConnectionChanged?.Invoke(true);
                        _connectCompletionSource?.TrySetResult(true);
                    }
                };
                _watcher.EnableRaisingEvents = true;
                // Chờ kết nối hoặc timeout
                var completedTask = await Task.WhenAny(_connectCompletionSource.Task, Task.Delay(timeoutMs));
                if (completedTask != _connectCompletionSource.Task)
                {
                    // Timeout
                    if (_watcher != null)
                    {
                        _watcher.EnableRaisingEvents = false;
                        _watcher.Dispose();
                        _watcher = null;
                    }
                    IsConnectedToSimatic = false;
                    OnSimaticConnectionChanged?.Invoke(false);
                    throw new TimeoutException("Kết nối Simatic bị timeout");
                }

            }
            catch (Exception ex)
            {
                throw new BusinessException("Lỗi khi kết nối với Simatic: " + ex.Message);
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

                if (!_measures.Any())
                {
                    MessageBoxHelper.ShowWarning("Dữ liệu scenario trống");
                    return;
                }

                // dispose watcher cũ
                if (_watcher != null)
                {
                    _watcher.EnableRaisingEvents = false;
                    _watcher.Dispose();
                    _watcher = null;
                }

                // lưu dải đo
                List<Measure> currentRange = new List<Measure>();
                double currentS = _measures[_currentIndex].S;
                int startIndex = 2; // track index của điểm bắt đầu trong dải đo
                int timeoutMs = UserSetting.Instance.TimeoutMilliseconds;

                // break khi eStop
                if (_measures == null) return;
                if (_currentIndex < 2) _currentIndex = 2;

                for (int i = _currentIndex; i < _measures.Count; i++)
                {
                    var measure = _measures[i];
                    var completionSource = new TaskCompletionSource<bool>();
                    bool isResultProcessed = false;
                    _watcher = new FileSystemWatcher(_exchangeFolder, $"2_S_IN.{_fileFormat}")
                    {
                        NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.Size
                    };

                    _watcher.Changed += async (s, e) =>
                    {
                        try
                        {
                            if (_measures == null) return;
                            if (isResultProcessed) return;
                            var results = await ReadSimaticResult(e.FullPath, _sensor);
                            var expectedK = _measures[i].k;
                            var matchingResult = results.FirstOrDefault(r => r.k == expectedK);

                            if (matchingResult != null && !currentRange.Any(m => m.k == expectedK))
                            {
                                ((FileSystemWatcher)s).EnableRaisingEvents = false;

                                isResultProcessed = true;

                                _measures[i].F = MeasureStatus.Completed;
                                _simaticResults.Add(matchingResult);
                                OnSimaticResultReceived?.Invoke(_measures[i]);

                                if (_measures[i].S != currentS)
                                {
                                    var fitting = DataProcess.FittingFC(currentRange.Count, startIndex);
                                    OnMeasureRangeCompleted?.Invoke(fitting, currentRange.LastOrDefault());
                                    currentRange.Clear();
                                    currentS = _measures[i].S;
                                    startIndex = i;
                                }

                                var measurePoint = DataProcess.OnePointMeasure(matchingResult, _inv, _sensor, _duct, _input);
                                var paramShow = DataProcess.ParaShow(matchingResult, _inv, _sensor, _duct, _input);
                                OnMeasurePointCompleted?.Invoke(measurePoint, paramShow);
                                if (!currentRange.Any(m => m.k == _measures[i].k))
                                {
                                    currentRange.Add(_measures[i]);
                                }
                                completionSource.TrySetResult(true);
                            }
                        }
                        catch (Exception ex)
                        {
#if DEBUG
                            throw new BusinessException("Lỗi khi xử lý kết quả từ Simatic: " + ex.Message);
#else
                            return;
#endif
                        }
                    };

                    _watcher.EnableRaisingEvents = true;

                    // gửi dòng hiện tại
                    WriteMeasureRow(_exchangeFolder, _measures[i]);

                    // chờ kết quả hoặc timeout
                    var completedTask = await Task.WhenAny(completionSource.Task, Task.Delay(timeoutMs));

                    // Dispose watcher cho dòng này
                    _watcher.EnableRaisingEvents = false;
                    _watcher.Dispose();
                    _watcher = null;

                    if (completedTask != completionSource.Task)
                    {
                        // Timeout
                        _measures[i].F = MeasureStatus.Error;
                        OnSimaticResultReceived?.Invoke(_measures[i]);
                    }
                    await Task.Delay(500);
                }
                if (currentRange.Count > 0)
                {
                    var fitting = DataProcess.FittingFC(currentRange.Count, startIndex);
                    OnMeasureRangeCompleted?.Invoke(fitting, currentRange.LastOrDefault());
                    currentRange.Clear();
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

        private void WriteMeasureRow(string exchangeFolder, Measure param)
        {
            string exchangeFilePath = Path.Combine(exchangeFolder, $"1_T_OUT.{_fileFormat}");

            if (_fileFormat == "csv")
            {
                // Ghi dữ liệu vào file CSV
                using (var writer = new StreamWriter(exchangeFilePath, append: true))
                {
                    writer.WriteLine($"{param.k}({param.S}({param.CV}");
                }
            }
            else
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using (var package = new ExcelPackage(new FileInfo(exchangeFilePath)))
                {
                    var ws = package.Workbook.Worksheets["1_T_OUT"] ?? package.Workbook.Worksheets.Add("1_T_OUT");
                    int row = param.k;
                    ws.Cells[row, 1].Value = param.k;
                    ws.Cells[row, 2].Value = param.S;
                    ws.Cells[row, 3].Value = param.CV;
                    ws.Cells[ws.Dimension.Address].AutoFitColumns();
                    package.Save();
                }
            }
        }

        private async Task<List<Measure>> ReadSimaticResult(string filePath, CamBien sen)
        {
            var measureResults = new List<Measure>();
            int retryCount = 5;
            int delayMs = 200;

            for (int i = 0; i < retryCount; i++)
            {
                try
                {

                    if (_fileFormat == "csv")
                    {
                        using (var reader = new StreamReader(filePath))
                        {
                            string? line;
                            while ((line = await reader.ReadLineAsync()) != null)
                            {
                                var values = line.Split(' ');
                                var result = new Measure
                                {
                                    k = int.Parse(values[0]),
                                    S = float.Parse(values[1]),
                                    CV = float.Parse(values[2]),
                                    NhietDoMoiTruong_sen = CalcSimatic(sen.NhietDoMoiTruongMin, sen.NhietDoMoiTruongMax, float.Parse(values[3])),
                                    DoAm_sen = CalcSimatic(sen.DoAmMoiTruongMin, sen.DoAmMoiTruongMax, float.Parse(values[4])),
                                    ApSuatkhiQuyen_sen = CalcSimatic(sen.ApSuatKhiQuyenMin, sen.ApSuatKhiQuyenMax, float.Parse(values[5])),
                                    ChenhLechApSuat_sen = CalcSimatic(sen.ChenhLechApSuatMin, sen.ChenhLechApSuatMax, float.Parse(values[6])),
                                    ApSuatTinh_sen = CalcSimatic(sen.ApSuatTinhMin, sen.ApSuatTinhMax, float.Parse(values[7])),
                                    DoRung_sen = CalcSimatic(sen.DoRungMin, sen.DoRungMax, float.Parse(values[8])),
                                    DoOn_sen = CalcSimatic(sen.DoOnMin, sen.DoOnMax, float.Parse(values[9])),
                                    SoVongQuay_sen = CalcSimatic(sen.SoVongQuayMin, sen.SoVongQuayMax, float.Parse(values[10])),
                                    Momen_sen = CalcSimatic(sen.MomenMin, sen.MomenMax, float.Parse(values[11])),
                                    DongDien_fb = CalcSimatic(sen.PhanHoiDongDienMin, sen.PhanHoiDongDienMax, float.Parse(values[12])),
                                    CongSuat_fb = CalcSimatic(sen.PhanHoiCongSuatMin, sen.PhanHoiCongSuatMax, float.Parse(values[13])),
                                    ViTriVan_fb = CalcSimatic(sen.PhanHoiViTriVanMin, sen.PhanHoiViTriVanMax, float.Parse(values[14]))
                                };

                                measureResults.Add(result);
                            }
                        }
                    }
                    else
                    {

                        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                        using (var package = new ExcelPackage(new FileInfo(filePath)))
                        {
                            var ws = package.Workbook.Worksheets.FirstOrDefault();
                            if (ws == null) return measureResults;
                            for (int row = 1; row <= ws.Dimension.End.Row; row++)
                            {
                                var result = new Measure
                                {
                                    k = int.TryParse(ws.Cells[row, 1].Text, out var k) ? k : 0,
                                    S = float.TryParse(ws.Cells[row, 2].Text, out var s) ? s : 0,
                                    CV = float.TryParse(ws.Cells[row, 3].Text, out var cv) ? cv : 0,
                                    NhietDoMoiTruong_sen = CalcSimatic(sen.NhietDoMoiTruongMin, sen.NhietDoMoiTruongMax, float.TryParse(ws.Cells[row, 4].Text, out var pv1) ? pv1 : 0),
                                    DoAm_sen = CalcSimatic(sen.DoAmMoiTruongMin, sen.DoAmMoiTruongMax, float.TryParse(ws.Cells[row, 5].Text, out var pv2) ? pv2 : 0),
                                    ApSuatkhiQuyen_sen = CalcSimatic(sen.ApSuatKhiQuyenMin, sen.ApSuatKhiQuyenMax, float.TryParse(ws.Cells[row, 6].Text, out var pv3) ? pv3 : 0),
                                    ChenhLechApSuat_sen = CalcSimatic(sen.ChenhLechApSuatMin, sen.ChenhLechApSuatMax, float.TryParse(ws.Cells[row, 7].Text, out var pv4) ? pv4 : 0),
                                    ApSuatTinh_sen = CalcSimatic(sen.ApSuatTinhMin, sen.ApSuatTinhMax, float.TryParse(ws.Cells[row, 8].Text, out var pv5) ? pv5 : 0),
                                    DoRung_sen = CalcSimatic(sen.DoRungMin, sen.DoRungMax, float.TryParse(ws.Cells[row, 9].Text, out var pv6) ? pv6 : 0),
                                    DoOn_sen = CalcSimatic(sen.DoOnMin, sen.DoOnMax, float.TryParse(ws.Cells[row, 10].Text, out var pv7) ? pv7 : 0),
                                    SoVongQuay_sen = CalcSimatic(sen.SoVongQuayMin, sen.SoVongQuayMax, float.TryParse(ws.Cells[row, 11].Text, out var pv8) ? pv8 : 0),
                                    Momen_sen = CalcSimatic(sen.MomenMin, sen.MomenMax, float.TryParse(ws.Cells[row, 12].Text, out var pv9) ? pv9 : 0),
                                    DongDien_fb = CalcSimatic(sen.PhanHoiDongDienMin, sen.PhanHoiDongDienMax, float.TryParse(ws.Cells[row, 13].Text, out var pv10) ? pv10 : 0),
                                    CongSuat_fb = CalcSimatic(sen.PhanHoiCongSuatMin, sen.PhanHoiCongSuatMax, float.TryParse(ws.Cells[row, 14].Text, out var pv11) ? pv11 : 0),
                                    ViTriVan_fb = CalcSimatic(sen.PhanHoiViTriVanMin, sen.PhanHoiViTriVanMax, float.TryParse(ws.Cells[row, 15].Text, out var pv12) ? pv12 : 0),
                                    TanSo_fb = CalcSimatic(sen.PhanHoiTanSoMin, sen.PhanHoiTanSoMax, float.TryParse(ws.Cells[row, 2].Text, out var pv13) ? pv13 : 0), // sử dụng cột S để lấy tần số
                                };
                                measureResults.Add(result);
                            }
                        }
                    }
                    return measureResults;

                }
                catch (IOException)
                {
                    await Task.Delay(delayMs);
                }
                catch (Exception ex)
                {
                    throw new Exception($"Lỗi đọc file 2_S_IN.{_fileFormat}: {ex.Message}");
                }
            }
            throw new Exception($"Không thể đọc file 2_S_IN.{_fileFormat} sau nhiều lần thử lại (file có thể đang bị lock bởi process khác)");
        }

        private float CalcSimatic(float minValue, float maxValue, float percent)
        {
            return minValue + (maxValue - minValue) * percent / 100f;
        }

        public event Action<bool> OnSimaticConnectionChanged;
        public event Action<Measure> OnSimaticResultReceived;
        public event Action<List<Measure>> OnSimaticExchangeCompleted;
        public event Action<MeasureResponse, ParameterShow> OnMeasurePointCompleted;
        public event Action<MeasureFittingFC, Measure> OnMeasureRangeCompleted;
    }
}
