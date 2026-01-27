using OfficeOpenXml;
using System.IO;
using TESMEA_TMS.DTOs;
using TESMEA_TMS.Helpers;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Wordprocessing;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using A = DocumentFormat.OpenXml.Drawing;
using System.Data;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using Break = DocumentFormat.OpenXml.Wordprocessing.Break;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using OrientationValues = DocumentFormat.OpenXml.Drawing.Charts.OrientationValues;
using Drawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;



namespace TESMEA_TMS.Services
{
    public interface IFileService
    {
        ThongSoDauVao ImportCalculation(string filePath);
        Task ExportExcelTestResult(string outputPath, string option, ThongSoDauVao tsdv, ThongTinDuAn project);
        Task ExportReportTestResult(string outputPath, string option, ThongSoDauVao tsdv, ThongTinDuAn project);

        // Report từ phần mềm cũ - scada
        //void ExportDatabase(string filePath);
        //void Report_1(string filePath);
        //void Report_2(string filePath);
        //void Report_3(string filePath);
        //void Report_4(string filePath);
        //void Report_5(string filePath);
        //void Report_tester(string filePath);

    }

    public class KetQuaTheoTanSo
    {
        public double TanSo { get; set; }
        public List<KetQuaTaiDieuKienDoKiem> DieuKienDoKiem { get; set; }
        public List<HieuChuanVeDieuKienTieuChuan> HieuChuanTieuChuan { get; set; }
        public List<HieuChuanVeDieuKienLamviec> HieuChuanLamViec { get; set; }
    }

    public class FileService : IFileService
    {
        private readonly ICalculationService _calculationService;

        public FileService(ICalculationService calculationService)
        {
            _calculationService = calculationService;
        }

        #region Xuất kết quả
        public async Task ExportExcelTestResult(string outputPath, string option, ThongSoDauVao tsdv, ThongTinDuAn project)
        {
            try
            {
                if (Common.IsFileLocked(outputPath))
                {
                    MessageBoxHelper.ShowWarning("File đang được mở bởi ứng dụng khác. Vui lòng tắt file trước khi xuất báo cáo!");
                    return;
                }

                if (tsdv == null)
                {
                    MessageBoxHelper.ShowWarning("Dữ liệu đầu vào không hợp lệ");
                    return;
                }
                string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "ketquadokiem_template.xlsx");
                if (!File.Exists(templatePath))
                {
                    MessageBoxHelper.ShowWarning("File mẫu kết quả không tồn tại");
                    return;
                }


                var ketQuaDoKiem = new KetQuaDoKiem();
                ketQuaDoKiem = await _calculationService.CalcutationTestResultAsync(tsdv);


                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using (var package = new ExcelPackage(new FileInfo(templatePath)))
                {
                    var kqDoKiem = await ExportCalculation(package, tsdv, ketQuaDoKiem);
                    // Lấy danh sách các tần số duy nhất
                    var freqGroups = tsdv.DanhSachThongSoDoKiem
                        .Select((item, idx) => new { item.TanSo_fb, Index = idx })
                        .GroupBy(x => x.TanSo_fb)
                        .Select(g =>
                        {
                            var index = g.Select(x => x.Index).ToList();
                            return new KetQuaTheoTanSo
                            {
                                TanSo = g.Key,
                                DieuKienDoKiem = index
                                    .Select(idx => kqDoKiem.DanhSachketQuaTaiDieuKienDoKiem.ElementAtOrDefault(idx))
                                    .Where(x => x != null)
                                    .ToList(),
                                HieuChuanTieuChuan = index
                                    .Select(idx => kqDoKiem.DanhSachhieuChuanVeDieuKienTieuChuan.ElementAtOrDefault(idx))
                                    .Where(x => x != null)
                                    .ToList(),
                                HieuChuanLamViec = index
                                    .Select(idx => kqDoKiem.DanhSachhieuChuanVeDieuKienLamviec.ElementAtOrDefault(idx))
                                    .Where(x => x != null)
                                    .ToList()
                            };
                        })
                        .ToList();

                    foreach(var item in freqGroups)
                    {
                        if (option == "FULL")
                        {
                            await ExportDesignCondition(package, item.DieuKienDoKiem, project, $"{item.TanSo}Hz - Design Condition");
                            await ExportNormalizedCondition(package, item.DieuKienDoKiem, item.HieuChuanTieuChuan, project, $"{item.TanSo}Hz - Normalized Condition");
                            await ExportOperatingCondition(package, item.DieuKienDoKiem, item.HieuChuanLamViec, project, $"{item.TanSo}Hz - Operating Condition");
                            await ExportFullCondition(package, item.DieuKienDoKiem, item.HieuChuanTieuChuan, item.HieuChuanLamViec, project, $"{item.TanSo}Hz - Full");
                        }
                        else
                        {
                            // Chỉ giữ lại sheet Calculation và sheet của option
                            var keepSheets = new List<string> { "Calculation", "" };
                            if (option == "DESIGN") keepSheets[1] = "Design Condition";
                            else if (option == "NORMALIZED") keepSheets[1] = "Normalized Condition";
                            else if (option == "OPERATION") keepSheets[1] = "Operating Condition";

                            // Xóa các sheet không cần thiết
                            for (int i = package.Workbook.Worksheets.Count - 1; i >= 0; i--)
                            {
                                var sheet = package.Workbook.Worksheets[i];
                                if (!keepSheets.Contains(sheet.Name))
                                    package.Workbook.Worksheets.Delete(sheet.Name);
                            }

                            // Fill dữ liệu cho sheet option
                            if (option == "DESIGN")
                                await ExportDesignCondition(package, item.DieuKienDoKiem, project, $"{item.TanSo}Hz - Design Condition");
                            else if (option == "NORMALIZED")
                                await ExportNormalizedCondition(package, item.DieuKienDoKiem, item.HieuChuanTieuChuan, project, $"{item.TanSo}Hz - Normalized Condition");
                            else if (option == "OPERATION")
                                await ExportOperatingCondition(package, item.DieuKienDoKiem, item.HieuChuanLamViec, project, $"{item.TanSo}Hz - Operating Condition");
                        }
                        foreach (var ws in package.Workbook.Worksheets)
                        {
                            var logo = ws.Drawings["logo"] as OfficeOpenXml.Drawing.ExcelPicture;
                            if (logo != null)
                            {
                                logo.SetSize((int)(7 * 96), (int)(0.88 * 96));
                            }
                        }
                    }
                   

                    package.SaveAs(new FileInfo(outputPath));
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        #region excel handlers
        public ThongSoDauVao ImportCalculation(string filePath)
        {
            try
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var ws = package.Workbook.Worksheets["Calculation"] ?? package.Workbook.Worksheets.FirstOrDefault();
                    if (ws == null)
                        throw new BusinessException("Không tìm thấy worksheet phù hợp, vui lòng kiểm tra lại");

                    // helper parse cell to double, throw if fail
                    float ParseCell(int row, int col)
                    {
                        var cellValue = ws.Cells[row, col].Value;
                        if (cellValue == null || !float.TryParse(cellValue.ToString(), out float result))
                        {
                            throw new BusinessException($"Thông số đầu vào không phù hợp: (dòng {row}, cột {col})");
                        }
                        return result;
                    }
                    double piValue = 3.14;
                    double duongKinhOngD5 = ParseCell(19, 4);
                    double duongKinhOngGioD3 = ParseCell(21, 4);
                    double tietDienOngD5 = piValue * Math.Pow(duongKinhOngD5 / 1000, 2) / 4;
                    double tietDienOngGioD3 = piValue * Math.Pow(duongKinhOngGioD3 / 1000, 2) / 4;

                    var thongSoDuongOngGio = new ThongSoDuongOngGio
                    {
                        DuongKinhOngD5 = duongKinhOngD5,
                        ChieuDaiOngGioTonThatL = ParseCell(20, 4),
                        DuongKinhOngGioD3 = duongKinhOngGioD3,
                        TietDienOngD5 = tietDienOngD5,
                        HeSoMaSatOngK = ParseCell(20, 8),
                        TietDienOngGioD3 = tietDienOngGioD3,
                    };

                    var thongSoCoBanCuaQuat = new ThongSoCoBanCuaQuat
                    {
                        SoVongQuayCuaQuatNLT = ParseCell(23, 4),
                        CongSuatDongCo = ParseCell(24, 4),
                        HeSoDongCo = ParseCell(25, 4),
                        Tanso = ParseCell(26, 4),
                        HieuSuatDongCo = ParseCell(23, 8),
                        DoNhotKhongKhi = ParseCell(24, 8),
                        ApSuatKhiQuyen = ParseCell(25, 8),
                        NhietDoLamViec = ParseCell(26, 8)
                    };

                    var thongSoDoKiem = new List<Measure>();
                    var stt = 1;
                    for(int col = 3; col <= ws.Dimension.End.Column; col++)
                    {
                        var kiemTraSoCell = ws.Cells[30, col].Value;
                        if (kiemTraSoCell == null) continue;

                        var measure = new Measure
                        {
                            k = stt,
                            NhietDoMoiTruong_sen = ParseCell(31, col),
                            DoAm_sen = ParseCell(32, col),
                            ApSuatkhiQuyen_sen = ParseCell(33, col),
                            ChenhLechApSuat_sen = ParseCell(34, col),
                            ApSuatTinh_sen = ParseCell(35, col),
                            DoRung_sen = ParseCell(36, col),
                            DoOn_sen = ParseCell(37, col),
                            SoVongQuay_sen = ParseCell(38, col),
                            Momen_sen = ParseCell(39, col),
                            DongDien_fb = ParseCell(40, col),
                            DienAp_fb = ParseCell(41, col),
                            CongSuat_fb = ParseCell(42, col),
                            ViTriVan_fb = ParseCell(43, col),
                            TanSo_fb = ParseCell(44, col),
                        };
                        thongSoDoKiem.Add(measure);
                        stt++;
                    }


                    return new ThongSoDauVao
                    {
                        ThongSoDuongOngGio = thongSoDuongOngGio,
                        ThongSoCoBanCuaQuat = thongSoCoBanCuaQuat,
                        DanhSachThongSoDoKiem = thongSoDoKiem
                    };
                }
            }
            catch(BusinessException ex)
            { throw; }
            catch (Exception ex)
            {
                throw new Exception($"Thông số đầu vào không phù hợp: {ex.Message}");
            }
        }
        public async Task<KetQuaDoKiem> ExportCalculation(ExcelPackage package, ThongSoDauVao tsdv, KetQuaDoKiem ketQua)
        {
            try
            {
                KetQuaTaiDieuKienDoKiem thongSoKetQuaDoKiem = new KetQuaTaiDieuKienDoKiem();
                var ws = package.Workbook.Worksheets["Calculation"];
                if (ws == null) return null;

                // Thông số đường ống gió
                ws.Cells[19, 4].Value = tsdv.ThongSoDuongOngGio.DuongKinhOngD5;
                ws.Cells[20, 4].Value = tsdv.ThongSoDuongOngGio.ChieuDaiOngGioTonThatL;
                ws.Cells[21, 4].Value = tsdv.ThongSoDuongOngGio.DuongKinhOngGioD3;
                ws.Cells[19, 8].Value = tsdv.ThongSoDuongOngGio.TietDienOngD5;
                ws.Cells[20, 8].Value = tsdv.ThongSoDuongOngGio.HeSoMaSatOngK;
                ws.Cells[21, 8].Value = tsdv.ThongSoDuongOngGio.TietDienOngGioD3;

                // Thông số cơ bản của quạt
                ws.Cells[23, 4].Value = tsdv.ThongSoCoBanCuaQuat.SoVongQuayCuaQuatNLT;
                ws.Cells[24, 4].Value = tsdv.ThongSoCoBanCuaQuat.CongSuatDongCo;
                ws.Cells[25, 4].Value = tsdv.ThongSoCoBanCuaQuat.HeSoDongCo;
                ws.Cells[26, 4].Value = tsdv.ThongSoCoBanCuaQuat.Tanso;
                ws.Cells[23, 8].Value = tsdv.ThongSoCoBanCuaQuat.HieuSuatDongCo;
                ws.Cells[24, 8].Value = tsdv.ThongSoCoBanCuaQuat.DoNhotKhongKhi;
                ws.Cells[25, 8].Value = tsdv.ThongSoCoBanCuaQuat.ApSuatKhiQuyen;
                ws.Cells[26, 8].Value = tsdv.ThongSoCoBanCuaQuat.NhietDoLamViec;

                // Thông sô đo kiểm
                for (int i = 0; i < tsdv.DanhSachThongSoDoKiem.Count; i++)
                {
                    var item = tsdv.DanhSachThongSoDoKiem[i];
                    //int row = 29 + i;
                    int col = 3 + i;
                    ws.Cells[29, col].Value = item.k;
                    ws.Cells[30, col].Value = item.NhietDoMoiTruong_sen;
                    ws.Cells[31, col].Value = item.DoAm_sen;
                    ws.Cells[32, col].Value = item.ApSuatkhiQuyen_sen;
                    ws.Cells[33, col].Value = item.ChenhLechApSuat_sen;
                    ws.Cells[34, col].Value = item.ApSuatTinh_sen;
                    ws.Cells[35, col].Value = item.DoRung_sen;
                    ws.Cells[36, col].Value = item.DoOn_sen;
                    ws.Cells[37, col].Value = item.SoVongQuay_sen;
                    ws.Cells[38, col].Value = item.Momen_sen;
                    ws.Cells[39, col].Value = item.DongDien_fb;
                    ws.Cells[40, col].Value = item.CongSuat_fb;
                    ws.Cells[41, col].Value = item.ViTriVan_fb;
                    ws.Cells[42, col].Value = item.DienAp_fb;
                    ws.Cells[43, col].Value = item.TanSo_fb;
                }


                var kqtsDoKiem = ketQua.DanhSachketQuaTaiDieuKienDoKiem;
                
                #region export kết quả đo kiểm
                for (int i = 0; i < kqtsDoKiem.Count; i++)
                {
                    var item = kqtsDoKiem[i];
                    int col = 3 + i;
                    ws.Cells[49, col].Value = item.STT;
                    ws.Cells[50, col].Value = item.NhietDoBauUot;
                    ws.Cells[51, col].Value = item.ApSuatBaoHoaPsat;
                    ws.Cells[52, col].Value = item.ApSuatRiengPhanPv;
                    ws.Cells[53, col].Value = item.KLRMoiTruong;
                    ws.Cells[54, col].Value = item.XacDinhRW;
                    ws.Cells[55, col].Value = item.ApSuatTaiDiemDoChenhLechApSuatP5;
                    ws.Cells[56, col].Value = item.KLRTaiDiemDoLuuLuongPL5;
                    ws.Cells[57, col].Value = item.DoNhotKhongKhi;
                    //int row1 = 57 + i;
                    //B6
                    ws.Cells[58, col].Value = item.HeSoLuuLuong;
                    ws.Cells[59, col].Value = item.LuuLuongKhoiLuong;
                    ws.Cells[60, col].Value = item.LuuLuongTheTich;
                    ws.Cells[61, col].Value = item.KLRTaiDiemDoApSuatPL3;
                    ws.Cells[62, col].Value = item.LuuLuongTheTichTaiPL3;
                    ws.Cells[63, col].Value = item.LuuLuongTheTichTheoRPM;
                    ws.Cells[64, col].Value = item.HieuChinhLuuLuongTheTichTheoRPM;

                    //int row2 = 74 + i;
                    // B12
                    ws.Cells[65, col].Value = item.VanTocDongKhi;
                    ws.Cells[66, col].Value = item.ApSuatDong;
                    ws.Cells[67, col].Value = item.TonThatDuongOng;
                    ws.Cells[68, col].Value = item.ApSuatTinh;
                    ws.Cells[69, col].Value = item.ApSuatTong;
                    ws.Cells[70, col].Value = item.CongSuatDongCoTaiDieuKienDoKiem;
                    ws.Cells[71, col].Value = item.CongSuatDongCoThucTe;
                    ws.Cells[72, col].Value = item.HieuSuatTinh;
                    ws.Cells[73, col].Value = item.HieuSuatTong;
                }

                #endregion
                #region export hiệu chỉnh điều kiện tiêu chuẩn
                var kqhcTieuChuan = ketQua.DanhSachhieuChuanVeDieuKienTieuChuan;
                for (int i = 0; i < kqhcTieuChuan.Count; i++)
                {
                    var item = kqhcTieuChuan[i];
                    //int row = 88 + i;
                    int col = 3 + i;
                    ws.Cells[78, col].Value = item.STT;
                    ws.Cells[79, col].Value = item.LuuLuongTieuChuan_m3s;
                    ws.Cells[80, col].Value = item.LuuLuongTieuChuan_m3h;
                    ws.Cells[81, col].Value = item.ApSuatTinhTieuChuan;
                    ws.Cells[82, col].Value = item.ApSuatDongTieuChuan;
                    ws.Cells[83, col].Value = item.ApSuatTongTieuChuan;
                    ws.Cells[84, col].Value = item.CongSuatHapThuTieuChuan;
                    ws.Cells[85, col].Value = item.HieuSuatTinh;
                    ws.Cells[86, col].Value = item.HieuSuatTong;
                }
                #endregion
                #region export hiệu chỉnh điều kiện làm việc
                var kqhcLamViec = ketQua.DanhSachhieuChuanVeDieuKienLamviec;
                for (int i = 0; i < kqhcLamViec.Count; i++)
                {
                    var item = kqhcLamViec[i];
                    //int row = 103 + i;
                    int col = 3 + i;
                    ws.Cells[90, col].Value = item.STT;
                    ws.Cells[91, col].Value = item.KLRTaiDieuKienLamViec;
                    ws.Cells[92, col].Value = item.LuuLuongLamViec_m3s;
                    ws.Cells[93, col].Value = item.LuuLuongLamViec_m3h;
                    ws.Cells[94, col].Value = item.ApSuatTinhLamViec;
                    ws.Cells[95, col].Value = item.ApSuatDongLamViec;
                    ws.Cells[96, col].Value = item.ApSuatTongLamViec;
                    ws.Cells[97, col].Value = item.CongSuatHapThuLamViec;
                    ws.Cells[98, col].Value = item.HieuSuatTinh;
                    ws.Cells[99, col].Value = item.HieuSuatTong;
                }
                #endregion
                return new KetQuaDoKiem
                {
                    DanhSachketQuaTaiDieuKienDoKiem = kqtsDoKiem,
                    DanhSachhieuChuanVeDieuKienTieuChuan = kqhcTieuChuan,
                    DanhSachhieuChuanVeDieuKienLamviec = kqhcLamViec
                };
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public async Task ExportDesignCondition(ExcelPackage package, List<KetQuaTaiDieuKienDoKiem> data, ThongTinDuAn project, string sheetName = "Design Condition")
        {
            try
            {
                double xMaxValue_powerchart = 0, yMaxValue_powerchart = 0, xMaxValue_effchart = 0, yMaxValue_effchart = 0, xMaxValue_pressurechart = 0, yMaxValue_pressurechart = 0;

                var ws = package.Workbook.Worksheets["Design Condition"];
                if (ws == null)
                {
                    throw new Exception("Không tìm thấy worksheet 'Design Condition'.");
                }

                await FillThongTinChung(ws, project.ThongTinChung);

                for (int i = 0; i < data.Count; i++)
                {
                    var item = data.ElementAtOrDefault(i);
                    ws.Cells[19, 4 + i].Value = item?.STT;
                    ws.Cells[20, 4 + i].Value = item?.HieuChinhLuuLuongTheTichTheoRPM;
                    ws.Cells[21, 4 + i].Value = item?.ApSuatTinh;
                    ws.Cells[22, 4 + i].Value = item?.ApSuatTong;
                    ws.Cells[23, 4 + i].Value = item?.CongSuatDongCoThucTe;
                    ws.Cells[24, 4 + i].Value = item?.HieuSuatTinh;
                    ws.Cells[25, 4 + i].Value = item?.HieuSuatTong;


                    if (item != null)
                    {
                        // Power Chart
                        xMaxValue_powerchart = Math.Max(xMaxValue_powerchart, (double)item.HieuChinhLuuLuongTheTichTheoRPM);
                        yMaxValue_powerchart = Math.Max(yMaxValue_powerchart, (double)item.CongSuatDongCoThucTe);

                        // Efficiency Chart
                        xMaxValue_effchart = Math.Max(xMaxValue_effchart, (double)item.HieuChinhLuuLuongTheTichTheoRPM);
                        yMaxValue_effchart = Math.Max(yMaxValue_effchart, (double)item.HieuSuatTinh);

                        // Pressure Chart
                        xMaxValue_pressurechart = Math.Max(xMaxValue_pressurechart, (double)item.HieuChinhLuuLuongTheTichTheoRPM);
                        yMaxValue_pressurechart = Math.Max(yMaxValue_pressurechart, (double)item.ApSuatTinh);
                    }
                }

                UpdateCharts(ws, xMaxValue_powerchart, yMaxValue_powerchart, xMaxValue_effchart, yMaxValue_effchart, xMaxValue_pressurechart, yMaxValue_pressurechart);
            }
            catch (Exception ex)
            {
                throw new Exception($"Lỗi khi xuất Design worksheet: {ex.Message}");
            }

        }

        // Xuất worksheet Normalized Condition
        public async Task ExportNormalizedCondition(ExcelPackage package, List<KetQuaTaiDieuKienDoKiem> doKiem, List<HieuChuanVeDieuKienTieuChuan> tieuChuan, ThongTinDuAn project, string sheetName = "Normalized Condition")
        {
            try
            {
                double xMaxValue_powerchart = 0, yMaxValue_powerchart = 0, xMaxValue_effchart = 0, yMaxValue_effchart = 0, xMaxValue_pressurechart = 0, yMaxValue_pressurechart = 0;
                var ws = package.Workbook.Worksheets["Normalized Condition"];
                if (ws == null)
                {
                    throw new Exception("Không tìm thấy worksheet 'Normalized Condition'.");
                }
                await FillThongTinChung(ws, project.ThongTinChung);

                for (int i = 0; i < doKiem.Count; i++)
                {
                    var item = tieuChuan.ElementAtOrDefault(i);
                    var item1 = doKiem.ElementAtOrDefault(i);

                    ws.Cells[19, 4 + i].Value = item1?.STT;
                    ws.Cells[20, 4 + i].Value = item1?.HieuChinhLuuLuongTheTichTheoRPM;
                    ws.Cells[21, 4 + i].Value = item?.ApSuatTinhTieuChuan;
                    ws.Cells[22, 4 + i].Value = item?.ApSuatTongTieuChuan;
                    ws.Cells[23, 4 + i].Value = item?.CongSuatHapThuTieuChuan;
                    ws.Cells[24, 4 + i].Value = item?.HieuSuatTinh;
                    ws.Cells[25, 4 + i].Value = item?.HieuSuatTong;

                    if (item != null && item1 != null)
                    {
                        // Power Chart
                        xMaxValue_powerchart = Math.Max(xMaxValue_powerchart, (double)item1.HieuChinhLuuLuongTheTichTheoRPM);
                        yMaxValue_powerchart = Math.Max(yMaxValue_powerchart, (double)item.CongSuatHapThuTieuChuan);

                        // Efficiency Chart
                        xMaxValue_effchart = Math.Max(xMaxValue_effchart, (double)item1.HieuChinhLuuLuongTheTichTheoRPM);
                        yMaxValue_effchart = Math.Max(yMaxValue_effchart, (double)item.HieuSuatTinh);

                        // Pressure Chart
                        xMaxValue_pressurechart = Math.Max(xMaxValue_pressurechart, (double)item1.HieuChinhLuuLuongTheTichTheoRPM);
                        yMaxValue_pressurechart = Math.Max(yMaxValue_pressurechart, (double)item.ApSuatTinhTieuChuan);
                    }
                }

                UpdateCharts(ws, xMaxValue_powerchart, yMaxValue_powerchart, xMaxValue_effchart, yMaxValue_effchart, xMaxValue_pressurechart, yMaxValue_pressurechart);
            }
            catch (Exception ex)
            {
                throw new Exception($"Lỗi khi xuất Normalized worksheet: {ex.Message}");
            }


        }

        // Xuất worksheet Operating Condition
        public async Task ExportOperatingCondition(ExcelPackage package, List<KetQuaTaiDieuKienDoKiem> doKiem, List<HieuChuanVeDieuKienLamviec> lamViec, ThongTinDuAn project, string sheetName = "Operating Condition")
        {
            try
            {
                double xMaxValue_powerchart = 0, yMaxValue_powerchart = 0, xMaxValue_effchart = 0, yMaxValue_effchart = 0, xMaxValue_pressurechart = 0, yMaxValue_pressurechart = 0;
                var ws = package.Workbook.Worksheets["Operating Condition"];
                if (ws == null)
                {
                    throw new Exception("Không tìm thấy worksheet 'Operating Condition'.");
                }

                await FillThongTinChung(ws, project.ThongTinChung);

                for (int i = 0; i < doKiem.Count; i++)
                {
                    var item = lamViec.ElementAtOrDefault(i);
                    var item1 = doKiem.ElementAtOrDefault(i);
                    ws.Cells[19, 4 + i].Value = item?.STT;
                    ws.Cells[20, 4 + i].Value = item?.LuuLuongLamViec_m3h;
                    ws.Cells[21, 4 + i].Value = item?.ApSuatTinhLamViec;
                    ws.Cells[22, 4 + i].Value = item?.ApSuatTongLamViec;
                    ws.Cells[23, 4 + i].Value = item1?.CongSuatDongCoThucTe;
                    ws.Cells[24, 4 + i].Value = item?.HieuSuatTinh;
                    ws.Cells[25, 4 + i].Value = item?.HieuSuatTong;

                    if (item != null && item1 != null)
                    {
                        // Power Chart
                        xMaxValue_powerchart = Math.Max(xMaxValue_powerchart, (double)item.LuuLuongLamViec_m3h);
                        yMaxValue_powerchart = Math.Max(yMaxValue_powerchart, (double)item1.CongSuatDongCoThucTe);

                        // Efficiency Chart
                        xMaxValue_effchart = Math.Max(xMaxValue_effchart, (double)item.LuuLuongLamViec_m3h);
                        yMaxValue_effchart = Math.Max(yMaxValue_effchart, (double)item.HieuSuatTinh);

                        // Pressure Chart
                        xMaxValue_pressurechart = Math.Max(xMaxValue_pressurechart, (double)item.LuuLuongLamViec_m3h);
                        yMaxValue_pressurechart = Math.Max(yMaxValue_pressurechart, (double)item.ApSuatTinhLamViec);
                    }
                }

                UpdateCharts(ws, xMaxValue_powerchart, yMaxValue_powerchart, xMaxValue_effchart, yMaxValue_effchart, xMaxValue_pressurechart, yMaxValue_pressurechart);
            }
            catch (Exception ex)
            {
                throw new Exception($"Lỗi khi xuất Operating worksheet: {ex.Message}");
            }

        }

        public async Task ExportFullCondition(ExcelPackage package, List<KetQuaTaiDieuKienDoKiem> doKiem, List<HieuChuanVeDieuKienTieuChuan> tieuChuan, List<HieuChuanVeDieuKienLamviec> lamViec, ThongTinDuAn project, string sheetName = "Full")
        {
            try
            {
                double xMaxValue_powerchart = 0, yMaxValue_powerchart = 0, xMaxValue_effchart = 0, yMaxValue_effchart = 0, xMaxValue_pressurechart = 0, yMaxValue_pressurechart = 0;
                double xMaxValue1_powerchart = 0, yMaxValue1_powerchart = 0, xMaxValue1_effchart = 0, yMaxValue1_effchart = 0, xMaxValue1_pressurechart = 0, yMaxValue1_pressurechart = 0;

                var ws = package.Workbook.Worksheets["Full"];
                if (ws == null)
                {
                    throw new Exception("Không tìm thấy worksheet 'Full'.");
                }

                await FillThongTinChung(ws, project.ThongTinChung);

                for (int i = 0; i < doKiem.Count; i++)
                {
                    var item = doKiem.ElementAtOrDefault(i);
                    ws.Cells[20, 4 + i].Value = item?.HieuChinhLuuLuongTheTichTheoRPM;
                    ws.Cells[21, 4 + i].Value = item?.ApSuatTinh;
                    ws.Cells[22, 4 + i].Value = item?.CongSuatDongCoThucTe;
                    ws.Cells[23, 4 + i].Value = item?.HieuSuatTinh;
                }

                for (int i = 0; i < doKiem.Count; i++)
                {
                    var item = tieuChuan.ElementAtOrDefault(i);
                    var item1 = doKiem.ElementAtOrDefault(i);

                    ws.Cells[25, 4 + i].Value = item1?.HieuChinhLuuLuongTheTichTheoRPM;
                    ws.Cells[26, 4 + i].Value = item?.ApSuatTinhTieuChuan;
                    ws.Cells[27, 4 + i].Value = item?.CongSuatHapThuTieuChuan;
                    ws.Cells[28, 4 + i].Value = item?.HieuSuatTinh;


                    if (item != null && item1 != null)
                    {
                        // Power Chart
                        xMaxValue_powerchart = Math.Max(xMaxValue_powerchart, (double)item1.HieuChinhLuuLuongTheTichTheoRPM);
                        yMaxValue_powerchart = Math.Max(yMaxValue_powerchart, (double)item.CongSuatHapThuTieuChuan);
                        // Efficiency Chart
                        xMaxValue_effchart = Math.Max(xMaxValue_effchart, (double)item1.HieuChinhLuuLuongTheTichTheoRPM);
                        yMaxValue_effchart = Math.Max(yMaxValue_effchart, (double)item.HieuSuatTinh);
                        // Pressure Chart
                        xMaxValue_pressurechart = Math.Max(xMaxValue_pressurechart, (double)item1.HieuChinhLuuLuongTheTichTheoRPM);
                        yMaxValue_pressurechart = Math.Max(yMaxValue_pressurechart, (double)item.ApSuatTinhTieuChuan);
                    }
                }

                for (int i = 0; i < doKiem.Count; i++)
                {
                    var item = lamViec.ElementAtOrDefault(i);
                    var item1 = doKiem.ElementAtOrDefault(i);
                    ws.Cells[30, 4 + i].Value = item?.LuuLuongLamViec_m3h;
                    ws.Cells[31, 4 + i].Value = item?.ApSuatTinhLamViec;
                    ws.Cells[32, 4 + i].Value = item1?.CongSuatDongCoThucTe;
                    ws.Cells[33, 4 + i].Value = item?.HieuSuatTinh;

                    if (item != null && item1 != null)
                    {
                        // Power Chart
                        xMaxValue1_powerchart = Math.Max(xMaxValue_powerchart, (double)item.LuuLuongLamViec_m3h);
                        yMaxValue1_powerchart = Math.Max(yMaxValue_powerchart, (double)item1.CongSuatDongCoThucTe);
                        // Efficiency Chart
                        xMaxValue1_effchart = Math.Max(xMaxValue_effchart, (double)item.LuuLuongLamViec_m3h);
                        yMaxValue1_effchart = Math.Max(yMaxValue_effchart, (double)item.HieuSuatTinh);
                        // Pressure Chart
                        xMaxValue1_pressurechart = Math.Max(xMaxValue_pressurechart, (double)item.LuuLuongLamViec_m3h);
                        yMaxValue1_pressurechart = Math.Max(yMaxValue_pressurechart, (double)item.ApSuatTinhLamViec);
                    }
                }


                // Chart ranges
                string xRangeStandard = "D25:M25";
                string yPowerStandard = "D27:M27";
                string yEfficiencyStandard = "D28:M28";
                string yPressureStandard = "D26:M26";

                string xRangeOperating = "D30:M30";
                string yPowerOperating = "D32:M32";
                string yEfficiencyOperating = "D33:M33";
                string yPressureOperating = "D31:M31";

                // Power chart

                var total_xMaxValue_powerchart = (xMaxValue_powerchart + xMaxValue1_powerchart) / 2;
                var total_yMaxValue_powerchart = (yMaxValue_powerchart + yMaxValue1_powerchart) / 2;

                var total_xMaxValue_effchart = (xMaxValue_effchart + xMaxValue1_effchart) / 2;
                var total_yMaxValue_effchart = (yMaxValue_effchart + yMaxValue1_effchart) / 2;

                var total_xMavValue_pressurechart = (xMaxValue_pressurechart + xMaxValue1_pressurechart) / 2;
                var total_yMaxValue_pressurechart = (yMaxValue_pressurechart + yMaxValue1_pressurechart) / 2;

                var chartPower = ws.Drawings["Power"] as OfficeOpenXml.Drawing.Chart.ExcelChart;
                if (chartPower != null && chartPower.Series.Count >= 2)
                {
                    double xMaxValue = Common.AdaptiveRoundUpByLength(total_xMaxValue_powerchart);
                    double yMaxValue = Common.AdaptiveRoundUpByLength(total_yMaxValue_powerchart);
                    int xMajorUnit = Common.AdaptiveRoundUpByLength(xMaxValue / 5);
                    int xMinorUnit = Math.Max(1, xMajorUnit / 5);
                    int yMajorUnit = Common.AdaptiveRoundUpByLength(yMaxValue / 5);
                    int yMinorUnit = Math.Max(1, yMajorUnit / 5);

                    chartPower.Series[0].XSeries = xRangeStandard;
                    chartPower.Series[0].Series = yPowerStandard;
                    chartPower.Series[0].Header = "Điều kiện tiêu chuẩn";
                    chartPower.Series[1].XSeries = xRangeOperating;
                    chartPower.Series[1].Series = yPowerOperating;
                    chartPower.Series[1].Header = "Điều kiện làm việc";

                    // Set axis
                    chartPower.XAxis.MinValue = 0.0;
                    chartPower.XAxis.MaxValue = xMaxValue;
                    chartPower.XAxis.MajorUnit = xMajorUnit;
                    chartPower.XAxis.MinorUnit = xMinorUnit;

                    chartPower.YAxis.MinValue = 0.0;
                    chartPower.YAxis.MaxValue = yMaxValue;
                    chartPower.YAxis.MajorUnit = yMajorUnit;
                    chartPower.YAxis.MinorUnit = yMinorUnit;

                }

                // Efficiency chart
                var chartEfficiency = ws.Drawings["Efficiency"] as OfficeOpenXml.Drawing.Chart.ExcelChart;
                if (chartEfficiency != null && chartEfficiency.Series.Count >= 2)
                {
                    double xMaxValue = Common.AdaptiveRoundUpByLength(total_xMaxValue_effchart);
                    double yMaxValue = Common.AdaptiveRoundUpByLength(total_yMaxValue_effchart);
                    int xMajorUnit = Common.AdaptiveRoundUpByLength(xMaxValue / 5);
                    int xMinorUnit = Math.Max(1, xMajorUnit / 5);
                    int yMajorUnit = Common.AdaptiveRoundUpByLength(yMaxValue / 5);
                    int yMinorUnit = Math.Max(1, yMajorUnit / 5);

                    chartEfficiency.Series[0].XSeries = xRangeStandard;
                    chartEfficiency.Series[0].Series = yEfficiencyStandard;
                    chartEfficiency.Series[0].Header = "Điều kiện tiêu chuẩn";
                    chartEfficiency.Series[1].XSeries = xRangeOperating;
                    chartEfficiency.Series[1].Series = yEfficiencyOperating;
                    chartEfficiency.Series[1].Header = "Điều kiện làm việc";

                    // Set axis
                    chartEfficiency.XAxis.MinValue = 0.0;
                    chartEfficiency.XAxis.MaxValue = xMaxValue;
                    chartEfficiency.XAxis.MajorUnit = xMajorUnit;
                    chartEfficiency.XAxis.MinorUnit = xMinorUnit;

                    chartEfficiency.YAxis.MinValue = 0.0;
                    chartEfficiency.YAxis.MaxValue = yMaxValue;
                    chartEfficiency.YAxis.MajorUnit = yMajorUnit;
                    chartEfficiency.YAxis.MinorUnit = yMinorUnit;

                }

                // Pressure chart
                var chartPressure = ws.Drawings["Pressure"] as OfficeOpenXml.Drawing.Chart.ExcelChart;
                if (chartPressure != null && chartPressure.Series.Count >= 2)
                {
                    double xMaxValue = Common.AdaptiveRoundUpByLength(total_xMavValue_pressurechart);
                    double yMaxValue = Common.AdaptiveRoundUpByLength(total_yMaxValue_pressurechart);
                    int xMajorUnit = Common.AdaptiveRoundUpByLength(xMaxValue / 5);
                    int xMinorUnit = Math.Max(1, xMajorUnit / 5);
                    int yMajorUnit = Common.AdaptiveRoundUpByLength(yMaxValue / 5);
                    int yMinorUnit = Math.Max(1, yMajorUnit / 5);

                    chartPressure.Series[0].XSeries = xRangeStandard;
                    chartPressure.Series[0].Series = yPressureStandard;
                    chartPressure.Series[0].Header = "Điều kiện tiêu chuẩn";
                    chartPressure.Series[1].XSeries = xRangeOperating;
                    chartPressure.Series[1].Series = yPressureOperating;
                    chartPressure.Series[1].Header = "Điều kiện làm việc";

                    // Set axis
                    chartPressure.XAxis.MinValue = 0.0;
                    chartPressure.XAxis.MaxValue = xMaxValue;
                    chartPressure.XAxis.MajorUnit = xMajorUnit;
                    chartPressure.XAxis.MinorUnit = xMinorUnit;

                    chartPressure.YAxis.MinValue = 0.0;
                    chartPressure.YAxis.MaxValue = yMaxValue;
                    chartPressure.YAxis.MajorUnit = yMajorUnit;
                    chartPressure.YAxis.MinorUnit = yMinorUnit;

                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Lỗi khi xuất full worksheet: {ex.Message}");
            }

        }

        private Task FillThongTinChung(ExcelWorksheet ws, ThongTinChung info)
        {
            try
            {
                ws.Cells[12, 3].Value = info.TenMauThu;
                ws.Cells[13, 3].Value = info.CoSoSanXuat;
                ws.Cells[14, 3].Value = info.KyHieu;
                ws.Cells[15, 3].Value = info.SoLuongMau;
                ws.Cells[16, 3].Value = info.TinhTrangMau;

                ws.Cells[12, 12].Value = info.NgayNhanYeuCau;
                ws.Cells[13, 12].Value = info.NgayNhanMau;
                ws.Cells[14, 12].Value = info.NgayThuNghiem;
                ws.Cells[15, 12].Value = info.NgayHoanThanh;
                ws.Cells[16, 12].Value = info.TieuChuanApDung;
                return Task.CompletedTask;
            }
            catch (Exception ex)
            {
                throw new Exception($"Lỗi khi fill thông tin chung: {ex.Message}");
            }
        }

        private void UpdateCharts(ExcelWorksheet ws, double xMaxValue_powerchart, double yMaxValue_powerchart, double xMaxValue_effchart, double yMaxValue_effchart, double xMaxValue_pressurechart, double yMaxValue_pressurechart)
        {
            // Power chart

            var chartPower = ws.Drawings["Power"] as OfficeOpenXml.Drawing.Chart.ExcelChart;
            if (chartPower != null)
            {
                double xMaxValue = Common.AdaptiveRoundUpByLength(xMaxValue_powerchart);
                double yMaxValue = Common.AdaptiveRoundUpByLength(yMaxValue_powerchart);
                int xMajorUnit = Common.AdaptiveRoundUpByLength(xMaxValue / 5);
                int xMinorUnit = Math.Max(1, xMajorUnit / 5);
                int yMajorUnit = Common.AdaptiveRoundUpByLength(yMaxValue / 5);
                int yMinorUnit = Math.Max(1, yMajorUnit / 5);

                if (chartPower.Series.Count == 0)
                    chartPower.Series.Add("D23:M23", "D20:M20");
                chartPower.Series[0].XSeries = "D20:M20";
                chartPower.Series[0].Series = "D23:M23";

                chartPower.XAxis.MinValue = 0.0;
                chartPower.XAxis.MaxValue = xMaxValue;
                chartPower.XAxis.MajorUnit = xMajorUnit;
                chartPower.XAxis.MinorUnit = xMinorUnit;

                chartPower.YAxis.MinValue = 0.0;
                chartPower.YAxis.MaxValue = yMaxValue;
                chartPower.YAxis.MajorUnit = yMajorUnit;
                chartPower.YAxis.MinorUnit = yMinorUnit;
            }

            // Efficiency chart
            var chartEfficiency = ws.Drawings["Efficiency"] as OfficeOpenXml.Drawing.Chart.ExcelChart;
            if (chartEfficiency != null)
            {
                double xMaxValue = Common.AdaptiveRoundUpByLength(xMaxValue_effchart);
                double yMaxValue = Common.AdaptiveRoundUpByLength(yMaxValue_effchart);
                int xMajorUnit = Common.AdaptiveRoundUpByLength(xMaxValue_effchart / 5);
                int xMinorUnit = Math.Max(1, xMajorUnit / 5);
                int yMajorUnit = Common.AdaptiveRoundUpByLength(yMaxValue_effchart / 5);
                int yMinorUnit = Math.Max(1, yMajorUnit / 5);

                if (chartEfficiency.Series.Count == 0)
                    chartEfficiency.Series.Add("D24:M24", "D20:M20");
                chartEfficiency.Series[0].XSeries = "D20:M20";
                chartEfficiency.Series[0].Series = "D24:M24";

                chartEfficiency.XAxis.MinValue = 0.0;
                chartEfficiency.XAxis.MaxValue = xMaxValue;
                chartEfficiency.XAxis.MajorUnit = xMajorUnit;
                chartEfficiency.XAxis.MinorUnit = xMinorUnit;

                chartEfficiency.YAxis.MinValue = 0.0;
                chartEfficiency.YAxis.MaxValue = yMaxValue;
                chartEfficiency.YAxis.MajorUnit = yMajorUnit;
                chartEfficiency.YAxis.MinorUnit = yMinorUnit;

            }

            // Pressure chart
            var chartPressure = ws.Drawings["Pressure"] as OfficeOpenXml.Drawing.Chart.ExcelChart;
            if (chartPressure != null)
            {

                double xMaxValue = Common.AdaptiveRoundUpByLength(xMaxValue_pressurechart);
                double yMinor = Common.AdaptiveRoundUpByLength(yMaxValue_pressurechart);
                int xMajorUnit = Common.AdaptiveRoundUpByLength(xMaxValue_pressurechart / 5);
                int yMaxValue = Math.Max(1, xMajorUnit / 5);
                int yMajorUnit = Common.AdaptiveRoundUpByLength(yMaxValue_pressurechart / 5);
                int yMinorUnit = Math.Max(1, yMajorUnit / 5);

                if (chartPressure.Series.Count == 0)
                    chartPressure.Series.Add("D21:M21", "D20:M20");
                chartPressure.Series[0].XSeries = "D20:M20";
                chartPressure.Series[0].Series = "D21:M21";

                chartPressure.XAxis.MinValue = 0.0;
                chartPressure.XAxis.MaxValue = xMaxValue;
                chartPressure.XAxis.MajorUnit = xMajorUnit;
                chartPressure.XAxis.MinorUnit = yMaxValue;

                chartPressure.YAxis.MinValue = 0.0;
                chartPressure.YAxis.MaxValue = yMinor;
                chartPressure.YAxis.MajorUnit = yMajorUnit;
                chartPressure.YAxis.MinorUnit = yMinorUnit;

            }
        }

        #endregion

        #endregion

        #region Xuất báo cáo
        public async Task ExportReportTestResult(string outputPath, string option, ThongSoDauVao tsdv, ThongTinDuAn project)
        {
            try
            {

                if (Common.IsFileLocked(outputPath))
                {
                    MessageBoxHelper.ShowWarning("File đang được mở bởi ứng dụng khác. Vui lòng tắt file trước khi xuất báo cáo!");
                    return;
                }

                if (tsdv == null)
                {
                    MessageBoxHelper.ShowWarning("Dữ liệu đầu vào không hợp lệ");
                    return;
                }


                string basePath = AppDomain.CurrentDomain.BaseDirectory;
                string templatePath = option == "FULL" ? Path.Combine(basePath, "Templates", "baocaodokiem_template_full.docx") : Path.Combine(basePath, "Templates", "baocaodokiem_template.docx");
                if (!File.Exists(templatePath))
                {
                    MessageBoxHelper.ShowWarning("File mẫu báo cáo không tồn tại");
                    return;
                }

                var kqdk = await _calculationService.CalcutationTestResultAsync(tsdv);
                BaoCao input = new BaoCao();
                input.ThongTinDuAn = project;
                switch (option)
                {
                    case "DESIGN":
                        input.BangKetQuaThuNghiem = kqdk.DanhSachketQuaTaiDieuKienDoKiem
                            ?.Select((x, idx) => new BangKetQuaThuNghiem(
                                (idx + 1).ToString(),
                                x.HieuChinhLuuLuongTheTichTheoRPM.ToString() ?? "",
                                x.ApSuatTinh.ToString() ?? "",
                                x.ApSuatTong.ToString() ?? "",
                                x.CongSuatDongCoThucTe.ToString() ?? "",
                                x.HieuSuatTinh.ToString() ?? "",
                                x.HieuSuatTong.ToString() ?? ""
                            )).ToList() ?? new List<BangKetQuaThuNghiem>();
                        break;
                    case "NORMALIZED":
                        input.BangKetQuaThuNghiem = kqdk.DanhSachhieuChuanVeDieuKienTieuChuan
                            ?.Select((x, idx) => new BangKetQuaThuNghiem(
                                (idx + 1).ToString(),
                                (kqdk.DanhSachketQuaTaiDieuKienDoKiem != null && kqdk.DanhSachketQuaTaiDieuKienDoKiem.Count > idx
                                    ? kqdk.DanhSachketQuaTaiDieuKienDoKiem[idx].HieuChinhLuuLuongTheTichTheoRPM.ToString()
                                    : "") ?? "",
                                x.ApSuatDongTieuChuan.ToString() ?? "",
                                x.ApSuatTongTieuChuan.ToString() ?? "",
                                x.CongSuatHapThuTieuChuan.ToString() ?? "",
                                x.HieuSuatTinh.ToString() ?? "",
                                x.HieuSuatTong.ToString() ?? ""
                            )).ToList() ?? new List<BangKetQuaThuNghiem>();
                        break;
                    case "OPERATING":
                        input.BangKetQuaThuNghiem = kqdk.DanhSachhieuChuanVeDieuKienLamviec
                            ?.Select((x, idx) => new BangKetQuaThuNghiem(
                                (idx + 1).ToString(),
                                x.LuuLuongLamViec_m3h.ToString() ?? "",
                                x.ApSuatTinhLamViec.ToString() ?? "",
                                x.ApSuatTongLamViec.ToString() ?? "",
                                (kqdk.DanhSachketQuaTaiDieuKienDoKiem != null && kqdk.DanhSachketQuaTaiDieuKienDoKiem.Count > idx
                                    ? kqdk.DanhSachketQuaTaiDieuKienDoKiem[idx].CongSuatDongCoThucTe.ToString()
                                    : "") ?? "",
                                x.HieuSuatTinh.ToString() ?? "",
                                x.HieuSuatTong.ToString() ?? ""
                            )).ToList() ?? new List<BangKetQuaThuNghiem>();
                        break;
                    default:
                        throw new Exception("Không có dữ liệu để xuất báo cáo");
                }

                if (input == null)
                {
                    throw new Exception("Báo cáo không có dữ liệu");
                }

                File.Copy(templatePath, outputPath, true);

                var data = new Dictionary<string, string>();
                if (input.ThongTinDuAn != null)
                {
                    foreach (var prop in typeof(ThongTinDuAn).GetProperties())
                    {
                        data[prop.Name] = prop.GetValue(input.ThongTinDuAn.ThongTinChung).ToString() ?? "";
                    }

                    foreach (var prop1 in typeof(ThongTinMauThuNghiem).GetProperties())
                    {
                        data[prop1.Name] = prop1.GetValue(input.ThongTinDuAn.ThongTinMauThuNghiem).ToString() ?? "";
                    }
                }

                using (var doc = WordprocessingDocument.Open(outputPath, true))
                {
                    // Fill bookmark
                    var bookmarks = doc.MainDocumentPart.RootElement.Descendants<BookmarkStart>();
                    foreach (var bm in bookmarks)
                    {
                        if (data.TryGetValue(bm.Name, out var value))
                        {
                            OpenXmlElement current = bm.NextSibling();
                            while (current != null && !current.Descendants<Text>().Any())
                                current = current.NextSibling();

                            var textElement = current?.Descendants<Text>().FirstOrDefault();
                            if (textElement != null)
                            {
                                textElement.Text = value;
                            }
                            else
                            {
                                var run = new Run(new Text(value));
                                bm.Parent.InsertAfter(run, bm);
                            }
                        }
                    }

                    var tables = doc.MainDocumentPart.Document.Body.Elements<Table>().ToList();

                    if (tables.Count > 2)
                    {
                        var table = tables[2];
                        var templateRow = table.Elements<TableRow>().ElementAt(3);

                        int rowCount = input?.BangKetQuaThuNghiem.Count ?? 0;
                        for (int i = 0; i < rowCount; i++)
                        {
                            var kq = input?.BangKetQuaThuNghiem[i];
                            var newRow = (TableRow)templateRow.CloneNode(true);
                            var cells = newRow.Elements<TableCell>().ToList();

                            cells[0].RemoveAllChildren<Paragraph>();
                            cells[0].Append(CreateCenteredParagraph(kq.STT));
                            cells[1].RemoveAllChildren<Paragraph>();
                            cells[1].Append(CreateCenteredParagraph(Math.Round(double.TryParse(kq.LuuLuong, out var v1) ? v1 : 0, 0).ToString()));
                            cells[2].RemoveAllChildren<Paragraph>();
                            cells[2].Append(CreateCenteredParagraph(Math.Round(double.TryParse(kq.ApSuatTinh, out var v2) ? v2 : 0, 0).ToString()));
                            cells[3].RemoveAllChildren<Paragraph>();
                            cells[3].Append(CreateCenteredParagraph(Math.Round(double.TryParse(kq.ApSuatTong, out var v3) ? v3 : 0, 0).ToString()));
                            cells[4].RemoveAllChildren<Paragraph>();
                            cells[4].Append(CreateCenteredParagraph(Math.Round(double.TryParse(kq.CongSuatTieuThu, out var v4) ? v4 : 0, 2).ToString()));
                            cells[5].RemoveAllChildren<Paragraph>();
                            cells[5].Append(CreateCenteredParagraph(Math.Round(double.TryParse(kq.HieuSuatTinh, out var v5) ? v5 : 0, 2).ToString()));
                            cells[6].RemoveAllChildren<Paragraph>();
                            cells[6].Append(CreateCenteredParagraph(Math.Round(double.TryParse(kq.HieuSuatTong, out var v6) ? v6 : 0, 2).ToString()));
                            table.AppendChild(newRow);
                        }
                        table.RemoveChild(templateRow);
                    }

                    // drawing chart
                    var body = doc.MainDocumentPart.Document.Body;
                    body.Append(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));
                    InsertScatterChart(doc, input.BangKetQuaThuNghiem, "LuuLuong", "CongSuatTieuThu", "Lưu lượng (m3/h)", "Công suất (kW)", 0.0, 50000.0, 10000.0, 0.0, 40.0, 5.0);
                    InsertScatterChart(doc, input.BangKetQuaThuNghiem, "LuuLuong", "HieuSuatTinh", "Lưu lượng (m3/h)", "Hiệu suất tĩnh (%)", 0.0, 50000.0, 10000.0, 0.0, 90.0, 10.0);
                    InsertScatterChart(doc, input.BangKetQuaThuNghiem, "LuuLuong", "ApSuatTinh", "Lưu lượng (m3/h)", "Áp suất tĩnh (Pa)", 0.0, 50000.0, 10000.0, 0.0, 5000.0, 500.0);

                    doc.MainDocumentPart.Document.Save();
                }
                return;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        #region report handler
        private Paragraph CreateCenteredParagraph(string text)
        {
            return new Paragraph(
                new ParagraphProperties(
                    new Justification() { Val = JustificationValues.Center }
                ),
                new Run(new Text(text))
            );
        }

        private void InsertScatterChart(
            WordprocessingDocument doc,
            List<BangKetQuaThuNghiem> data,
            string xField,
            string yField,
            string xTitle,
            string yTitle,
            double xMin,
            double xMax,
            double xMajor,
            double yMin,
            double yMax,
            double yMajor
        )
        {
            var mainPart = doc.MainDocumentPart;
            var chartPart = mainPart.AddNewPart<ChartPart>();
            string chartPartId = mainPart.GetIdOfPart(chartPart);

            var xValues = data.Select(d => double.TryParse(d.GetType().GetProperty(xField)?.GetValue(d)?.ToString(), out var v) ? v : 0).ToList();
            var yValues = data.Select(d => double.TryParse(d.GetType().GetProperty(yField)?.GetValue(d)?.ToString(), out var v) ? v : 0).ToList();

            var xNumberLiteral = new NumberLiteral(
                new FormatCode("General"),
                new PointCount() { Val = (uint)xValues.Count }
            );
            foreach (var (v, i) in xValues.Select((v, i) => (v, i)))
            {
                xNumberLiteral.Append(new NumericPoint() { Index = (uint)i, NumericValue = new NumericValue(v.ToString()) });
            }

            var yNumberLiteral = new NumberLiteral(
                new FormatCode("General"),
                new PointCount() { Val = (uint)yValues.Count }
            );
            foreach (var (v, i) in yValues.Select((v, i) => (v, i)))
            {
                yNumberLiteral.Append(new NumericPoint() { Index = (uint)i, NumericValue = new NumericValue(v.ToString()) });
            }

            // Tạo ScatterChart
            var scatterChart = new ScatterChart(
                new ScatterStyle() { Val = ScatterStyleValues.LineMarker },
                new ScatterChartSeries(
                    new A.Charts.Index() { Val = (uint)0 },
                    new Order() { Val = (uint)0 },
                    new SeriesText(new NumericValue() { Text = "Chart" }),
                    new XValues(xNumberLiteral),
                    new YValues(yNumberLiteral),
                    new Marker(
                        new Symbol() { Val = MarkerStyleValues.Circle },
                         new A.Charts.Size() { Val = 7 },
                        new ChartShapeProperties(
        new A.SolidFill(new A.RgbColorModelHex() { Val = "000000" }),
        new A.Outline(new A.SolidFill(new A.RgbColorModelHex() { Val = "000000" }))
    )
                    ),
                    new ChartShapeProperties(
                        new A.Outline(
                            new A.SolidFill(new A.RgbColorModelHex() { Val = "000000" })
                        )
                    )
                ),
                new AxisId() { Val = 48650112u },
                new AxisId() { Val = 48672768u }
            );

            // Tiêu đề trục X
            var xAxisTitle = new Title(
                new ChartText(
                    new RichText(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(
                           new A.Run(
                    new A.RunProperties
                    {
                        Language = "vi-VN",
                        FontSize = 1100,
                        Bold = true
                    },
                    new A.Text() { Text = Common.ToSuperscriptUnit(xTitle) }
                )
                        )
                    )
                ),
                new Overlay() { Val = false }
            );

            // Tiêu đề trục Y
            var yAxisTitle = new Title(
                new ChartText(
                    new RichText(
                        new A.BodyProperties() { Rotation = -5400000 },

                        new A.ListStyle(),
                        new A.Paragraph(
                           new A.Run(
                    new A.RunProperties
                    {
                        Language = "vi-VN",
                        FontSize = 1100,
                        Bold = true
                    },
                    new A.Text() { Text = Common.ToSuperscriptUnit(yTitle) }
                )
                        )
                    )
                ),
                new Overlay() { Val = false }
            );

            var catAx = new ValueAxis(
               new AxisId() { Val = 48650112u },
               new Scaling(
                   new A.Charts.Orientation() { Val = OrientationValues.MinMax },
                   new MinAxisValue() { Val = xMin },
                   new MaxAxisValue() { Val = xMax }
               ),
               new Delete() { Val = false },
               new AxisPosition() { Val = AxisPositionValues.Bottom },
               new MajorGridlines(),
               new A.Charts.NumberingFormat() { FormatCode = "General", SourceLinked = true },
               new MajorUnit() { Val = xMajor },
               new TickLabelPosition() { Val = TickLabelPositionValues.NextTo },
               new CrossingAxis() { Val = 48672768u },
               new Crosses() { Val = CrossesValues.AutoZero },
               new CrossBetween() { Val = CrossBetweenValues.Between },
               xAxisTitle
           );

            var valAx = new ValueAxis(
                new AxisId() { Val = 48672768u },
                new Scaling(
                    new A.Charts.Orientation() { Val = OrientationValues.MinMax },
                    new MinAxisValue() { Val = yMin },
                    new MaxAxisValue() { Val = yMax }
                ),
                new Delete() { Val = false },
                new AxisPosition() { Val = AxisPositionValues.Left },
                new MajorGridlines(),
                new A.Charts.NumberingFormat() { FormatCode = "General", SourceLinked = true },
                new MajorUnit() { Val = yMajor },
                new TickLabelPosition() { Val = TickLabelPositionValues.NextTo },
                new CrossingAxis() { Val = 48650112u },
                new Crosses() { Val = CrossesValues.AutoZero },
                new CrossBetween() { Val = CrossBetweenValues.Between },
                yAxisTitle
            );

            var chart = new Chart(
                  new AutoTitleDeleted() { Val = true },
                  new PlotArea(scatterChart, catAx, valAx),
                  new PlotVisibleOnly() { Val = true }
              );

            chartPart.ChartSpace = new ChartSpace(chart);

            var drawing = new Drawing(
                new Inline(
                    new Extent() { Cx = 5486400, Cy = 3200400 },
                    new EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                    new DocProperties() { Id = (UInt32Value)1U, Name = "Chart name" },
                    new NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks() { NoChangeAspect = true }),
                    new A.Graphic(
                        new A.GraphicData(
                            new ChartReference() { Id = chartPartId }
                        )
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }
                    )
                )
            );
            var centeredChartParagraph = new Paragraph(
                new ParagraphProperties(
                    new Justification() { Val = JustificationValues.Center }
                ),
                drawing
            );

            var body = mainPart.Document.Body;
            body.Append(centeredChartParagraph);
        }

        public Task ExportReport_full(string templatePath, string outputPath, ThongTinDuAn project, KetQuaDoKiem kqdk)
        {
            try
            {
                if (Common.IsFileLocked(outputPath))
                {
                    throw new Exception("File đang được mở bởi ứng dụng khác. Vui lòng tắt file trước khi xuất báo cáo!");
                }

                File.Copy(templatePath, outputPath, true);

                // Chuẩn bị dữ liệu cho 3 bảng
                var bangThietKe = kqdk.DanhSachketQuaTaiDieuKienDoKiem
                    ?.Select((x, idx) => new BangKetQuaThuNghiem(
                        (idx + 1).ToString(),
                        x.HieuChinhLuuLuongTheTichTheoRPM.ToString() ?? "",
                        x.ApSuatTinh.ToString() ?? "",
                        "", // Không dùng ApSuatTong cho bảng này
                        x.CongSuatDongCoThucTe.ToString() ?? "",
                        x.HieuSuatTinh.ToString() ?? "",
                        "" // Không dùng HieuSuatTong cho bảng này
                    )).ToList() ?? new List<BangKetQuaThuNghiem>();

                var bangTieuChuan = kqdk.DanhSachhieuChuanVeDieuKienTieuChuan
                    ?.Select((x, idx) => new BangKetQuaThuNghiem(
                        (idx + 1).ToString(),
                        x.LuuLuongTieuChuan_m3h.ToString() ?? "",
                        x.ApSuatTinhTieuChuan.ToString() ?? "",
                        "", // Không dùng ApSuatTong cho bảng này
                        x.CongSuatHapThuTieuChuan.ToString() ?? "",
                        x.HieuSuatTinh.ToString() ?? "",
                        "" // Không dùng HieuSuatTong cho bảng này
                    )).ToList() ?? new List<BangKetQuaThuNghiem>();

                var bangLamViec = kqdk.DanhSachhieuChuanVeDieuKienLamviec
                    ?.Select((x, idx) => new BangKetQuaThuNghiem(
                        (idx + 1).ToString(),
                        x.LuuLuongLamViec_m3h.ToString() ?? "",
                        x.ApSuatTinhLamViec.ToString() ?? "",
                        "", // Không dùng ApSuatTong cho bảng này
                        x.CongSuatHapThuLamViec.ToString() ?? "",
                        x.HieuSuatTinh.ToString() ?? "",
                        "" // Không dùng HieuSuatTong cho bảng này
                    )).ToList() ?? new List<BangKetQuaThuNghiem>();

                using (var doc = WordprocessingDocument.Open(outputPath, true))
                {
                    // fill thông tin chung
                    var data = new Dictionary<string, string>();
                    if (project != null)
                    {
                        if (project.ThongTinChung != null)
                        {
                            foreach (var prop in typeof(ThongTinDuAn).GetProperties())
                            {
                                var value = prop.GetValue(project.ThongTinChung);
                                data[prop.Name] = value != null ? value.ToString() : "";
                            }
                        }
                        if (project.ThongTinMauThuNghiem != null)
                        {
                            foreach (var prop in typeof(ThongTinMauThuNghiem).GetProperties())
                            {
                                var value = prop.GetValue(project.ThongTinMauThuNghiem);
                                data[prop.Name] = value != null ? value.ToString() : "";
                            }
                        }
                    }

                    var bookmarks = doc.MainDocumentPart.RootElement.Descendants<BookmarkStart>();
                    foreach (var bm in bookmarks)
                    {
                        if (data.TryGetValue(bm.Name, out var value))
                        {
                            OpenXmlElement current = bm.NextSibling();
                            while (current != null && !current.Descendants<Text>().Any())
                                current = current.NextSibling();

                            var textElement = current?.Descendants<Text>().FirstOrDefault();
                            if (textElement != null)
                            {
                                textElement.Text = value;
                            }
                            else
                            {
                                var run = new Run(new Text(value));
                                bm.Parent.InsertAfter(run, bm);
                            }
                        }
                    }


                    var tables = doc.MainDocumentPart.Document.Body.Elements<Table>().ToList();

                    // Fill bảng thiết kế
                    if (tables.Count > 2)
                        FillTable(tables[2], bangThietKe);

                    // Fill bảng tiêu chuẩn
                    if (tables.Count > 3)
                        FillTable(tables[3], bangTieuChuan);

                    // Fill bảng làm việc
                    if (tables.Count > 4)
                        FillTable(tables[4], bangLamViec);

                    // Vẽ chart cho từng bảng (chỉ lấy dữ liệu thiết kế)
                    var body = doc.MainDocumentPart.Document.Body;
                    body.Append(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));
                    InsertScatterChart_full(doc, bangTieuChuan, bangLamViec, "LuuLuong", "CongSuatTieuThu", "Lưu lượng (m3/h)", "Công suất (kW)", 0.0, 50000.0, 10000.0, 0.0, 40.0, 5.0);
                    InsertScatterChart_full(doc, bangTieuChuan, bangLamViec, "LuuLuong", "HieuSuatTinh", "Lưu lượng (m3/h)", "Hiệu suất tĩnh (%)", 0.0, 50000.0, 10000.0, 0.0, 90.0, 10.0);
                    InsertScatterChart_full(doc, bangTieuChuan, bangLamViec, "LuuLuong", "ApSuatTinh", "Lưu lượng (m3/h)", "Áp suất tĩnh (Pa)", 0.0, 50000.0, 10000.0, 0.0, 5000.0, 500.0);

                    doc.MainDocumentPart.Document.Save();
                }
                return Task.CompletedTask;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        private void FillTable(Table table, List<BangKetQuaThuNghiem> data)
        {
            var templateRow = table.Elements<TableRow>().ElementAt(3);
            for (int i = 0; i < data.Count; i++)
            {
                var kq = data[i];
                var newRow = (TableRow)templateRow.CloneNode(true);
                var cells = newRow.Elements<TableCell>().ToList();


                cells[0].RemoveAllChildren<Paragraph>();
                cells[0].Append(CreateCenteredParagraph(kq.STT));
                cells[1].RemoveAllChildren<Paragraph>();
                cells[1].Append(CreateCenteredParagraph(Math.Round(double.TryParse(kq.LuuLuong, out var v1) ? v1 : 0, 0).ToString()));
                cells[2].RemoveAllChildren<Paragraph>();
                cells[2].Append(CreateCenteredParagraph(Math.Round(double.TryParse(kq.ApSuatTinh, out var v2) ? v2 : 0, 0).ToString()));
                cells[3].RemoveAllChildren<Paragraph>();
                cells[3].Append(CreateCenteredParagraph(Math.Round(double.TryParse(kq.CongSuatTieuThu, out var v3) ? v3 : 0, 2).ToString()));
                cells[4].RemoveAllChildren<Paragraph>();
                cells[4].Append(CreateCenteredParagraph(Math.Round(double.TryParse(kq.HieuSuatTinh, out var v4) ? v4 : 0, 2).ToString()));
                table.AppendChild(newRow);
            }
            table.RemoveChild(templateRow);
        }
        private void InsertScatterChart_full(
             WordprocessingDocument doc,
             List<BangKetQuaThuNghiem> dataTieuChuan,
             List<BangKetQuaThuNghiem> dataLamViec,
             string xField,
             string yField,
             string xTitle,
             string yTitle,
             double xMin,
             double xMax,
             double xMajor,
             double yMin,
             double yMax,
             double yMajor
         )
        {
            var mainPart = doc.MainDocumentPart;
            var chartPart = mainPart.AddNewPart<ChartPart>();
            string chartPartId = mainPart.GetIdOfPart(chartPart);

            // Series 1: Tiêu chuẩn
            var xValues1 = dataTieuChuan.Select(d => double.TryParse(d.GetType().GetProperty(xField)?.GetValue(d)?.ToString(), out var v) ? v : 0).ToList();
            var yValues1 = dataTieuChuan.Select(d => double.TryParse(d.GetType().GetProperty(yField)?.GetValue(d)?.ToString(), out var v) ? v : 0).ToList();

            var xNumberLiteral1 = new NumberLiteral(new FormatCode("General"), new PointCount() { Val = (uint)xValues1.Count });
            foreach (var (v, i) in xValues1.Select((v, i) => (v, i)))
                xNumberLiteral1.Append(new NumericPoint() { Index = (uint)i, NumericValue = new NumericValue(v.ToString()) });

            var yNumberLiteral1 = new NumberLiteral(new FormatCode("General"), new PointCount() { Val = (uint)yValues1.Count });
            foreach (var (v, i) in yValues1.Select((v, i) => (v, i)))
                yNumberLiteral1.Append(new NumericPoint() { Index = (uint)i, NumericValue = new NumericValue(v.ToString()) });

            // Series 2: Làm việc
            var xValues2 = dataLamViec.Select(d => double.TryParse(d.GetType().GetProperty(xField)?.GetValue(d)?.ToString(), out var v) ? v : 0).ToList();
            var yValues2 = dataLamViec.Select(d => double.TryParse(d.GetType().GetProperty(yField)?.GetValue(d)?.ToString(), out var v) ? v : 0).ToList();

            var xNumberLiteral2 = new NumberLiteral(new FormatCode("General"), new PointCount() { Val = (uint)xValues2.Count });
            foreach (var (v, i) in xValues2.Select((v, i) => (v, i)))
                xNumberLiteral2.Append(new NumericPoint() { Index = (uint)i, NumericValue = new NumericValue(v.ToString()) });

            var yNumberLiteral2 = new NumberLiteral(new FormatCode("General"), new PointCount() { Val = (uint)yValues2.Count });
            foreach (var (v, i) in yValues2.Select((v, i) => (v, i)))
                yNumberLiteral2.Append(new NumericPoint() { Index = (uint)i, NumericValue = new NumericValue(v.ToString()) });

            var scatterChart = new ScatterChart(
                new ScatterStyle() { Val = ScatterStyleValues.LineMarker },
                // Series 1: Tiêu chuẩn
                new ScatterChartSeries(
                    new A.Charts.Index() { Val = (uint)0 },
                    new Order() { Val = (uint)0 },
                    new SeriesText(new NumericValue() { Text = "Điều kiện tiêu chuẩn" }),
                    new XValues(xNumberLiteral1),
                    new YValues(yNumberLiteral1),
                    new Marker(
                        new Symbol() { Val = MarkerStyleValues.Circle },
                        new A.Charts.Size() { Val = 7 },
                       new ChartShapeProperties(
        new A.SolidFill(new A.RgbColorModelHex() { Val = "000000" }),
        new A.Outline(new A.SolidFill(new A.RgbColorModelHex() { Val = "000000" }))
    )
                    ),
                    new ChartShapeProperties(
                        new A.Outline(new A.SolidFill(new A.RgbColorModelHex() { Val = "000000" }))
                    )
                ),
                // Series 2: Làm việc
                new ScatterChartSeries(
                    new A.Charts.Index() { Val = (uint)1 },
                    new Order() { Val = (uint)1 },
                    new SeriesText(new NumericValue() { Text = "Điều kiện làm việc" }),
                    new XValues(xNumberLiteral2),
                    new YValues(yNumberLiteral2),
                    new Marker(
                        new Symbol() { Val = MarkerStyleValues.Diamond },
                        new A.Charts.Size() { Val = 7 },
                        new ChartShapeProperties(
        new A.SolidFill(new A.RgbColorModelHex() { Val = "000000" }),
        new A.Outline(new A.SolidFill(new A.RgbColorModelHex() { Val = "000000" }))
    )
                    ),
                    new ChartShapeProperties(
                        new A.Outline(new A.SolidFill(new A.RgbColorModelHex() { Val = "000000" }))
                    )
                ),
                new AxisId() { Val = 48650112u },
                new AxisId() { Val = 48672768u }
            );

            // Tiêu đề trục X
            var xAxisTitle = new Title(
                new ChartText(
                    new RichText(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(
                           new A.Run(
                    new A.RunProperties
                    {
                        Language = "vi-VN",
                        FontSize = 1100,
                        Bold = true
                    },
                    new A.Text() { Text = Common.ToSuperscriptUnit(xTitle) }
                )
                        )
                    )
                ),
                new Overlay() { Val = false }
            );

            // Tiêu đề trục Y
            var yAxisTitle = new Title(
                new ChartText(
                    new RichText(
                        new A.BodyProperties() { Rotation = -5400000 },

                        new A.ListStyle(),
                        new A.Paragraph(
                           new A.Run(
                    new A.RunProperties
                    {
                        Language = "vi-VN",
                        FontSize = 1100,
                        Bold = true
                    },
                    new A.Text() { Text = Common.ToSuperscriptUnit(yTitle) }
                )
                        )
                    )
                ),
                new Overlay() { Val = false }
            );

            var catAx = new ValueAxis(
                new AxisId() { Val = 48650112u },
                new Scaling(
                    new A.Charts.Orientation() { Val = OrientationValues.MinMax },
                    new MinAxisValue() { Val = xMin },
                    new MaxAxisValue() { Val = xMax }
                ),
                new Delete() { Val = false },
                new AxisPosition() { Val = AxisPositionValues.Bottom },
                new MajorGridlines(),
                new A.Charts.NumberingFormat() { FormatCode = "General", SourceLinked = true },
                new MajorUnit() { Val = xMajor },
                new TickLabelPosition() { Val = TickLabelPositionValues.NextTo },
                new CrossingAxis() { Val = 48672768u },
                new Crosses() { Val = CrossesValues.AutoZero },
                new CrossBetween() { Val = CrossBetweenValues.Between },
                xAxisTitle
            );

            var valAx = new ValueAxis(
                new AxisId() { Val = 48672768u },
                new Scaling(
                    new A.Charts.Orientation() { Val = OrientationValues.MinMax },
                    new MinAxisValue() { Val = yMin },
                    new MaxAxisValue() { Val = yMax }
                ),
                new Delete() { Val = false },
                new AxisPosition() { Val = AxisPositionValues.Left },
                new MajorGridlines(),
                new A.Charts.NumberingFormat() { FormatCode = "General", SourceLinked = true },
                new MajorUnit() { Val = yMajor },
                new TickLabelPosition() { Val = TickLabelPositionValues.NextTo },
                new CrossingAxis() { Val = 48650112u },
                new Crosses() { Val = CrossesValues.AutoZero },
                new CrossBetween() { Val = CrossBetweenValues.Between },
                yAxisTitle
            );



            var chart = new Chart(
                new AutoTitleDeleted() { Val = true },
                new PlotArea(scatterChart, catAx, valAx),
                new Legend(
                    new LegendPosition() { Val = LegendPositionValues.Bottom },
                    new Layout(),
                    new Overlay() { Val = false }
                ),
                new PlotVisibleOnly() { Val = true }
            );

            chartPart.ChartSpace = new ChartSpace(chart);

            var drawing = new Drawing(
                new Inline(
                    new Extent() { Cx = 5486400, Cy = 3200400 },
                    new EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                    new DocProperties() { Id = (UInt32Value)1U, Name = "Chart name" },
                    new NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks() { NoChangeAspect = true }),
                    new A.Graphic(
                        new A.GraphicData(
                            new ChartReference() { Id = chartPartId }
                        )
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }
                    )
                )
            );

            var centeredChartParagraph = new Paragraph(
                new ParagraphProperties(
                    new Justification() { Val = JustificationValues.Center }
                ),
                drawing
            );
            var body = mainPart.Document.Body;
            body.Append(centeredChartParagraph);
        }
        #endregion

        #endregion

        //#region xuất báo cáo phiên bản cũ

        //string[] _Info = new string[15];
        //string[,] _Para2 = new string[15, 10];
        //string[] _Para3 = new string[15];
        //string[] _Para4 = new string[5];

        //string[] __Info = new string[15];
        //string[,] __Para2 = new string[10, 10];
        //string[] __Para3 = new string[15];
        //string[] __Para4 = new string[5];

        //string[] Info = new string[20];// thêm info thì phải mở rộng mảng tăng giá trị lên lúc đấy không phải là 20 nữa có thể là một số lớn hơn 
        //string[,] Para1 = new string[15, 15];
        //string[,] Para2 = new string[15, 15];
        //string[,] Para3 = new string[15, 15];
        //string[] Para4 = new string[15];
        //string[] Para5 = new string[5];
        //private void TransferData_0()
        //{
        //    //Thông tin đặt hàng
        //    //Model - Serial No - Nhiệt độ - Công suất lắp đặt - Tốc độ - Tên đông cơ
        //    _Info[0] = DataProcess.Model;                            // Model
        //    _Info[1] = DataProcess.SerialNo;                         // SerialNo
        //    _Info[2] = DataProcess.Ta.ToString();                    // Nhiệt độ
        //    _Info[3] = DataProcess.MotorPower.ToString();                       // Công suất lắp đặt
        //    _Info[4] = DataProcess.MotorSpeed.ToString();            // Tốc độ động cơ
        //    _Info[5] = DataProcess.MotorName;                        // Tên động cơ
        //    _Info[6] = DataProcess.n1.ToString();                    // Tốc độ thiết kế
        //    _Info[12] = DataProcess.Tw.ToString();
        //    _Info[13] = DataProcess.Ta.ToString();
        //    _Info[9] = DataProcess.Idm.ToString();
        //    // _Info[14] = DataProcess.Vdm.ToString();                        
        //    // _Info[15] = DataProcess.Pdm.ToString();
        //    _Info[10] = DataProcess.e_motor.ToString();
        //    _Info[11] = DataProcess.CosPhi.ToString();


        //    for (int j = 0; j < 10; j++)
        //    {
        //        //Thông số đo kiểm
        //        //Nhiệt độ môi trường - Tỷ trọng đo kiểm - Tốc độ đo kiểm - Lưu lượng - Áp suất tĩnh - Áp suất tổng - Công suất tiêu thụ - Hiệu suất tĩnh - Hiệu suất tổng 
        //        _Para2[0, j] = DataProcess.TaPoint[j].ToString();    // Nhiệt độ môi trường
        //        _Para2[1, j] = DataProcess.rhoaPoint[j].ToString();  // Tỷ trọng đo kiểm
        //        _Para2[2, j] = DataProcess.n2Point[j].ToString();    // Tốc độ đo kiểm
        //        _Para2[3, j] = DataProcess.FlowPoint[j].ToString();  // Lưu lượng
        //        _Para2[4, j] = DataProcess.PsPoint[j].ToString();    // Áp suất tĩnh
        //        _Para2[5, j] = DataProcess.PtPoint[j].ToString();    // Áp suất tổng
        //        _Para2[6, j] = DataProcess.PwPoint[j].ToString();    // Công suất tiêu thụ
        //        _Para2[7, j] = DataProcess.EsPoint[j].ToString();    // Hiệu suất tĩnh
        //        _Para2[8, j] = DataProcess.EtPoint[j].ToString();    // Hiệu suất tổng
        //        _Para2[9, j] = DataProcess.rho3[j].ToString();       // Tỷ trọng kk điểm đo áp suất
        //        _Para2[10, j] = DataProcess.Pr[j].ToString();        // Công suất trên trục

        //    }

        //    //Thông số chạy thử cơ khí
        //    //Độ ồn - Độ rung ngang 1 - Độ rung ngang 2 - Độ rung đứng 1 - Độ rung đứng 2 - Độ rung dọc 1 - Độ rung dọc 2 - Nhiệt độ ngang 1 - Nhiệt độ ngang 2 - Nhiệt độ đứng 1 - Nhiệt độ đứng 2 - Nhiệt độ dọc 1 - Nhiệt độ dọc 2
        //    _Para3[0] = DataProcess.Noise.ToString();                  // Độ ồn
        //    _Para3[1] = DataProcess.BearingVia_H.ToString();           // Độ rung - ngang   gối 1
        //    _Para3[2] = DataProcess.BearingVia_H.ToString();           // Độ rung - ngang   gối 2
        //    _Para3[3] = DataProcess.BearingVia.ToString();             // Độ rung - đứng    gối 1
        //    _Para3[4] = DataProcess.BearingVia.ToString();             // Độ rung - đứng    gối 2
        //    _Para3[5] = DataProcess.BearingVia_V.ToString();           // Độ rung - dọc     gối 1
        //    _Para3[6] = DataProcess.BearingVia_V.ToString();           // Độ rung - dọc     gối 2
        //    _Para3[7] = DataProcess.BearingTemp_H.ToString();          // Nhiệt độ - ngang  gối 1
        //    _Para3[8] = DataProcess.BearingTemp_H.ToString();          // Nhiệt độ - ngang  gối 2
        //    _Para3[9] = DataProcess.BearingTemp.ToString();            // Nhiệt độ - đứng   gối 1
        //    _Para3[10] = DataProcess.BearingTemp.ToString();           // Nhiệt độ - đứng   gối 2
        //    _Para3[11] = DataProcess.BearingTemp_V.ToString();         // Nhiệt độ - dọc    gối 1
        //    _Para3[12] = DataProcess.BearingTemp_V.ToString();         // Nhiệt độ - dọc    gối 2

        //    //Kết luận
        //    //Người thực hiện - Người chứng kiến - Người chấp thuận
        //    _Para4[0] = DataProcess.TestPerson;                        // Người thực hiện
        //    _Para4[1] = DataProcess.TestWitness;                       // Người chứng kiến
        //    _Para4[2] = DataProcess.Approved;                          // Người chấp thuận
        //}

        //public void ExportDatabase(string savePath)
        //{
        //    try
        //    {
        //        TransferData_0();
        //        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        //        using (var package = new ExcelPackage())
        //        {
        //            var ws = package.Workbook.Worksheets.Add("Sheet1");

        //            // Thông tin đặt hàng
        //            ws.Cells[1, 1].Value = "Thông tin quạt";
        //            ws.Cells[1, 2].Value = "Model";
        //            ws.Cells[1, 3].Value = "Serial No";
        //            ws.Cells[1, 4].Value = "Nhiệt độ môi trường";
        //            ws.Cells[1, 5].Value = "Công suất lắp đặt";
        //            ws.Cells[1, 6].Value = "Tốc độ";
        //            ws.Cells[1, 7].Value = "Động cơ";
        //            ws.Cells[1, 8].Value = "Tốc độ thiết kế";
        //            ws.Cells[1, 9].Value = "Dòng định mức";
        //            ws.Cells[1, 10].Value = "Hiệu suất động cơ";
        //            ws.Cells[1, 11].Value = "Hệ số cos phi";
        //            for (int i = 2; i < 9; i++)
        //            {
        //                ws.Cells[2, i].Value = _Info[i - 2];
        //            }

        //            // Thông tin đo kiểm
        //            string[] thongSo = {
        //        "Nhiệt độ môi trường", "Tỷ trọng đo kiểm", "Tốc độ đo kiểm", "Lưu lượng",
        //        "Áp suất tĩnh", "Áp suất tổng", "Công suất tiêu thụ", "Hiệu suất tĩnh",
        //        "Hiệu suất tổng", "Tỷ trọng kk điểm đo áp suất", "Công suất trên trục"
        //    };
        //            for (int i = 0; i < thongSo.Length; i++)
        //            {
        //                ws.Cells[5 + i, 1].Value = thongSo[i];
        //            }
        //            for (int i = 5; i < 16; i++)
        //            {
        //                for (int j = 2; j < 12; j++)
        //                {
        //                    ws.Cells[i, j].Value = _Para2[i - 5, j - 2];
        //                }
        //            }

        //            // Thông số chạy thử cơ khí
        //            ws.Cells[20, 1].Value = "Độ ồn";
        //            ws.Cells[22, 2].Value = "Phương ngang 1";
        //            ws.Cells[22, 3].Value = "Phương ngang 2";
        //            ws.Cells[22, 4].Value = "Phương đứng 1";
        //            ws.Cells[22, 5].Value = "Phương đứng 2";
        //            ws.Cells[22, 6].Value = "Phương dọc 1";
        //            ws.Cells[22, 7].Value = "Phương dọc 2";
        //            ws.Cells[23, 1].Value = "Độ rung gối trục";
        //            ws.Cells[24, 1].Value = "Nhiệt độ gối trục";
        //            ws.Cells[20, 2].Value = _Para3[0];
        //            for (int j = 2; j < 8; j++)
        //            {
        //                ws.Cells[23, j].Value = _Para3[j - 1];
        //            }
        //            for (int j = 2; j < 8; j++)
        //            {
        //                ws.Cells[24, j].Value = _Para3[j + 5];
        //            }

        //            // Kết luận
        //            ws.Cells[26, 1].Value = "Thông tin quản trị";
        //            ws.Cells[26, 2].Value = "Người thực hiện";
        //            ws.Cells[26, 3].Value = "Người chứng kiến";
        //            ws.Cells[26, 4].Value = "Người Phê duyệt";
        //            for (int j = 2; j < 5; j++)
        //            {
        //                ws.Cells[27, j].Value = _Para4[j - 2];
        //            }

        //            ws.Cells.AutoFitColumns();

        //            // Tạo tên file
        //            var name = _Info[0];
        //            for (int i = 0; i < name.Length; i++)
        //            {
        //                if (name.Substring(i, 1) == ".")
        //                {
        //                    name = name.Remove(i, 1).Insert(i, "_");
        //                }
        //            }
        //            string filename = name + "_" +
        //                              DateTime.Now.ToString("yyyy") + "_" +
        //                              DateTime.Now.ToString("MM") + "_" +
        //                              DateTime.Now.ToString("dd") + "_" +
        //                              DateTime.Now.ToString("HH") + "_" +
        //                              DateTime.Now.ToString("mm") + "_" +
        //                              DateTime.Now.ToString("ss");
        //            string FilePath = savePath + @"\" + filename + ".xlsx";
        //            package.SaveAs(new FileInfo(FilePath));
        //            MessageBox.Show("Export Database Successfully to \n" + FilePath, "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //            DataProcess.Done = 1;
        //        }
        //    }
        //    catch (Exception exml)
        //    {
        //        MessageBox.Show("Export Database fail \n" + exml.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //    }
        //}


        ////==================== Import Database ==================
        //private static DataSet ds;
        //private static DataTable dt;
        //static int _flag = 0;
        //public static DataTable ImportDatabase()
        //{
        //    DataTable dt = new DataTable();
        //    try
        //    {
        //        using (OpenFileDialog ofd = new OpenFileDialog()
        //        {
        //            Filter = "Excel Workbook|*.xlsx",
        //            ValidateNames = true
        //        })
        //        {
        //            if (ofd.ShowDialog() == DialogResult.OK)
        //            {
        //                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        //                using (var package = new ExcelPackage(new FileInfo(ofd.FileName)))
        //                {
        //                    // Đọc sheet đầu tiên
        //                    var ws = package.Workbook.Worksheets[0];
        //                    if (ws == null)
        //                        throw new Exception("Không tìm thấy sheet trong file Excel.");

        //                    int startRow = ws.Dimension.Start.Row;
        //                    int endRow = ws.Dimension.End.Row;
        //                    int startCol = ws.Dimension.Start.Column;
        //                    int endCol = ws.Dimension.End.Column;

        //                    // Tạo cột cho DataTable
        //                    for (int col = startCol; col <= endCol; col++)
        //                    {
        //                        dt.Columns.Add("Col" + col, typeof(string));
        //                    }

        //                    // Đọc từng dòng và thêm vào DataTable
        //                    for (int row = startRow; row <= endRow; row++)
        //                    {
        //                        DataRow dr = dt.NewRow();
        //                        for (int col = startCol; col <= endCol; col++)
        //                        {
        //                            dr[col - startCol] = ws.Cells[row, col].Text;
        //                        }
        //                        dt.Rows.Add(dr);
        //                    }
        //                }
        //            }
        //            else
        //            {
        //                _flag = 1;
        //            }
        //        }
        //    }
        //    catch (Exception exml)
        //    {
        //        MessageBox.Show("Import Database fail \n" + exml.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //    return dt;
        //}


        //public static void TransferData_0_0(DataGridView dgv)
        //{
        //    if (_flag != 1)
        //    {
        //        try
        //        {
        //            //Thông tin đặt hàng
        //            //Model - Serial No - Nhiệt độ - Công suất lắp đặt - Tốc độ - Tên đông cơ
        //            DataProcess.Model = dgv.Rows[1].Cells[1].Value.ToString();                       // Model
        //            DataProcess.SerialNo = dgv.Rows[1].Cells[2].Value.ToString();                    // SerialNo
        //            DataProcess.Ta = float.Parse(dgv.Rows[1].Cells[3].Value.ToString());             // Nhiệt độ
        //            DataProcess.MotorPower = dgv.Rows[1].Cells[4].Value.ToString();                  // Công suất lắp đặt
        //            DataProcess.MotorSpeed = short.Parse(dgv.Rows[1].Cells[5].Value.ToString());     // Tốc độ
        //            DataProcess.MotorName = dgv.Rows[1].Cells[6].Value.ToString();                   // Tên động cơ
        //            DataProcess.n1 = short.Parse(dgv.Rows[1].Cells[7].Value.ToString());             // Tốc độ thiết kế
        //            DataProcess.Idm = float.Parse(dgv.Rows[1].Cells[8].Value.ToString());             // Dòng điện định mức 
        //            DataProcess.e_motor = float.Parse(dgv.Rows[1].Cells[9].Value.ToString());         // Hiệu suất động cơ 
        //            DataProcess.CosPhi = float.Parse(dgv.Rows[1].Cells[10].Value.ToString());         // Cos phi


        //            for (int j = 1; j < 11; j++)
        //            {
        //                //Thông số đo kiểm
        //                //Nhiệt độ môi trường - Tỷ trọng đo kiểm - Tốc độ đo kiểm - Lưu lượng - Áp suất tĩnh - Áp suất tổng - Công suất tiêu thụ - Hiệu suất tĩnh - Hiệu suất tổng  
        //                string temp1 = dgv.Rows[6].Cells[j].Value.ToString();
        //                DataProcess.TaPoint[j - 1] = float.Parse(temp1);       // Nhiệt độ môi trường
        //                string temp2 = dgv.Rows[7].Cells[j].Value.ToString();
        //                DataProcess.IdmPoint[j - 1] = float.Parse(temp2);     // Dòng điện 
        //                string temp3 = dgv.Rows[8].Cells[j].Value.ToString();
        //                DataProcess.UdmPoint[j - 1] = float.Parse(temp3);        // Điện áp 
        //                string temp4 = dgv.Rows[9].Cells[j].Value.ToString();
        //                DataProcess.n2Point[j - 1] = float.Parse(temp4);     // Tốc độ thực 
        //                string temp5 = dgv.Rows[10].Cells[j].Value.ToString();
        //                DataProcess.PsPoint[j - 1] = float.Parse(temp5);       // Áp suất tĩnh
        //                string temp6 = dgv.Rows[11].Cells[j].Value.ToString();
        //                DataProcess.PtPoint[j - 1] = float.Parse(temp6);       // Áp suất tổng
        //                string temp7 = dgv.Rows[12].Cells[j].Value.ToString();
        //                DataProcess.PwPoint[j - 1] = float.Parse(temp7);       // Công suất tiêu thụ
        //                string temp8 = dgv.Rows[13].Cells[j].Value.ToString();
        //                DataProcess.EsPoint[j - 1] = float.Parse(temp8);       // Hiệu suất tĩnh
        //                string temp9 = dgv.Rows[14].Cells[j].Value.ToString();
        //                DataProcess.EtPoint[j - 1] = float.Parse(temp9);       // Hiệu suất tổng
        //                string temp10 = dgv.Rows[15].Cells[j].Value.ToString();
        //                DataProcess.rho3[j - 1] = float.Parse(temp9);       // Hiệu suất tổng
        //                string temp11 = dgv.Rows[16].Cells[j].Value.ToString();
        //                DataProcess.Pr[j - 1] = float.Parse(temp9);       // Hiệu suất tổng

        //            }

        //            //Thông số chạy thử cơ khí
        //            //Độ ồn - Độ rung ngang 1 - Độ rung ngang 2 - Độ rung đứng 1 - Độ rung đứng 2 - Độ rung dọc 1 - Độ rung dọc 2 - Nhiệt độ ngang 1 - Nhiệt độ ngang 2 - Nhiệt độ đứng 1 - Nhiệt độ đứng 2 - Nhiệt độ dọc 1 - Nhiệt độ dọc 2
        //            DataProcess.Noise = float.Parse(dgv.Rows[19].Cells[1].Value.ToString());                  // Độ ồn
        //            DataProcess.BearingVia_H = float.Parse(dgv.Rows[22].Cells[1].Value.ToString());           // Độ rung - ngang   gối 1
        //            DataProcess.BearingVia_H = float.Parse(dgv.Rows[22].Cells[2].Value.ToString());           // Độ rung - ngang   gối 2
        //            DataProcess.BearingVia = float.Parse(dgv.Rows[22].Cells[3].Value.ToString());             // Độ rung - đứng    gối 1
        //            DataProcess.BearingVia = float.Parse(dgv.Rows[22].Cells[4].Value.ToString());             // Độ rung - đứng    gối 2
        //            DataProcess.BearingVia_V = float.Parse(dgv.Rows[22].Cells[5].Value.ToString());           // Độ rung - dọc     gối 1
        //            DataProcess.BearingVia_V = float.Parse(dgv.Rows[22].Cells[6].Value.ToString());           // Độ rung - dọc     gối 2
        //            DataProcess.BearingTemp_H = float.Parse(dgv.Rows[23].Cells[1].Value.ToString());          // Nhiệt độ - ngang  gối 1
        //            DataProcess.BearingTemp_H = float.Parse(dgv.Rows[23].Cells[2].Value.ToString());          // Nhiệt độ - ngang  gối 2
        //            DataProcess.BearingTemp = float.Parse(dgv.Rows[23].Cells[3].Value.ToString());            // Nhiệt độ - đứng   gối 1
        //            DataProcess.BearingTemp = float.Parse(dgv.Rows[23].Cells[4].Value.ToString());            // Nhiệt độ - đứng   gối 2
        //            DataProcess.BearingTemp_V = float.Parse(dgv.Rows[23].Cells[5].Value.ToString());          // Nhiệt độ - dọc    gối 1
        //            DataProcess.BearingTemp_V = float.Parse(dgv.Rows[23].Cells[6].Value.ToString());          // Nhiệt độ - dọc    gối 2

        //            //Kết luận
        //            //Người thực hiện - Người chứng kiến - Người chấp thuận
        //            DataProcess.TestPerson = dgv.Rows[26].Cells[1].Value.ToString();       //Người thực hiện
        //            DataProcess.TestWitness = dgv.Rows[26].Cells[2].Value.ToString();      //Người chứng kiến
        //            DataProcess.Approved = dgv.Rows[26].Cells[3].Value.ToString();         //Người phê duyệt


        //            MessageBox.Show("Import Successfully!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //            // Check
        //            DataProcess.ImportCheck = 1;
        //        }
        //        catch (Exception exml)
        //        {
        //            MessageBox.Show("Import Database fail \n" + exml.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        }
        //    }
        //    else _flag = 0;

        //}


        ////========== Truyền dữ liệu vào mảng để xuất báo cáo ===========
        ////Report_1
        //private void TransferData_1()
        //{
        //    //Thông tin đặt hàng
        //    Info[0] = DataProcess.Customer;                      // Khách hàng
        //    Info[1] = DataProcess.Model;                         // Model
        //    Info[2] = DataProcess.SerialNo;                      // SerialNo
        //    Info[3] = DataProcess.Ta.ToString();                 // Nhiệt độ
        //    Info[4] = DataProcess.Project;                       // Dự án
        //    Info[5] = DataProcess.MotorPower.ToString();                    // Công suất lắp đặt
        //    Info[6] = DataProcess.MotorSpeed.ToString();         // Tốc độ
        //    Info[7] = DataProcess.MotorName;                     // Tên động cơ
        //    Info[8] = DataProcess.ReportNo;                      // STT
        //    Info[12] = DataProcess.Tw.ToString();
        //    Info[13] = DataProcess.Ta.ToString();
        //    Info[9] = DataProcess.Idm.ToString();
        //    Info[14] = DataProcess.Vdm.ToString();
        //    Info[15] = DataProcess.Pdm.ToString();
        //    Info[10] = DataProcess.e_motor.ToString();
        //    Info[11] = DataProcess.CosPhi.ToString();

        //    for (int j = 0; j < 10; j++)
        //    {
        //        //Thông số đo kiểm
        //        Para1[0, j] = DataProcess.TaPoint[j].ToString();    // Nhiệt độ môi trường
        //        Para1[1, j] = DataProcess.IdmPoint[j].ToString();  // Tỷ trọng đo kiểm
        //        Para1[2, j] = DataProcess.UdmPoint[j].ToString();    // Tốc độ đo kiểm
        //        Para1[3, j] = DataProcess.n2Point[j].ToString();  // Tốc độ thực 
        //        Para1[4, j] = DataProcess.PsPoint[j].ToString();    // Áp suất tĩnh
        //        Para1[5, j] = DataProcess.PtPoint[j].ToString();    // Áp suất tổng
        //        Para1[6, j] = DataProcess.T_Point[j].ToString();    // Momen xoắn trên trục
        //        Para1[7, j] = DataProcess.Pr[j].ToString();         // Công suất tiêu thụ
        //        //Para1[7, j] = DataProcess.EsPoint[j].ToString();    // Hiệu suất tĩnh
        //        //Para1[8, j] = DataProcess.EtPoint[j].ToString();    // Hiệu suất tổng

        //        //Thông số đo quy đổi về điều kiện tiêu chuẩn
        //        Para2[0, j] = DataProcess.Std_FlowPoint[j].ToString();  // Lưu lượng
        //        Para2[1, j] = DataProcess.Std_PsPoint[j].ToString();    // Áp suất tĩnh
        //        Para2[2, j] = DataProcess.Std_PtPoint[j].ToString();    // Áp suất tổng
        //        Para2[3, j] = DataProcess.Std_PrPoint[j].ToString();    // Công suất tiêu thụ
        //        Para2[4, j] = DataProcess.PrtPoint[j].ToString();       // Công suất tiêu thụ tính theo T
        //        Para2[5, j] = DataProcess.Std_EsPoint[j].ToString();    // Hiệu suất tĩnh
        //        Para2[6, j] = DataProcess.Std_EstPoint[j].ToString();   // Hiệu suất tĩnh tính theo T
        //        Para2[7, j] = DataProcess.Std_EtPoint[j].ToString();    // Hiệu suất tổng
        //        Para2[8, j] = DataProcess.Std_EttPoint[j].ToString();   // Hiệu suất tổng tính theo T

        //        //Thông số đo quy đổi về điều kiện làm việc
        //        Para3[0, j] = DataProcess.Tw.ToString();                // Nhiệt độ khí
        //        Para3[1, j] = DataProcess.rhow.ToString();              // Tỷ trọng khí
        //        Para3[2, j] = DataProcess.n1.ToString();                // Tốc độ guồng cánh
        //        Para3[3, j] = DataProcess.Ope_FlowPoint[j].ToString();  // Lưu lượng
        //        Para3[4, j] = DataProcess.Ope_PsPoint[j].ToString();    // Áp suất tĩnh
        //        Para3[5, j] = DataProcess.Ope_PtPoint[j].ToString();    // Áp suất tổng
        //        Para3[6, j] = DataProcess.Ope_PrPoint[j].ToString();    // Công suất tiêu thụ
        //        Para3[7, j] = DataProcess.PrtPoint[j].ToString();       // Công suất tiêu thụ tính theo T
        //        Para3[8, j] = DataProcess.Ope_EsPoint[j].ToString();    // Hiệu suất tĩnh
        //        Para3[9, j] = DataProcess.Ope_EstPoint[j].ToString();   // Hiệu suất tĩnh tính theo T
        //        Para3[10, j] = DataProcess.Ope_EtPoint[j].ToString();   // Hiệu suất tổng
        //        Para3[11, j] = DataProcess.Ope_EttPoint[j].ToString();  // Hiệu suất tổng tính theo T
        //    }

        //    //Thông số chạy thử cơ khí
        //    Para4[0] = DataProcess.Noise.ToString();           // Độ ồn
        //    Para4[1] = DataProcess.BearingVia_H.ToString();    // Độ rung - ngang   gối 1
        //    Para4[2] = DataProcess.BearingVia_H.ToString();    // Độ rung - ngang   gối 2
        //    Para4[3] = DataProcess.BearingVia.ToString();      // Độ rung - đứng    gối 1
        //    Para4[4] = DataProcess.BearingVia.ToString();      // Độ rung - đứng    gối 2
        //    Para4[5] = DataProcess.BearingVia_V.ToString();    // Độ rung - dọc     gối 1
        //    Para4[6] = DataProcess.BearingVia_V.ToString();    // Độ rung - dọc     gối 2
        //    Para4[7] = DataProcess.BearingTemp_H.ToString();   // Nhiệt độ - ngang  gối 1
        //    Para4[8] = DataProcess.BearingTemp_H.ToString();   // Nhiệt độ - ngang  gối 2
        //    Para4[9] = DataProcess.BearingTemp.ToString();     // Nhiệt độ - đứng   gối 1
        //    Para4[10] = DataProcess.BearingTemp.ToString();    // Nhiệt độ - đứng   gối 2
        //    Para4[11] = DataProcess.BearingTemp_V.ToString();  // Nhiệt độ - dọc    gối 1
        //    Para4[12] = DataProcess.BearingTemp_V.ToString();  // Nhiệt độ - dọc    gối 2

        //    //Kết luận
        //    Para5[0] = DataProcess.TestPerson;                 // Người thực hiện
        //    Para5[1] = DataProcess.TestWitness;                // Người chứng kiến
        //    Para5[2] = DataProcess.Approved;                   // Người chấp thuận
        //}

        ////Report_2
        //private void TransferData_2()
        //{
        //    //Thông tin đặt hàng
        //    Info[0] = DataProcess.Customer;                      // Khách hàng
        //    Info[1] = DataProcess.Model;                         // Model
        //    Info[2] = DataProcess.SerialNo;                      // SerialNo
        //    Info[3] = DataProcess.Ta.ToString();                 // Nhiệt độ
        //    Info[4] = DataProcess.Project;                       // Dự án
        //    Info[5] = DataProcess.MotorPower;                    // Công suất lắp đặt
        //    Info[6] = DataProcess.MotorSpeed.ToString();         // Tốc độ
        //    Info[7] = DataProcess.MotorName;                     // Tên động cơ
        //    Info[8] = DataProcess.ReportNo;                      // STT

        //    for (int j = 0; j < 10; j++)
        //    {
        //        //Thông số đo kiểm
        //        Para1[0, j] = DataProcess.TaPoint[j].ToString();    // Nhiệt độ môi trường
        //        Para1[1, j] = DataProcess.rhoaPoint[j].ToString();  // Tỷ trọng đo kiểm
        //        Para1[2, j] = DataProcess.n2Point[j].ToString();    // Tốc độ đo kiểm
        //        Para1[3, j] = DataProcess.FlowPoint[j].ToString();  // Lưu lượng
        //        Para1[4, j] = DataProcess.PsPoint[j].ToString();    // Áp suất tĩnh
        //        Para1[5, j] = DataProcess.T_Point[j].ToString();    // Momen xoắn trên trục
        //        Para1[6, j] = DataProcess.Pr[j].ToString();         // Công suất tiêu thụ

        //        //Thông số đo quy đổi về điều kiện tiêu chuẩn
        //        Para2[0, j] = DataProcess.Std_FlowPoint[j].ToString();  // Lưu lượng
        //        Para2[1, j] = DataProcess.Std_PsPoint[j].ToString();    // Áp suất tĩnh
        //        Para2[2, j] = DataProcess.Std_PrPoint[j].ToString();    // Công suất tiêu thụ
        //        Para2[3, j] = DataProcess.PrtPoint[j].ToString();       // Công suất tiêu thụ tính theo T

        //        //Thông số đo quy đổi về điều kiện làm việc
        //        Para3[0, j] = DataProcess.Tw.ToString();                // Nhiệt độ khí
        //        Para3[1, j] = DataProcess.rhow.ToString();              // Tỷ trọng khí
        //        Para3[2, j] = DataProcess.n1.ToString();                // Tốc độ guồng cánh
        //        Para3[3, j] = DataProcess.Ope_FlowPoint[j].ToString();  // Lưu lượng
        //        Para3[4, j] = DataProcess.Ope_PsPoint[j].ToString();    // Áp suất tĩnh
        //        Para3[5, j] = DataProcess.Ope_PrPoint[j].ToString();    // Công suất tiêu thụ
        //        Para3[6, j] = DataProcess.PrtPoint[j].ToString();       // Công suất tiêu thụ tính theo T
        //    }

        //    //Thông số chạy thử cơ khí
        //    Para4[0] = DataProcess.Noise.ToString();           // Độ ồn
        //    Para4[1] = DataProcess.BearingVia_H.ToString();    // Độ rung - ngang   gối 1
        //    Para4[2] = DataProcess.BearingVia_H.ToString();    // Độ rung - ngang   gối 2
        //    Para4[3] = DataProcess.BearingVia.ToString();      // Độ rung - đứng    gối 1
        //    Para4[4] = DataProcess.BearingVia.ToString();      // Độ rung - đứng    gối 2
        //    Para4[5] = DataProcess.BearingVia_V.ToString();    // Độ rung - dọc     gối 1
        //    Para4[6] = DataProcess.BearingVia_V.ToString();    // Độ rung - dọc     gối 2
        //    Para4[7] = DataProcess.BearingTemp_H.ToString();   // Nhiệt độ - ngang  gối 1
        //    Para4[8] = DataProcess.BearingTemp_H.ToString();   // Nhiệt độ - ngang  gối 2
        //    Para4[9] = DataProcess.BearingTemp.ToString();     // Nhiệt độ - đứng   gối 1
        //    Para4[10] = DataProcess.BearingTemp.ToString();    // Nhiệt độ - đứng   gối 2
        //    Para4[11] = DataProcess.BearingTemp_V.ToString();  // Nhiệt độ - dọc    gối 1
        //    Para4[12] = DataProcess.BearingTemp_V.ToString();  // Nhiệt độ - dọc    gối 2

        //    //Kết luận
        //    Para5[0] = DataProcess.TestPerson;                 // Người thực hiện
        //    Para5[1] = DataProcess.TestWitness;                // Người chứng kiến
        //    Para5[2] = DataProcess.Approved;                   // Người chấp thuận
        //}

        ////Report_3
        //private void TransferData_3()
        //{
        //    //Thông tin đặt hàng
        //    Info[0] = DataProcess.Customer;                      // Khách hàng
        //    Info[1] = DataProcess.Model;                         // Model
        //    Info[2] = DataProcess.SerialNo;                      // SerialNo
        //    Info[3] = DataProcess.Ta.ToString();                 // Nhiệt độ
        //    Info[4] = DataProcess.Project;                       // Dự án
        //    Info[5] = DataProcess.MotorPower;                    // Công suất lắp đặt
        //    Info[6] = DataProcess.MotorSpeed.ToString();         // Tốc độ
        //    Info[7] = DataProcess.MotorName;                     // Tên động cơ
        //    Info[8] = DataProcess.ReportNo;                      // STT
        //    for (int j = 0; j < 10; j++)
        //    {
        //        //Thông số đo kiểm
        //        Para1[0, j] = DataProcess.TaPoint[j].ToString();    // Nhiệt độ môi trường
        //        Para1[1, j] = DataProcess.rhoaPoint[j].ToString();  // Tỷ trọng đo kiểm
        //        Para1[2, j] = DataProcess.n2Point[j].ToString();    // Tốc độ đo kiểm
        //        Para1[3, j] = DataProcess.FlowPoint[j].ToString();  // Lưu lượng
        //        Para1[4, j] = DataProcess.PtPoint[j].ToString();    // Áp suất tổng
        //        Para1[5, j] = DataProcess.T_Point[j].ToString();    // Momen xoắn trên trục
        //        Para1[6, j] = DataProcess.Pr[j].ToString();         // Công suất tiêu thụ

        //        //Thông số đo quy đổi về điều kiện tiêu chuẩn
        //        Para2[0, j] = DataProcess.Std_FlowPoint[j].ToString();  // Lưu lượng
        //        Para2[1, j] = DataProcess.Std_PtPoint[j].ToString();    // Áp suất tổng
        //        Para2[2, j] = DataProcess.Std_PrPoint[j].ToString();    // Công suất tiêu thụ
        //        Para2[3, j] = DataProcess.PrtPoint[j].ToString();       // Công suất tiêu thụ tính theo T

        //        //Thông số đo quy đổi về điều kiện làm việc
        //        Para3[0, j] = DataProcess.Tw.ToString();                // Nhiệt độ khí
        //        Para3[1, j] = DataProcess.rhow.ToString();              // Tỷ trọng khí
        //        Para3[2, j] = DataProcess.n1.ToString();                // Tốc độ guồng cánh
        //        Para3[3, j] = DataProcess.Ope_FlowPoint[j].ToString();  // Lưu lượng
        //        Para3[4, j] = DataProcess.Ope_PtPoint[j].ToString();    // Áp suất tổng
        //        Para3[5, j] = DataProcess.Ope_PrPoint[j].ToString();    // Công suất tiêu thụ
        //        Para3[6, j] = DataProcess.PrtPoint[j].ToString();       // Công suất tiêu thụ tính theo T
        //    }

        //    //Thông số chạy thử cơ khí
        //    Para4[0] = DataProcess.Noise.ToString();           // Độ ồn
        //    Para4[1] = DataProcess.BearingVia_H.ToString();    // Độ rung - ngang   gối 1
        //    Para4[2] = DataProcess.BearingVia_H.ToString();    // Độ rung - ngang   gối 2
        //    Para4[3] = DataProcess.BearingVia.ToString();      // Độ rung - đứng    gối 1
        //    Para4[4] = DataProcess.BearingVia.ToString();      // Độ rung - đứng    gối 2
        //    Para4[5] = DataProcess.BearingVia_V.ToString();    // Độ rung - dọc     gối 1
        //    Para4[6] = DataProcess.BearingVia_V.ToString();    // Độ rung - dọc     gối 2
        //    Para4[7] = DataProcess.BearingTemp_H.ToString();   // Nhiệt độ - ngang  gối 1
        //    Para4[8] = DataProcess.BearingTemp_H.ToString();   // Nhiệt độ - ngang  gối 2
        //    Para4[9] = DataProcess.BearingTemp.ToString();     // Nhiệt độ - đứng   gối 1
        //    Para4[10] = DataProcess.BearingTemp.ToString();    // Nhiệt độ - đứng   gối 2
        //    Para4[11] = DataProcess.BearingTemp_V.ToString();  // Nhiệt độ - dọc    gối 1
        //    Para4[12] = DataProcess.BearingTemp_V.ToString();  // Nhiệt độ - dọc    gối 2

        //    //Kết luận
        //    Para5[0] = DataProcess.TestPerson;                 // Người thực hiện
        //    Para5[1] = DataProcess.TestWitness;                // Người chứng kiến
        //    Para5[2] = DataProcess.Approved;                   // Người chấp thuận
        //}

        ////Report_4
        //private void TransferData_4()
        //{
        //    //Thông tin đặt hàng
        //    Info[0] = DataProcess.Customer;                      // Khách hàng
        //    Info[1] = DataProcess.Model;                         // Model
        //    Info[2] = DataProcess.SerialNo;                      // SerialNo
        //    Info[3] = DataProcess.Ta.ToString();                 // Nhiệt độ
        //    Info[4] = DataProcess.Project;                       // Dự án
        //    Info[5] = DataProcess.MotorPower;                    // Công suất lắp đặt
        //    Info[6] = DataProcess.MotorSpeed.ToString();         // Tốc độ
        //    Info[7] = DataProcess.MotorName;                     // Tên động cơ
        //    Info[8] = DataProcess.ReportNo;                      // STT

        //    for (int j = 0; j < 10; j++)
        //    {
        //        //Thông số đo kiểm
        //        Para1[0, j] = DataProcess.TaPoint[j].ToString();    // Nhiệt độ môi trường
        //        Para1[1, j] = DataProcess.rhoaPoint[j].ToString();  // Tỷ trọng đo kiểm
        //        Para1[2, j] = DataProcess.n2Point[j].ToString();    // Tốc độ đo kiểm
        //        Para1[3, j] = DataProcess.FlowPoint[j].ToString();  // Lưu lượng
        //        Para1[4, j] = DataProcess.PsPoint[j].ToString();    // Áp suất tĩnh
        //        Para1[5, j] = DataProcess.T_Point[j].ToString();    // Momen xoắn trên trục
        //        Para1[6, j] = DataProcess.Pr[j].ToString();         // Công suất tiêu thụ

        //        //Thông số đo quy đổi về điều kiện tiêu chuẩn
        //        Para2[0, j] = DataProcess.Std_FlowPoint[j].ToString();  // Lưu lượng
        //        Para2[1, j] = DataProcess.Std_PsPoint[j].ToString();    // Áp suất tĩnh
        //        Para2[2, j] = DataProcess.Std_PrPoint[j].ToString();    // Công suất tiêu thụ
        //        Para2[3, j] = DataProcess.PrtPoint[j].ToString();       // Công suất tiêu thụ tính theo T
        //        Para2[4, j] = DataProcess.Std_EsPoint[j].ToString();    // Hiệu suất tĩnh
        //        Para2[5, j] = DataProcess.Std_EstPoint[j].ToString();   // Hiệu suất tĩnh tính theo T

        //        //Thông số đo quy đổi về điều kiện làm việc
        //        Para3[0, j] = DataProcess.Tw.ToString();                // Nhiệt độ khí
        //        Para3[1, j] = DataProcess.rhow.ToString();              // Tỷ trọng khí
        //        Para3[2, j] = DataProcess.n1.ToString();                // Tốc độ guồng cánh
        //        Para3[3, j] = DataProcess.Ope_FlowPoint[j].ToString();  // Lưu lượng
        //        Para3[4, j] = DataProcess.Ope_PsPoint[j].ToString();    // Áp suất tĩnh
        //        Para3[5, j] = DataProcess.Ope_PrPoint[j].ToString();    // Công suất tiêu thụ
        //        Para3[6, j] = DataProcess.PrtPoint[j].ToString();       // Công suất tiêu thụ tính theo T
        //        Para3[7, j] = DataProcess.Ope_EsPoint[j].ToString();    // Hiệu suất tĩnh
        //        Para3[8, j] = DataProcess.Ope_EstPoint[j].ToString();   // Hiệu suất tĩnh tính theo T
        //    }

        //    //Thông số chạy thử cơ khí
        //    Para4[0] = DataProcess.Noise.ToString();           // Độ ồn
        //    Para4[1] = DataProcess.BearingVia_H.ToString();    // Độ rung - ngang   gối 1
        //    Para4[2] = DataProcess.BearingVia_H.ToString();    // Độ rung - ngang   gối 2
        //    Para4[3] = DataProcess.BearingVia.ToString();      // Độ rung - đứng    gối 1
        //    Para4[4] = DataProcess.BearingVia.ToString();      // Độ rung - đứng    gối 2
        //    Para4[5] = DataProcess.BearingVia_V.ToString();    // Độ rung - dọc     gối 1
        //    Para4[6] = DataProcess.BearingVia_V.ToString();    // Độ rung - dọc     gối 2
        //    Para4[7] = DataProcess.BearingTemp_H.ToString();   // Nhiệt độ - ngang  gối 1
        //    Para4[8] = DataProcess.BearingTemp_H.ToString();   // Nhiệt độ - ngang  gối 2
        //    Para4[9] = DataProcess.BearingTemp.ToString();     // Nhiệt độ - đứng   gối 1
        //    Para4[10] = DataProcess.BearingTemp.ToString();    // Nhiệt độ - đứng   gối 2
        //    Para4[11] = DataProcess.BearingTemp_V.ToString();  // Nhiệt độ - dọc    gối 1
        //    Para4[12] = DataProcess.BearingTemp_V.ToString();  // Nhiệt độ - dọc    gối 2

        //    //Kết luận
        //    Para5[0] = DataProcess.TestPerson;                 // Người thực hiện
        //    Para5[1] = DataProcess.TestWitness;                // Người chứng kiến
        //    Para5[2] = DataProcess.Approved;                   // Người chấp thuận
        //}

        ////Report_5
        //private void TransferData_5()
        //{
        //    //Thông tin đặt hàng
        //    Info[0] = DataProcess.Customer;                      // Khách hàng
        //    Info[1] = DataProcess.Model;                         // Model
        //    Info[2] = DataProcess.SerialNo;                      // SerialNo
        //    Info[3] = DataProcess.Ta.ToString();                 // Nhiệt độ
        //    Info[4] = DataProcess.Project;                       // Dự án
        //    Info[5] = DataProcess.MotorPower;                    // Công suất lắp đặt
        //    Info[6] = DataProcess.MotorSpeed.ToString();         // Tốc độ
        //    Info[7] = DataProcess.MotorName;                     // Tên động cơ
        //    Info[8] = DataProcess.ReportNo;                      // STT

        //    for (int j = 0; j < 10; j++)
        //    {
        //        //Thông số đo kiểm
        //        Para1[0, j] = DataProcess.TaPoint[j].ToString();    // Nhiệt độ môi trường
        //        Para1[1, j] = DataProcess.rhoaPoint[j].ToString();  // Tỷ trọng đo kiểm
        //        Para1[2, j] = DataProcess.n2Point[j].ToString();    // Tốc độ đo kiểm
        //        Para1[3, j] = DataProcess.FlowPoint[j].ToString();  // Lưu lượng
        //        Para1[4, j] = DataProcess.PtPoint[j].ToString();    // Áp suất tổng
        //        Para1[5, j] = DataProcess.T_Point[j].ToString();    // Momen xoắn trên trục
        //        Para1[6, j] = DataProcess.Pr[j].ToString();         // Công suất tiêu thụ

        //        //Thông số đo quy đổi về điều kiện tiêu chuẩn
        //        Para2[0, j] = DataProcess.Std_FlowPoint[j].ToString();  // Lưu lượng
        //        Para2[1, j] = DataProcess.Std_PtPoint[j].ToString();    // Áp suất tổng
        //        Para2[2, j] = DataProcess.Std_PrPoint[j].ToString();    // Công suất tiêu thụ
        //        Para2[3, j] = DataProcess.PrtPoint[j].ToString();       // Công suất tiêu thụ tính theo T
        //        Para2[4, j] = DataProcess.Std_EtPoint[j].ToString();    // Hiệu suất tổng
        //        Para2[5, j] = DataProcess.Std_EttPoint[j].ToString();   // Hiệu suất tổng tính theo T

        //        //Thông số đo quy đổi về điều kiện làm việc
        //        Para3[0, j] = DataProcess.Tw.ToString();                // Nhiệt độ khí
        //        Para3[1, j] = DataProcess.rhow.ToString();              // Tỷ trọng khí
        //        Para3[2, j] = DataProcess.n1.ToString();                // Tốc độ guồng cánh
        //        Para3[3, j] = DataProcess.Ope_FlowPoint[j].ToString();  // Lưu lượng
        //        Para3[4, j] = DataProcess.Ope_PtPoint[j].ToString();    // Áp suất tổng
        //        Para3[5, j] = DataProcess.Ope_PrPoint[j].ToString();    // Công suất tiêu thụ
        //        Para3[6, j] = DataProcess.PrtPoint[j].ToString();       // Công suất tiêu thụ tính theo T
        //        Para3[7, j] = DataProcess.Ope_EtPoint[j].ToString();    // Hiệu suất tổng
        //        Para3[8, j] = DataProcess.Ope_EttPoint[j].ToString();   // Hiệu suất tổng tính theo T
        //    }

        //    //Thông số chạy thử cơ khí
        //    Para4[0] = DataProcess.Noise.ToString();           // Độ ồn
        //    Para4[1] = DataProcess.BearingVia_H.ToString();    // Độ rung - ngang   gối 1
        //    Para4[2] = DataProcess.BearingVia_H.ToString();    // Độ rung - ngang   gối 2
        //    Para4[3] = DataProcess.BearingVia.ToString();      // Độ rung - đứng    gối 1
        //    Para4[4] = DataProcess.BearingVia.ToString();      // Độ rung - đứng    gối 2
        //    Para4[5] = DataProcess.BearingVia_V.ToString();    // Độ rung - dọc     gối 1
        //    Para4[6] = DataProcess.BearingVia_V.ToString();    // Độ rung - dọc     gối 2
        //    Para4[7] = DataProcess.BearingTemp_H.ToString();   // Nhiệt độ - ngang  gối 1
        //    Para4[8] = DataProcess.BearingTemp_H.ToString();   // Nhiệt độ - ngang  gối 2
        //    Para4[9] = DataProcess.BearingTemp.ToString();     // Nhiệt độ - đứng   gối 1
        //    Para4[10] = DataProcess.BearingTemp.ToString();    // Nhiệt độ - đứng   gối 2
        //    Para4[11] = DataProcess.BearingTemp_V.ToString();  // Nhiệt độ - dọc    gối 1
        //    Para4[12] = DataProcess.BearingTemp_V.ToString();  // Nhiệt độ - dọc    gối 2

        //    //Kết luận
        //    Para5[0] = DataProcess.TestPerson;                 // Người thực hiện
        //    Para5[1] = DataProcess.TestWitness;                // Người chứng kiến
        //    Para5[2] = DataProcess.Approved;                   // Người chấp thuận
        //}

        ////Report_Tester
        //private void TransferData_tester()
        //{
        //    //Thông tin đặt hàng
        //    Info[0] = DataProcess.Customer;                      // Khách hàng
        //    Info[1] = DataProcess.Model;                         // Model
        //    Info[2] = DataProcess.SerialNo;                      // SerialNo
        //    Info[3] = DataProcess.Ta.ToString();                 // Nhiệt độ
        //    Info[4] = DataProcess.Project;                       // Dự án
        //    Info[5] = DataProcess.MotorPower;                    // Công suất lắp đặt
        //    Info[6] = DataProcess.MotorSpeed.ToString();         // Tốc độ
        //    Info[7] = DataProcess.MotorName;                     // Tên động cơ
        //    Info[8] = DataProcess.ReportNo;                      // STT
        //    Info[9] = DataProcess.Idm.ToString();
        //    Info[10] = DataProcess.e_motor.ToString();
        //    Info[11] = DataProcess.CosPhi.ToString();
        //    Info[12] = DataProcess.Tw.ToString();
        //    Info[13] = DataProcess.Td.ToString();
        //    Info[14] = DataProcess.Vdm.ToString();
        //    Info[15] = DataProcess.Pdm.ToString();




        //    for (int j = 0; j < 10; j++)
        //    {
        //        //Thông số đo kiểm
        //        Para1[0, j] = DataProcess.TaPoint[j].ToString();       // Nhiệt độ môi trường
        //        Para1[1, j] = DataProcess.PaPoint[j].ToString();       // Áp suất khí quyển
        //        //Para1[2, j] = DataProcess.huPoint[j].ToString();       // Độ ẩm kk
        //        Para1[2, j] = DataProcess.rhow.ToString();             // Tỷ trọng khí điều kiện thực tế
        //        Para1[3, j] = DataProcess.rhoaPoint[j].ToString();     // Tỷ trọng đo kiểm
        //        Para1[4, j] = DataProcess.n2Point[j].ToString();       // Tốc độ guồng cánh thực
        //        Para1[5, j] = DataProcess.n1Point[j].ToString();       // Tốc độ guồng cánh thiết kế
        //        Para1[6, j] = DataProcess.deltaP_Point[j].ToString();  // Chênh áp điểm đo Lưu lượng
        //        Para1[7, j] = DataProcess.Pe3Point[j].ToString();      // Chênh áp điểm đo Áp suất
        //        Para1[8, j] = DataProcess.TdPoint[j].ToString();       // Nhiệt độ điểm đo
        //        Para1[9, j] = DataProcess.T_Point[j].ToString();       // Momen xoắn trên trục
        //        Para1[10, j] = DataProcess.PwPoint[j].ToString();      // Công suất tiêu thụ
        //        Para1[11, j] = DataProcess.e_motor.ToString();         // Hiệu suất động cơ
        //        Para1[12, j] = DataProcess.e_noitruc.ToString();       // Hiệu suất nối trục
        //        Para1[13, j] = DataProcess.e_goitruc.ToString();       // Hiệu suất gối trục
        //        Para1[14, j] = DataProcess.e_botruyen.ToString();      // Hiệu suất bộ truyền

        //        //Thông số đo quy đổi về điều kiện tiêu chuẩn
        //        Para2[0, j] = DataProcess.Std_FlowPoint[j].ToString();  // Lưu lượng
        //        Para2[1, j] = DataProcess.Std_PsPoint[j].ToString();    // Áp suất tĩnh
        //        Para2[2, j] = DataProcess.Std_PtPoint[j].ToString();    // Áp suất tổng
        //        Para2[3, j] = DataProcess.Std_PrPoint[j].ToString();    // Công suất tiêu thụ
        //        Para2[4, j] = DataProcess.PrtPoint[j].ToString();       // Công suất tiêu thụ tính theo T
        //        Para2[5, j] = DataProcess.Std_EsPoint[j].ToString();    // Hiệu suất tĩnh
        //        Para2[6, j] = DataProcess.Std_EstPoint[j].ToString();   // Hiệu suất tĩnh tính theo T
        //        Para2[7, j] = DataProcess.Std_EtPoint[j].ToString();    // Hiệu suất tổng
        //        Para2[8, j] = DataProcess.Std_EttPoint[j].ToString();   // Hiệu suất tổng tính theo T

        //        //Thông số đo quy đổi về điều kiện làm việc
        //        Para3[0, j] = DataProcess.Tw.ToString();                // Nhiệt độ khí
        //        Para3[1, j] = DataProcess.rhow.ToString();              // Tỷ trọng khí
        //        Para3[2, j] = DataProcess.n1.ToString();                // Tốc độ guồng cánh
        //        Para3[3, j] = DataProcess.Ope_FlowPoint[j].ToString();  // Lưu lượng
        //        Para3[4, j] = DataProcess.Ope_PsPoint[j].ToString();    // Áp suất tĩnh
        //        Para3[5, j] = DataProcess.Ope_PtPoint[j].ToString();    // Áp suất tổng
        //        Para3[6, j] = DataProcess.Ope_PrPoint[j].ToString();    // Công suất tiêu thụ
        //        Para3[7, j] = DataProcess.PrtPoint[j].ToString();       // Công suất tiêu thụ tính theo T
        //        Para3[8, j] = DataProcess.Ope_EsPoint[j].ToString();    // Hiệu suất tĩnh
        //        Para3[9, j] = DataProcess.Ope_EstPoint[j].ToString();   // Hiệu suất tĩnh tính theo T
        //        Para3[10, j] = DataProcess.Ope_EtPoint[j].ToString();   // Hiệu suất tổng
        //        Para3[11, j] = DataProcess.Ope_EttPoint[j].ToString();  // Hiệu suất tổng tính theo T
        //    }

        //    //Thông số chạy thử cơ khí
        //    Para4[0] = DataProcess.Noise.ToString();           // Độ ồn
        //    Para4[1] = DataProcess.BearingVia_H.ToString();    // Độ rung - ngang   gối 1
        //    Para4[2] = DataProcess.BearingVia_H.ToString();    // Độ rung - ngang   gối 2
        //    Para4[3] = DataProcess.BearingVia.ToString();      // Độ rung - đứng    gối 1
        //    Para4[4] = DataProcess.BearingVia.ToString();      // Độ rung - đứng    gối 2
        //    Para4[5] = DataProcess.BearingVia_V.ToString();    // Độ rung - dọc     gối 1
        //    Para4[6] = DataProcess.BearingVia_V.ToString();    // Độ rung - dọc     gối 2
        //    Para4[7] = DataProcess.BearingTemp_H.ToString();   // Nhiệt độ - ngang  gối 1
        //    Para4[8] = DataProcess.BearingTemp_H.ToString();   // Nhiệt độ - ngang  gối 2
        //    Para4[9] = DataProcess.BearingTemp.ToString();     // Nhiệt độ - đứng   gối 1
        //    Para4[10] = DataProcess.BearingTemp.ToString();    // Nhiệt độ - đứng   gối 2
        //    Para4[11] = DataProcess.BearingTemp_V.ToString();  // Nhiệt độ - dọc    gối 1
        //    Para4[12] = DataProcess.BearingTemp_V.ToString();  // Nhiệt độ - dọc    gối 2

        //    //Kết luận
        //    Para5[0] = DataProcess.TestPerson;                 // Người thực hiện
        //    Para5[1] = DataProcess.TestWitness;                // Người chứng kiến
        //    Para5[2] = DataProcess.Approved;                   // Người chấp thuận
        //}

        ////============== Export to PDF ================
        //string _templatePath = "";  //template report
        ////string _templatePath = @"C:\PDFs\TEST-REPORT-NHIET-DO-LAM-VIEC-KHAC-20-DO.xlsx";
        //public void Report_1(string savePath)
        //{
        //    // Info = Khách hàng - Model - SerialNo - Nhiệt Độ - Dự án - Công suất lắp đặt - Tốc độ - Động cơ -STT
        //    // Para1 = Nhiệt độ môi trường - Tỷ trọng đo kiểm - Tốc độ đo kiểm - Lưu lượng - 
        //    //        Áp suất tĩnh - Áp suất tổng - Công suất tiêu thụ - Hiệu suất tĩnh - Hiệu suất tổng
        //    // Para2 = Lưu lượng - Áp suất tĩnh - Áp suất tổng - Công suất tiêu thụ - Hiệu suất tĩnh - Hiệu suất tổng
        //    // Para3 = Nhiệt độ khí - Tỷ trọng khí - Tốc độ guồng cạnh - Lưu lượng -  Áp suất tĩnh - Áp suất tổng
        //    //         Công suất tiêu thụ - Hiệu suất tĩnh - Hiệu suất tổng
        //    // Para4 = Độ ồn - Phương ngang1 - Phương ngang2 - Phương đứng1 - Phương đứng2 - Phương dọc1 - Phương dọc2 
        //    //                 Nhiệt ngang1 - Nhiệt ngang2 - Nhiệt đứng1 - Nhiệt đứng2 - Nhiệt dọc1 - Nhiệt dọc2
        //    // Para5 = Người thực hiện - Người chứng kiến - Người phê duyệt   
        //    try
        //    {
        //        TransferData_1();
        //        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        //        using (var package = new ExcelPackage(new FileInfo(_templatePath)))
        //        {
        //            var ws = package.Workbook.Worksheets[1];


        //            // Thông tin đặt hàng
        //            ws.Cells[12, 5].Value = Info[0];
        //            ws.Cells[14, 5].Value = Info[1];
        //            ws.Cells[13, 11].Value = Info[2];
        //            ws.Cells[12, 11].Value = Info[4];
        //            ws.Cells[16, 11].Value = Info[5];
        //            ws.Cells[20, 5].Value = Info[6];
        //            ws.Cells[15, 11].Value = Info[7];
        //            ws.Cells[18, 5].Value = Info[12];
        //            ws.Cells[19, 5].Value = Info[13];
        //            ws.Cells[17, 11].Value = Info[9];
        //            ws.Cells[18, 11].Value = Info[10];
        //            ws.Cells[19, 11].Value = Info[11];
        //            ws.Cells[16, 5].Value = Info[14];
        //            ws.Cells[17, 5].Value = Info[15];

        //            ws.Cells[11, 3].Value = Info[4] + "-" + Info[8];
        //            ws.Cells[11, 10].Value = DateTime.Now.ToString("dd") + "   /";
        //            ws.Cells[11, 11].Value = DateTime.Now.ToString("MM") + "   /";
        //            ws.Cells[11, 12].Value = DateTime.Now.ToString("yyyy");


        //            // Thông số đo kiểm
        //            for (int i = 23; i < 31; i++)
        //            {
        //                for (int j = 0; j < 10; j++)
        //                {
        //                    ws.Cells[i, j + 4].Value = Para1[i - 23, j];
        //                }
        //            }
        //            //Thông số đo kiểm đã quy đổi về điều kiện tiêu chuẩn
        //            for (int i = 33; i < 42; i++)
        //            {
        //                for (int j = 0; j < 10; j++)
        //                {
        //                    ws.Cells[i, j + 4].Value = Para2[i - 33, j];
        //                }
        //            }

        //            //Thông số đo kiểm đã quy đổi về điều kiện làm việc 
        //            for (int i = 44; i < 56; i++)
        //            {
        //                for (int j = 0; j < 10; j++)
        //                {
        //                    ws.Cells[i, j + 4].Value = Para3[i - 44, j];
        //                }
        //            }

        //            // Thêm hình ảnh
        //            var pwPic = ws.Drawings.AddPicture("Pw_Picture", new FileInfo(Path.Combine(savePath, "Pw_Picture.jpg")));
        //            pwPic.SetPosition(58, 8, 1, 25); // Vị trí tương đối, có thể cần chỉnh lại

        //            var prEfPic = ws.Drawings.AddPicture("PrEf_Picture", new FileInfo(Path.Combine(savePath, "PrEf_Picture.jpg")));
        //            prEfPic.SetPosition(79, 12, 1, 25);

        //            //Thông số chạy thử cơ khí
        //            ws.Cells[108, 4].Value = Para4[0];    // Độ ồn
        //            ws.Cells[110, 4].Value = Para4[1];    // Ngang 1
        //            ws.Cells[110, 6].Value = Para4[2];    // Ngang 2
        //            ws.Cells[111, 4].Value = Para4[3];    // đứng 1
        //            ws.Cells[111, 6].Value = Para4[4];    // đứng 2
        //            ws.Cells[112, 4].Value = Para4[5];    // dọc 1
        //            ws.Cells[113, 6].Value = Para4[6];    // dọc 2
        //            ws.Cells[110, 9].Value = Para4[7];    // Nhiệt Ngang 1
        //            ws.Cells[110, 11].Value = Para4[8];   // Nhiệt Ngang 2
        //            ws.Cells[111, 9].Value = Para4[9];    // Nhiệt đứng 1
        //            ws.Cells[111, 11].Value = Para4[10];  // Nhiệt đứng 2
        //            ws.Cells[112, 9].Value = Para4[11];   // Nhiệt dọc 1
        //            ws.Cells[112, 11].Value = Para4[12];  // Nhiệt dọc 2

        //            //Kết luận
        //            ws.Cells[121, 2].Value = Para5[0];   // Người thực hiện
        //            ws.Cells[121, 4].Value = Para5[1];   // Người chứng kiến
        //            ws.Cells[121, 9].Value = Para5[2];  // Phê duyệt

        //            var name = Info[4] + "-" + Info[8];
        //            for (int i = 0; i < name.Length; i++)
        //            {
        //                if (name.Substring(i, 1) == "/")
        //                {
        //                    name = name.Remove(i, 1).Insert(i, "_");
        //                }
        //            }

        //            string filename = name + "_" +
        //                      DateTime.Now.ToString("yyyy") + "_" +
        //                      DateTime.Now.ToString("MM") + "_" +
        //                      DateTime.Now.ToString("dd") + "_" +
        //                      DateTime.Now.ToString("HH") + "_" +
        //                      DateTime.Now.ToString("mm") + "_" +
        //                      DateTime.Now.ToString("ss");
        //            string FilePath = savePath + @"\" + filename + ".xlsx";
        //            package.SaveAs(new FileInfo(FilePath));
        //            MessageBox.Show("Export Report Successfully to \n" + FilePath, "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //            DataProcess.Done = 1;
        //        }
        //    }
        //    catch (Exception exml)
        //    {
        //        MessageBox.Show("Export Report fail \n" + exml.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}

        //public void Report_2(string savePath)
        //{
        //    //Info = Khách hàng - Model - SerialNo - Nhiệt Độ - Dự án - Công suất lắp đặt - Tốc độ - Động cơ -STT
        //    //Para1 = Nhiệt độ môi trường - Tỷ trọng đo kiểm - Tốc độ đo kiểm - Lưu lượng - 
        //    //       Áp suất tĩnh - Công suất tiêu thụ
        //    //Para2 = Lưu lượng - Áp suất tĩnh - Công suất tiêu thụ
        //    //Para3 = Nhiệt độ khí - Tỷ trọng khí - Tốc độ guồng cạnh - Lưu lượng -  Áp suất tĩnh
        //    //        Công suất tiêu thụ
        //    //Para4 = Độ ồn - Phương ngang1 - Phương ngang2 - Phương đứng1 - Phương đứng2 - Phương dọc1 - Phương dọc2 
        //    //                Nhiệt ngang1 - Nhiệt ngang2 - Nhiệt đứng1 - Nhiệt đứng2 - Nhiệt dọc1 - Nhiệt dọc2
        //    try
        //    {
        //        TransferData_2();
        //        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        //        using (var package = new ExcelPackage(new FileInfo(_templatePath)))
        //        {
        //            var ws = package.Workbook.Worksheets[2];

        //            // Thông tin đặt hàng
        //            ws.Cells[12, 3].Value = Info[0];
        //            ws.Cells[13, 3].Value = Info[1];
        //            ws.Cells[14, 3].Value = Info[2];
        //            ws.Cells[15, 3].Value = Info[3];
        //            ws.Cells[12, 11].Value = Info[4];
        //            ws.Cells[13, 11].Value = Info[5];
        //            ws.Cells[14, 11].Value = Info[6];
        //            ws.Cells[15, 11].Value = Info[7];
        //            ws.Cells[11, 3].Value = Info[4] + "-" + Info[8];
        //            ws.Cells[11, 10].Value = DateTime.Now.ToString("dd") + "   /";
        //            ws.Cells[11, 11].Value = DateTime.Now.ToString("MM") + "   /";
        //            ws.Cells[11, 12].Value = DateTime.Now.ToString("yyyy");

        //            // Thông số đo kiểm
        //            for (int i = 18; i < 25; i++)
        //                for (int j = 0; j < 10; j++)
        //                    ws.Cells[i, j + 4].Value = Para1[i - 18, j];

        //            // Thông số quy đổi điều kiện tiêu chuẩn
        //            for (int i = 27; i < 31; i++)
        //                for (int j = 0; j < 10; j++)
        //                    ws.Cells[i, j + 4].Value = Para2[i - 27, j];

        //            // Thông số quy đổi điều kiện làm việc
        //            for (int i = 33; i < 40; i++)
        //                for (int j = 0; j < 10; j++)
        //                    ws.Cells[i, j + 4].Value = Para3[i - 33, j];

        //            // Thêm hình ảnh
        //            var pwPic = ws.Drawings.AddPicture("Pw_Picture", new FileInfo(Path.Combine(savePath, "Pw_Picture.jpg")));
        //            pwPic.SetPosition(41, 8, 1, 25);
        //            var prEfPic = ws.Drawings.AddPicture("PrEf_Picture", new FileInfo(Path.Combine(savePath, "PrEf_Picture.jpg")));
        //            prEfPic.SetPosition(58, 12, 1, 25);

        //            // Thông số chạy thử cơ khí
        //            ws.Cells[82, 4].Value = Para4[0];
        //            ws.Cells[84, 4].Value = Para4[1];
        //            ws.Cells[84, 6].Value = Para4[2];
        //            ws.Cells[85, 4].Value = Para4[3];
        //            ws.Cells[85, 6].Value = Para4[4];
        //            ws.Cells[86, 4].Value = Para4[5];
        //            ws.Cells[86, 6].Value = Para4[6];
        //            ws.Cells[84, 9].Value = Para4[7];
        //            ws.Cells[84, 11].Value = Para4[8];
        //            ws.Cells[85, 9].Value = Para4[9];
        //            ws.Cells[85, 11].Value = Para4[10];
        //            ws.Cells[86, 9].Value = Para4[11];
        //            ws.Cells[86, 11].Value = Para4[12];

        //            // Kết luận
        //            ws.Cells[97, 2].Value = Para5[0];
        //            ws.Cells[97, 4].Value = Para5[1];
        //            ws.Cells[97, 10].Value = Para5[2];

        //            var name = Info[4] + "-" + Info[8];
        //            for (int i = 0; i < name.Length; i++)
        //                if (name.Substring(i, 1) == "/")
        //                    name = name.Remove(i, 1).Insert(i, "_");
        //            string filename = name + "_" +
        //                              DateTime.Now.ToString("yyyy") + "_" +
        //                              DateTime.Now.ToString("MM") + "_" +
        //                              DateTime.Now.ToString("dd") + "_" +
        //                              DateTime.Now.ToString("HH") + "_" +
        //                              DateTime.Now.ToString("mm") + "_" +
        //                              DateTime.Now.ToString("ss");
        //            string FilePath = savePath + @"\" + filename + ".xlsx";
        //            package.SaveAs(new FileInfo(FilePath));
        //            MessageBox.Show("Export Report Successfully to \n" + FilePath, "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //            DataProcess.Done = 1;
        //        }
        //    }
        //    catch (Exception exml)
        //    {
        //        MessageBox.Show("Export Report fail \n" + exml.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}

        //public void Report_3(string savePath)
        //{
        //    //Info = Khách hàng - Model - SerialNo - Nhiệt Độ - Dự án - Công suất lắp đặt - Tốc độ - Động cơ -STT
        //    //Para1 = Nhiệt độ môi trường - Tỷ trọng đo kiểm - Tốc độ đo kiểm - Lưu lượng - 
        //    //       Áp suất tổng - Công suất tiêu tổng
        //    //Para2 = Lưu lượng - Áp suất tổng - Công suất tiêu thụ
        //    //Para3 = Nhiệt độ khí - Tỷ trọng khí - Tốc độ guồng cạnh - Lưu lượng -  Áp suất tổng
        //    //        Công suất tiêu thụ
        //    //Para4 = Độ ồn - Phương ngang1 - Phương ngang2 - Phương đứng1 - Phương đứng2 - Phương dọc1 - Phương dọc2 
        //    //                Nhiệt ngang1 - Nhiệt ngang2 - Nhiệt đứng1 - Nhiệt đứng2 - Nhiệt dọc1 - Nhiệt dọc2
        //    try
        //    {
        //        TransferData_3();
        //        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        //        using (var package = new ExcelPackage(new FileInfo(_templatePath)))
        //        {
        //            var ws = package.Workbook.Worksheets[3];

        //            ws.Cells[12, 4].Value = Info[0];
        //            ws.Cells[14, 4].Value = Info[1];
        //            ws.Cells[13, 11].Value = Info[2];
        //            ws.Cells[12, 11].Value = Info[4];
        //            ws.Cells[16, 11].Value = Info[5];
        //            ws.Cells[20, 4].Value = Info[6];
        //            ws.Cells[15, 11].Value = Info[7];
        //            ws.Cells[11, 3].Value = Info[4] + "-" + Info[8];
        //            ws.Cells[11, 10].Value = DateTime.Now.ToString("dd") + "   /";
        //            ws.Cells[11, 11].Value = DateTime.Now.ToString("MM") + "   /";
        //            ws.Cells[11, 12].Value = DateTime.Now.ToString("yyyy");

        //            for (int i = 18; i < 25; i++)
        //                for (int j = 0; j < 10; j++)
        //                    ws.Cells[i, j + 4].Value = Para1[i - 18, j];

        //            for (int i = 27; i < 31; i++)
        //                for (int j = 0; j < 10; j++)
        //                    ws.Cells[i, j + 4].Value = Para2[i - 27, j];

        //            for (int i = 44; i < 55; i++)
        //                for (int j = 0; j < 10; j++)
        //                    ws.Cells[i, j + 4].Value = Para3[i - 44, j];

        //            var pwPic = ws.Drawings.AddPicture("Pw_Picture", new FileInfo(Path.Combine(savePath, "Pw_Picture.jpg")));
        //            pwPic.SetPosition(46, 8, 1, 25);
        //            var prEfPic = ws.Drawings.AddPicture("PrEf_Picture", new FileInfo(Path.Combine(savePath, "PrEf_Picture.jpg")));
        //            prEfPic.SetPosition(58, 12, 1, 25);

        //            ws.Cells[82, 4].Value = Para4[0];
        //            ws.Cells[84, 4].Value = Para4[1];
        //            ws.Cells[84, 6].Value = Para4[2];
        //            ws.Cells[85, 4].Value = Para4[3];
        //            ws.Cells[85, 6].Value = Para4[4];
        //            ws.Cells[86, 4].Value = Para4[5];
        //            ws.Cells[86, 6].Value = Para4[6];
        //            ws.Cells[84, 9].Value = Para4[7];
        //            ws.Cells[84, 11].Value = Para4[8];
        //            ws.Cells[85, 9].Value = Para4[9];
        //            ws.Cells[85, 11].Value = Para4[10];
        //            ws.Cells[86, 9].Value = Para4[11];
        //            ws.Cells[86, 11].Value = Para4[12];

        //            ws.Cells[124, 2].Value = Para5[0];
        //            ws.Cells[124, 4].Value = Para5[1];
        //            ws.Cells[124, 10].Value = Para5[2];

        //            var name = Info[4] + "-" + Info[8];
        //            for (int i = 0; i < name.Length; i++)
        //                if (name.Substring(i, 1) == "/")
        //                    name = name.Remove(i, 1).Insert(i, "_");
        //            string filename = name + "_" +
        //                              DateTime.Now.ToString("yyyy") + "_" +
        //                              DateTime.Now.ToString("MM") + "_" +
        //                              DateTime.Now.ToString("dd") + "_" +
        //                              DateTime.Now.ToString("HH") + "_" +
        //                              DateTime.Now.ToString("mm") + "_" +
        //                              DateTime.Now.ToString("ss");
        //            string FilePath = savePath + @"\" + filename + ".xlsx";
        //            package.SaveAs(new FileInfo(FilePath));
        //            MessageBox.Show("Export Report Successfully to \n" + FilePath, "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //            DataProcess.Done = 1;
        //        }
        //    }
        //    catch (Exception exml)
        //    {
        //        MessageBox.Show("Export Report fail \n" + exml.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}

        //public void Report_4(string savePath)
        //{
        //    //Info = Khách hàng - Model - SerialNo - Nhiệt Độ - Dự án - Công suất lắp đặt - Tốc độ - Động cơ -STT
        //    //Para1 = Nhiệt độ môi trường - Tỷ trọng đo kiểm - Tốc độ đo kiểm - Lưu lượng - 
        //    //       Áp suất tổng - Hiệu suất tĩnh - Công suất tiêu thụ
        //    //Para2 = Lưu lượng - Áp suất tổng - Hiệu suất tĩnh - Công suất tiêu thụ
        //    //Para3 = Nhiệt độ khí - Tỷ trọng khí - Tốc độ guồng cạnh - Lưu lượng -  Áp suất tổng
        //    //        - Hiệu suất tĩnh - Công suất tiêu thụ
        //    //Para4 = Độ ồn - Phương ngang1 - Phương ngang2 - Phương đứng1 - Phương đứng2 - Phương dọc1 - Phương dọc2 
        //    //                Nhiệt ngang1 - Nhiệt ngang2 - Nhiệt đứng1 - Nhiệt đứng2 - Nhiệt dọc1 - Nhiệt dọc2
        //    try
        //    {
        //        TransferData_4();
        //        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        //        using (var package = new ExcelPackage(new FileInfo(_templatePath)))
        //        {
        //            var ws = package.Workbook.Worksheets[4];

        //            ws.Cells[12, 3].Value = Info[0];
        //            ws.Cells[13, 3].Value = Info[1];
        //            ws.Cells[14, 3].Value = Info[2];
        //            ws.Cells[15, 3].Value = Info[3];
        //            ws.Cells[12, 11].Value = Info[4];
        //            ws.Cells[13, 11].Value = Info[5];
        //            ws.Cells[14, 11].Value = Info[6];
        //            ws.Cells[15, 11].Value = Info[7];
        //            ws.Cells[11, 3].Value = Info[4] + "-" + Info[8];
        //            ws.Cells[11, 10].Value = DateTime.Now.ToString("dd") + "   /";
        //            ws.Cells[11, 11].Value = DateTime.Now.ToString("MM") + "   /";
        //            ws.Cells[11, 12].Value = DateTime.Now.ToString("yyyy");

        //            for (int i = 18; i < 25; i++)
        //                for (int j = 0; j < 10; j++)
        //                    ws.Cells[i, j + 4].Value = Para1[i - 18, j];

        //            for (int i = 27; i < 33; i++)
        //                for (int j = 0; j < 10; j++)
        //                    ws.Cells[i, j + 4].Value = Para2[i - 27, j];

        //            for (int i = 33; i < 40; i++)
        //                for (int j = 0; j < 10; j++)
        //                    ws.Cells[i, j + 4].Value = Para3[i - 33, j];

        //            var pwPic = ws.Drawings.AddPicture("Pw_Picture", new FileInfo(Path.Combine(savePath, "Pw_Picture.jpg")));
        //            pwPic.SetPosition(45, 8, 1, 25);
        //            var prEfPic = ws.Drawings.AddPicture("PrEf_Picture", new FileInfo(Path.Combine(savePath, "PrEf_Picture.jpg")));
        //            prEfPic.SetPosition(62, 12, 1, 25);

        //            ws.Cells[86, 4].Value = Para4[0];
        //            ws.Cells[88, 4].Value = Para4[1];
        //            ws.Cells[88, 6].Value = Para4[2];
        //            ws.Cells[89, 4].Value = Para4[3];
        //            ws.Cells[89, 6].Value = Para4[4];
        //            ws.Cells[90, 4].Value = Para4[5];
        //            ws.Cells[90, 6].Value = Para4[6];
        //            ws.Cells[88, 9].Value = Para4[7];
        //            ws.Cells[88, 11].Value = Para4[8];
        //            ws.Cells[89, 9].Value = Para4[9];
        //            ws.Cells[89, 11].Value = Para4[10];
        //            ws.Cells[90, 9].Value = Para4[11];
        //            ws.Cells[90, 11].Value = Para4[12];

        //            ws.Cells[101, 2].Value = Para5[0];
        //            ws.Cells[101, 4].Value = Para5[1];
        //            ws.Cells[101, 10].Value = Para5[2];

        //            var name = Info[4] + "-" + Info[8];
        //            for (int i = 0; i < name.Length; i++)
        //                if (name.Substring(i, 1) == "/")
        //                    name = name.Remove(i, 1).Insert(i, "_");
        //            string filename = name + "_" +
        //                              DateTime.Now.ToString("yyyy") + "_" +
        //                              DateTime.Now.ToString("MM") + "_" +
        //                              DateTime.Now.ToString("dd") + "_" +
        //                              DateTime.Now.ToString("HH") + "_" +
        //                              DateTime.Now.ToString("mm") + "_" +
        //                              DateTime.Now.ToString("ss");
        //            string FilePath = savePath + @"\" + filename + ".xlsx";
        //            package.SaveAs(new FileInfo(FilePath));
        //            MessageBox.Show("Export Report Successfully to \n" + FilePath, "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //            DataProcess.Done = 1;
        //        }
        //    }
        //    catch (Exception exml)
        //    {
        //        MessageBox.Show("Export Report fail \n" + exml.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }

        //}

        //public void Report_5(string savePath)
        //{
        //    //Info = Khách hàng - Model - SerialNo - Nhiệt Độ - Dự án - Công suất lắp đặt - Tốc độ - Động cơ -STT
        //    //Para1 = Nhiệt độ môi trường - Tỷ trọng đo kiểm - Tốc độ đo kiểm - Lưu lượng - 
        //    //       Áp suất tổng - Hiệu suất tổng - Công suất tiêu thụ
        //    //Para2 = Lưu lượng - Áp suất tổng - Hiệu suất tổng - Công suất tiêu thụ
        //    //Para3 = Nhiệt độ khí - Tỷ trọng khí - Tốc độ guồng cạnh - Lưu lượng -  Áp suất tổng
        //    //        - Hiệu suất tổng - Công suất tiêu thụ
        //    //Para4 = Độ ồn - Phương ngang1 - Phương ngang2 - Phương đứng1 - Phương đứng2 - Phương dọc1 - Phương dọc2 
        //    //                Nhiệt ngang1 - Nhiệt ngang2 - Nhiệt đứng1 - Nhiệt đứng2 - Nhiệt dọc1 - Nhiệt dọc2
        //    try
        //    {
        //        TransferData_5();
        //        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        //        using (var package = new ExcelPackage(new FileInfo(_templatePath)))
        //        {
        //            var ws = package.Workbook.Worksheets[5];

        //            ws.Cells[12, 3].Value = Info[0];
        //            ws.Cells[13, 3].Value = Info[1];
        //            ws.Cells[14, 3].Value = Info[2];
        //            ws.Cells[15, 3].Value = Info[3];
        //            ws.Cells[12, 11].Value = Info[4];
        //            ws.Cells[13, 11].Value = Info[5];
        //            ws.Cells[14, 11].Value = Info[6];
        //            ws.Cells[15, 11].Value = Info[7];
        //            ws.Cells[11, 3].Value = Info[4] + "-" + Info[8];
        //            ws.Cells[11, 10].Value = DateTime.Now.ToString("dd") + "   /";
        //            ws.Cells[11, 11].Value = DateTime.Now.ToString("MM") + "   /";
        //            ws.Cells[11, 12].Value = DateTime.Now.ToString("yyyy");

        //            for (int i = 18; i < 25; i++)
        //                for (int j = 0; j < 10; j++)
        //                    ws.Cells[i, j + 4].Value = Para1[i - 18, j];

        //            for (int i = 27; i < 33; i++)
        //                for (int j = 0; j < 10; j++)
        //                    ws.Cells[i, j + 4].Value = Para2[i - 27, j];

        //            for (int i = 35; i < 44; i++)
        //                for (int j = 0; j < 10; j++)
        //                    ws.Cells[i, j + 4].Value = Para3[i - 35, j];

        //            var pwPic = ws.Drawings.AddPicture("Pw_Picture", new FileInfo(Path.Combine(savePath, "Pw_Picture.jpg")));
        //            pwPic.SetPosition(45, 8, 1, 25);
        //            var prEfPic = ws.Drawings.AddPicture("PrEf_Picture", new FileInfo(Path.Combine(savePath, "PrEf_Picture.jpg")));
        //            prEfPic.SetPosition(62, 12, 1, 25);

        //            ws.Cells[86, 4].Value = Para4[0];
        //            ws.Cells[88, 4].Value = Para4[1];
        //            ws.Cells[88, 6].Value = Para4[2];
        //            ws.Cells[89, 4].Value = Para4[3];
        //            ws.Cells[89, 6].Value = Para4[4];
        //            ws.Cells[90, 4].Value = Para4[5];
        //            ws.Cells[90, 6].Value = Para4[6];
        //            ws.Cells[88, 9].Value = Para4[7];
        //            ws.Cells[88, 11].Value = Para4[8];
        //            ws.Cells[89, 9].Value = Para4[9];
        //            ws.Cells[89, 11].Value = Para4[10];
        //            ws.Cells[90, 9].Value = Para4[11];
        //            ws.Cells[90, 11].Value = Para4[12];

        //            ws.Cells[101, 2].Value = Para5[0];
        //            ws.Cells[101, 4].Value = Para5[1];
        //            ws.Cells[101, 10].Value = Para5[2];

        //            var name = Info[4] + "-" + Info[8];
        //            for (int i = 0; i < name.Length; i++)
        //                if (name.Substring(i, 1) == "/")
        //                    name = name.Remove(i, 1).Insert(i, "_");
        //            string filename = name + "_" +
        //                              DateTime.Now.ToString("yyyy") + "_" +
        //                              DateTime.Now.ToString("MM") + "_" +
        //                              DateTime.Now.ToString("dd") + "_" +
        //                              DateTime.Now.ToString("HH") + "_" +
        //                              DateTime.Now.ToString("mm") + "_" +
        //                              DateTime.Now.ToString("ss");
        //            string FilePath = savePath+ @"\" + filename + ".xlsx";
        //            package.SaveAs(new FileInfo(FilePath));
        //            MessageBox.Show("Export Report Successfully to \n" + FilePath, "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //            DataProcess.Done = 1;
        //        }
        //    }
        //    catch (Exception exml)
        //    {
        //        MessageBox.Show("Export Report fail \n" + exml.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}

        //public void Report_tester(string savePath)
        //{
        //    //Info = Khách hàng - Model - SerialNo - Nhiệt Độ - Dự án - Công suất lắp đặt - Tốc độ - Động cơ -STT
        //    //Para1 = Nhiệt độ môi trường - Tỷ trọng đo kiểm - Tốc độ đo kiểm - Lưu lượng - 
        //    //       Áp suất tổng - Hiệu suất tổng - Công suất tiêu thụ
        //    //Para2 = Lưu lượng - Áp suất tổng - Hiệu suất tổng - Công suất tiêu thụ
        //    //Para3 = Nhiệt độ khí - Tỷ trọng khí - Tốc độ guồng cạnh - Lưu lượng -  Áp suất tổng
        //    //        - Hiệu suất tổng - Công suất tiêu thụ
        //    //Para4 = Độ ồn - Phương ngang1 - Phương ngang2 - Phương đứng1 - Phương đứng2 - Phương dọc1 - Phương dọc2 
        //    //                Nhiệt ngang1 - Nhiệt ngang2 - Nhiệt đứng1 - Nhiệt đứng2 - Nhiệt dọc1 - Nhiệt dọc2
        //    try
        //    {
        //        TransferData_tester();
        //        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        //        using (var package = new ExcelPackage(new FileInfo(_templatePath)))
        //        {
        //            var ws = package.Workbook.Worksheets[6];

        //            ws.Cells[12, 3].Value = Info[0];
        //            ws.Cells[13, 3].Value = Info[1];
        //            ws.Cells[14, 3].Value = Info[2];
        //            ws.Cells[15, 3].Value = Info[3];
        //            ws.Cells[12, 11].Value = Info[4];
        //            ws.Cells[13, 11].Value = Info[5];
        //            ws.Cells[14, 11].Value = Info[6];
        //            ws.Cells[15, 11].Value = Info[7];
        //            ws.Cells[11, 3].Value = Info[4] + "-" + Info[8];
        //            ws.Cells[11, 10].Value = DateTime.Now.ToString("dd") + "   /";
        //            ws.Cells[11, 11].Value = DateTime.Now.ToString("MM") + "   /";
        //            ws.Cells[11, 12].Value = DateTime.Now.ToString("yyyy");

        //            for (int i = 18; i < 33; i++)
        //                for (int j = 0; j < 10; j++)
        //                    ws.Cells[i, j + 4].Value = Para1[i - 18, j];

        //            for (int i = 35; i < 44; i++)
        //                for (int j = 0; j < 10; j++)
        //                    ws.Cells[i, j + 4].Value = Para2[i - 35, j];

        //            for (int i = 46; i < 57; i++)
        //                for (int j = 0; j < 10; j++)
        //                    ws.Cells[i, j + 4].Value = Para3[i - 46, j];

        //            var pwPic = ws.Drawings.AddPicture("Pw_Picture", new FileInfo(Path.Combine(savePath, "Pw_Picture.jpg")));
        //            pwPic.SetPosition(59, 8, 1, 25);
        //            var prEfPic = ws.Drawings.AddPicture("PrEf_Picture", new FileInfo(Path.Combine(savePath, "PrEf_Picture.jpg")));
        //            prEfPic.SetPosition(76, 12, 1, 25);

        //            ws.Cells[100, 4].Value = Para4[0];
        //            ws.Cells[102, 4].Value = Para4[1];
        //            ws.Cells[102, 6].Value = Para4[2];
        //            ws.Cells[103, 4].Value = Para4[3];
        //            ws.Cells[103, 6].Value = Para4[4];
        //            ws.Cells[104, 4].Value = Para4[5];
        //            ws.Cells[104, 6].Value = Para4[6];
        //            ws.Cells[102, 9].Value = Para4[7];
        //            ws.Cells[102, 11].Value = Para4[8];
        //            ws.Cells[103, 9].Value = Para4[9];
        //            ws.Cells[103, 11].Value = Para4[10];
        //            ws.Cells[104, 9].Value = Para4[11];
        //            ws.Cells[104, 11].Value = Para4[12];

        //            ws.Cells[115, 2].Value = Para5[0];
        //            ws.Cells[115, 4].Value = Para5[1];
        //            ws.Cells[115, 10].Value = Para5[2];

        //            var name = Info[4] + "-" + Info[8];
        //            for (int i = 0; i < name.Length; i++)
        //                if (name.Substring(i, 1) == "/")
        //                    name = name.Remove(i, 1).Insert(i, "_");
        //            string filename = name + "_" +
        //                              DateTime.Now.ToString("yyyy") + "_" +
        //                              DateTime.Now.ToString("MM") + "_" +
        //                              DateTime.Now.ToString("dd") + "_" +
        //                              DateTime.Now.ToString("HH") + "_" +
        //                              DateTime.Now.ToString("mm") + "_" +
        //                              DateTime.Now.ToString("ss");
        //            string FilePath = savePath + @"\" + filename + ".xlsx";
        //            package.SaveAs(new FileInfo(FilePath));
        //            MessageBox.Show("Export Report Successfully to \n" + FilePath, "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //            DataProcess.Done = 1;
        //        }
        //    }
        //    catch (Exception exml)
        //    {
        //        MessageBox.Show("Export Report fail \n" + exml.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}
        //#endregion
    
    }
}
