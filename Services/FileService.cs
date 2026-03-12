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
using System.Globalization;



namespace TESMEA_TMS.Services
{
    public interface IFileService
    {
        ThongSoDauVao ImportCalculation(string filePath);
        Task ExportExcelTestResult(string outputPath, string option, ThongTinDuAn project, ThongSoDauVao input, KetQuaDoKiem res);
        Task ExportReportTestResult(string outputPath, string option, ThongTinDuAn project, ThongSoDauVao input, KetQuaDoKiem res);
        Task ExportReportTestResult_full(string outputPath, ThongTinDuAn project, ThongSoDauVao input, KetQuaDoKiem res);

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
        public async Task ExportExcelTestResult(string outputPath, string option, ThongTinDuAn project, ThongSoDauVao tsdv, KetQuaDoKiem ketQuaDoKiem)
        {
            try
            {
                if (Common.IsFileLocked(outputPath))
                {
                    MessageBoxHelper.ShowWarning("File đang được mở bởi ứng dụng khác. Vui lòng tắt file trước khi xuất báo cáo!");
                    return;
                }
                string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "ketquadokiem_template1.xlsx");
                if (!File.Exists(templatePath))
                {
                    MessageBoxHelper.ShowWarning("File mẫu kết quả không tồn tại");
                    return;
                }


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

                    foreach (var item in freqGroups)
                    {
                        if (option == "FULL")
                        {
                            await ExportDesignCondition(package, item.DieuKienDoKiem, project, $"{Math.Round(item.TanSo, 0)}Hz - Design Condition");
                            await ExportNormalizedCondition(package, item.DieuKienDoKiem, item.HieuChuanTieuChuan, project, $"{Math.Round(item.TanSo, 0)}Hz - Normalized Condition");
                            await ExportOperatingCondition(package, item.DieuKienDoKiem, item.HieuChuanLamViec, project, $"{Math.Round(item.TanSo, 0)}Hz - Operating Condition");
                            await ExportFullCondition(package, item.DieuKienDoKiem, item.HieuChuanTieuChuan, item.HieuChuanLamViec, project, $"{Math.Round(item.TanSo, 0)}Hz - Full");
                        }
                        else
                        {
                            // Chỉ giữ lại sheet Calculation và sheet của option
                            var keepSheets = new List<string> { "Calculation", "" };
                            if (option == "DESIGN") keepSheets[1] = "Design Condition";
                            else if (option == "NORMALIZED") keepSheets[1] = "Normalized Condition";
                            else if (option == "OPERATION") keepSheets[1] = "Operating Condition";

                            //// Xóa các sheet không cần thiết
                            //for (int i = package.Workbook.Worksheets.Count - 1; i >= 0; i--)
                            //{
                            //    var sheet = package.Workbook.Worksheets[i];
                            //    if (!keepSheets.Contains(sheet.Name))
                            //        package.Workbook.Worksheets.Delete(sheet.Name);
                            //}

                            // Fill dữ liệu cho sheet option
                            if (option == "DESIGN")
                                await ExportDesignCondition(package, item.DieuKienDoKiem, project, $"{Math.Round(item.TanSo, 0)}Hz - Design Condition");
                            else if (option == "NORMALIZED")
                                await ExportNormalizedCondition(package, item.DieuKienDoKiem, item.HieuChuanTieuChuan, project, $"{Math.Round(item.TanSo, 0)}Hz - Normalized Condition");
                            else if (option == "OPERATION")
                                await ExportOperatingCondition(package, item.DieuKienDoKiem, item.HieuChuanLamViec, project, $"{Math.Round(item.TanSo, 0)}Hz - Operating Condition");
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
                    package.Workbook.Worksheets.Delete("Design Condition");
                    package.Workbook.Worksheets.Delete("Normalized Condition");
                    package.Workbook.Worksheets.Delete("Operating Condition");
                    package.Workbook.Worksheets.Delete("Full");


                    //if (option == "FULL")
                    //{
                    //    package.Workbook.Worksheets.Delete("Design Condition");
                    //    package.Workbook.Worksheets.Delete("Normalized Condition");
                    //    package.Workbook.Worksheets.Delete("Operating Condition");
                    //    package.Workbook.Worksheets.Delete("Full");
                    //}
                    //else
                    //{
                    //    if(option == "DESIGN")
                    //        package.Workbook.Worksheets.Delete("Design Condition");
                    //    else if (option == "NORMALIZED")
                    //        package.Workbook.Worksheets.Delete("Normalized Condition");
                    //    else if (option == "OPERATION")
                    //        package.Workbook.Worksheets.Delete("Operating Condition");
                    //}
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
                    for (int col = 3; col <= ws.Dimension.End.Column; col++)
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
            catch (BusinessException ex)
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
                    ws.Cells[29, col].Value = i + 1;
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

                var templateSheet = package.Workbook.Worksheets["Design Condition"];
                if (templateSheet == null)
                {
                    throw new Exception("Không tìm thấy worksheet template 'Design Condition'");
                }
                var ws = package.Workbook.Worksheets.Add(sheetName, templateSheet);

                await FillThongTinChung(ws, project.ThongTinChung);

                for (int i = 0; i < data.Count; i++)
                {
                    var item = data.ElementAtOrDefault(i);
                    ws.Cells[19, 4 + i].Value = i + 1;
                    ws.Cells[20, 4 + i].Value = item?.HieuChinhLuuLuongTheTichTheoRPM;
                    ws.Cells[21, 4 + i].Value = item?.ApSuatTinh;
                    ws.Cells[22, 4 + i].Value = item?.ApSuatTong;
                    ws.Cells[23, 4 + i].Value = item?.CongSuatDongCoTaiDieuKienDoKiem;
                    ws.Cells[24, 4 + i].Value = item?.HieuSuatTinh;
                    ws.Cells[25, 4 + i].Value = item?.HieuSuatTong;


                    if (item != null)
                    {
                        // Power Chart
                        xMaxValue_powerchart = Math.Max(xMaxValue_powerchart, (double)item.HieuChinhLuuLuongTheTichTheoRPM);
                        yMaxValue_powerchart = Math.Max(yMaxValue_powerchart, (double)item.CongSuatDongCoTaiDieuKienDoKiem);

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
                var templateSheet = package.Workbook.Worksheets["Normalized Condition"];
                if (templateSheet == null)
                {
                    throw new Exception("Không tìm thấy worksheet template 'Normalized Condition'");
                }
                var ws = package.Workbook.Worksheets.Add(sheetName, templateSheet);

                await FillThongTinChung(ws, project.ThongTinChung);

                for (int i = 0; i < doKiem.Count; i++)
                {
                    var item = tieuChuan.ElementAtOrDefault(i);
                    var item1 = doKiem.ElementAtOrDefault(i);

                    ws.Cells[19, 4 + i].Value = i + 1;
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
                var templateSheet = package.Workbook.Worksheets["Operating Condition"];
                if (templateSheet == null)
                {
                    throw new Exception("Không tìm thấy worksheet template 'Operating Condition'.");
                }
                var ws = package.Workbook.Worksheets.Add(sheetName, templateSheet);
                await FillThongTinChung(ws, project.ThongTinChung);

                for (int i = 0; i < doKiem.Count; i++)
                {
                    var item = lamViec.ElementAtOrDefault(i);
                    var item1 = doKiem.ElementAtOrDefault(i);
                    ws.Cells[19, 4 + i].Value = i + 1;
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
                var templateSheet = package.Workbook.Worksheets["Full"];
                if (templateSheet == null)
                {
                    throw new Exception("Không tìm thấy worksheet 'Full'.");
                }
                var ws = package.Workbook.Worksheets.Add(sheetName, templateSheet);
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
        public async Task ExportReportTestResult(string outputPath, string option, ThongTinDuAn project, ThongSoDauVao tsdv, KetQuaDoKiem kqdk)
        {
            try
            {

                if (Common.IsFileLocked(outputPath))
                {
                    MessageBoxHelper.ShowWarning("File đang được mở bởi ứng dụng khác. Vui lòng tắt file trước khi xuất báo cáo!");
                    return;
                }

                string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "baocaodokiem_template.docx");
                if (!File.Exists(templatePath))
                {
                    MessageBoxHelper.ShowWarning("File mẫu báo cáo không tồn tại");
                    return;
                }

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
                                x.CongSuatDongCoTaiDieuKienDoKiem.ToString() ?? "",
                                x.HieuSuatTinh.ToString() ?? "",
                                x.HieuSuatTong.ToString() ?? "",
                                x.TanSo.ToString() ?? ""
                            )).ToList() ?? new List<BangKetQuaThuNghiem>();
                        break;
                    case "NORMALIZED":
                        input.BangKetQuaThuNghiem = kqdk.DanhSachhieuChuanVeDieuKienTieuChuan
                            ?.Select((x, idx) => new BangKetQuaThuNghiem(
                                (idx + 1).ToString(),
                                (kqdk.DanhSachketQuaTaiDieuKienDoKiem != null && kqdk.DanhSachketQuaTaiDieuKienDoKiem.Count > idx
                                    ? kqdk.DanhSachketQuaTaiDieuKienDoKiem[idx].HieuChinhLuuLuongTheTichTheoRPM.ToString()
                                    : "") ?? "",
                                x.ApSuatTinhTieuChuan.ToString() ?? "",
                                x.ApSuatTongTieuChuan.ToString() ?? "",
                                x.CongSuatHapThuTieuChuan.ToString() ?? "",
                                x.HieuSuatTinh.ToString() ?? "",
                                x.HieuSuatTong.ToString() ?? "",
                                x.TanSo.ToString() ?? ""
                            )).ToList() ?? new List<BangKetQuaThuNghiem>();
                        break;
                    case "OPERATION":
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
                                x.HieuSuatTong.ToString() ?? "",
                                x.TanSo.ToString() ?? ""
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
                    foreach (var prop in typeof(ThamSo).GetProperties())
                    {
                        data[prop.Name] = prop.GetValue(input.ThongTinDuAn.ThamSo).ToString() ?? "";
                    }
                    foreach (var prop in typeof(ThongTinChung).GetProperties())
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
                    InsertScatterChart(doc, input.BangKetQuaThuNghiem, "LuuLuong", "CongSuatTieuThu", "Lưu lượng (m3/h)", "Công suất (kW)", 1U);
                    InsertScatterChart(doc, input.BangKetQuaThuNghiem, "LuuLuong", "HieuSuatTinh", "Lưu lượng (m3/h)", "Hiệu suất tĩnh (%)", 2U);
                    InsertScatterChart(doc, input.BangKetQuaThuNghiem, "LuuLuong", "ApSuatTinh", "Lưu lượng (m3/h)", "Áp suất tĩnh (Pa)", 3U);

                    doc.MainDocumentPart.Document.Save();
                }
                return;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        private void InsertScatterChart(
            WordprocessingDocument doc,
            List<BangKetQuaThuNghiem> data,
            string xField,
            string yField,
            string xTitle,
            string yTitle,
            uint chartId
        )
        {
            try
            {
                var mainPart = doc.MainDocumentPart;
                var chartPart = mainPart.AddNewPart<ChartPart>();
                string chartPartId = mainPart.GetIdOfPart(chartPart);
                var inv = CultureInfo.InvariantCulture;

                var freqGroups = data.GroupBy(item => item.TanSo)
                             .OrderBy(g => g.Key)
                             .ToList();

                var scatterChart = new ScatterChart(
                    new ScatterStyle() { Val = ScatterStyleValues.LineMarker },
                    new VaryColors() { Val = false }
                );

                uint seriesIndex = 0;
                foreach (var group in freqGroups)
                {
                    var groupList = group.ToList();
                    var series = new ScatterChartSeries(
                                    new A.Charts.Index() { Val = seriesIndex },
                                    new Order() { Val = seriesIndex },
                                    new SeriesText(new NumericValue() { Text = group.Key.ToString() + " Hz" }),
                                    new ChartShapeProperties(new A.Outline(new A.SolidFill(new A.RgbColorModelHex() { Val = "000000" }), new A.Round())),
                                    new Marker(new Symbol() { Val = MarkerStyleValues.Circle }, new A.Charts.Size() { Val = 5 },
                                                        new ChartShapeProperties(
                                                            new A.SolidFill(new A.RgbColorModelHex() { Val = "000000" }),
                                                            new A.Outline(new A.SolidFill(new A.RgbColorModelHex() { Val = "000000" }))
                                                            )),
                                    new XValues(CreateNumLst(groupList, xField)),
                                    new YValues(CreateNumLst(groupList, yField)),
                                    new Smooth() { Val = false }
                                );

                    scatterChart.Append(series);
                    seriesIndex++;
                }

                scatterChart.Append(new AxisId() { Val = 48650112u });
                scatterChart.Append(new AxisId() { Val = 48672768u });

                double maxPosX = data.Any() ? data.Max(d => Convert.ToDouble(d.GetType().GetProperty(xField)?.GetValue(d) ?? 0)) : 0;
                double maxPosY = data.Any() ? data.Max(d => Convert.ToDouble(d.GetType().GetProperty(yField)?.GetValue(d) ?? 0)) : 0;

                var catAx = CreateAxis(48650112u, 48672768u, xTitle, 0, Common.RoundUpToNearest((float)maxPosX * 1.1f,  1000), false);
                var valAx = CreateAxis(48672768u, 48650112u, yTitle, 0, yField == "HieuSuatTinh" ? 100 : Common.RoundUpToNearest((float)maxPosY * 1.1f, yField == "CongSuatTieuThu" ? 10 : 1000), true);

                var chart = new Chart(
                    new AutoTitleDeleted() { Val = true },
                    new PlotArea(scatterChart, catAx, valAx),
                    new PlotVisibleOnly() { Val = true }
                );
                chartPart.ChartSpace = new ChartSpace(new EditingLanguage() { Val = "vi-VN" }, chart);

                var drawing = new Drawing(new Inline(
                    new Extent() { Cx = 5486400, Cy = 3200400 },
                    new DocProperties() { Id = chartId, Name = "Chart " + chartId },
                    new NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks() { NoChangeAspect = true }),
                    new A.Graphic(new A.GraphicData(new ChartReference() { Id = chartPartId }) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" })
                ));

                var p = new Paragraph(new ParagraphProperties(new Justification() { Val = JustificationValues.Center }), drawing);
                var body = mainPart.Document.Body;
                var sectPr = body.Elements<SectionProperties>().LastOrDefault();
                if (sectPr != null) body.InsertBefore(p, sectPr);
                else body.Append(p);
            }
            catch(Exception ex)
            {
                throw ex;
            }
        }

        public Task ExportReportTestResult_full(string outputPath, ThongTinDuAn project, ThongSoDauVao tsdv, KetQuaDoKiem kqdk)
        {
            try
            {
                if (Common.IsFileLocked(outputPath))
                {
                    throw new BusinessException("File đang được mở bởi ứng dụng khác. Vui lòng tắt file trước khi xuất báo cáo!");
                }
                string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "baocaodokiem_template_full.docx");
                if (!File.Exists(templatePath))
                {
                    throw new BusinessException("File mẫu báo cáo không tồn tại");
                }

                BaoCao input = new BaoCao();
                input.ThongTinDuAn = project;

                File.Copy(templatePath, outputPath, true);

                // Chuẩn bị dữ liệu cho 3 bảng
                var bangThietKe = kqdk.DanhSachketQuaTaiDieuKienDoKiem
                    ?.Select((x, idx) => new BangKetQuaThuNghiem(
                        (idx + 1).ToString(),
                        x.HieuChinhLuuLuongTheTichTheoRPM.ToString() ?? "",
                        x.ApSuatTinh.ToString() ?? "",
                        "", // Không dùng ApSuatTong cho bảng này
                        x.CongSuatDongCoTaiDieuKienDoKiem.ToString() ?? "",
                        x.HieuSuatTinh.ToString() ?? "",
                        "",
                        x.TanSo.ToString() ?? ""
                    )).ToList() ?? new List<BangKetQuaThuNghiem>();

                var bangTieuChuan = kqdk.DanhSachhieuChuanVeDieuKienTieuChuan
                    ?.Select((x, idx) => new BangKetQuaThuNghiem(
                        (idx + 1).ToString(),
                        x.LuuLuongTieuChuan_m3h.ToString() ?? "",
                        x.ApSuatTinhTieuChuan.ToString() ?? "",
                        "", // Không dùng ApSuatTong cho bảng này
                        x.CongSuatHapThuTieuChuan.ToString() ?? "",
                        x.HieuSuatTinh.ToString() ?? "",
                        "",
                        x.TanSo.ToString() ?? ""
                    )).ToList() ?? new List<BangKetQuaThuNghiem>();

                var bangLamViec = kqdk.DanhSachhieuChuanVeDieuKienLamviec
                    ?.Select((x, idx) => new BangKetQuaThuNghiem(
                        (idx + 1).ToString(),
                        x.LuuLuongLamViec_m3h.ToString() ?? "",
                        x.ApSuatTinhLamViec.ToString() ?? "",
                        "", // Không dùng ApSuatTong cho bảng này
                        x.CongSuatHapThuLamViec.ToString() ?? "",
                        x.HieuSuatTinh.ToString() ?? "",
                        "",
                        x.TanSo.ToString() ?? ""
                    )).ToList() ?? new List<BangKetQuaThuNghiem>();


                var data = new Dictionary<string, string>();
                if (input.ThongTinDuAn != null)
                {
                    foreach (var prop in typeof(ThamSo).GetProperties())
                    {
                        data[prop.Name] = prop.GetValue(input.ThongTinDuAn.ThamSo).ToString() ?? "";
                    }
                    foreach (var prop in typeof(ThongTinChung).GetProperties())
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
                    InsertScatterChart_full(doc, bangThietKe, bangLamViec, "LuuLuong", "CongSuatTieuThu", "Lưu lượng (m3/h)", "Công suất (kW)", 1U);
                    InsertScatterChart_full(doc, bangThietKe, bangLamViec, "LuuLuong", "HieuSuatTinh", "Lưu lượng (m3/h)", "Hiệu suất tĩnh (%)", 2U);
                    InsertScatterChart_full(doc, bangThietKe, bangLamViec, "LuuLuong", "ApSuatTinh", "Lưu lượng (m3/h)", "Áp suất tĩnh (Pa)", 3U);

                    doc.MainDocumentPart.Document.Save();
                }
                return Task.CompletedTask;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
       
        private void InsertScatterChart_full(
            WordprocessingDocument doc,
            List<BangKetQuaThuNghiem> data, // design condition
            List<BangKetQuaThuNghiem> opeData, // operating condition
            string xField,
            string yField,
            string xTitle,
            string yTitle,
            uint chartId
        )
        {
            try
            {
                var mainPart = doc.MainDocumentPart;
                var chartPart = mainPart.AddNewPart<ChartPart>();
                string chartPartId = mainPart.GetIdOfPart(chartPart);

                var freqGroups = data.Select(d => d.TanSo)
                                          .Union(data.Select(d => d.TanSo))
                                          .OrderBy(f => f)
                                          .ToList();

                var scatterChart = new ScatterChart(
                    new ScatterStyle() { Val = ScatterStyleValues.LineMarker },
                    new VaryColors() { Val = true }
                );
                uint seriesIndex = 0;

                foreach (var freq in freqGroups)
                {
                    var desCond = data.Where(d => d.TanSo == freq).ToList();
                    if (desCond.Any())
                    {
                        string desName = (seriesIndex == 0) ? "Design condition" : $"DES {freq}";
                        var desSeries = new ScatterChartSeries(
                            new A.Charts.Index() { Val = seriesIndex },
                            new Order() { Val = seriesIndex },
                            new SeriesText(new NumericValue() { Text = desName }),
                            new ChartShapeProperties(new A.Outline(new A.SolidFill(new A.RgbColorModelHex() { Val = "FF0000" }), new A.Round())),
                            new Marker(new Symbol() { Val = MarkerStyleValues.Circle }, new A.Charts.Size() { Val = 5 },
                                                new ChartShapeProperties(
                                                    new A.SolidFill(new A.RgbColorModelHex() { Val = "FF0000" }), 
                                                    new A.Outline(new A.SolidFill(new A.RgbColorModelHex() { Val = "FF0000" }))
                                                    )),
                            new XValues(CreateNumLst(desCond, xField)),
                            new YValues(CreateNumLst(desCond, yField)),
                            new Smooth() { Val = false }
                        );
                        scatterChart.Append(desSeries);
                        seriesIndex++;
                    }

                    var opeCond = opeData.Where(d => d.TanSo == freq).ToList();
                    if (opeCond.Any())
                    {
                        string opeName = (seriesIndex == 1) ? "Operating condition" : $"OPE {freq}";
                        var opeSeries = new ScatterChartSeries(
                            new A.Charts.Index() { Val = seriesIndex },
                            new Order() { Val = seriesIndex },
                            new SeriesText(new NumericValue() { Text = opeName }),
                            new ChartShapeProperties(new A.Outline(new A.SolidFill(new A.RgbColorModelHex() { Val = "000000" }), new A.Round())),
                            new Marker(new Symbol() { Val = MarkerStyleValues.Square }, new A.Charts.Size() { Val = 5 },
                            new ChartShapeProperties(
                                new A.SolidFill(new A.RgbColorModelHex() { Val = "000000" }),
                                new A.Outline(new A.SolidFill(new A.RgbColorModelHex() { Val = "000000" }))
                            )),
                            new XValues(CreateNumLst(opeCond, xField)),
                            new YValues(CreateNumLst(opeCond, yField)),
                            new Smooth() { Val = false }
                        );
                        scatterChart.Append(opeSeries);
                        seriesIndex++;
                    }
                }

                scatterChart.Append(new AxisId() { Val = 48650112u });
                scatterChart.Append(new AxisId() { Val = 48672768u });

                var allData = data.Concat(opeData).ToList();
                double maxPosX = allData.Any() ? allData.Max(d => Convert.ToDouble(d.GetType().GetProperty(xField)?.GetValue(d) ?? 0)) : 0;
                double maxPosY = allData.Any() ? allData.Max(d => Convert.ToDouble(d.GetType().GetProperty(yField)?.GetValue(d) ?? 0)) : 0;

                var catAx = CreateAxis(48650112u, 48672768u, xTitle, 0, Common.RoundUpToNearest((float)maxPosX * 1.1f, 1000), false);
                var valAx = CreateAxis(48672768u, 48650112u, yTitle, 0, yField == "HieuSuatTinh" ? 100 : Common.RoundUpToNearest((float)maxPosY * 1.1f, yField == "CongSuatTieuThu" ? 10 : 1000), true);

                var legend = new Legend(
                    new LegendPosition() { Val = LegendPositionValues.Bottom },
                    new Overlay() { Val = false }
                );

                for (uint i = 2; i < seriesIndex; i++)
                {
                    legend.Append(new LegendEntry(
                        new A.Charts.Index() { Val = i },
                        new Delete() { Val = true }
                    ));
                }

                var chart = new Chart(
                    new AutoTitleDeleted() { Val = true },
                    new PlotArea(scatterChart, catAx, valAx),
                    legend,
                    new PlotVisibleOnly() { Val = true }
                );

                chartPart.ChartSpace = new ChartSpace(new EditingLanguage() { Val = "vi-VN" }, chart);

                var drawing = new Drawing(new Inline(
                    new Extent() { Cx = 5486400, Cy = 3200400 },
                    new DocProperties() { Id = chartId, Name = "Chart " + chartId },
                    new NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks() { NoChangeAspect = true }),
                    new A.Graphic(new A.GraphicData(new ChartReference() { Id = chartPartId }) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" })
                ));

                var p = new Paragraph(new ParagraphProperties(new Justification() { Val = JustificationValues.Center }), drawing);

                var body = mainPart.Document.Body;
                var sectPr = body.Elements<SectionProperties>().LastOrDefault();
                if (sectPr != null) body.InsertBefore(p, sectPr);
                else body.Append(p);
            }
            catch (Exception ex) { throw ex; }
        }

        private ValueAxis CreateAxis(uint id, uint crossId, string titleText, double min, double max, bool isVertical)
        {
            var bodyProps = isVertical ? new A.BodyProperties() { Rotation = -5400000, Vertical = A.TextVerticalValues.Horizontal } : new A.BodyProperties();
            var axis = new ValueAxis(
                new AxisId() { Val = id },
                new Scaling(new A.Charts.Orientation() { Val = OrientationValues.MinMax }, new MinAxisValue() { Val = min }),
                new Delete() { Val = false },
                new AxisPosition() { Val = isVertical ? AxisPositionValues.Left : AxisPositionValues.Bottom },
                new MajorGridlines(),
                new Title(new ChartText(
                    new RichText(bodyProps, new A.ListStyle(),
                        new A.Paragraph(new A.Run(new A.RunProperties { Language = "vi-VN", FontSize = 1100, Bold = true }, new A.Text() { Text = Common.ToSuperscriptUnit(titleText) }),
                        new A.EndParagraphRunProperties() { Language = "vi-VN" })
                )),
                new Overlay() { Val = false }),
                new A.Charts.NumberingFormat() { FormatCode = "General", SourceLinked = true },
                new TickLabelPosition() { Val = TickLabelPositionValues.NextTo },
                new CrossingAxis() { Val = crossId },
                new Crosses() { Val = CrossesValues.AutoZero },
                new CrossBetween() { Val = CrossBetweenValues.Between }
            );

            if (max > 0) axis.Scaling.Append(new MaxAxisValue() { Val = max });
            return axis;
        }

        private NumberLiteral CreateNumLst(List<BangKetQuaThuNghiem> data, string field)
        {
            var inv = CultureInfo.InvariantCulture;
            var values = data.Select(d => double.TryParse(d.GetType().GetProperty(field)?.GetValue(d)?.ToString(), out var v) ? v : 0).ToList();
            var lit = new NumberLiteral(new FormatCode("General"), new PointCount() { Val = (uint)values.Count });
            for (int i = 0; i < values.Count; i++)
            {
                lit.Append(new NumericPoint() { Index = (uint)i, NumericValue = new NumericValue(values[i].ToString(inv)) });
            }
            return lit;
        }

        private Paragraph CreateCenteredParagraph(string text)
        {
            return new Paragraph(
                new ParagraphProperties(
                    new Justification() { Val = JustificationValues.Center }
                ),
                new Run(new Text(text))
            );
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

        #endregion
    }
}
