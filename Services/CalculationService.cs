using TESMEA_TMS.DTOs;

namespace TESMEA_TMS.Services
{
    public interface ICalculationService
    {
        Task<KetQuaDoKiem> CalcutationTestResultAsync(ThongSoDauVao input);
    }
    public class CalculationService : ICalculationService
    {
        public async Task<KetQuaDoKiem> CalcutationTestResultAsync(ThongSoDauVao thongSo)
        {
            var kqtsDoKiem = await TinhToanTaiDieuKienDoKiem(thongSo);
            var kqhcTieuChuan = await TinhToanHieuChinhDieuKienTieuChuan(kqtsDoKiem);
            var kqhcLamViec = await TinhToanHieuChinhDieuKienLamViec(thongSo.ThongSoCoBanCuaQuat, thongSo.DanhSachThongSoDoKiem, kqtsDoKiem);
            return new KetQuaDoKiem
            {
                DanhSachketQuaTaiDieuKienDoKiem = kqtsDoKiem,
                DanhSachhieuChuanVeDieuKienTieuChuan = kqhcTieuChuan,
                DanhSachhieuChuanVeDieuKienLamviec = kqhcLamViec
            };
        }

        // Hàm tính toàn tại điều kiện đo kiểm
        private Task<List<KetQuaTaiDieuKienDoKiem>> TinhToanTaiDieuKienDoKiem(ThongSoDauVao thongSo)
        {
            try
            {
                List<KetQuaTaiDieuKienDoKiem> dstskqDoKiem = new List<KetQuaTaiDieuKienDoKiem>();
                foreach (var item in thongSo.DanhSachThongSoDoKiem)
                {
                    double apSuatTinhHieuChinhVeDieukienThietKe = item.ApSuatTinh * Math.Pow(thongSo.ThongSoCoBanCuaQuat.SoVongQuayCuaQuatNLT / item.SoVongQuayNTT, 2);

                    double nhietDoBauUot = CalcNhietDoBauUot(item.NhietDoBauKho, item.DoAmTuongDoi);
                    double apSuatBaoHoa = CalcApSuatBaoHoa(nhietDoBauUot);
                    double apSuatRiengPhanPv = CalcApSuatRiengPhanPv(thongSo.ThongSoCoBanCuaQuat.ApSuatKhiQuyen, item.NhietDoBauKho, nhietDoBauUot, apSuatBaoHoa);
                    double klrMoiTruong = CalcKLRMoiTruong(thongSo.ThongSoCoBanCuaQuat.ApSuatKhiQuyen, item.NhietDoBauKho, apSuatRiengPhanPv);
                    double xacDinhRW = CalcXacDinhRW(thongSo.ThongSoCoBanCuaQuat.ApSuatKhiQuyen, apSuatRiengPhanPv);
                    double apSuatTaiDiemDoChenhLechApSuatP5 = CalcApSuatTaiDiemDoChenhLechApSuatPL5(thongSo.ThongSoCoBanCuaQuat.ApSuatKhiQuyen, item.ChenhLechApSuat);
                    double klrTaiDiemDoLuuLuongPL5 = CalcKLRTaiDiemDoLuuLuongPL5(item.NhietDoBauKho, xacDinhRW, apSuatTaiDiemDoChenhLechApSuatP5);
                    double doNhotKhongKhi = CalcDoNhotKhongKhi(item.NhietDoBauKho);

                    // B6-B11: Hệ số lưu lượng, Lưu lượng khối lượng, Lưu lượng thể tích, KLR tại điểm đo áp suất PL3, Lưu lượng thể tích tại PL3, Lưu lượng thể tích theo RPM, Hiệu chỉnh lưu lượng thể tích theo RPM
                    double heSoLuuLuong = 1; // Assuming this is a constant or calculated elsewhere
                    double luuLuongKhoiLuong = CalcLuuLuongKhoiLuong(thongSo.ThongSoDuongOngGio.DuongKinhOngD5, item.ChenhLechApSuat, klrTaiDiemDoLuuLuongPL5, heSoLuuLuong);
                    double luuLuongTheTich = CalcLuuLuongTheTich(klrTaiDiemDoLuuLuongPL5, luuLuongKhoiLuong);
                    double klrTaiDiemDoApSuatPL3 = CalcKlrDiemDoApSuatPL3(thongSo.ThongSoCoBanCuaQuat.ApSuatKhiQuyen, item.NhietDoBauKho, item.ApSuatTinh, xacDinhRW);
                    double luuLuongTheTichTaiPL3 = CalcLuuLuongTheTichTaiPL3(luuLuongKhoiLuong, klrTaiDiemDoApSuatPL3);
                    double luuLuongTheTichTheoRPM = CalcLuuLuongTheTichTheoRPM(thongSo.ThongSoCoBanCuaQuat.SoVongQuayCuaQuatNLT, item.SoVongQuayNTT, luuLuongTheTichTaiPL3);
                    double hieuChinhLuuLuongTheTichTheoRPM = CalcHieuChinhLuuLuongTheTichTheoRPM(luuLuongTheTichTheoRPM);
                    // B12-B19: Vận tốc dòng khí, Áp suất động, Tổn thất đường ống, Áp suất tĩnh tổng, Công suất động cơ tại điều kiện đo, Công suất động cơ thực tế, Hiệu suất tính tổng, Hiệu suất tổng
                    double vanTocDongKhi = CalcVanTocDongKhi(thongSo.ThongSoDuongOngGio.TietDienOngGioD3, luuLuongTheTichTheoRPM);
                    double apSuatDong = CalcApSuatDong(klrTaiDiemDoApSuatPL3, vanTocDongKhi);
                    double tonThatDuongOng = CalcTonThatDuongOng(thongSo.ThongSoDuongOngGio.ChieuDaiOngGioTonThatL, thongSo.ThongSoDuongOngGio.DuongKinhOngGioD3, thongSo.ThongSoDuongOngGio.HeSoMaSatOngK, apSuatDong);


                    // 0808: fix thay công thức: thay áp suất tĩnh -> áp suất tĩnh hiệu chỉnh tại điều kiện đo kiểm
                    double apSuatTinh = CalcApSuatTinh(apSuatTinhHieuChinhVeDieukienThietKe, apSuatDong, tonThatDuongOng);


                    double apSuatTong = CalcApSuatTong(apSuatDong, apSuatTinh);
                    double congSuatDongCoTaiDieuKienDo = CalcCongSuatDongCoDoTaiDiemDo(thongSo.ThongSoCoBanCuaQuat.HeSoDongCo, thongSo.ThongSoCoBanCuaQuat.HieuSuatDongCo, item.DongLamViec, item.DienAp);
                    double congSuatDongCoThucTe = CalcCongSuatDongCoThucTe(thongSo.ThongSoCoBanCuaQuat.SoVongQuayCuaQuatNLT, item.SoVongQuayNTT, congSuatDongCoTaiDieuKienDo);
                    double hieuSuatTinh = CalcHieuSuatTinh(luuLuongTheTichTheoRPM, apSuatTinh, congSuatDongCoThucTe);
                    double hieuSuatTong = CalcHieuSuatTong(luuLuongTheTichTheoRPM, apSuatTong, congSuatDongCoThucTe);
                    // Assign calculated values to the result object

                    KetQuaTaiDieuKienDoKiem kqDoKiem = new KetQuaTaiDieuKienDoKiem
                    {
                        STT = item.KiemTraSo,

                        // Bước 1 - 5
                        NhietDoBauUot = nhietDoBauUot,
                        ApSuatBaoHoaPsat = apSuatBaoHoa,
                        ApSuatRiengPhanPv = apSuatRiengPhanPv,
                        KLRMoiTruong = klrMoiTruong,
                        XacDinhRW = xacDinhRW,
                        ApSuatTaiDiemDoChenhLechApSuatP5 = apSuatTaiDiemDoChenhLechApSuatP5,
                        KLRTaiDiemDoLuuLuongPL5 = klrTaiDiemDoLuuLuongPL5,
                        DoNhotKhongKhi = doNhotKhongKhi,
                        // Bước 6 - 11
                        HeSoLuuLuong = heSoLuuLuong,
                        LuuLuongKhoiLuong = luuLuongKhoiLuong,
                        LuuLuongTheTich = luuLuongTheTich,
                        KLRTaiDiemDoApSuatPL3 = klrTaiDiemDoApSuatPL3,
                        LuuLuongTheTichTaiPL3 = luuLuongTheTichTaiPL3,
                        LuuLuongTheTichTheoRPM = luuLuongTheTichTheoRPM,
                        HieuChinhLuuLuongTheTichTheoRPM = hieuChinhLuuLuongTheTichTheoRPM,
                        // Bước 12 - 19
                        VanTocDongKhi = vanTocDongKhi,
                        ApSuatDong = apSuatDong,
                        TonThatDuongOng = tonThatDuongOng,
                        ApSuatTinh = apSuatTinh,
                        ApSuatTong = apSuatTong,
                        CongSuatDongCoTaiDieuKienDoKiem = congSuatDongCoTaiDieuKienDo,
                        CongSuatDongCoThucTe = congSuatDongCoThucTe,
                        HieuSuatTinh = hieuSuatTinh,
                        HieuSuatTong = hieuSuatTong

                    };
                    dstskqDoKiem.Add(kqDoKiem);
                }
                return Task.FromResult(dstskqDoKiem);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        // Hàm tính toán hiệu chỉnh ở điều kiện tiêu chuẩn
        private Task<List<HieuChuanVeDieuKienTieuChuan>> TinhToanHieuChinhDieuKienTieuChuan(List<KetQuaTaiDieuKienDoKiem> dstskqDoKiem)
        {
            try
            {
                List<HieuChuanVeDieuKienTieuChuan> dskqTieuChuan = new List<HieuChuanVeDieuKienTieuChuan>();
                foreach (var item in dstskqDoKiem)
                {
                    double luuLuong_m3s = item.LuuLuongTheTichTheoRPM;
                    double apSuatTinh = item.ApSuatTinh * (1.204 / item.KLRTaiDiemDoApSuatPL3);
                    double apSuatDong = item.ApSuatDong;
                    double apSuatTong = apSuatTinh + apSuatDong;
                    double congSuatHapThu = item.CongSuatDongCoThucTe;
                    HieuChuanVeDieuKienTieuChuan result = new HieuChuanVeDieuKienTieuChuan
                    {
                        STT = item.STT,
                        LuuLuongTieuChuan_m3s = luuLuong_m3s,
                        LuuLuongTieuChuan_m3h = item.HieuChinhLuuLuongTheTichTheoRPM,
                        ApSuatTinhTieuChuan = apSuatTinh,
                        ApSuatDongTieuChuan = apSuatDong,
                        ApSuatTongTieuChuan = apSuatTong,
                        CongSuatHapThuTieuChuan = congSuatHapThu,
                        HieuSuatTinh = luuLuong_m3s * apSuatTinh * 100 / (1000 * congSuatHapThu),
                        HieuSuatTong = luuLuong_m3s * apSuatTong * 100 / (1000 * congSuatHapThu)
                    };
                    dskqTieuChuan.Add(result);
                }
                return Task.FromResult(dskqTieuChuan);
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }


        // Hàm tính toán hiệu chỉnh ở điều kiện làm việc
        private Task<List<HieuChuanVeDieuKienLamviec>> TinhToanHieuChinhDieuKienLamViec(ThongSoCoBanCuaQuat tsCoBanCuaQuat, List<ThongSoDoKiem> dsThongSoDoKiem, List<KetQuaTaiDieuKienDoKiem> dstskqDoKiem)
        {
            try
            {
                List<HieuChuanVeDieuKienLamviec> dskqLamViec = new List<HieuChuanVeDieuKienLamviec>();
                foreach (var item in dstskqDoKiem)
                {
                    double klr = item.KLRTaiDiemDoApSuatPL3 * (273 + dsThongSoDoKiem[item.STT - 1].NhietDoBauKho) / (273 + tsCoBanCuaQuat.NhietDoLamViec);
                    double luuLuong_m3s = item.LuuLuongTheTichTheoRPM;
                    double apSuatTinh = item.ApSuatTinh * (klr / item.KLRTaiDiemDoApSuatPL3);
                    double apSuatDong = item.ApSuatDong;
                    double apSuatTong = apSuatTinh + apSuatDong;
                    double congSuatHapThu = item.CongSuatDongCoThucTe;
                    HieuChuanVeDieuKienLamviec result = new HieuChuanVeDieuKienLamviec
                    {
                        STT = item.STT,
                        KLRTaiDieuKienLamViec = klr,
                        LuuLuongLamViec_m3s = luuLuong_m3s,
                        LuuLuongLamViec_m3h = luuLuong_m3s * 3600,
                        ApSuatTinhLamViec = apSuatTinh,
                        ApSuatDongLamViec = apSuatDong,
                        ApSuatTongLamViec = apSuatTong,
                        CongSuatHapThuLamViec = congSuatHapThu,
                        HieuSuatTinh = luuLuong_m3s * apSuatTinh * 100 / (1000 * congSuatHapThu),
                        HieuSuatTong = luuLuong_m3s * apSuatTong * 100 / (1000 * congSuatHapThu)
                    };
                    dskqLamViec.Add(result);
                }
                return Task.FromResult(dskqLamViec);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        #region Công thức tính toán chi tiết
        double piValue = 3.14;
        private double CalcNhietDoBauUot(double nhietDoBauKho, double doAmTuongDoi)
        {
            return nhietDoBauKho * Math.Atan(0.151977 * Math.Pow(doAmTuongDoi + 8.313659, 0.5))
                + Math.Atan(nhietDoBauKho + doAmTuongDoi)
                - Math.Atan(doAmTuongDoi - 1.676331)
                + 0.00391838 * Math.Pow(doAmTuongDoi, 1.5) * Math.Atan(0.023101 * doAmTuongDoi)
                - 4.686035;
        }

        private double CalcApSuatBaoHoa(double nhietDoBauUot)
        {
            return 610.8
                + 44.442 * nhietDoBauUot
                + 1.4133 * Math.Pow(nhietDoBauUot, 2)
                + 0.02768 * Math.Pow(nhietDoBauUot, 3)
                + 2.55667e-4 * Math.Pow(nhietDoBauUot, 4)
                + 2.89166e-6 * Math.Pow(nhietDoBauUot, 5);
        }

        private double CalcApSuatRiengPhanPv(double apSuatKhiQuyen, double nhietDoBauKho, double nhietDoBauUot, double apSuatBaoHoa)
        {
            return apSuatBaoHoa - apSuatKhiQuyen * 1000 * 6.66e-4 * (nhietDoBauKho - nhietDoBauUot);
        }

        private double CalcKLRMoiTruong(double apSuatKhiQuyen, double nhietDoBauKho, double apSuatRiengPhanPv)
        {
            return 3.484 * (apSuatKhiQuyen * 1000 - 0.378 * apSuatRiengPhanPv) / 1000 / (nhietDoBauKho + 273);
        }

        private double CalcXacDinhRW(double apSuatKhiQuyen, double apSuatRiengPhanPv)
        {
            return 287 / (1 - 0.378 * (apSuatRiengPhanPv / 1000 / apSuatKhiQuyen));
        }

        private double CalcApSuatTaiDiemDoChenhLechApSuatPL5(double apSuatKhiQuyen, double chenhLechApSuat)
        {
            return apSuatKhiQuyen * 1000 - chenhLechApSuat;
        }

        private double CalcKLRTaiDiemDoLuuLuongPL5(double nhietDoBauKho, double xacDinhRW, double apSuatTaiDiemDoChenhLechApSuatPL5)
        {
            return apSuatTaiDiemDoChenhLechApSuatPL5 / (xacDinhRW * (nhietDoBauKho + 273.15));
        }

        private double CalcDoNhotKhongKhi(double nhietDoBauKho)
        {
            return (17.23 + 0.048 * nhietDoBauKho) * 1e-6;
        }

        private double CalcLuuLuongKhoiLuong(double duongKinhOngD5, double chenhLechApSuat, double klrTaiDiemDoPL5, double hsLuuLuong)
        {
            return hsLuuLuong * (piValue * Math.Pow(duongKinhOngD5 / 1000, 2) / 4) * Math.Sqrt(2 * klrTaiDiemDoPL5 * chenhLechApSuat);
        }
        private double CalcLuuLuongTheTich(double klrTaiDiemDoPL5, double luuLuongKhoiLuong)
        {
            return luuLuongKhoiLuong / klrTaiDiemDoPL5;
        }

        private double CalcKlrDiemDoApSuatPL3(double apSuatKhiQuyen, double nhietDoBauKho, double thongSoApSuatTinh, double xacDinhRW)
        {
            return (apSuatKhiQuyen * 1000 - thongSoApSuatTinh) / (xacDinhRW * (nhietDoBauKho + 273.15));
        }
        private double CalcLuuLuongTheTichTaiPL3(double luuLuongKhoiLuong, double klrTaiDiemDoPL3)
        {
            return luuLuongKhoiLuong / klrTaiDiemDoPL3;
        }
        private double CalcLuuLuongTheTichTheoRPM(double soVongQuayCuaQuat, double soVongQuay, double luuLuongTheTichTaiPL3)
        {
            return luuLuongTheTichTaiPL3 * soVongQuayCuaQuat / soVongQuay;
        }
        private double CalcHieuChinhLuuLuongTheTichTheoRPM(double luuLuongTheTichTheoRPM)
        {
            return luuLuongTheTichTheoRPM * 3600;
        }
        private double CalcVanTocDongKhi(double tietDienOngGio, double luuLuongTheTichTheoRPM)
        {
            return luuLuongTheTichTheoRPM / tietDienOngGio;
        }
        private double CalcApSuatDong(double klrTaiDiemDoPL3, double vanTocDongKhi)
        {
            return 0.5 * Math.Pow(vanTocDongKhi, 2) * klrTaiDiemDoPL3;
        }
        private double CalcTonThatDuongOng(double chieuDaiOngGioTonThat, double duongKinhOngGioD3, double heSoMaSatOngK, double apSuatDong)
        {
            return heSoMaSatOngK * chieuDaiOngGioTonThat / duongKinhOngGioD3 * apSuatDong;
        }
        private double CalcApSuatTinh(double thongSoApSuatTinh, double apSuatDong, double tonThatDuongOng)
        {
            return thongSoApSuatTinh - apSuatDong + tonThatDuongOng;
        }
        private double CalcApSuatTong(double apSuatDong, double apSuatTinh)
        {
            return apSuatDong + apSuatTinh;
        }
        private double CalcCongSuatDongCoDoTaiDiemDo(double heSoDongCo, double hieuSuatDongCo, double dongLamViec, double dienAp)
        {
            return Math.Sqrt(3) * dongLamViec * dienAp * heSoDongCo / 1000 * hieuSuatDongCo / 100;
        }
        private double CalcCongSuatDongCoThucTe(double soVongQuayCuaQuat, double soVongQuay, double congSuatDongCoTaiDiemDo)
        {
            return congSuatDongCoTaiDiemDo * Math.Pow(soVongQuayCuaQuat / soVongQuay, 3);
        }
        private double CalcHieuSuatTinh(double luuLuongTheTichTheoRPM, double apSuatTinh, double congSuatDongCoThucTe)
        {
            return luuLuongTheTichTheoRPM * apSuatTinh * 100 / (1000 * congSuatDongCoThucTe);
        }
        private double CalcHieuSuatTong(double luuLuongTheTichTheoRPM, double apSuatTong, double congSuatDongCoThucTe)
        {
            return apSuatTong * luuLuongTheTichTheoRPM * 100 / (1000 * congSuatDongCoThucTe);
        }

        private double CalcTietDienOngD5(double duongKinhOngD5)
        {
            return piValue * Math.Pow(duongKinhOngD5 / 1000, 2) / 4;
        }

        private double CalcTietDienOngGioD3(double duongKinhOngGioD3)
        {
            return piValue * Math.Pow(duongKinhOngGioD3 / 1000, 2) / 4;
        }
        #endregion
    }
}
