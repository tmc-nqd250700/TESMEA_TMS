namespace TESMEA_TMS.DTOs
{
    public class KetQuaDoKiem
    {
        public List<KetQuaTaiDieuKienDoKiem> DanhSachketQuaTaiDieuKienDoKiem { get; set; }
        public List<HieuChuanVeDieuKienTieuChuan> DanhSachhieuChuanVeDieuKienTieuChuan { get; set; }
        public List<HieuChuanVeDieuKienLamviec> DanhSachhieuChuanVeDieuKienLamviec { get; set; }
    }

    public class KetQuaTaiDieuKienDoKiem
    {
        public int STT { get; set; }
        // B1-B4: Nhiệt độ bầu ướt, Áp suất bảo hòa, Áp suất riêng phần Pv, KLR môi trường
        public double NhietDoBauUot { get; set; } // nhiệt độ bầu ướt
        public double ApSuatBaoHoaPsat { get; set; } // áp suất bão hòa Psat
        public double ApSuatRiengPhanPv { get; set; } // áp suất riêng phần Pv
        public double KLRMoiTruong { get; set; } // KLR môi trường (khối lượng riêng của không khí tại điều kiện đo)
        public double XacDinhRW { get; set; } // xác định RW 
        public double ApSuatTaiDiemDoChenhLechApSuatP5 { get; set; } // áp suất tại điểm đo chênh lệch áp suất P5
        public double KLRTaiDiemDoLuuLuongPL5 { get; set; } // KLR tại điểm đo lưu lượng PL5
        public double DoNhotKhongKhi { get; set; }  // độ nhớt không khí
        // B5-B11: Hệ số lưu lượng, Lưu lượng khối lượng, Lưu lượng thể tích, KLR tại điểm đo áp suất PL3, Lưu lượng thể tích tại PL3, Lưu lượng thể tích theo RPM, Hiệu chỉnh lưu lượng thể tích theo RPM

        public double HeSoLuuLuong { get; set; } // hệ số lưu lượng
        public double LuuLuongKhoiLuong { get; set; } // lưu lượng khối lượng
        public double LuuLuongTheTich { get; set; } // lưu lượng thể tích
        public double KLRTaiDiemDoApSuatPL3 { get; set; } // KLR tại điểm đo áp suất PL3
        public double LuuLuongTheTichTaiPL3 { get; set; } // lưu lượng thể tích tại PL3
        public double LuuLuongTheTichTheoRPM { get; set; } // lưu lượng thể tích theo RPM
        public double HieuChinhLuuLuongTheTichTheoRPM { get; set; } // hiệu chỉnh lưu lượng thể tích theo RPM

        // B12-B19: Vận tốc dòng khí, Áp suất động, Tổn thất đường ống, Áp suất tĩnh tổng, Công suất động cơ tại điều kiện đo, Công suất động cơ thực tế, Hiệu suất tính tổng, Hiệu suất tổng
        public double VanTocDongKhi { get; set; }
        public double ApSuatDong { get; set; }
        public double TonThatDuongOng { get; set; }
        public double ApSuatTinh { get; set; }
        public double ApSuatTong { get; set; }
        public double CongSuatDongCoTaiDieuKienDoKiem { get; set; }
        public double CongSuatDongCoThucTe { get; set; }
        public double HieuSuatTinh { get; set; }
        public double HieuSuatTong { get; set; }
    }

    public class HieuChuanVeDieuKienTieuChuan
    {
        public int STT { get; set; }
        public double LuuLuongTieuChuan_m3s { get; set; }
        public double LuuLuongTieuChuan_m3h { get; set; }
        public double ApSuatTinhTieuChuan { get; set; }
        public double ApSuatDongTieuChuan { get; set; }
        public double ApSuatTongTieuChuan { get; set; }
        public double CongSuatHapThuTieuChuan { get; set; }
        public double HieuSuatTinh { get; set; }
        public double HieuSuatTong { get; set; }
    }

    public class HieuChuanVeDieuKienLamviec
    {
        public int STT { get; set; }
        public double KLRTaiDieuKienLamViec { get; set; }
        public double LuuLuongLamViec_m3s { get; set; }
        public double LuuLuongLamViec_m3h { get; set; }
        public double ApSuatTinhLamViec { get; set; }
        public double ApSuatDongLamViec { get; set; }
        public double ApSuatTongLamViec { get; set; }
        public double CongSuatHapThuLamViec { get; set; }
        public double HieuSuatTinh { get; set; }
        public double HieuSuatTong { get; set; }
    }
}
