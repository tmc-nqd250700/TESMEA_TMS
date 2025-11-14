using TESMEA_TMS.Configs;

namespace TESMEA_TMS.DTOs
{
    public class ThongTinDuAn
    {
        public ThamSo ThamSo { get; set; }
        public ThongTinChung ThongTinChung { get; set; }
        public ThongTinMauThuNghiem ThongTinMauThuNghiem { get; set; }

        public ThongTinDuAn() { }
        public ThongTinDuAn(ThamSo thamSo, ThongTinChung thongTinChung, ThongTinMauThuNghiem thongTinMauThuNghiem)
        {
            ThamSo = thamSo;
            ThongTinChung = thongTinChung;
            ThongTinMauThuNghiem = thongTinMauThuNghiem;
        }
    }

    public class ThamSo
    {
        public string TenDuAn { get; set; }
        //public string DuLieuKiemThu { get; set; }
        public string KichBan { get; set; }
        public string KieuKiemThu { get; set; }
        public string DuongDanLuuDuAn { get; set; }
    }

    // Thông tin của khách hàng
    public class ThongTinChung
    {
        public string TenMauThu { get; set; }
        public string CoSoSanXuat { get; set; }
        public string KyHieu { get; set; }
        public int SoLuongMau { get; set; } = 1;
        public string TinhTrangMau { get; set; } = UserSetting.Instance.Language == "en" ? "New" : "Mới";
        public DateTime NgayNhanYeuCau { get; set; } = DateTime.Now;
        public DateTime NgayNhanMau { get; set; } = DateTime.Now;
        public DateTime NgayThuNghiem { get; set; } = DateTime.Now;
        public DateTime NgayHoanThanh { get; set; } = DateTime.Now;
        public string TieuChuanApDung { get; set; } = "ISO 5801:2017 - TCVN 9439:2013";
    }

    // Thông tin mẫu thử nghiệm
    public class ThongTinMauThuNghiem
    {
        public float LuuLuongKhiThietKe { get; set; }
        public float ApSuatThietKe { get; set; }
        public float NhietDoThietKeLamViec { get; set; } = 20;
        public float CongSuatDongCo { get; set; } = 11;
        public float TocDoThietKeCuaQuat { get; set; } = 2950;
        public float TanSoDongCoTheoThietKe { get; set; }
        public float HeSoCongSuatDongCo { get; set; } = 0.88f; // cosphi
        public float HieuSuatDongCo { get; set; } = 100;
        public float DongDienDinhMucCuaDongCo { get; set; } = 21.5f;
        public float DienApDongCo { get; set; } = 390;

        // add
        public float TocDoDongCo { get; set; } // tốc độ động cơ
        public string HangDongCo { get; set; } = "HEM"; // hãng động cơ
    }
}
