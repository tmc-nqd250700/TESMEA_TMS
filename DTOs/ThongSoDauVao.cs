namespace TESMEA_TMS.DTOs
{
    // Tổng hợp các thông số đầu vào đo kiểm
    public class ThongSoDauVao
    {
        public ThongSoDuongOngGio ThongSoDuongOngGio { get; set; }
        public ThongSoCoBanCuaQuat ThongSoCoBanCuaQuat { get; set; }
        public List<ThongSoDoKiem> DanhSachThongSoDoKiem { get; set; }
    }
    public class ThongSoDuongOngGio
    {
        public double DuongKinhOngD5 { get; set; }
        public double ChieuDaiOngGioTonThatL { get; set; }
        public double DuongKinhOngGioD3 { get; set; }
        public double TietDienOngD5 { get; set; }
        public double HeSoMaSatOngK { get; set; }
        public double TietDienOngGioD3 { get; set; }


    }
    public class ThongSoCoBanCuaQuat
    {
        public double SoVongQuayCuaQuatNLT { get; set; }
        public double CongSuatDongCo { get; set; }
        public double HeSoDongCo { get; set; }
        public double Tanso { get; set; }
        public double HieuSuatDongCo { get; set; }
        public double DoNhotKhongKhi { get; set; }
        public double ApSuatKhiQuyen { get; set; }
        public double NhietDoLamViec { get; set; }
    }

    public class ThongSoDoKiem
    {
        public int KiemTraSo { get; set; }
        public double NhietDoBauKho { get; set; }
        public double DoAmTuongDoi { get; set; }
        public double SoVongQuayNTT { get; set; }
        public double ChenhLechApSuat { get; set; }
        public double ApSuatTinh { get; set; }
        public double DongLamViec { get; set; }
        public double DienAp { get; set; }
        public double TanSo { get; set; }
    }
}
