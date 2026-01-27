using System.ComponentModel.DataAnnotations;

namespace TESMEA_TMS.Models.Entities
{
    // Thông số nhập
    public class Library : BaseEntity
    {
        [Key]
        public Guid LibId { get; set; }
        public string LibName { get; set; }
    }

    // Biến tần
    public class BienTan
    {
        [Key]
        public Guid LibId { get; set; }
        public float DienApVao { get; set; } // Điện áp vào V
        public float DongDienVao { get; set; } // Dòng điện vào A
        public float TanSoVao { get; set; } // Tần số vào Hz
        public float CongSuatVao { get; set; } // Công suất vào kVA
        public float DienApRa { get; set; } // Điện áp ra V
        public float DongDienRa { get; set; } // Dòng điện ra A
        public float TanSoRa { get; set; } // Tần số ra Hz
        public float CongSuatTongRa { get; set; } // Công suất tổng ra kVA
        public float CongSuatHieuDungRa { get; set; } // Công suất tiêu dụng ra kW

        // thêm
        public float HieuSuatBoTruyen { get; set; }
        public float HieuSuatNoiTruc { get; set; }
        public float HieuSuatGoiTruc { get; set; }
    }


    // Dải đo cảm biến (đơn vị đo - SI)
    public class CamBien
    {
        [Key]
        public Guid LibId { get; set; }

        // Nhiệt độ môi trường C
        public bool IsImportNhietDoMoiTruong { get; set; }
        public float NhietDoMoiTruongValue { get; set; }
        public float NhietDoMoiTruongMin { get; set; }
        public float NhietDoMoiTruongMax { get; set; }

        // Độ ẩm môi trường %
        public bool IsImportDoAmMoiTruong { get; set; }
        public float DoAmMoiTruongValue { get; set; }
        public float DoAmMoiTruongMin { get; set; }
        public float DoAmMoiTruongMax { get; set; }

        // Áp suất khí quyển Pa
        public bool IsImportApSuatKhiQuyen { get; set; }
        public float ApSuatKhiQuyenValue { get; set; }
        public float ApSuatKhiQuyenMin { get; set; }
        public float ApSuatKhiQuyenMax { get; set; }

        // Áp suất chênh lệch Pa
        public bool IsImportChenhLechApSuat { get; set; }
        public float ChenhLechApSuatValue { get; set; }
        public float ChenhLechApSuatMin { get; set; }
        public float ChenhLechApSuatMax { get; set; }

        // Áp suất tĩnh Pa
        public bool IsImportApSuatTinh { get; set; }
        public float ApSuatTinhValue { get; set; }
        public float ApSuatTinhMin { get; set; }
        public float ApSuatTinhMax { get; set; }

        // Độ rung
        public bool IsImportDoRung { get; set; }
        public float DoRungValue { get; set; }
        public float DoRungMin { get; set; }
        public float DoRungMax { get; set; }

        // Độ ồn
        public bool IsImportDoOn { get; set; }
        public float DoOnValue { get; set; }
        public float DoOnMin { get; set; }
        public float DoOnMax { get; set; }

        // Số vòng quay
        public bool IsImportSoVongQuay { get; set; }
        public float SoVongQuayValue { get; set; }
        public float SoVongQuayMin { get; set; }
        public float SoVongQuayMax { get; set; }

        // Momen xoắn
        public bool IsImportMomen { get; set; }
        public float MomenValue { get; set; }
        public float MomenMin { get; set; }
        public float MomenMax { get; set; }

        // Phản hồi dòng điện
        public bool IsImportPhanHoiDongDien { get; set; }
        public float PhanHoiDongDienValue { get; set; }
        public float PhanHoiDongDienMin { get; set; }
        public float PhanHoiDongDienMax { get; set; }

        // Phản hồi công suất
        public bool IsImportPhanHoiCongSuat { get; set; }
        public float PhanHoiCongSuatValue { get; set; }
        public float PhanHoiCongSuatMin { get; set; }
        public float PhanHoiCongSuatMax { get; set; }

        // Phản hồi vị trí van
        public bool IsImportPhanHoiViTriVan { get; set; }
        public float PhanHoiViTriVanValue { get; set; }
        public float PhanHoiViTriVanMin { get; set; }
        public float PhanHoiViTriVanMax { get; set; }

        // Phản hồi điện áp
        public bool IsImportPhanHoiDienAp { get; set; }
        public float PhanHoiDienApValue { get; set; }
        public float PhanHoiDienApMin { get; set; }
        public float PhanHoiDienApMax { get; set; }

        // Nhiệt độ gối trục
        public bool IsImportNhietDoGoiTruc { get; set; }
        public float NhietDoGoiTrucValue { get; set; }
        public float NhietDoGoiTrucMin { get; set; }
        public float NhietDoGoiTrucMax { get; set; }

        // Nhiệt độ bầu khô
        public bool IsImportNhietDoBauKho { get; set; }
        public float NhietDoBauKhoValue { get; set; }
        public float NhietDoBauKhoMin { get; set; }
        public float NhietDoBauKhoMax { get; set; }

        // Cảm biến lưu lượng
        public bool IsImportCamBienLuuLuong { get; set; }
        public float CamBienLuuLuongValue { get; set; }
        public float CamBienLuuLuongMin { get; set; }
        public float CamBienLuuLuongMax { get; set; }

        // Phản hồi tần số
        public bool IsImportPhanHoiTanSo { get; set; }
        public float PhanHoiTanSoValue { get; set; }
        public float PhanHoiTanSoMin { get; set; }
        public float PhanHoiTanSoMax { get; set; }
    }

    //  Bảng thông số ống gió và van điều khiển
    public class OngGio
    {
        [Key]
        public Guid LibId { get; set; }

        public float DuongKinhOngD5 { get; set; } // Đường kính ống D5
        public float ChieuDaiConQuat { get; set; } // Chiều dài ống gió tổn thất L
        public float DuongKinhOngD3 { get; set; } // Đường kính ống D3
        public float DuongKinhLoPhut { get; set; } // Đường kính lỗ phụt
        public float HeSoMaSat { get; set; } = 0.025f; // Hệ số ma sát ống K
        public float DuongKinhMiengQuat { get; set; }
        public float TietDienOngD5 { get; set; } // Tiết diện ống D5
        public float TietDienOngD3 { get; set; } // Tiết diện ống gió D3


        //public float DuongKinhOngGio { get; set; } // đường kính ống gió
        //public float ChieuDaiOngGioSauQuat { get; set; } // %
        //public float ChieuDaiOngGioTruocQuat { get; set; } // %
        
    }
}
