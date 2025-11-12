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

        // nhiệt độ môi trường C
        public float NhietDoMoiTruongMin { get; set; }
        public float NhietDoMoiTruongMax { get; set; }

        // độ ẩm môi trường %
        public float DoAmMoiTruongMin { get; set; }
        public float DoAmMoiTruongMax { get; set; }

        // áp suất khí quyển Pa
        public float ApSuatKhiQuyenMin { get; set; }
        public float ApSuatKhiQuyenMax { get; set; }

        //  áp suất chênh lệch Pa
        public float ChenhLechApSuatMin { get; set; }
        public float ChenhLechApSuatMax { get; set; }

        // áp suất tĩnh Pa
        public float ApSuatTinhMin { get; set; }
        public float ApSuatTinhMax { get; set; }

        // độ rung
        public float DoRungMin { get; set; }
        public float DoRungMax { get; set; }

        // độ ồn
        public float DoOnMin { get; set; }
        public float DoOnMax { get; set; }

        // số vòng quay
        public float SoVongQuayMin { get; set; }
        public float SoVongQuayMax { get; set; }

        // momen xoắn
        public float MomenMin { get; set; }
        public float MomenMax { get; set; }
        // phản hồi dòng điện
        public float PhanHoiDongDienMin { get; set; }
        public float PhanHoiDongDienMax { get; set; }

        // phản hồi công suất
        public float PhanHoiCongSuatMin { get; set; }
        public float PhanHoiCongSuatMax { get; set; }

        // phản hồi vị trí van
        public float PhanHoiViTriVanMin { get; set; }
        public float PhanHoiViTriVanMax { get; set; }

        // phản hồi điện áp
        public float PhanHoiDienApMin { get; set; }
        public float PhanHoiDienApMax { get; set; }

        // Nhiệt độ gối trục
        public float NhietDoGoiTrucMin { get; set; }
        public float NhietDoGoiTrucMax { get; set; }

        // Nhiệt độ bầu khô
        public float NhietDoBauKhoMin { get; set; }
        public float NhietDoBauKhoMax { get; set; }

        // Cảm biến lưu lượng
        public float CamBienLuuLuongMin { get; set; }
        public float CamBienLuuLuongMax { get; set; }

        public float PhanHoiTanSoMin { get; set; }
        public float PhanHoiTanSoMax { get; set; }
    }

    //  Bảng thông số ống gió và van điều khiển
    public class OngGio
    {
        [Key]
        public Guid LibId { get; set; }
        public float DuongKinhOngGio { get; set; } // đường kính ống gió
        public float ChieuDaiOngGioSauQuat { get; set; } // %
        public float ChieuDaiOngGioTruocQuat { get; set; } // %
        public float DuongKinhLoPhut { get; set; } // Đường kính lỗ phụt

        public float DuongKinhMiengQuat { get; set; }
        public float ChieuDaiConQuat { get; set; }
    }
}
