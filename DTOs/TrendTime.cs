namespace TESMEA_TMS.DTOs
{
    public class TrendTime
    {
        public int Index { get; set; }
        public float Time { get; set; }
        public float NhietDoMoiTruong_sen { get; set; } // nhiệt độ bầu khô
        public float DoAm_sen { get; set; }
        public float ApSuatkhiQuyen_sen { get; set; }
        public float ChenhLechApSuat_sen { get; set; }
        public float ApSuatTinh_sen { get; set; }
        public float DoRung_sen { get; set; }
        public float DoOn_sen { get; set; }
        public float SoVongQuay_sen { get; set; } // tốc độ thực của guồn cánh
        public float Momen_sen { get; set; }
        public float DongDien_fb { get; set; }
        public float CongSuat_fb { get; set; }
        public float ViTriVan_fb { get; set; }
    }
}
