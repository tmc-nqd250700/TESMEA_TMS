namespace TESMEA_TMS.DTOs
{
    #region báo cáo từ tool tính toán cho tesmea
    public class BaoCao
    {
        public ThongTinDuAn ThongTinDuAn { get; set; }
        public List<BangKetQuaThuNghiem> BangKetQuaThuNghiem { get; set; }
    }

    public class BangKetQuaThuNghiem
    {
        public string STT { get; set; }
        public string LuuLuong { get; set; }
        public string ApSuatTinh { get; set; }
        public string ApSuatTong { get; set; }
        public string CongSuatTieuThu { get; set; }
        public string HieuSuatTinh { get; set; }
        public string HieuSuatTong { get; set; }

        public BangKetQuaThuNghiem(string stt, string luuLuong, string apSuatTinh, string apSuatTong, string congSuatTieuThu, string hieuSuatTinh, string hieuSuatTong)
        {
            STT = stt;
            LuuLuong = luuLuong;
            ApSuatTinh = apSuatTinh;
            ApSuatTong = apSuatTong;
            CongSuatTieuThu = congSuatTieuThu;
            HieuSuatTinh = hieuSuatTinh;
            HieuSuatTong = hieuSuatTong;
        }
    }
    #endregion


    #region báo cáo từ phần mềm đo kiểm cũ
    //public class BaoCao_scada()
    //{
    //    // Managerment Information Input
    //    public static string ReportNo;                      // Số phiếu đo
    //    public static string Customer;                      // Khách hàng
    //    public static string Project;                       // Dự án 
    //    public static string TestPerson;                    // Người thực hiện
    //    public static string TestWitness;                   // Người chứng kiến
    //    public static string Approved;                      // Người phê duyệt
    //    public static int FanType;                          // Loại quạt 1.Ly tâm, 2. Hướng trục
    //    public static string Model;                         // Model quạt
    //    public static string SerialNo;                      // Serial No
    //    public static int Methods;                          // Phương pháp đo 1.A, 2.B, 3.C, 4.D
    //    public static string MotorName;                     // Tên động cơ
    //    public static int rpNo1;                            // Số thứ tự phiếu đo quạt ly tâm
    //    public static int rpNo2;                            // Số thứ tự phiếu đo quạt hướng trục


    //    // Motor and Transmision Parameter Input - Đầu vào tham số động cơ và truyền động
    //    public static float e_motor;           // Hiệu suất động cơ
    //    public static float e_noitruc;         // Hiệu suất nối trục
    //    public static float e_goitruc;         // Hiệu suất gối trục
    //    public static float e_botruyen;        // Hiệu suất bộ truyền
    //    public static float e_Total;           // Hiệu suất tổng
    //    public static float n1;                // Tốc độ guồng cánh nhập vào
    //    public static string MotorPower;       // Công suất động cơ nhập vào
    //    public static float MotorSpeed;        // Tốc độ động cơ nhập vào
    //    public static float Udm;               // Điện áp định mức nhập vào
    //    public static float Idm;               // Dòng điện định mức nhập vào
    //    public static float CosPhi = 0.88f;    // Cosphi động cơ
    //    public static float Vdm;               // lưu lượng định mức 
    //    public static float Pdm;               // áp suất định mức 

    //    // Geometry parameters Input - thông số ống gió và van điều khiển
    //    public static float D5;                // Đường kính ống đo kiểm
    //    public static float d;                 // Đường kính miệng quạt
    //    public static float L34;               // Chiều dài côn quạt

    //    // Working condition - điều kiện làm việc
    //    public static float Tw;                // Nhiệt độ làm việc
    //    public static float rhow;              // Tỷ trọng kk làm việc

    //    // Environment Parameter Input or Sensor - Đầu vào tham số môi trường hoặc cảm biến
    //    public static float Ta;                // Nhiệt độ môi trường
    //    public static float Pa;                // Áp suất khí quyển khu vực đo
    //    public static float hu;                // Độ ẩm không khí

    //    // Technical Parameter of Fan Input or Sensor: Thông số kỹ thuật của đầu vào quạt hoặc cảm biến
    //    public static float deltap;            // Chênh lệch áp suất điểm đo lưu lượng
    //    public static float Pe3;               // Chênh lệch áp suất điểm đo áp suất
    //    public static float Td;                // Nhiệt độ bầu khô
    //    public static float Pw;                // Công suất tiêu thụ của động cơ
    //    public static float Noise;             // Đồ ồn
    //    public static float n2;                // Tốc độ thực của guồng cánh
    //    public static float BearingVia;        // Độ rung gối đứng
    //    public static float BearingVia_H;      // Độ rung gối ngang
    //    public static float BearingVia_V;      // Độ rung gối dọc
    //    public static float BearingTemp;       // Nhiệt đô gối đứng
    //    public static float BearingTemp_H;     // Nhiệt đô gối ngang
    //    public static float BearingTemp_V;     // Nhiệt đô gối dọc
    //    public static float T;                 // Momen xoắn trên trục quạt

    //    // Inverter Feedback Parameter (Modbus) - Thông số phản hồi biến tần
    //    public static float Power_fb;          // Công suất
    //    public static float Voltage_fb;        // Điện áp
    //    public static float Current_fb;        // Dòng điện
    //    public static float Freq_fb;           // Tần số
    //    public static float Speed_fb;          // Tốc độ
    //    public static short Status_fb;         // Trạng thái
    //}
    #endregion
}
