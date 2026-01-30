using MathNet.Numerics;
using System.IO;
using System.Text.Json;
using TESMEA_TMS.Configs;
using TESMEA_TMS.DTOs;
using TESMEA_TMS.Models.Entities;

namespace TESMEA_TMS.Helpers
{
    public static class DataProcess
    {
        #region
        //public static int ImportCheck = 0;
        //public static int Done = 0;

        // Managerment Information Input
        //public static string ReportNo;                      // Số phiếu đo
        //public static string Customer;                      // Khách hàng
        //public static string Project;                       // Dự án 
        //public static string TestPerson;                    // Người thực hiện
        //public static string TestWitness;                   // Người chứng kiến
        //public static string Approved;                      // Người phê duyệt
        //public static int FanType;                          // Loại quạt 1.Ly tâm, 2. Hướng trục
        //public static string Model;                         // Model quạt
        //public static string SerialNo;                      // Serial No
        //public static int Methods;                          // Phương pháp đo 1.A, 2.B, 3.C, 4.D
        //public static string MotorName;                     // Tên động cơ
        //public static int rpNo1;                            // Số thứ tự phiếu đo quạt ly tâm
        //public static int rpNo2;                            // Số thứ tự phiếu đo quạt hướng trục


        // Motor and Transmision Parameter Input - Đầu vào tham số động cơ và truyền động
        //public static float e_motor;           // Hiệu suất động cơ
        //public static float e_noitruc;         // Hiệu suất nối trục
        //public static float e_goitruc;         // Hiệu suất gối trục
        //public static float e_botruyen;        // Hiệu suất bộ truyền
        //public static float e_Total;           // Hiệu suất tổng
        //public static float n1;                // Tốc độ guồng cánh nhập vào
        //public static string MotorPower;       // Công suất động cơ nhập vào
        //public static float MotorSpeed;        // Tốc độ động cơ nhập vào
        //public static float Udm;               // Điện áp định mức nhập vào
        //public static float Idm;               // Dòng điện định mức nhập vào
        //public static float CosPhi = 0.88f;    // Cosphi động cơ
        //public static float Vdm;               // lưu lượng định mức 
        //public static float Pdm;               // áp suất định mức 

        // Geometry parameters Input - thông số ống gió và van điều khiển
        //public static float D5;                // Đường kính ống đo kiểm
        //public static float d;                 // Đường kính miệng quạt
        //public static float L34;               // Chiều dài côn quạt

        //// Working condition - điều kiện làm việc
        //public static float Tw;                // Nhiệt độ làm việc
        //public static float rhow;              // Tỷ trọng kk làm việc

        //// Environment Parameter Input or Sensor - Đầu vào tham số môi trường hoặc cảm biến
        //public static float Ta;                // Nhiệt độ môi trường
        //public static float Pa;                // Áp suất khí quyển khu vực đo
        //public static float hu;                // Độ ẩm không khí

        // Technical Parameter of Fan Input or Sensor: Thông số kỹ thuật của đầu vào quạt hoặc cảm biến
        //public static float deltap;            // Chênh lệch áp suất điểm đo lưu lượng
        //public static float Pe3;               // Chênh lệch áp suất điểm đo áp suất
        //public static float Td;                // Nhiệt độ bầu khô
        //public static float Pw;                // Công suất tiêu thụ của động cơ
        //public static float Noise;             // Đồ ồn
        //public static float n2;                // Tốc độ thực của guồng cánh
        //public static float BearingVia;        // Độ rung gối đứng
        //public static float BearingVia_H;      // Độ rung gối ngang
        //public static float BearingVia_V;      // Độ rung gối dọc
        //public static float BearingTemp;       // Nhiệt đô gối đứng
        //public static float BearingTemp_H;     // Nhiệt đô gối ngang
        //public static float BearingTemp_V;     // Nhiệt đô gối dọc
        //public static float T;                 // Momen xoắn trên trục quạt

        //// Inverter Feedback Parameter (Modbus) - Thông số phản hồi biến tần
        //public static float Power_fb;          // Công suất
        //public static float Voltage_fb;        // Điện áp
        //public static float Current_fb;        // Dòng điện
        //public static float Freq_fb;           // Tần số
        //public static float Speed_fb;          // Tốc độ
        //public static short Status_fb;         // Trạng thái
        #endregion


        // Caculate Constant & Variable
        public static float rhokk = 1.204f;    // Tỷ trọng không khí điều kiện tiêu chuẩn

        //// Parameter Caculate for Show - thông số tính toán để hiển thị
        //public static float Freq_show;         // Tần số 
        //public static float Current_show;      // Dòng điện
        //public static float Pw_show;           // Công suất
        //public static float Speed_show;        // Tốc độ
        //public static float TempB_show;        // Nhiệt độ gối
        //public static float T_Show;            // Momen xoắn trên trục quạt
        //public static float Ps_show;           // Áp suất tĩnh
        //public static float Pt_show;           // Áp suất tổng
        //public static float Flow_show;         // Lưu lượng
        //public static float Ta_show;           // Nhiệt độ T3
        //public static float Td_show;           // Nhiệt độ bầu khô
        //public static float ViaB_show;         // Nhiệt độ bầu khô
        //public static float Prt_show;          // Công suất quạt tính bằng momen xoắn

        /* DataGridView - Report - DataBase */

        // Measurement parameter
        public static float[] TaPoint;        // Nhiệt độ môi trường
        public static float[] PaPoint;        // Áp suất khí quyển
        public static float[] huPoint;        // Độ ẩm không khí
        public static float[] rhoaPoint;      // Tỷ trọng đo kiểm
        public static float[] n2Point;        // Tốc độ guồng cánh thực tế
        public static float[] n1Point;        // Tốc độ guồng cánh nhập vào
        public static float[] deltaP_Point;   // Chênh lệch áp suất tại điểm đo lưu lượng
        public static float[] Pe3Point;        // Chênh lệch áp suất tại điểm đo áp suất
        public static float[] TdPoint;        // Nhiệt độ tuyệt đối (nhiệt độ bầu khô)
        public static float[] T_Point;        // Moomen xoắn trên trục quạt
        public static float[] PwPoint;        // Công suất tiêu thụ
        public static float[] IdmPoint;        // Dòng điện 
        public static float[] UdmPoint;       // Điện áp 
        // Dữ liệu tính toán chưa quy đổi
        public static float[] FlowPoint;      // Lưu lượng
        public static float[] PsPoint;        // Áp suất tĩnh
        public static float[] PtPoint;        // Áp suất tổng
        public static float[] EsPoint;        // Hiệu suất tĩnh
        public static float[] EtPoint;        // Hiệu suất tổng
        public static float[] PrtPoint;       // Công suất quạt tính bằng momen xoắn

        // Test data at Standard condition
        public static float[] Std_FlowPoint;  // Lưu lượng
        public static float[] Std_PsPoint;    // Áp suất tĩnh
        public static float[] Std_PtPoint;    // Áp suất tổng
        public static float[] Std_PrPoint;    // Công suất tiêu thụ
        public static float[] Std_EsPoint;    // Hiệu suất tĩnh
        public static float[] Std_EtPoint;    // Hiệu suất tổng
        public static float[] Std_EstPoint;   // Hiệu suất tĩnh tính theo momen xoắn
        public static float[] Std_EttPoint;   // Hiệu suất tĩnh tính theo momen xoắn

        // Test data at Operating condition
        public static float[] Ope_TaPoint;    // Nhiệt độ khí
        public static float[] Ope_rhoaPoint;  // Tỷ trọng khí
        public static float[] Ope_SpPoint;    // Tốc độ guồng cánh

        public static float[] Ope_FlowPoint;  // Lưu lượng
        public static float[] Ope_PsPoint;    // Áp suất tĩnh
        public static float[] Ope_PtPoint;    // Áp suất tổng
        public static float[] Ope_PrPoint;    // Công suất tiêu thụ
        public static float[] Ope_EsPoint;    // Hiệu suất tĩnh
        public static float[] Ope_EtPoint;    // Hiệu suất tổng
        public static float[] Ope_EstPoint;   // Hiệu suất tĩnh tính theo momen xoắn
        public static float[] Ope_EttPoint;   // Hiệu suất tĩnh tính theo momen xoắn

        //// Fitting
        //public static double[] FlowPoint_ft = new double[10];
        //public static double[] PsPoint_ft = new double[10];
        //public static double[] PtPoint_ft = new double[10];
        //public static double[] EsPoint_ft = new double[10];
        //public static double[] EtPoint_ft = new double[10];
        //public static double[] PrPoint_ft = new double[10];
        //public static double[] PrtPoint_ft = new double[10];
        //public static double[] EstPoint_ft = new double[10];
        //public static double[] EttPoint_ft = new double[10];

        // Trend
        //public static float FlowTrend;
        //public static float PsTrend;
        //public static float PtTrend;
        //public static float EsTrend;
        //public static float EtTrend;
        //public static float PrTrend;
        //public static float PrtTrend;
        //public static float EstTrend;
        //public static float EttTrend;

        // chỉ số điểm đo
        //public static int j = 0;

        // Biến tính toán nội suy
        public static float[] rho3;
        public static float[] Pr;            // Công suất trên trục của guồng cánh 


        public static void Initialize(int range)
        {
            if (TaPoint == null || TaPoint.Length != range)
            {
                TaPoint = new float[range];
                PaPoint = new float[range];
                huPoint = new float[range];
                rhoaPoint = new float[range];
                n2Point = new float[range];
                n1Point = new float[range];
                deltaP_Point = new float[range];
                Pe3Point = new float[range];
                TdPoint = new float[range];
                T_Point = new float[range];
                PwPoint = new float[range];
                IdmPoint = new float[range];
                UdmPoint = new float[range];
                FlowPoint = new float[range];
                PsPoint = new float[range];
                PtPoint = new float[range];
                EsPoint = new float[range];
                EtPoint = new float[range];
                PrtPoint = new float[range];
                // Điều kiện tiêu chuẩn
                Std_FlowPoint = new float[range];
                Std_PsPoint = new float[range];
                Std_PtPoint = new float[range];
                Std_PrPoint = new float[range];
                Std_EsPoint = new float[range];
                Std_EtPoint = new float[range];
                Std_EstPoint = new float[range];
                Std_EttPoint = new float[range];
                // Điều kiện hoạt động
                Ope_TaPoint = new float[range];
                Ope_rhoaPoint = new float[range];
                Ope_SpPoint = new float[range];
                Ope_FlowPoint = new float[range];
                Ope_PsPoint = new float[range];
                Ope_PtPoint = new float[range];
                Ope_PrPoint = new float[range];
                Ope_EsPoint = new float[range];
                Ope_EtPoint = new float[range];
                Ope_EstPoint = new float[range];
                Ope_EttPoint = new float[range];
                rho3 = new float[range];
                Pr = new float[range];
            }
        }

        public static MeasureResponse OnePointMeasure(Measure measure, BienTan inv, CamBien sen, OngGio duct, ThongTinMauThuNghiem input)
        {
            void LogCalculation(string message)
            {
                string logPath = Path.Combine(UserSetting.TOMFAN_folder, "OnePointMeasureLog.txt");
                using (var writer = new StreamWriter(logPath, true))
                {
                    writer.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}");
                }
            }
            try
            {
                // Phân cách mỗi lần tính toán
                LogCalculation("========================================");
                LogCalculation($"Bắt đầu tính toán cho điểm đo k = {measure.k}");


                // Auxiliary variable  
                float deltap = 0;      // Chênh lệch áp suất điểm đo lưu lượng
                float pe3 = 0;         // Chênh lệch áp suất điểm đo áp suất
                float Ta = 0;          // Nhiệt độ môi trường
                float Pa = 0;          // Áp suất khí quyển khu vực đo
                float hu = 0;          // Độ ẩm không khí
                float Td = 0;          // Nhiệt độ bầu khô
                float n2 = 0;          // Tốc độ thực của guồng cánh
                float T = 0;          // Momen xoắn trên trục quạt
                float pv = 0;             // Áp suất hơi riêng phần
                float Tw = 0;             // Nhiệt độ bầu ướt
                float psat = 0;           // Áp suất hơi bão hòa của hơi nước trong không khí ẩm
                float Aw = 0;
                float Rw = 0;             // Hằng số lí tưởng
                float rhox = 0;           // Khối lượng riêng không khí tại điểm đo lưu lượng
                float emax = 0;           // Hệ số hỗn hợp
                float e = 0;
                float ae = 0;
                float m = 0;
                float mm = 0;
                float qm = 0;             // Lưu lượng khối lượng
                float q = 0;              // Lưu lượng thể tích
                float qV = 0;             // Lưu lượng thể tích sau khi hiệu chỉnh về điều kiện thiết kế
                float v3 = 0;             // Vận tốc dòng khí
                float A3 = 0;             // Tiết diện đường ống
                float Re = 0;             // Hằng số Reynolds
                float M = 0;              // độ nhớt
                float aett = 0;           // Hệ số hỗn hợp
                float pv3 = 0;            // Áp suất động dòng khí
                float csi13 = 0;          // Hệ số ma sát
                float pf3 = 0;            // Tổn thất áp suất
                float pcb = 0;            // Tổn thất áp suất cục bộ
                float phi = 0;
                float Psu = 0;            // Công suất tĩnh của dòng khí tc
                float Pu = 0;             // Tổng công suất của dòng khí tc
                float Psulv = 0;          // Công suất tĩnh của dòng khí lv
                float Pulv = 0;           // Tổng công suất của dòng khí lv
                float L1_3 = 0;           // Chiều dài tổn thất
                float D3 = 0;             // Đường kính ống tại điểm đo áp suất tĩnh

                #region phần code mới thêm, lấy từ cảm biến và dữ liệu biến tần lấy từ csdl

                float Power_fb = 0; // công suất phản hồi
                float Voltage_fb = 0; // điện áp phản hồi
                float Current_fb = 0; // dòng điện phản hồi

                // thông số động cơ và truyền động
                float e_motor = input.HieuSuatDongCo;
                float n1 = input.TocDoThietKeCuaQuat;
                float CosPhi = input.HeSoCongSuatDongCo;

                // biến tần
                float e_noitruc = inv.HieuSuatNoiTruc;
                float e_goitruc = inv.HieuSuatGoiTruc;
                float e_botruyen = inv.HieuSuatBoTruyen;

                // thông số ống gió và van điều khiển
                float D5 = duct.DuongKinhOngD5;
                float d = duct.DuongKinhMiengQuat;
                float L34 = duct.ChieuDaiConQuat;

                // điều kiện làm việc
                float rhow = 1.204f;
                MeasureResponse measurePoint = new MeasureResponse();
                #endregion

                deltap = measure.ChenhLechApSuat_sen * 750;
                pe3 = measure.ApSuatTinh_sen;
                Ta = measure.NhietDoMoiTruong_sen;
                Pa = measure.ApSuatkhiQuyen_sen;
                hu = measure.DoAm_sen;
                Td = measure.NhietDoMoiTruong_sen; // nhiệt độ bầu khô = nhiệt độ môi trường
                T = measure.Momen_sen;

                // tốc độ thực của guồng cánh = tốc độ định mức * % tần số trong kịch bản
                //n2 = measure.SoVongQuay_sen;
                n2 = (n1 * (float)measure.S) / 100;

                Power_fb = measure.CongSuat_fb;
                Current_fb = measure.DongDien_fb;
                Voltage_fb = measure.DienAp_fb;

                LogCalculation($"Giá trị chênh lệch áp suất deltap: {deltap}");
                LogCalculation($"Giá trị áp suất tĩnh pe3: {pe3}");
                LogCalculation($"Giá trị nhiệt độ môi trường Ta: {Ta}");
                LogCalculation($"Giá trị áp suất khí quyển pa: {Pa}");
                LogCalculation($"Giá trị độ ẩm không khí hu: {hu}");
                LogCalculation($"Giá trị nhiệt độ bầu khô Td: {Td}");
                LogCalculation($"Giá trị tốc độ guồng cánh n2: {n2}");
                LogCalculation($"Giá trị momen xoắn T: {T}");
                LogCalculation($"Giá trị công suất phản hồi Power_fb: {Power_fb}");

                int j = measure.k - 1;
                LogCalculation($"Chỉ số điểm đo j: {j}");

                // hiệu suất tổng
                float e_Total = (e_motor / 100) * (e_goitruc / 100) * (e_noitruc / 100) * (e_botruyen / 100);
                LogCalculation($"Hiệu suất tổng e_Total: {e_Total}");
                L1_3 = 3 * D5 + L34;
                LogCalculation($"Chiều dài tổn thất L1_3: {L1_3}");
                //D3 = D5;
                D3 = duct.DuongKinhOngD3;
                LogCalculation($"Đường kính ống tại điểm đo áp suất tĩnh D3: {D3}");
                //// Công suất tiêu thụ
                ////Delta inv
                //if (InvType == 0) Pw = (float)(Math.Sqrt(3) * (Voltage_fb / 10) * ((Current_fb * 0.86) / 100) * CosPhi / 1000);
                ////Yaskawa inv
                //else if (InvType == 1) Pw = Power_fb;
                //float Pw = Power_fb;
                float Pw = Power_fb;
                LogCalculation($"Công suất tiêu thụ Pw: {Pw}");
                LogCalculation($"Hiệu suất động cơ e_motor: {e_motor}");
                LogCalculation($"Tốc độ thiết kế của quạt n1: {n1}");
                LogCalculation($"Hệ số công suất động cơ CosPhi: {CosPhi}");
                LogCalculation($"Hiệu suất nối trục e_noitruc: {e_noitruc}");
                LogCalculation($"Hiệu suất gối trục e_goitruc: {e_goitruc}");
                LogCalculation($"Hiệu suất bộ truyền e_botruyen: {e_botruyen}");
                LogCalculation($"Đường kính ống gió D5: {D5}");
                LogCalculation($"Đường kính miệng quạt d: {d}");
                LogCalculation($"Chiều dài côn quạt L34: {L34}");
                LogCalculation($"Điện áp vào biến tần Voltage_fb: {Voltage_fb}");
                // Tính toán giá trị tại 1 điểm đo   
                while (1 > 0)
                {
                    // Bước 1: Tính toán khối lượng riêng không khí tại khu vực đo kiểm 'rhoa'
                    // Nhiệt độ bầu ướt
                    Tw = (float)(Td * Math.Atan(0.151977 * Math.Pow((hu + 8.313659), 0.5)) + Math.Atan(Td + hu) - Math.Atan(hu - 1.676331) + 0.00391838 * Math.Pow(hu, 1.5) * Math.Atan(0.023101 * hu) - 4.686035);
                    LogCalculation($"Nhiệt độ bầu ướt Tw: {Tw}");
                    // Áp suất hơi bão hòa của hơi nước trong không khí ẩm từ 0 đến 100
                    psat = (float)(610.8 + 44.442 * Tw + 1.4133 * Math.Pow(Tw, 2) + 0.02768 * Math.Pow(Tw, 3) + 0.000255667 * Math.Pow(Tw, 4) + 2.89166 * Math.Pow(10, -6) * Math.Pow(Tw, 5));
                    LogCalculation($"Áp suất hơi bão hòa của hơi nước trong không khí ẩm psat: {psat}");
                    // Áp suất hơi riêng phần
                    Aw = (float)(6.66 * Math.Pow(10, -4));  //với Tw tu 0 den 150
                    LogCalculation($"Áp suất hơi riêng phần với nhiệt đồ bầu ướt Aw: {Aw}");
                    pv = (float)(psat - Pa * Aw * (Ta - Tw) * (1 + 0.00115 * Tw));
                    LogCalculation($"Áp suất hơi riêng phần pv: {pv}");
                    // Khối lượng riêng không khí tại khu vực đo kiểm
                    rhoaPoint[j] = (Pa - 0.378f * pv) / (287 * (Ta + 273.15f));
                    LogCalculation($"Khối lượng riêng không khí tại khu vực đo kiểm  rhoaPoint[{j}] : {rhoaPoint[j]}");
                    // Bước 2: Tính toán khối lượng riêng không khí tại điểm đo lưu lượng 'rhox'
                    // Hằng số lí tưởng
                    Rw = Pa / (rhoaPoint[j] * (Ta + 273.15f));
                    LogCalculation($"Hằng số lí tưởng Rw: {Rw}");
                    // Khối lượng riêng không khí tại điểm đo lưu lượng
                    rhox = (Pa - deltap) / (Rw * (Td + 273.15f));
                    LogCalculation($"Khối lượng riêng không khí tại điểm đo lưu lượng rhox: {rhox}");
                    // Bước 3-5: Tính toán lưu lượng thể tích hiệu chỉnh về điều kiện thiết kế 'qV'
                    if (D5 > 500 && D5 < 2000)
                    {
                        emax = (float)((0.9131 + 0.0623 * (D5 / 1000) - 0.01567 * (D5 / 1000) * (D5 / 1000)) * 10000 + 1);
                        LogCalculation($"Hệ số hỗn hợp emax: {emax}");
                        m = emax - 9300;
                        LogCalculation($"Giá trị m: {m}");
                        mm = (float)(Math.Round(m));
                        LogCalculation($"Giá trị mm: {mm}");
                        for (int i = 1; i <= mm; i++)
                        {
                            e = emax - i;
                            LogCalculation($"Giá trị e: {e}");
                            ae = e / 10000;
                            LogCalculation($"Giá trị ae: {ae}");
                            // Lưu lượng khối lượng của dòng khí
                            qm = (float)(ae * Math.PI * (D5 / 1000) * (D5 / 1000) / 4 * Math.Pow((2 * rhox * deltap), 0.5));
                            LogCalculation($"Lưu lượng khối lượng của dòng khí qm: {qm}");
                            // Khối lượng riêng của không khí tại vị trí đo áp suất tĩnh
                            rho3[j] = (Pa - pe3) / (Rw * (Ta + 273.15f));
                            LogCalculation($"Khối lượng riêng của không khí tại vị trí đo áp suất tĩnh rho3[{j}]: {rho3[j]}");
                            // Tính toán lưu lượng thể tích
                            q = qm / rho3[j];
                            LogCalculation($"Lưu lượng thể tích q: {q}");
                            // Tính toán lưu lượng thể tích không khí vận chuyển trong đường ống
                            // Sau khi hiệu chỉnh về điều kiện thiết kế 'qV'
                            qV = q * n2 / n1;
                            LogCalculation($"Lưu lượng thể tích sau khi hiệu chỉnh về điều kiện thiết kế qV: {qV}");
                            //Vận tốc dòng khí
                            A3 = (float)(Math.PI * (D5 / 1000) * (D5 / 1000) / 4);
                            LogCalculation($"Tiết diện đường ống A3: {A3}");
                            v3 = q / A3;
                            LogCalculation($"Vận tốc dòng khí v3: {v3}");
                            // Độ nhớt 
                            M = (float)((17.1 + 0.048 * Ta) * Math.Pow(10, -6));
                            LogCalculation($"Độ nhớt M: {M}");
                            // Hằng số Reynolds
                            Re = v3 * (D5 / 1000) * rhox / M;
                            LogCalculation($"Hằng số Reynolds Re: {Re}");
                            // Hệ số hỗn hợp tính toán
                            aett = (float)((-0.00963 + 0.04783 * (D5 / 1000) - 0.01286 * (D5 / 1000) * (D5 / 1000)) * Math.Log10(Re) + 0.9715 - 0.205 * (D5 / 1000) + 0.05533 * (D5 / 1000) * (D5 / 1000));
                            LogCalculation($"Hệ số hỗn hợp tính toán aett: {aett}");
                            if (aett >= emax || ae - aett <= 0.0000009) break;

                            else i++;

                        } // end for

                    } // end if

                    else
                    {
                        // Hệ số hỗn hợp
                        for (int i = 1; i <= 1000; i++)
                        {
                            e = 9401 - i;
                            LogCalculation($"Giá trị e: {e}");
                            ae = e / 10000;
                            LogCalculation($"Giá trị ae: {ae}");
                            // Lưu lượng khối lượng của dòng khí
                            qm = (float)(ae * Math.PI * (D5 / 1000) * (D5 / 1000) / 4 * Math.Pow((2 * rhox * deltap), 0.5));
                            LogCalculation($"Lưu lượng khối lượng của dòng khí: {qm}");
                            // Khối lượng riêng của không khí tại vị trí đo áp suất tĩnh
                            rho3[j] = (Pa - pe3) / (Rw * (Ta + 273.15f));
                            LogCalculation($"Khối lượng riêng của không khí tại vị trí đo áp suất tĩnh rho3[{j}] : {rho3[j]}");
                            // Tính toán lưu lượng thể tích
                            q = qm / rho3[j];
                            LogCalculation($"Lưu lượng thể tích q: {q}");
                            // Tính toán lưu lượng thể tích không khí vận chuyển trong đường ống
                            // Sau khi hiệu chỉnh về điều kiện thiết kế 'qV'
                            qV = q * n2 / n1;
                            LogCalculation($"Lưu lượng thể tích sau khi hiệu chỉnh về điều kiện thiết kế qV: {qV}");
                            //Vận tốc dòng khí
                            A3 = (float)(Math.PI * (D5 / 1000) * (D5 / 1000) / 4);
                            LogCalculation($"Tiết diện đường ống A3: {A3}");
                            v3 = q / A3;
                            LogCalculation($"Vận tốc dòng khí v3: {v3}");
                            // Độ nhớt 
                            M = (float)((17.1 + 0.048 * Ta) * Math.Pow(10, -6));
                            LogCalculation($"Độ nhớt M: {M}");
                            // Hằng số Reynolds
                            Re = v3 * (D5 / 1000) * rhox / M;
                            LogCalculation($"Hằng số Reynolds Re: {Re}");
                            // Hệ số hỗn hợp tính toán
                            aett = (float)(0.01 * Math.Log10(Re) + 0.887);
                            LogCalculation($"Hệ số hỗn hợp tính toán aett: {aett}");
                            if (aett >= 0.94 || ae - aett <= 0.0000009) break;

                            else i++;

                        }

                    } // end else

                    // Lưu lượng chuyển sang m3/h
                    FlowPoint[j] = qV * 3600;
                    LogCalculation($"Lưu lượng FlowPoint[{j}]: {FlowPoint[j]}");
                    Std_FlowPoint[j] = FlowPoint[j];
                    LogCalculation($"Lưu lượng Std_FlowPoint[{j}]: {Std_FlowPoint[j]}");
                    Ope_FlowPoint[j] = FlowPoint[j];
                    LogCalculation($"Lưu lượng Ope_FlowPoint[{j}]: {Ope_FlowPoint[j]}");

                    // Tính toán áp suất tĩnh
                    // Áp suất động của dòng khí
                    pv3 = 0.5f * rho3[j] * v3 * v3;
                    LogCalculation($"Áp suất động của dòng khí pv3: {pv3}");
                    // hệ số ma sát
                    csi13 = (float)(0.414 * Math.Pow(Re, -0.174));
                    LogCalculation($"hệ số ma sát csi13: {csi13}");
                    // Tổn thất áp suất
                    pf3 = csi13 * (L1_3 + L34) / D3 * pv3;
                    LogCalculation($"Tổn thất áp suất pf3: {pf3}");
                    // Tổn thất áp suất cục bộ
                    phi = (float)(Math.Atan((D3 / 2 - d / 2) / L34) * (180 / Math.PI));
                    LogCalculation("Tổn thất áp suất cục bộ phi: {phi}");

                    if (phi < 30 && phi > 0) pcb = 0.05f * pv3;
                    else pcb = 0;
                    LogCalculation($"Tổn thất áp suất cục bộ pcb: {pcb}");

                    // Áp suất tĩnh
                    PsPoint[j] = pe3 - pv3 + pf3 + pcb;
                    LogCalculation($"Áp suất tĩnh PsPoint[{j}]: {PsPoint[j]}");

                    // Áp suất tĩnh hiệu chuẩn
                    Std_PsPoint[j] = (float)((pe3 - pv3 + pf3) * (n2 / n1) * (n2 / n1) * rhokk / rho3[j]);  // Đk tiêu chuẩn
                    LogCalculation($"Áp suất tĩnh điều kiện tiêu chuẩn Std_PsPoint[{j}]: {Std_PsPoint[j]}");
                    Ope_PsPoint[j] = (float)((pe3 - pv3 + pf3) * (n2 / n1) * (n2 / n1) * rhow / rho3[j]);   // Đk làm việc
                    LogCalculation($"Áp suất tĩnh điều kiện làm việc Ope_PsPoint[{j}]: {Ope_PsPoint[j]}");

                    // Áp suất tổng
                    PtPoint[j] = PsPoint[j] + pv3 + pf3;
                    LogCalculation($"Áp suất tổng PtPoint[{j}]: {PtPoint[j]}");

                    // Áp suất tổng hiệu chuẩn
                    Std_PtPoint[j] = PtPoint[j] * (n2 / n1) * (n2 / n1) * (rhokk / rho3[j]);   // Đk tiêu chuẩn
                    LogCalculation($"Áp suất tổng điều kiện tiêu chuẩn Std_PtPoint[{j}]: {Std_PtPoint[j]}");
                    Ope_PtPoint[j] = PtPoint[j] * (n2 / n1) * (n2 / n1) * (rhow / rho3[j]);    // Đk làm việc
                    LogCalculation($"Áp suất tổng điều kiện làm việc Ope_PtPoint[{j}]: {Ope_PtPoint[j]}");

                    // Công suất tĩnh của dòng khí
                    Psu = Std_PsPoint[j] * qV / 1000;
                    LogCalculation($"Công suất tĩnh của dòng khí tc: {Psu}");
                    Psulv = Ope_PsPoint[j] * qV / 1000;
                    LogCalculation($"Công suất tĩnh của dòng khí lv: {Psulv}");

                    // Tổng công suất của dòng khí
                    Pu = Std_PtPoint[j] * qV / 1000;
                    LogCalculation($"Tổng công suất của dòng khí tc: {Pu}");
                    Pulv = Ope_PtPoint[j] * qV / 1000;
                    LogCalculation($"Tổng công suất của dòng khí lv: {Pulv}");

                    // Công suất trên trục của guồng cánh
                    Pr[j] = Pw * e_Total;
                    LogCalculation($"Công suất trên trục của guồng cánh Pr[{j}]: {Pr[j]}");
                    // Công suất trên trục qua hiệu chỉnh
                    Std_PrPoint[j] = Pr[j] * (n2 / n1) * (n2 / n1) * (n2 / n1) * (rhokk / rho3[j]);   // Đk tiêu chuẩn
                    LogCalculation($"Công suất trên trục điều kiện tiêu chuẩn Std_PrPoint[{j}]: {Std_PrPoint[j]}");
                    Ope_PrPoint[j] = Pr[j] * (n2 / n1) * (n2 / n1) * (n2 / n1) * (rhow / rho3[j]);    // Đk làm việc
                    LogCalculation($"Công suất trên trục điều kiện làm việc Ope_PrPoint[{j}]: {Ope_PrPoint[j]}");

                    // Hiệu suất tĩnh
                    EsPoint[j] = Psu / Std_PrPoint[j] * 100;          // Đk tiêu chuẩn
                    LogCalculation($"Hiệu suất tĩnh điều kiện tiêu chuẩn EsPoint[{j}]: {EsPoint[j]}");
                    Std_EsPoint[j] = EsPoint[j];
                    LogCalculation($"Hiệu suất tĩnh điều kiện tiêu chuẩn Std_EsPoint[{j}]: {Std_EsPoint[j]}");
                    Ope_EsPoint[j] = Psulv / Ope_PrPoint[j] * 100;    // Đk làm việc
                    LogCalculation($"Hiệu suất tĩnh điều kiện làm việc Ope_EsPoint[{j}]: {Ope_EsPoint[j]}");

                    // Hiệu suất tổng
                    EtPoint[j] = Pu / Std_PrPoint[j] * 100;          // Đk tiêu chuẩn
                    LogCalculation($"Hiệu suất tổng điều kiện tiêu chuẩn EtPoint[{j}]: {EtPoint[j]}");
                    Std_EtPoint[j] = EtPoint[j];
                    LogCalculation($"Hiệu suất tổng điều kiện tiêu chuẩn Std_EtPoint[{j}]: {Std_EtPoint[j]}");
                    Ope_EtPoint[j] = Pulv / Ope_PrPoint[j] * 100;    // Đk làm việc
                    LogCalculation($"Hiệu suất tổng điều kiện làm việc Ope_EtPoint[{j}]: {Ope_EtPoint[j]}");

                    // Tính toán hiệu suất theo momen xoắn
                    if (T > 0)
                    {
                        LogCalculation("Tính toán hiệu suất theo momen xoắn.");
                        // Công suất tính theo T
                        float Prt = T * n2 / 9550;
                        LogCalculation($"Công suất tính theo momen xoắn Prt: {Prt}");
                        PrtPoint[j] = (float)Math.Round(Prt, 2);
                        LogCalculation($"Công suất tính theo momen xoắn PrtPoint[{j}]: {PrtPoint[j]}");

                        // Hiệu suất tĩnh tính theo T
                        Std_EstPoint[j] = Psu / Prt * 100;
                        LogCalculation($"Hiệu suất tĩnh tính theo momen xoắn Std_EstPoint[{j}]: {Std_EstPoint[j]}");
                        Ope_EstPoint[j] = Psulv / Prt * 100;
                        LogCalculation($"Hiệu suất tĩnh tính theo momen xoắn Ope_EstPoint[{j}]: {Ope_EstPoint[j]}");

                        // Hiệu suất tổng tính theo T
                        Std_EttPoint[j] = Pu / Prt * 100;
                        LogCalculation($"Hiệu suất tổng tính theo momen xoắn Std_EttPoint[{j}]: {Std_EttPoint[j]}");
                        Ope_EttPoint[j] = Pulv / Prt * 100;
                        LogCalculation($"Hiệu suất tổng tính theo momen xoắn Ope_EttPoint[{j}]: {Ope_EttPoint[j]}");
                    }

                    measurePoint = new MeasureResponse
                    {
                        STT = measure.k,
                        Airflow = Ope_FlowPoint[j],
                        Ps = Ope_PsPoint[j],
                        Pt = Ope_PtPoint[j],
                        SEff = Ope_EsPoint[j],
                        TEff = Ope_EtPoint[j],
                        Power = Ope_PrPoint[j],
                        Prt = PrtPoint[j],
                        Est = Ope_EstPoint[j],
                        Ett = Ope_EttPoint[j]
                    };

                    LogCalculation($"Hoàn thành tính toán tại một điểm đo: {JsonSerializer.Serialize(measurePoint)}");

                    // report
                    TaPoint[j] = (float)Math.Round(Ta, 2);
                    rhoaPoint[j] = (float)Math.Round(rhoaPoint[j], 2);
                    n2Point[j] = (float)Math.Round(n2, 2);
                    //test
                    PaPoint[j] = Pa;
                    huPoint[j] = hu;
                    n1Point[j] = n1;
                    deltaP_Point[j] = deltap;
                    Pe3Point[j] = pe3;
                    TdPoint[j] = (float)Math.Round(Td, 2);
                    T_Point[j] = (float)Math.Round(T, 2);
                    PwPoint[j] = (float)Math.Round(Pw, 2);
                    IdmPoint[j] = (float)Math.Round(Current_fb / 100, 3);
                    UdmPoint[j] = (float)Math.Round(Voltage_fb / 10, 3);

                    //------------End While------------//
                    j++;
                    LogCalculation("========================================================");
                    break;
                }
                //if (j == 10)
                //{
                //    j = 0;
                //}
                return measurePoint;
            }
            catch (Exception ex)
            {
#if DEBUG
                LogCalculation("Lỗi khi tính toán tại một điểm đo: " + ex.Message);
                return new MeasureResponse
                {
                    STT = measure.k,
                    Airflow = 0,
                    Ps = 0,
                    Pt = 0,
                    SEff = 0,
                    TEff = 0,
                    Power = 0,
                    Prt = 0,
                    Est = 0,
                    Ett = 0
                };
#else
                return null;
#endif

            }
        }

        public static MeasureFittingFC FittingFC(int range, int index)
        {
            double[] x = new double[range];
            double[] yPs = new double[range];
            double[] yPt = new double[range];
            double[] yEs = new double[range];
            double[] yEt = new double[range];
            double[] yPw = new double[range];
            double[] yPrt = new double[range];
            double[] yEst = new double[range];
            double[] yEtt = new double[range];
            var fc = new MeasureFittingFC(range);


            for (int i = 0; i < range; i++)
            {
                x[i] = Ope_FlowPoint[index + i];
                yPs[i] = Ope_PsPoint[index + i];
                yPt[i] = Ope_PtPoint[index + i];
                yEs[i] = Ope_EsPoint[index + i];
                yEt[i] = Ope_EtPoint[index + i];
                yPw[i] = Ope_PrPoint[index + i];
                yPrt[i] = PrtPoint[index + i];
                yEst[i] = Ope_EstPoint[index + i];
                yEtt[i] = Ope_EttPoint[index + i];


                fc.Ope_FlowPoint[i] = Ope_FlowPoint[index + i];
                fc.Ope_PsPoint[i] = Ope_PsPoint[index + i];
                fc.Ope_PtPoint[i] = Ope_PtPoint[index + i];
                fc.Ope_EsPoint[i] = Ope_EsPoint[index + i];
                fc.Ope_EtPoint[i] = Ope_EtPoint[index + i];
                fc.Ope_PrPoint[i] = Ope_PrPoint[index + i];
                fc.PrtPoint[i] = PrtPoint[index + i];
                fc.Ope_EstPoint[i] = Ope_EstPoint[index + i];
                fc.Ope_EttPoint[i] = Ope_EttPoint[index + i];

            }

            double[] polyPw = Fit.Polynomial(x, yPw, 3);    //1
            double[] polyPs = Fit.Polynomial(x, yPs, 2);    //2
            double[] polyPt = Fit.Polynomial(x, yPt, 2);    //3
            double[] polyEs = Fit.Polynomial(x, yEs, 2);    //4
            double[] polyEt = Fit.Polynomial(x, yEt, 2);    //5
            double[] polyPrt = Fit.Polynomial(x, yPrt, 3);  //6
            double[] polyEst = Fit.Polynomial(x, yEst, 2);  //7
            double[] polyEtt = Fit.Polynomial(x, yEtt, 2);  //8

            double a1 = polyPw[0]; double b1 = polyPw[1]; double c1 = polyPw[2]; double d1 = polyPw[3];
            double a2 = polyPs[0]; double b2 = polyPs[1]; double c2 = polyPs[2];
            double a3 = polyPt[0]; double b3 = polyPt[1]; double c3 = polyPt[2];
            double a4 = polyEs[0]; double b4 = polyEs[1]; double c4 = polyEs[2];
            double a5 = polyEt[0]; double b5 = polyEt[1]; double c5 = polyEt[2];
            double a6 = polyPrt[0]; double b6 = polyPrt[1]; double c6 = polyPrt[2]; double d6 = polyPrt[3];
            double a7 = polyEst[0]; double b7 = polyEst[1]; double c7 = polyEst[2];
            double a8 = polyEtt[0]; double b8 = polyEtt[1]; double c8 = polyEtt[2];
            Sort(x);

            
            for (int i = 0; i < range; i++)
            {
                fc.FlowPoint_ft[i] = x[i];
                fc.PrPoint_ft[i] = a1 + b1 * x[i] + c1 * Math.Pow(x[i], 2) + d1 * Math.Pow(x[i], 3);
                fc.PsPoint_ft[i] = a2 + b2 * x[i] + c2 * Math.Pow(x[i], 2);
                fc.PtPoint_ft[i] = a3 + b3 * x[i] + c3 * Math.Pow(x[i], 2);
                fc.EsPoint_ft[i] = a4 + b4 * x[i] + c4 * Math.Pow(x[i], 2);
                fc.EtPoint_ft[i] = a5 + b5 * x[i] + c5 * Math.Pow(x[i], 2);
                fc.PrtPoint_ft[i] = a6 + b6 * x[i] + c6 * Math.Pow(x[i], 2) + d6 * Math.Pow(x[i], 3);
                fc.EstPoint_ft[i] = a7 + b7 * x[i] + c7 * Math.Pow(x[i], 2);
                fc.EttPoint_ft[i] = a8 + b8 * x[i] + c8 * Math.Pow(x[i], 2);
            }
            return fc;
        }

        // Hàm thông số hiển thị
        public static ParameterShow ParaShow(Measure measure, BienTan inv, CamBien sen, OngGio duct, ThongTinMauThuNghiem input)
        {
            float pv = 0;             // Áp suất hơi riêng phần
            float Tw = 0;             // Nhiệt độ bầu ướt
            float psat = 0;           // Áp suất hơi bão hòa của hơi nước trong không khí ẩm
            float Aw = 0;
            float Rw = 0;             // Hằng số lí tưởng
            float rhox = 0;           // Khối lượng riêng không khí tại điểm đo lưu lượng
            float emax = 0;           // Hệ số hỗn hợp
            float e = 0;
            float ae = 0;
            float m = 0;
            float mm = 0;
            float qm = 0;             // Lưu lượng khối lượng
            float q = 0;              // Lưu lượng thể tích
            float qV = 0;             // Lưu lượng thể tích sau khi hiệu chỉnh về điều kiện thiết kế
            float v3 = 0;             // Vận tốc dòng khí
            float A3 = 0;             // Tiết diện đường ống
            float Re = 0;             // Hằng số Reynolds
            float M = 0;              // độ nhớt
            float aett = 0;           // Hệ số hỗn hợp
            float pv3 = 0;            // Áp suất động dòng khí
            float csi13 = 0;          // Hệ số ma sát
            float pf3 = 0;            // Tổn thất áp suất
            float pcb = 0;            // Tổn thất áp suất cục bộ
            float phi = 0;
            float rho3 = 0;
            float rhoatem = 0;
            float L1_3 = 0;           // Chiều dài tổn thất
            float D3 = 0;             // Đường kính ống tại điểm đo áp suất tĩnh


            // thông số động cơ và truyền động
            float e_motor = input.HieuSuatDongCo;
            float n1 = input.TocDoThietKeCuaQuat;
            float CosPhi = input.HeSoCongSuatDongCo;

            // thông số ống gió và van điều khiển
            float D5 = duct.DuongKinhOngD5;
            float d = duct.DuongKinhMiengQuat;
            float L34 = duct.ChieuDaiConQuat;

            float Ta = measure.NhietDoMoiTruong_sen;
            // thông số phản hồi
            float Power_fb = measure.CongSuat_fb;
            float Current_fb = measure.DongDien_fb;
            float Voltage_fb = measure.DienAp_fb;


            //// note
            ///
            float deltap = measure.ChenhLechApSuat_sen * 750; // chênh lệch áp suất điểm đo lưu lượng
            float Pe3 = measure.ApSuatTinh_sen; // chênh lệch áp suất điểm đo áp suất tĩnh
            // nhiệt độ môi trường
            float Pa = measure.ApSuatkhiQuyen_sen; // áp suất khí quyển khu vực đo
            float hu = measure.DoAm_sen; // độ ẩm không khí
            float Td = measure.NhietDoMoiTruong_sen; // nhiệt độ bầu khô
            //float n2 = measure.SoVongQuay_sen; // tốc độ guồng cánh (số vòng quay)
            // tốc độ thực của guồng cánh = tốc độ định mức * % tần số trong kịch bản
            //n2 = measure.SoVongQuay_sen;
            float n2 = (n1 * (float)measure.S) / 100;
            float T = measure.Momen_sen; // momen xoắn

            float Freq_fb = measure.TanSo_fb; // tần số
            float BearingTemp = 0; // nhiệt độ gối đứng
            float BearingVia = measure.DoRung_sen; // độ rung gối đứng

            // Tính toán đầu
            L1_3 = 3 * D5 + L34;
            //D3 = D5;
            D3 = duct.DuongKinhOngD3;

            // Bước 1: Tính toán khối lượng riêng không khí tại khu vực đo kiểm 'rhoa'
            // Nhiệt độ bầu ướt
            Tw = (float)(Td * Math.Atan(0.151977 * Math.Pow((hu + 8.313659), 0.5)) + Math.Atan(Td + hu) - Math.Atan(hu - 1.676331) + 0.00391838 * Math.Pow(hu, 1.5) * Math.Atan(0.023101 * hu) - 4.686035);
            // Áp suất hơi bão hòa của hơi nước trong không khí ẩm từ 0 đến 100
            psat = (float)(610.8 + 44.442 * Tw + 1.4133 * Math.Pow(Tw, 2) + 0.02768 * Math.Pow(Tw, 3) + 0.000255667 * Math.Pow(Tw, 4) + 2.89166 * Math.Pow(10, -6) * Math.Pow(Tw, 5));
            // Áp suất hơi riêng phần
            Aw = (float)(6.66 * Math.Pow(10, -4));  //với Tw tu 0 den 150
            pv = (float)(psat - Pa * Aw * (Ta - Tw) * (1 + 0.00115 * Tw));
            // Khối lượng riêng không khí tại khu vực đo kiểm
            rhoatem = (Pa - 0.378f * pv) / (287 * (Ta + 273.15f));

            // Bước 2: Tính toán khối lượng riêng không khí tại điểm đo lưu lượng 'rhox'
            // Hằng số lí tưởng
            Rw = Pa / (rhoatem * (Ta + 273.15f));
            // Khối lượng riêng không khí tại điểm đo lưu lượng
            rhox = (Pa - deltap) / (Rw * (Td + 273.15f));

            // Bước 3-5: Tính toán lưu lượng thể tích hiệu chỉnh về điều kiện thiết kế 'qV'
            if (D5 > 500 && D5 < 2000)
            {
                emax = (float)((0.9131 + 0.0623 * (D5 / 1000) - 0.01567 * (D5 / 1000) * (D5 / 1000)) * 10000 + 1);
                m = emax - 9300;
                mm = (float)(Math.Round(m));

                for (int i = 1; i <= mm; i++)
                {
                    e = emax - i;
                    ae = e / 10000;
                    // Lưu lượng khối lượng của dòng khí
                    qm = (float)(ae * Math.PI * (D5 / 1000) * (D5 / 1000) / 4 * Math.Pow((2 * rhox * deltap), 0.5));
                    // Khối lượng riêng của không khí tại vị trí đo áp suất tĩnh
                    rho3 = (Pa - Pe3) / (Rw * (Ta + 273.15f));
                    // Tính toán lưu lượng thể tích
                    q = qm / rho3;
                    // Tính toán lưu lượng thể tích không khí vận chuyển trong đường ống
                    // Sau khi hiệu chỉnh về điều kiện thiết kế 'qV'
                    qV = q * n2 / n1;
                    //Vận tốc dòng khí
                    A3 = (float)(Math.PI * (D5 / 1000) * (D5 / 1000) / 4);
                    v3 = q / A3;
                    // Độ nhớt 
                    M = (float)((17.1 + 0.048 * Ta) * Math.Pow(10, -6));
                    // Hằng số Reynolds
                    Re = v3 * (D5 / 1000) * rhox / M;
                    // Hệ số hỗn hợp tính toán
                    aett = (float)((-0.00963 + 0.04783 * (D5 / 1000) - 0.01286 * (D5 / 1000) * (D5 / 1000)) * Math.Log10(Re) + 0.9715 - 0.205 * (D5 / 1000) + 0.05533 * (D5 / 1000) * (D5 / 1000));

                    if (aett >= emax || ae - aett <= 0.0000009) break;

                    else i++;

                } // end for

            } // end if

            else
            {
                // Hệ số hỗn hợp
                for (int i = 1; i <= 1000; i++)
                {
                    e = 9401 - i;
                    ae = e / 10000;
                    // Lưu lượng khối lượng của dòng khí
                    qm = (float)(ae * Math.PI * (D5 / 1000) * (D5 / 1000) / 4 * Math.Pow((2 * rhox * deltap), 0.5));
                    // Khối lượng riêng của không khí tại vị trí đo áp suất tĩnh
                    rho3 = (Pa - Pe3) / (Rw * (Ta + 273.15f));
                    // Tính toán lưu lượng thể tích
                    q = qm / rho3;
                    // Tính toán lưu lượng thể tích không khí vận chuyển trong đường ống
                    // Sau khi hiệu chỉnh về điều kiện thiết kế 'qV'
                    qV = q * n2 / n1;
                    //Vận tốc dòng khí
                    A3 = (float)(Math.PI * (D5 / 1000) * (D5 / 1000) / 4);
                    v3 = q / A3;
                    // Độ nhớt 
                    M = (float)((17.1 + 0.048 * Ta) * Math.Pow(10, -6));
                    // Hằng số Reynolds
                    Re = v3 * (D5 / 1000) * rhox / M;
                    // Hệ số hỗn hợp tính toán
                    aett = (float)(0.01 * Math.Log10(Re) + 0.887);

                    if (aett >= 0.94 || ae - aett <= 0.0000009) break;

                    else i++;

                }

            } // end else

            

            // Tính toán áp suất tĩnh
            // Áp suất động của dòng khí
            pv3 = 0.5f * rho3 * v3 * v3;
            // hệ số ma sát
            csi13 = (float)(0.414 * Math.Pow(Re, -0.174));
            // Tổn thất áp suất
            pf3 = csi13 * (L1_3 + L34) / D3 * pv3;
            // Tổn thất áp suất cục bộ
            phi = (float)(Math.Atan((D3 / 2 - d / 2) / L34) * (180 / Math.PI));

            if (phi < 30 && phi > 0) pcb = 0.05f * pv3;

            else pcb = 0;

            //// Lưu lượng chuyển sang m3/h
            //Flow_show = qV * 3600;
            //// Áp suất tĩnh
            //Ps_show = Pe3 - pv3 + pf3 + pcb;
            //// Áp suất tổng
            //Pt_show = Ps_show + pv3 + pf3;
            //// Tính toán hiệu suất theo momen xoắn
            //Prt_show = T * n2 / 9550;
            //// Công suất tiêu thụ
            //////Delta inv
            ////if (InvType == 0) Pw_show = (float)(Math.Sqrt(3) * (Voltage_fb / 10) * ((Current_fb * 0.86) / 100) * CosPhi / 1000);
            //////Yaskawa inv
            ////else if (InvType == 1) Pw_show = Power_fb;


            //Pw_show = Power_fb;
            //// Parameter Show
            //Freq_show = (Freq_fb) / 100;
            //Freq_show = (float)Math.Round(Freq_show, 2);
            //Current_show = (float)(Current_fb / 100);
            //Current_show = (float)Math.Round(Current_show, 2);
            //Pw_show = (float)Math.Round(Pw_show, 2);
            //Speed_show = Freq_show * n1 / 50;
            //Speed_show = (float)Math.Round(Speed_show, 2);
            //TempB_show = (float)Math.Round(BearingTemp, 2);
            //T_Show = (float)Math.Round(T, 2);
            //Ps_show = (float)Math.Round(Ps_show, 2);
            //Pt_show = (float)Math.Round(Pt_show, 2);
            //Flow_show = (float)Math.Round(Flow_show, 2);
            //Ta_show = (float)Math.Round(Ta, 2);
            //Td_show = (float)Math.Round(Td, 2);
            //ViaB_show = (float)Math.Round(BearingVia, 2);
            //Prt_show = (float)Math.Round(Prt_show, 2);

            return new ParameterShow
            {
                Freq_show = (float)Math.Round(Freq_fb, 2),
                Current_show = (float)Math.Round(Current_fb / 100, 2),
                Pw_show = (float)Math.Round(Power_fb, 2),
                Speed_show = (float)Math.Round((float)Math.Round(Freq_fb, 2) * n1 / 50, 2),
                TempB_show = (float)Math.Round(BearingTemp, 2),
                T_Show = (float)Math.Round(T, 2),
                Ps_show = (float)Math.Round(Pe3 - pv3 + pf3 + pcb, 2),
                Pt_show = (float)Math.Round((Pe3 - pv3 + pf3 + pcb) + pv3 + pf3, 2), 
                Flow_show = (float)Math.Round(qV * 3600, 2),
                Ta_show = (float)Math.Round(Ta, 2),
                Td_show = (float)Math.Round(Td, 2),
                ViaB_show = (float)Math.Round(BearingVia, 2),
                Prt_show = (float)Math.Round(T * n2 / 9550, 2),
                deltap = deltap,
                Pe3 = Pe3,
                Ta = Ta
            };
        }

        // Hàm test tính toán
        public static float qm = 0;
        public static float qV = 0;
        public static float Rw = 0;
        public static float Tu = 0;
        public static float pv = 0;
        public static float As3 = 0;
        public static float V3 = 0;
        public static float pv3 = 0;
        public static float pf3 = 0;
        public static float pcb = 0;
        public static float rhox = 0;
        public static float ae = 0;
        public static float M = 0;
        public static float Re = 0;
        public static float csi3 = 0;


        // Hàm phụ trợ
        static void Swap(double[] array, int i, int m)
        {
            double temp = array[i];
            array[i] = array[m];
            array[m] = temp;
        }

        static void Sort(double[] array)
        {
            double[] minValue = new double[array.Length];
            for (int i = 0; i < array.Length - 1; i++)
            {
                int m = i;
                minValue[i] = array[i];
                for (int j = i + 1; j < array.Length; j++)
                {
                    if (array[j].CompareTo(minValue[i]) < 0)
                    {
                        m = j;
                        minValue[i] = array[j];
                    }
                }
                Swap(array, i, m);
            }
        }
    }
}
