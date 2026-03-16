using System.IO;
using System.Text.RegularExpressions;
using TESMEA_TMS.Configs;

namespace TESMEA_TMS.Helpers
{
    public static class Common
    {
        public static bool IsFileLocked(string filePath)
        {
            if (!File.Exists(filePath))
                return false; // Nếu file chưa tồn tại thì không bị khóa

            FileStream stream = null;
            try
            {
                stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                // Nếu mở được stream thì file không bị khóa
                return false;
            }
            catch (IOException)
            {
                // Nếu bị IOException thì file đang bị lock (đang mở ở ứng dụng khác)
                return true;
            }
            finally
            {
                stream?.Close();
            }
        }

        /// <summary>
        /// Chuyển đổi các chuỗi đơn vị "m2" và "m3" thành "m²" và "m³" (Unicode).
        /// </summary>
        public static string ToSuperscriptUnit(string input)
        {
            if (string.IsNullOrEmpty(input))
                return input;

            input = Regex.Replace(input, @"m2\b", "m²", RegexOptions.IgnoreCase);
            input = Regex.Replace(input, @"m3\b", "m³", RegexOptions.IgnoreCase);

            return input;
        }

        public static int AdaptiveRoundUpByLength(double value)
        {
            int absValue = Math.Abs((int)value);
            int digitCount = absValue == 0 ? 1 : (int)Math.Floor(Math.Log10(absValue)) + 1;

            int factor = digitCount switch
            {
                1 => 5,     // 1 chữ số: làm tròn lên 5
                2 => 10,    // 2 chữ số: lên 10
                3 => 100,   // 3 chữ số: lên 100
                4 => 1000,  // 4 chữ số: lên 1000
                5 => 10000, // 5 chữ số: lên 10000
                6 => 100000, // 6 chữ số: lên 100000
                _ => 1000000 // Lớn hơn: lên 1000000
            };

            return (value % factor == 0) ? (int)value : (int)(value + factor - (value % factor));
        }

        public static int RoundUpToNearest(float value, int step = 50)
        {
            return (int)(Math.Ceiling(value / (double)step) * step);
        }

        public static void ShowMessageBoxHelper(string type)
        {
            var isEn = UserSetting.Instance.Language == "en";
            switch (type)
            {
                case "btnNhietDoMoiTruong":
                    MessageBoxHelper.ShowInformation(isEn ? "Ambient temperature is the temperature of the surrounding measurement environment. The sensor's measuring range is configured according to the catalog" : "Nhiệt độ của môi trường nơi đo kiểm, dải đo được cấu hình theo catalog");
                    break;
                case "btnDoAmMoiTruong":
                    MessageBoxHelper.ShowInformation(isEn ? "Ambient humidity is the humidity of the surrounding measurement environment. The sensor's measuring range is configured according to the catalog" : "Độ ẩm của môi trường nơi đo kiểm, dải đo được cấu hình theo catalog");
                    break;
                case "btnApSuatKhiQuyen":
                    MessageBoxHelper.ShowInformation(isEn ? "Atmospheric pressure is the pressure of the surrounding air at the measurement location. The measuring range is configured according to the catalog" : "Áp suất khí quyển là áp suất của không khí nơi đo kiểm, dải đo được cấu hình theo catalog"); 
                    break;
                case "btnChenhLechApSuat":
                    MessageBoxHelper.ShowInformation(isEn ? "Differential pressure at the flow measurement point, obtained from the sensor. The measuring range is configured according to the sensor catalog" : "Chênh lệch áp suất tải điểm đo lưu lượng thu được, dải đo được cấu hình theo catalog"); 
                    break;
                case "btnApSuatTinh":
                    MessageBoxHelper.ShowInformation(isEn ? "Static pressure data obtained from the sensor. The measuring range is configured according to the sensor catalog" : "Áp suất tĩnh thu được, dải đo được cấu hình theo catalog");
                    break;
                case "btnDoRung":
                    MessageBoxHelper.ShowInformation(isEn ? "Mechanical vibration of the bearing. The measuring range is configured according to the sensor catalog" : "Độ rung cơ học của gối trục, dải đo được cấu hình theo catalog");
                    break;
                case "btnSoVongQuay":
                    MessageBoxHelper.ShowInformation(isEn ? "Motor speed. The measuring range is configured according to the sensor catalog" : "Tốc độ của động cơ, dải đo được cấu hình theo catalog");
                    break;
                case "btnMomen":
                    break;
                case "btnPhanHoiDongDien":
                    MessageBoxHelper.ShowInformation(isEn ? "Current data obtained from the inverter. The measuring range is configured according to the motor nameplate" : "Dòng điện thu được từ biến tần, dải đo cấu hình theo nhãn của động cơ");
                    break;
                case "btnPhanHoiCongSuat":
                    MessageBoxHelper.ShowInformation(isEn ? "Power data obtained from the inverter. The measuring range is based on the inverter or motor power rating (if the inverter supports automatic power calibration for the motor)" : "Công suất thu được từ biến tần, dải đo theo công suất của biến tần hoặc động cơ (nếu biến tần hỗ trợ tự hiệu chỉnh về công suất động cơ)");
                    break;
                case "btnPhanHoiViTriVan":
                    MessageBoxHelper.ShowInformation(isEn ? "Valve position feedback at the measurement point. The measuring range is 0–100%" : "Phản hồi vị trí van so tại điểm đo, dải đo từ 0-100%");
                    break;
                case "btnNhietDoGoiTruc":
                    MessageBoxHelper.ShowInformation(isEn ? "Infrared temperature data obtained from the sensor. The measuring range is configured according to the sensor catalog" : "Nhiệt độ hồng ngoại thu được, đải đo được cấu hình theo catalog");
                    break;
                case "btnPhanHoiTanSo":
                    MessageBoxHelper.ShowInformation(isEn ? "Frequency feedback data obtained from the sensor. The measuring range is configured according to the sensor catalog" : "Phản hồi tần số thu được, dải đo được cấu hình theo catalog");
                    break;



                // hiệu suất biến tần
                case "btnHieuSuatBoTruyen":
                    MessageBoxHelper.ShowInformation(isEn ? "Transmission efficiency is the power transfer ratio between the inverter output power and the motor input power" : "Hiệu suất truyền động là tỷ lệ truyền tải giữa công suất đầu ra của biến tần và công suất đầu vào của động cơ");
                    break;
                case "btnHieuSuatNoiTruc":
                    MessageBoxHelper.ShowInformation(isEn ? "Coupling efficiency is the ratio of power transmitted to the shaft through the coupling" : "Hiệu suất nối trục là tỷ lệ truyền tải công suất đến trục qua khớp nối");
                    break;
                case "btnHieuSuatGoiTruc":
                    MessageBoxHelper.ShowInformation(isEn ? "Bearing efficiency is the ratio representing the reliability and degree of power loss due to friction in the rolling bearing support during operation" : "Hiệu suất gối trục là tỷ lệ độ tin cậy và mức độ hao hụt công suất do ma sát của gối đỡ vòng bi khi làm việc");
                    break;


                // ống gió
                case "btnDuongKinhOngD5":
                    MessageBoxHelper.ShowInformation(isEn ? "D5 duct diameter is the inlet diameter" : "Đường kính ống D5 là đường kính miệng hút");
                    break;
                case "btnChieuDaiConQuat":
                    MessageBoxHelper.ShowInformation(isEn ? "Fan cone length" : "Chiều dài côn quạt");
                    break;
                case "btnDuongKinhOngD3":
                    MessageBoxHelper.ShowInformation(isEn ? "D3 duct diameter is the diameter of the static pressure measurement pipe" : "Đường kính ống D3 là đường kính ống đo áp suất tĩnh");
                    break;
                case "btnDuongKinhLoPhut":
                    MessageBoxHelper.ShowInformation(isEn ? "The diameter of the auxiliary hole is the diameter of the auxiliary pipe (applies when measuring type B)" : "Đường kính lỗ phụt là đường kính của ống phụt (áp dụng khi đo kiểu B)");
                    break;

                // thông số quạt
                case "btnCongSuatDongCo":
                    MessageBoxHelper.ShowInformation(isEn ? "Motor power is mentioned on the nameplate of the motor" : "Công suất động cơ là công suất in trên nhãn của động cơ");
                    break;
                case "btnTocDoThietKeCuaQuat":
                    MessageBoxHelper.ShowInformation(isEn ? "Motor speed is mentioned on the nameplate of the motor" : "Tốc độ động cơ là tốc độ in trên nhãn của động cơ");
                    break;
                case "btnHeSoCongSuatDongCo":
                    MessageBoxHelper.ShowInformation(isEn ? "Cos (phi) is mentioned on the nameplate of the motor" : "Hệ số Cos (phi) là hệ số Cos (phi) in trên nhãn của động cơ");
                    break;
                case "btnHieuSuatDongCo":
                    MessageBoxHelper.ShowInformation(isEn ? "Motor efficiency is mentioned on the nameplate of the motor" : "Hiệu suất động cơ là hiệu suất in trên nhãn của động cơ");
                    break;
                case "btnDongDienDinhMucCuaDongCo":
                    MessageBoxHelper.ShowInformation(isEn ? "Rated current is mentioned on the nameplate of the motor" : "Dòng điện định mức là dòng điện in trên nhãn của động cơ");
                    break;
                case "btnDienApDongCo":
                    MessageBoxHelper.ShowInformation(isEn ? "Motor voltage is mentioned on the nameplate of the motor" : "Điện áp động cơ là điện áp in trên nhãn của động cơ");
                    break;
                default:
                    break;
            }
        }
    }
}
