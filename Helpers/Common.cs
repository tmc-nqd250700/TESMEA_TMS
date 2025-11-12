using System.IO;
using System.Text.RegularExpressions;

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
    }
}
