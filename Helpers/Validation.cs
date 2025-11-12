using System.Text.RegularExpressions;

namespace TESMEA_TMS.Helpers
{
    public class Validation
    {
        public Validation()
        {

        }
        public static bool IsValidPassword(string password)
        {
            string specialChars = "!@#$%^&*()_+-=[]{}|;':\",.<>/?";

            if (string.IsNullOrEmpty(password))
                throw new BusinessException("Mật khẩu không được phép để trống");
            if (password.Length < 8)
                throw new BusinessException("Mật khẩu phải có độ dài tối thiểu 8 ký tự");
            if (!password.Any(char.IsUpper))
                throw new BusinessException("Mật khẩu phải có ít nhất 1 ký tự viết hoa");
            if (!password.Any(char.IsLower))
                throw new BusinessException("Mật khẩu phải có ít nhất 1 ký tự viết thường");
            if (!password.Any(char.IsDigit))
                throw new BusinessException("Mật khẩu phải có ít nhất 1 chữ số");
            if (!password.Any(c => specialChars.Contains(c)))
                throw new BusinessException("Mật khẩu phải có ít nhất 1 ký tự đặc biệt");
            if (password.ToLower().Distinct().Count() == 1)
                throw new BusinessException("Mật khẩu không được lặp lại 1 ký tự duy nhất");
            return true;
        }
        public static bool IsValidEmail(string email)
        {
            var trimmedEmail = email.Trim();

            if (trimmedEmail.EndsWith("."))
            {
                return false;
            }
            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return addr.Address == trimmedEmail;
            }
            catch
            {
                return false;
            }
        }

        public static bool IsPhoneNumber(string number)
        {
            number = number.Replace(" ", "").Replace("-", "");
            return Regex.IsMatch(number, @"^((\+84|0)[3|5|7|8|9][0-9]{8})$");
            //return Regex.Match(number, @"^(\+[0-9]{9})$").Success;
        }

        public static bool IsValidGuid(object input)
        {
            if (input == null)
                return false;

            if (input is Guid guid)
                return guid != Guid.Empty;

            if (input is string s)
                return Guid.TryParse(s, out var g) && g != Guid.Empty;

            return false;
        }

        public static bool IsValidDouble(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return false;
            if (double.TryParse(input, out double result))
                return !double.IsNaN(result) && !double.IsInfinity(result);
            return false;
        }

        public static bool IsValidFloat(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return false;
            if (float.TryParse(input, out float result))
                return !float.IsNaN(result) && !float.IsInfinity(result);
            return false;
        }

        public static bool IsValidDecimal(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return false;
            return decimal.TryParse(input, out _);
        }
    }

}
