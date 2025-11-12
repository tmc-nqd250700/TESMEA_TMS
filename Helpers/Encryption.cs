using System.Text;
using System.Security.Cryptography;
namespace TESMEA_TMS.Helpers
{
    public class Encryption
    {
        public static string GetHashCode(string i_str_ClearText)
        {
            MD5 mD = new MD5CryptoServiceProvider();
            byte[] bytes = Encoding.Default.GetBytes(i_str_ClearText);
            byte[] inArray = mD.ComputeHash(bytes);
            return Convert.ToBase64String(inArray);
        }

        public static string Encrypt(string toEncrypt, bool useHashing)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(toEncrypt);
            string s = "@Tomeco2024";
            byte[] key;
            if (useHashing)
            {
                MD5CryptoServiceProvider mD5CryptoServiceProvider = new MD5CryptoServiceProvider();
                key = mD5CryptoServiceProvider.ComputeHash(Encoding.UTF8.GetBytes(s));
                mD5CryptoServiceProvider.Clear();
            }
            else
            {
                key = Encoding.UTF8.GetBytes(s);
            }

            TripleDESCryptoServiceProvider tripleDESCryptoServiceProvider = new TripleDESCryptoServiceProvider();
            tripleDESCryptoServiceProvider.Key = key;
            tripleDESCryptoServiceProvider.Mode = CipherMode.ECB;
            tripleDESCryptoServiceProvider.Padding = PaddingMode.PKCS7;
            ICryptoTransform cryptoTransform = tripleDESCryptoServiceProvider.CreateEncryptor();
            byte[] array = cryptoTransform.TransformFinalBlock(bytes, 0, bytes.Length);
            tripleDESCryptoServiceProvider.Clear();
            return Convert.ToBase64String(array, 0, array.Length);
        }

        public static string Decrypt(string cipherString, bool useHashing)
        {
            byte[] array = Convert.FromBase64String(cipherString);
            string s = "@Tomeco2024";
            byte[] key;
            if (useHashing)
            {
                MD5CryptoServiceProvider mD5CryptoServiceProvider = new MD5CryptoServiceProvider();
                key = mD5CryptoServiceProvider.ComputeHash(Encoding.UTF8.GetBytes(s));
                mD5CryptoServiceProvider.Clear();
            }
            else
            {
                key = Encoding.UTF8.GetBytes(s);
            }

            TripleDESCryptoServiceProvider tripleDESCryptoServiceProvider = new TripleDESCryptoServiceProvider();
            tripleDESCryptoServiceProvider.Key = key;
            tripleDESCryptoServiceProvider.Mode = CipherMode.ECB;
            tripleDESCryptoServiceProvider.Padding = PaddingMode.PKCS7;
            ICryptoTransform cryptoTransform = tripleDESCryptoServiceProvider.CreateDecryptor();
            byte[] bytes = cryptoTransform.TransformFinalBlock(array, 0, array.Length);
            tripleDESCryptoServiceProvider.Clear();
            return Encoding.UTF8.GetString(bytes);
        }
    }
}
