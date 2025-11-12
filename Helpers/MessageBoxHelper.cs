using System.Windows;
using MessageBox = System.Windows.MessageBox;

namespace TESMEA_TMS.Helpers
{
    public static class MessageBoxHelper
    {
        public static void ShowSuccess(string message)
        {
            MessageBox.Show(message, "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        public static void ShowInformation(string message)
        {
            MessageBox.Show(message, "", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        public static void ShowWarning(string message)
        {
            MessageBox.Show(message, "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
        }

        public static void ShowError(string message)
        {
            MessageBox.Show(message, "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        public static bool ShowQuestion(string message)
        {
            var result = MessageBox.Show(message, "Xác nhận", MessageBoxButton.YesNo, MessageBoxImage.Question);
            return result == MessageBoxResult.Yes;
        }
    }
}
