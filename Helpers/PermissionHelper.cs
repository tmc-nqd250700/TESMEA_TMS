using System.Windows;
using TESMEA_TMS.Configs;

namespace TESMEA_TMS.Helpers
{
    public class PermissionHelper
    {
        public static bool HasPermission(string module, string action)
        {
            var claim = $"{module}.{action}";
            return CurrentUser.Instance.Claims?.Any(c => c == claim) == true;
        }
    }
    public static class PermissionVisibility
    {
        public static readonly DependencyProperty PermissionTagProperty =
            DependencyProperty.RegisterAttached(
                "PermissionTag",
                typeof(string),
                typeof(PermissionVisibility),
                new PropertyMetadata(null, OnPermissionTagChanged));

        public static void SetPermissionTag(UIElement element, string value)
            => element.SetValue(PermissionTagProperty, value);

        public static string GetPermissionTag(UIElement element)
            => (string)element.GetValue(PermissionTagProperty);

        private static void OnPermissionTagChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is UIElement element && e.NewValue is string tag)
            {
                // Tách module và action từ tag, ví dụ: "UserManagement.View"
                var parts = tag.Split('.');
                if (parts.Length == 2)
                {
                    bool hasPermission = PermissionHelper.HasPermission(parts[0], parts[1]);
                    element.Visibility = hasPermission ? Visibility.Visible : Visibility.Collapsed;
                }
                else
                {
                    element.Visibility = Visibility.Collapsed;
                }
            }
        }
    }
}
