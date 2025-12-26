using Microsoft.Extensions.DependencyInjection;
using OfficeOpenXml;
using System.Globalization;
using System.IO;
using System.Windows;
using TESMEA_TMS.Configs;
using TESMEA_TMS.Helpers;
using TESMEA_TMS.Services;
using TESMEA_TMS.ViewModels;
using Application = System.Windows.Application;

namespace TESMEA_TMS;

/// <summary>
/// Interaction logic for App.xaml
/// </summary>
public partial class App : Application
{
    public static IServiceProvider? ServiceProvider { get; private set; }
    protected override void OnStartup(StartupEventArgs e)
    {
        // Config global language
        var setting = UserSetting.Load();
        CultureInfo culture;
        if (setting.Language == "vi")
        {
            culture = new CultureInfo("vi");
        }
        else if (setting.Language == "en")
        {
            culture = new CultureInfo("en");
        }
        else
        {
            culture = CultureInfo.CurrentCulture;
        }


        Thread.CurrentThread.CurrentCulture = culture;
        Thread.CurrentThread.CurrentUICulture = culture;
        WPFLocalizeExtension.Engine.LocalizeDictionary.Instance.Culture = culture;
        System.Threading.Thread.CurrentThread.CurrentUICulture = culture;
        base.OnStartup(e);
        try
        {
            var serviceProvider = ServiceConfigurator.Configure();
            App.ServiceProvider = serviceProvider;
            // Khởi tạo ViewModelLocator với ServiceProvider
            ViewModelLocator.Initialize(serviceProvider);
            // configure global exception handling
            ConfigureExceptionHandling();
            var navigationService = serviceProvider.GetRequiredService<IAppNavigationService>();
            var settingPath = UserSetting.GetDefaultFilePath();
            if (!File.Exists(settingPath))
            {
                var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "1.0.0";
                setting.CurrentVersion = version;

                var exePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                var lastWrite = File.GetLastWriteTime(exePath);
                setting.LastUpdate = lastWrite.ToString("dd/MM/yyyy");
                setting.Save();
            }

            if (!Directory.Exists(UserSetting.GetLocalAppPath()))
                Directory.CreateDirectory(UserSetting.GetLocalAppPath());


            //// folder TOMFAN lưu các file exchange, trendline
            //var exchangeFolder = UserSetting.TOMFAN_folder;
            //if (!Directory.Exists(exchangeFolder))
            //{
            //    Directory.CreateDirectory(exchangeFolder);
            //    var dirInfo = new DirectoryInfo(exchangeFolder);
            //    dirInfo.Attributes |= FileAttributes.Hidden;

            //    // Set folder administrators only permissions
            //    //var dirSecurity = dirInfo.GetAccessControl();
            //    //dirSecurity.SetAccessRuleProtection(true, false);
            //    //var adminSid = new SecurityIdentifier(WellKnownSidType.BuiltinAdministratorsSid, null);
            //    //var accessRule = new FileSystemAccessRule(
            //    //    adminSid,
            //    //    FileSystemRights.FullControl,
            //    //    InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
            //    //    PropagationFlags.None,
            //    //    AccessControlType.Allow);
            //    //dirSecurity.SetAccessRule(accessRule);
            //    //dirInfo.SetAccessControl(dirSecurity);

            //    // Create folder trend and 2 files to exchange
            //    if (!Directory.Exists(Path.Combine(exchangeFolder, "Trend")))
            //    {
            //        Directory.CreateDirectory(Path.Combine(exchangeFolder, "Trend"));
            //    }
            //    using (var writer2 = new StreamWriter(Path.Combine(exchangeFolder, "1_T_OUT.csv")))
            //    {
            //    }

            //    using (var writer2 = new StreamWriter(Path.Combine(exchangeFolder, "2_S_IN.csv")))
            //    {
            //    }

            //    string xlsxPath = Path.Combine(exchangeFolder, "1_T_OUT.xlsx");
            //    if (!File.Exists(xlsxPath))
            //    {
            //        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            //        using (var package = new ExcelPackage(new FileInfo(xlsxPath)))
            //        {
            //            package.Workbook.Worksheets.Add("1_T_OUT");
            //            package.Save();
            //        }
            //    }
            //}
            //else
            //{
            //    // check if files exist then clear content, otherwise create new files
            //    var files = Directory.GetFiles(exchangeFolder, "*.csv");
            //    if (files.Length == 0)
            //    {
            //        using (var writer2 = new StreamWriter(Path.Combine(exchangeFolder, "1_T_OUT.csv")))
            //        {
            //        }

            //        using (var writer2 = new StreamWriter(Path.Combine(exchangeFolder, "2_S_IN.csv")))
            //        {
            //        }
            //    }
            //    else
            //    {
            //        foreach (var file in Directory.GetFiles(exchangeFolder))
            //        {
            //            try { File.WriteAllText(file, string.Empty); } catch {  }
            //        }

            //    }

            //    string xlsxPath = Path.Combine(exchangeFolder, "1_T_OUT.xlsx");
            //    if (!File.Exists(xlsxPath))
            //    {
            //        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            //        using (var package = new ExcelPackage(new FileInfo(xlsxPath)))
            //        {
            //            package.Workbook.Worksheets.Add("1_T_OUT");
            //            package.Save();
            //        }
            //    }

            //    // trend folder
            //    if (!Directory.Exists(Path.Combine(exchangeFolder, "Trend")))
            //    {
            //        Directory.CreateDirectory(Path.Combine(exchangeFolder, "Trend"));
            //    }
            //    else
            //    {
            //        // delete all files in trend folder
            //        foreach (var file in Directory.GetFiles(Path.Combine(exchangeFolder, "Trend")))
            //        {
            //            try { File.Delete(file); } catch {  }
            //        }
            //    }
            //}



                // navigate to login
                
            navigationService.ShowLoginWindow();
        }
        catch (Exception ex)
        {
#if DEBUG
            throw ex;
#else
            System.Windows.MessageBox.Show("Có lỗi xảy ra. Vui lòng khởi động lại hoặc liên hệ bộ phận hỗ trợ", "Lỗi hệ thống", MessageBoxButton.OK, MessageBoxImage.Error);
            System.Windows.Application.Current.Shutdown();
            return;
#endif
        }
    }
    private void ConfigureExceptionHandling()
    {
        // Handle UI thread exceptions
        AppDomain.CurrentDomain.UnhandledException += (s, exArgs) =>
        {
            if (exArgs.ExceptionObject is Exception ex)
            {
                if (ex is BusinessException)
                {
                    System.Windows.MessageBox.Show(
                    ex.Message,
                    "Có lỗi xảy ra",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                }
                else
                {
#if DEBUG
                    System.Windows.MessageBox.Show(
                        $"Lỗi hệ thống: {ex.Message}\n{ex.StackTrace}",
                        "Lỗi hệ thống",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
#else
                     System.Windows.MessageBox.Show(
                                            $"Lỗi hệ thống: {ex.Message}",
                                            "Lỗi hệ thống",
                                            MessageBoxButton.OK,
                                            MessageBoxImage.Error);
#endif
                }
            }
        };

        // Handle exceptions in async void methods
        TaskScheduler.UnobservedTaskException += (s, exArgs) =>
        {
            if (exArgs.Exception is Exception ex)
            {
                if (ex is BusinessException)
                {
                    System.Windows.MessageBox.Show(
                    ex.Message,
                    "Có lỗi xảy ra",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                    exArgs.SetObserved();
                }
                else
                {
#if DEBUG
                    System.Windows.MessageBox.Show(
                        $"Lỗi hệ thống: {ex.Message}\n{ex.StackTrace}",
                        "Lỗi hệ thống",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
#else
                     System.Windows.MessageBox.Show(
                                            $"Lỗi hệ thống: {ex.Message}",
                                            "Lỗi hệ thống",
                                            MessageBoxButton.OK,
                                            MessageBoxImage.Error);
#endif
                    exArgs.SetObserved();
                }
            }
        };

        // Handle exceptions in WPF bindings
        this.DispatcherUnhandledException += (s, exArgs) =>
        {
            System.Windows.MessageBox.Show(
                $"Lỗi hệ thống: {exArgs.Exception.Message}",
                "Lỗi hệ thống",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            exArgs.Handled = true;
        };
    }

    protected override async void OnExit(ExitEventArgs e)
    {
        try
        {
            var gcService = App.ServiceProvider?.GetService<IGarbageCollectionService>();
            if (gcService != null)
                await gcService.RunOnAppExitAsync();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error during app exit: {ex.Message}");
        }
        finally
        {
            base.OnExit(e);
        }
    }
}

