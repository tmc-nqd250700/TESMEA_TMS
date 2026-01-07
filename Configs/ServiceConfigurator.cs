using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Configuration;
using Microsoft.EntityFrameworkCore;
using TESMEA_TMS.Models.Infrastructure;
using TESMEA_TMS.ViewModels;
using TESMEA_TMS.Services;
using System.IO;

namespace TESMEA_TMS.Configs
{
    public static class ServiceConfigurator
    {
        /// <summary>
        /// Config and register services into the DI container
        /// </summary>
        /// <returns></returns>
        public static IServiceProvider Configure()
        {
            var services = new ServiceCollection();
            var optionsBuilder = new DbContextOptionsBuilder<AppDbContext>();
            // Đăng ký cấu hình
            var configuration = new ConfigurationBuilder()
                .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
                //.AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .Build();
            services.AddSingleton<IConfiguration>(configuration);

            //SQLitePCL.Batteries_V2.Init(); // sqlcipher
            // Đăng ký DbContext
            //var connectionString = configuration.GetConnectionString("Default");
            //if (string.IsNullOrEmpty(connectionString))
            //{
            //    throw new InvalidOperationException("Connection string not found.");
            //}
            //optionsBuilder.UseSqlite(connectionString);

            services.AddDbContext<AppDbContext>(options =>
            {
                //options.UseSqlite(configuration.GetConnectionString("Default"));
                //var dbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Db", "tesmea_tms.db");
                //var connectionString = $"Data Source={dbPath}";
                //options.UseSqlite(connectionString);




                var dataFolder = Path.Combine(
    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
    "TESMEA_TMS", "Db");
                Directory.CreateDirectory(dataFolder);

                var dbPath = Path.Combine(dataFolder, "tesmea_tms.db");
                if (!File.Exists(dbPath))
                {
                    var sourceDb = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Db", "tesmea_tms.db");
                    File.Copy(sourceDb, dbPath);
                }
                var connectionString = $"Data Source={dbPath}";
                options.UseSqlite(connectionString);
            }, ServiceLifetime.Transient);

            services.AddMemoryCache();

            // Đăng ký service
            services.AddSingleton<ICurrentUser, CurrentUser>();
            services.AddSingleton<IAppNavigationService, AppNavigationService>();
            services.AddSingleton<IAuthenticationService, AuthenticationService>();
            services.AddSingleton<ICalculationService, CalculationService>();
            services.AddSingleton<IExternalAppService, ExternalAppService>();
            services.AddSingleton<IFileService, FileService>();
            services.AddSingleton<IGarbageCollectionService, GarbageCollectionService>();
            services.AddSingleton<IParameterService, ParameterService>();

            // viewmodels
            services.AddSingleton<ViewModelLocator>();
            services.AddSingleton<LoginViewModel>();
            services.AddSingleton<ChooseExternalAppViewModel>();
            services.AddSingleton<MainViewModel>();
            services.AddSingleton<ProjectViewModel>();
            services.AddSingleton<ProgressSplashViewModel>();
            services.AddSingleton<SettingViewModel>();
            services.AddSingleton<LibraryViewModel>();
            services.AddSingleton<MeasureViewModel>();
            services.AddSingleton<CalculationViewModel>();
            services.AddSingleton<ScenarioViewModel>();

            return services.BuildServiceProvider();
        }
    }
}
