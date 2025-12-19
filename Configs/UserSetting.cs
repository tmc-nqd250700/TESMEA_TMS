using System.IO;
using System.Text.Json;

namespace TESMEA_TMS.Configs
{
    public class UserSetting
    {
        public string LastUserName { get; set; }
        public string Language { get; set; } = "vi";
        public string CurrentVersion { get; set; }
        public string LastUpdate { get; set; }
        public string SimaticPath { get; set; } = "";
        public string WinccExePath { get; set; } = "";
        //public string SimaticProjectPath { get; set; } = "";
        public string DbPath { get; set; } = "";
        public string Scada_ReportTemplatePath { get; set; } = "";
        public int TimeoutMilliseconds { get; set; } = 10000;
        #region private methods
        private static string GetUserSetting()
        {
            var folder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "TESMEA_TMS");
            if (!Directory.Exists(folder))
                Directory.CreateDirectory(folder);
            return Path.Combine(folder, "usersetting.json");
        }

        private static UserSetting _instance;
        private static readonly object _lock = new();

        public static UserSetting Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lock)
                    {
                        if (_instance == null)
                            _instance = Load();
                    }
                }
                return _instance;
            }
        }

        public static void Reload()
        {
            lock (_lock)
            {
                _instance = Load();
            }
        }
        #endregion

        #region public methods
        public static UserSetting Load()
        {
            var path = GetUserSetting();
            if (File.Exists(path))
            {
                var json = File.ReadAllText(path);
                return JsonSerializer.Deserialize<UserSetting>(json) ?? new UserSetting();
            }
            return new UserSetting();
        }

        public void Save()
        {
            var path = GetUserSetting();
            var json = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(path, json);
            // Cập nhật lại instance sau khi save
            Reload();
        }

        public static string GetDefaultFilePath()
        {
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "TESMEA_TMS", "usersetting.json");
        }

       

        //public static string GetTestDataFilePath()
        //{
        //    return Path.Combine(GetLocalAppPath(), "TestData");
        //}


        public static string GetLocalAppPath()
        {
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "TESMEA_TMS");
        }

        public static string TOMFAN_folder = @"C:\TOMFAN";
       
        #endregion
    }
}
