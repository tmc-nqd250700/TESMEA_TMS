using System.IO;
using TESMEA_TMS.Configs;

namespace TESMEA_TMS.Services
{
    public interface IGarbageCollectionService
    {
        Task ClearResourcesAsync();
        Task DeleteUserTempAsync();
        Task RunOnAppExitAsync();
    }
    public class GarbageCollectionService : IGarbageCollectionService
    {
        private readonly string _localAppPath;
        private readonly IExternalAppService _externalAppService;
        public GarbageCollectionService(IExternalAppService externalAppService)
        {
            _localAppPath = UserSetting.GetLocalAppPath();
            _externalAppService = externalAppService;
        }

        public async Task ClearResourcesAsync()
        {
            await _externalAppService.StopAppAsync();
            //GC.Collect();
        }

        public async Task DeleteUserTempAsync()
        {
            DeleteAllContents(_localAppPath);
            DeleteAllContents(UserSetting.TOMFAN_folder);
            await Task.CompletedTask;
        }

        private void DeleteAllContents(string folderPath)
        {
            if (Directory.Exists(folderPath))
            {
                // Xóa tất cả file
                foreach (var file in Directory.GetFiles(folderPath))
                {
                    try { File.Delete(file); } catch { /* ignore */ }
                }
                // Xóa tất cả thư mục con
                foreach (var dir in Directory.GetDirectories(folderPath))
                {
                    try { Directory.Delete(dir, true); } catch { /* ignore */ }
                }
            }
        }

        public async Task RunOnAppExitAsync()
        {
            await DeleteUserTempAsync();
            await ClearResourcesAsync();
        }
    }
}
