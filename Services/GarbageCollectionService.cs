using DocumentFormat.OpenXml.Packaging;
using OfficeOpenXml;
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
            //// delete local app folders, files
            //DeleteAllContents(_localAppPath);

            //// delete trend folder
            //ClearFileContents(UserSetting.TOMFAN_folder);
            //foreach (var file in Directory.GetFiles(Path.Combine(UserSetting.TOMFAN_folder, "Trend")))
            //{
            //    try { File.Delete(file); } catch {  }
            //}
            //await Task.CompletedTask;
        }

        private void DeleteAllContents(string folderPath)
        {
            if (Directory.Exists(folderPath))
            {
                // Xóa tất cả file
                foreach (var file in Directory.GetFiles(folderPath))
                {
                    try { File.Delete(file); } catch {  }
                }
                // Xóa tất cả thư mục con
                foreach (var dir in Directory.GetDirectories(folderPath))
                {
                    try { Directory.Delete(dir, true); } catch {  }
                }
            }
        }

        public void ClearFileContents(string folderPath)
        {
            if (Directory.Exists(folderPath))
            {
                foreach (var file in Directory.GetFiles(folderPath))
                {
                    try
                    {
                        var extension = Path.GetExtension(file).ToLower();
                        switch (extension)
                        {
                            case ".csv":
                                File.WriteAllText(file, string.Empty);
                                break;

                            case ".xlsx":
                                ClearExcelFile(file);
                                break;

                            case ".docx":
                                ClearWordFile(file);
                                break;

                            default:
                                System.Diagnostics.Debug.WriteLine($"Unsupported file format: {extension}");
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Failed to clear content of file {file}: {ex.Message}");
                    }
                }
            }
        }

        private void ClearExcelFile(string filePath)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                foreach (var worksheet in package.Workbook.Worksheets)
                {
                    worksheet.Cells.Clear(); // Clear all cells in the worksheet
                }
                package.Save();
            }
        }

        private void ClearWordFile(string filePath)
        {
            using (var wordDoc = WordprocessingDocument.Open(filePath, true))
            {
                var body = wordDoc.MainDocumentPart.Document.Body;
                if (body != null)
                {
                    body.RemoveAllChildren(); // Remove all content from the document body
                    wordDoc.MainDocumentPart.Document.Save();
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
