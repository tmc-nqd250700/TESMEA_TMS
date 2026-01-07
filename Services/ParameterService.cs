using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.DependencyModel;
using TESMEA_TMS.DTOs;
using TESMEA_TMS.Helpers;
using TESMEA_TMS.Models.Entities;
using TESMEA_TMS.Models.Infrastructure;

namespace TESMEA_TMS.Services
{
    // Service dùng chung cho các entity thông số: scenario, library, ...
    public interface IParameterService
    {
        #region Scenario
        Task<List<Scenario>> GetScenariosAsync();
        Task<Scenario> GetScenarioAsync(string id);
        Task<List<ScenarioParam>> GetScenarioDetailAsync(Guid id);
        Task UpdateScenarioAsync(ScenarioUpdateDto scenarioToUpdate);
        Task DeleteScenarioAsync(Guid id);
        #endregion

        #region library gồm thông số biến tần, ống gió và cảm biến
        Task<List<Models.Entities.Library>> GetLibrariesAsync();
        Task<(BienTan, CamBien, OngGio)> GetLibraryByIdAsync(Guid id);
        Task AddLibraryAsync(Models.Entities.Library inputParam, BienTan bienTan, CamBien camBien, OngGio ongGio);
        Task UpdateLibraryAsync(Guid libId, BienTan bienTan, CamBien camBien, OngGio ongGio);
        Task DeleteLibraryAsync(Guid libId);
        #endregion
    }
    public class ParameterService : RepositoryBase, IParameterService
    {
        private readonly AppDbContext _dbContext;
        public ParameterService(AppDbContext dbContext)
        {
            _dbContext = dbContext;
        }


        // tạm: vì sqlite không thao tác được với guid nên tạm thời convert sang string để tìm kiếm
        #region Scneario implementation
        public async Task<List<Scenario>> GetScenariosAsync()
        {
            return await ExecuteAsync(async () =>
            {
                return await _dbContext.Scenarios.ToListAsync();
            });
        }

        public async Task<Scenario> GetScenarioAsync(string id)
        {
            return await ExecuteAsync(async () =>
            {
                return await _dbContext.Scenarios.Where(x => x.ScenarioId.ToString() == id).FirstOrDefaultAsync();
            });
        }

        public async Task<List<ScenarioParam>> GetScenarioDetailAsync(Guid id)
        {
            return await ExecuteAsync(async () =>
            {
                
                return await _dbContext.ScenarioParams.Where(x => x.ScenarioId == id).ToListAsync();
            });
        }

        public async Task UpdateScenarioAsync(ScenarioUpdateDto scenario)
        {
            await ExecuteAsync(async () =>
            {
                var exist = await _dbContext.Scenarios.Where(x => x.ScenarioId == scenario.Scenario.ScenarioId).FirstOrDefaultAsync();
                if (exist == null)
                {
                    var paramsToAdd = scenario.Params.Select(p => p.ToEntity()).ToList();
                    if (paramsToAdd.Count == 0)
                    {
                        throw new BusinessException($"Danh sách thông số kịch bản không được để trống");
                    }
                    if (string.IsNullOrEmpty(scenario.Scenario.ScenarioName))
                    {
                        throw new BusinessException("Tên kịch bản không được để trống");
                    }
                    foreach (var item in paramsToAdd) item.ScenarioId = scenario.Scenario.ScenarioId;
                    await _dbContext.Scenarios.AddAsync(scenario.Scenario.ToEntity());
                    await _dbContext.ScenarioParams.AddRangeAsync(paramsToAdd);
                }
                else
                {
                    var paramsToUpdate = scenario.Params.Select(p => p.ToEntity()).Where(x => x.ScenarioId == scenario.Scenario.ScenarioId).ToList();
                    if (paramsToUpdate.Count == 0)
                    {
                        throw new BusinessException($"Danh sách thông số kịch bản không được để trống");
                    }
                    // Xóa scenario đã tồn tại
                    if (scenario.Scenario.IsMarkedForDeletion && !scenario.Scenario.IsNew)
                    {
                        _dbContext.Scenarios.Remove(exist);
                    }
                    // Cập nhật
                    else
                    {
                        var existParams = await _dbContext.ScenarioParams.Where(x => x.ScenarioId == scenario.Scenario.ScenarioId).ToListAsync();
                        _dbContext.ScenarioParams.RemoveRange(existParams);

                        exist.ScenarioName = scenario.Scenario.ScenarioName;
                        exist.StandardDeviation = scenario.Scenario.StandardDeviation;
                        exist.TimeRange = scenario.Scenario.TimeRange;
                        _dbContext.Scenarios.Update(exist);
                        _dbContext.ScenarioParams.UpdateRange(paramsToUpdate);
                    }
                }
                
                await _dbContext.SaveChangesAsync();
            });
        }

        public async Task DeleteScenarioAsync(Guid id)
        {
            await ExecuteAsync(async () =>
            {
                var exist = await _dbContext.Scenarios.Where(x => x.ScenarioId == id).FirstOrDefaultAsync();
                if(exist != null)
                {
                    _dbContext.Scenarios.Remove(exist);
                    await _dbContext.SaveChangesAsync();
                }
            });
        }
        #endregion

        #region thông số đầu vào
        public async Task<List<Models.Entities.Library>> GetLibrariesAsync()
        {
            return await ExecuteAsync(async () =>
            {
                return await _dbContext.Libraries.ToListAsync();
            });
        }

        public async Task<(BienTan, CamBien, OngGio)> GetLibraryByIdAsync(Guid id)
        {
            try
            {
                return await ExecuteAsync(async () =>
                {
                    var bienTan = await _dbContext.BienTans.FirstOrDefaultAsync(x=>x.LibId == id);
                    var camBien = await _dbContext.CamBiens.FirstOrDefaultAsync(x=>x.LibId == id);
                    var ongGio = await _dbContext.OngGios.FirstOrDefaultAsync(x=>x.LibId == id);
                    return (bienTan, camBien, ongGio);
                });
            }
            catch(Exception ex)
            {
                return (null, null, null);
            }
        }

        public async Task AddLibraryAsync(Models.Entities.Library lib, BienTan bienTan, CamBien camBien, OngGio ongGio)
        {
            await ExecuteAsync(async () =>
            {
                if(await _dbContext.Libraries.Where(x=>x.LibName.ToLower().Equals(lib.LibName.ToLower())).FirstOrDefaultAsync() != null)
                {
                    throw new BusinessException("Tên library đã tồn tại");
                }
                await _dbContext.Libraries.AddAsync(lib);
                await _dbContext.BienTans.AddAsync(bienTan);
                await _dbContext.CamBiens.AddAsync(camBien);
                await _dbContext.OngGios.AddAsync(ongGio);
                await _dbContext.SaveChangesAsync();

            });
        }
        public async Task UpdateLibraryAsync(Guid id, BienTan bienTan, CamBien camBien, OngGio ongGio)
        {
            await ExecuteAsync(async () =>
            {
                var existLib = await _dbContext.Libraries.Where(x => x.LibId == id).FirstOrDefaultAsync();
                if(existLib == null)
                {
                    throw new BusinessException("Library không tồn tại");
                }

                existLib.ModifiedDate = DateTime.Now;
                existLib.ModifiedUser = Environment.UserName;
                _dbContext.Libraries.Update(existLib);
                var exist = await GetLibraryByIdAsync(id);
                if (exist.Item1 != null)
                {
                    exist.Item1.DienApRa = bienTan.DienApVao;
                    exist.Item1.DongDienVao = bienTan.DongDienVao;
                    exist.Item1.TanSoVao = bienTan.TanSoVao;
                    exist.Item1.CongSuatVao = bienTan.CongSuatVao;
                    exist.Item1.DienApRa = bienTan.DienApRa;
                    exist.Item1.DongDienRa = bienTan.DongDienRa;
                    exist.Item1.TanSoRa = bienTan.TanSoRa;
                    exist.Item1.CongSuatTongRa = bienTan.CongSuatTongRa;
                    exist.Item1.CongSuatHieuDungRa = bienTan.CongSuatHieuDungRa;
                    exist.Item1.HieuSuatBoTruyen = bienTan.HieuSuatBoTruyen;
                    exist.Item1.HieuSuatNoiTruc = bienTan.HieuSuatNoiTruc;
                    exist.Item1.HieuSuatGoiTruc = bienTan.HieuSuatGoiTruc;
                    _dbContext.BienTans.Update(exist.Item1);
                }
                if (exist.Item2 != null)
                {
                    exist.Item2.NhietDoMoiTruongMin = camBien.NhietDoMoiTruongMin;
                    exist.Item2.NhietDoMoiTruongMax = camBien.NhietDoMoiTruongMax;
                    exist.Item2.DoAmMoiTruongMin = camBien.DoAmMoiTruongMin;
                    exist.Item2.DoAmMoiTruongMax = camBien.DoAmMoiTruongMax;
                    exist.Item2.ApSuatKhiQuyenMin = camBien.ApSuatKhiQuyenMin;
                    exist.Item2.ApSuatKhiQuyenMax = camBien.ApSuatKhiQuyenMax;
                    exist.Item2.ChenhLechApSuatMin = camBien.ChenhLechApSuatMin;
                    exist.Item2.ChenhLechApSuatMax = camBien.ChenhLechApSuatMax;
                    exist.Item2.ApSuatTinhMin = camBien.ApSuatTinhMin;
                    exist.Item2.ApSuatTinhMax = camBien.ApSuatTinhMax;
                    exist.Item2.DoRungMin = camBien.DoRungMin;
                    exist.Item2.DoRungMax = camBien.DoRungMax;
                    exist.Item2.DoOnMin = camBien.DoOnMin;
                    exist.Item2.DoOnMax = camBien.DoOnMax;
                    exist.Item2.SoVongQuayMin = camBien.SoVongQuayMin;
                    exist.Item2.SoVongQuayMax = camBien.SoVongQuayMax;
                    exist.Item2.MomenMin = camBien.MomenMin;
                    exist.Item2.MomenMax = camBien.MomenMax;
                    exist.Item2.PhanHoiDongDienMin = camBien.PhanHoiDongDienMin;
                    exist.Item2.PhanHoiDongDienMax = camBien.PhanHoiDongDienMax;
                    exist.Item2.PhanHoiCongSuatMin = camBien.PhanHoiCongSuatMin;
                    exist.Item2.PhanHoiCongSuatMax = camBien.PhanHoiCongSuatMax;
                    exist.Item2.PhanHoiViTriVanMin = camBien.PhanHoiViTriVanMin;
                    exist.Item2.PhanHoiViTriVanMax = camBien.PhanHoiViTriVanMax;
                    exist.Item2.PhanHoiDienApMin = camBien.PhanHoiDienApMin;
                    exist.Item2.PhanHoiDienApMax = camBien.PhanHoiDienApMax;
                    exist.Item2.NhietDoGoiTrucMin = camBien.NhietDoGoiTrucMin;
                    exist.Item2.NhietDoGoiTrucMax = camBien.NhietDoGoiTrucMax;
                    exist.Item2.NhietDoBauKhoMin = camBien.NhietDoBauKhoMin;
                    exist.Item2.NhietDoBauKhoMax = camBien.NhietDoBauKhoMax;
                    exist.Item2.CamBienLuuLuongMin = camBien.CamBienLuuLuongMin;
                    exist.Item2.CamBienLuuLuongMax = camBien.CamBienLuuLuongMax;
                    exist.Item2.PhanHoiTanSoMin = camBien.PhanHoiTanSoMin;
                    exist.Item2.PhanHoiTanSoMax = camBien.PhanHoiTanSoMax;
                    _dbContext.CamBiens.Update(exist.Item2);
                }
                if (exist.Item3 != null)
                {
                    exist.Item3.DuongKinhOngGio = ongGio.DuongKinhOngGio;
                    exist.Item3.ChieuDaiOngGioSauQuat = ongGio.ChieuDaiOngGioSauQuat;
                    exist.Item3.ChieuDaiOngGioTruocQuat = ongGio.ChieuDaiOngGioTruocQuat;
                    exist.Item3.DuongKinhLoPhut = ongGio.DuongKinhLoPhut;
                    exist.Item3.DuongKinhMiengQuat = ongGio.DuongKinhMiengQuat;
                    exist.Item3.ChieuDaiConQuat = ongGio.ChieuDaiConQuat;
                    _dbContext.OngGios.Update(exist.Item3);
                }
                await _dbContext.SaveChangesAsync();

            });
        }
        public async Task DeleteLibraryAsync(Guid libId)
        {
            await ExecuteAsync(async () =>
            {
                var exist = await _dbContext.Libraries.Where(x => x.LibId == libId).FirstOrDefaultAsync();
                if(exist != null)
                {
                    _dbContext.Libraries.Remove(exist);
                    await _dbContext.SaveChangesAsync();
                }    
            });
        }
        #endregion
    }
}
