using Microsoft.EntityFrameworkCore;
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
        Task<List<Library>> GetLibrariesAsync();
        Task<(BienTan, CamBien, OngGio)> GetLibraryByIdAsync(Guid id);
        Task AddLibraryAsync(Library inputParam, BienTan bienTan, CamBien camBien, OngGio ongGio);
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
        public async Task<List<Library>> GetLibrariesAsync()
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

        public async Task AddLibraryAsync(Library lib, BienTan bienTan, CamBien camBien, OngGio ongGio)
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
                   
                _dbContext.Libraries.Update(existLib);
                var exist = await GetLibraryByIdAsync(id);
                if (exist.Item1 != null)
                {
                    _dbContext.BienTans.Update(bienTan);
                }
                if (exist.Item2 != null)
                {
                    _dbContext.CamBiens.Update(camBien);
                }
                if (exist.Item3 != null)
                {
                    _dbContext.OngGios.Update(ongGio);
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
