using Microsoft.EntityFrameworkCore;
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
        Task UpdateScenarioAsync(List<Scenario> scenarios, List<ScenarioParam> scenarioParams);
        Task DeleteScenarioAsync(Guid id);
        #endregion

        #region library gồm thông số biến tần, ống gió và cảm biến
        Task<List<Library>> GetLibrariesAsync();
        Task<CamBien> GetCamBienByIdAsync(string id);
        Task<(BienTan, CamBien, OngGio)> GetLibraryByIdAsync(Guid id);
        Task AddInputParamAsync(Library inputParam, BienTan bienTan, CamBien camBien, OngGio ongGio);
        Task UpdateInputParamAsync(Guid paramId, BienTan bienTan, CamBien camBien, OngGio ongGio);
        Task DeleteInputParamAsync(Guid paramId);
        #endregion
    }
    public class ParameterService : RepositoryBase, IParameterService
    {
        private readonly AppDbContext _dbContext;
        public ParameterService(AppDbContext dbContext)
        {
            _dbContext = dbContext;
        }

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
                // tạm: vì sqlite không tìm được với guid
                return await _dbContext.ScenarioParams.Where(x => x.ScenarioId.ToString() == id.ToString()).ToListAsync();
            });
        }

        public async Task UpdateScenarioAsync(List<Scenario> scenarios, List<ScenarioParam> scenarioParams)
        {
            await ExecuteAsync(async () =>
            {
                foreach(var item in scenarios)
                {
                    var exist = await _dbContext.Scenarios.FindAsync(item.ScenarioId);
                    if (exist == null)
                    {
                        // Tạo mới
                        await _dbContext.Scenarios.AddAsync(item);
                    }
                    else
                    {
                        // Cập nhật
                        _dbContext.Entry(exist).CurrentValues.SetValues(item);
                    }
                }
                await _dbContext.SaveChangesAsync();
            });
        }

        public async Task DeleteScenarioAsync(Guid id)
        {
            await ExecuteAsync(async () =>
            {
                _dbContext.Scenarios.Remove(await _dbContext.Scenarios.FindAsync(id));
                await _dbContext.SaveChangesAsync();
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

        public async Task<CamBien> GetCamBienByIdAsync(string libId)
        {
            try
            {
                return await ExecuteAsync(async () =>
                {
                    var camBien = await _dbContext.CamBiens.FirstOrDefaultAsync(x => x.LibId.ToString() == libId);
                    return camBien;
                });
            }
            catch (Exception ex)
            {
                return null;
            }
        }
        public async Task<(BienTan, CamBien, OngGio)> GetLibraryByIdAsync(Guid id)
        {
            try
            {
                return await ExecuteAsync(async () =>
                {
                    var bienTan = await _dbContext.BienTans.FirstOrDefaultAsync(x=>x.LibId.ToString() == id.ToString());
                    var camBien = await _dbContext.CamBiens.FirstOrDefaultAsync(x=>x.LibId.ToString() == id.ToString());
                    var ongGio = await _dbContext.OngGios.FirstOrDefaultAsync(x=>x.LibId.ToString() == id.ToString());
                    return (bienTan, camBien, ongGio);
                });
            }
            catch(Exception ex)
            {
                return (null, null, null);
            }
        }

        public async Task AddInputParamAsync(Library lib, BienTan bienTan, CamBien camBien, OngGio ongGio)
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

            });
        }
        public async Task UpdateInputParamAsync(Guid id, BienTan bienTan, CamBien camBien, OngGio ongGio)
        {
            await ExecuteAsync(async () =>
            {
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
        public async Task DeleteInputParamAsync(Guid paramId)
        {
            await ExecuteAsync(async () =>
            {
                _dbContext.Libraries.Remove(await _dbContext.Libraries.FindAsync(paramId));
                await _dbContext.SaveChangesAsync();
            });
        }
        #endregion
    }
}
