using Microsoft.EntityFrameworkCore;
using TESMEA_TMS.Configs;
using TESMEA_TMS.DTOs;
using TESMEA_TMS.Helpers;
using TESMEA_TMS.Models.Entities;
using TESMEA_TMS.Models.Infrastructure;

namespace TESMEA_TMS.Services
{
    public interface IAuthenticationService
    {
        Task<bool> LoginAsync(string userName, string password);
        Task<bool> RegisterAsync(UserAccount userAccount);
        Task<bool> ChangePasswordAsync(ChangePasswordDto input);
    }
    public class AuthenticationService : RepositoryBase, IAuthenticationService
    {
        private readonly AppDbContext _dbContext;
        public AuthenticationService(AppDbContext dbContext)
        {
            _dbContext = dbContext;
        }

        private async Task<List<string>> GetPermissionClaimsByRoleAsync(Guid roleId)
        {
            var permissions = await (from rpm in _dbContext.RolePermissions
                                     join p in _dbContext.Permissions on rpm.PermissionId equals p.PermissionId
                                     where rpm.RoleId == roleId && rpm.IsGranted == true
                                     select p)
                                         .OrderBy(x => x.ModuleId)
                                         .OrderBy(x => x.PermissionCode)
                                         .ToListAsync();
            if (!permissions.Any() || permissions == null)
            {
                throw new BusinessException("Role chưa được phân quyền");
            }
            var permissionClaims = new List<string>();
            var modules = await _dbContext.Modules.ToListAsync();
            foreach (var permission in permissions)
            {
                var module = modules.FirstOrDefault(m => m.ModuleId == permission.ModuleId);
                if (module != null && !string.IsNullOrEmpty(permission.PermissionCode))
                {
                    // Tạo claim theo định dạng "ModuleName.PermissionCode"
                    string claim = $"{module.ModuleName}.{permission.PermissionCode}";
                    permissionClaims.Add(claim);
                }
            }
            return permissionClaims;
        }

        public async Task<bool> LoginAsync(string userName, string password)
        {
            try
            {
                var userAccount = await _dbContext.UserAccounts.FirstOrDefaultAsync(x => x.UserName.ToUpper() == userName.ToUpper().Trim());
                if(userAccount == null)
                {
                    throw new BusinessException("Người dùng không tồn tại");
                }
                if (userAccount.PasswordHash.Equals(Encryption.Encrypt(password, true)))
                {
                    var role = await _dbContext.Roles.FirstOrDefaultAsync(x => x.RoleId == userAccount.RoleId);
                    if (role == null)
                    {
                        throw new BusinessException("Vai trò không tồn tại");
                    }
                    var claims = await GetPermissionClaimsByRoleAsync(role.RoleId);
                    var sessionId = Guid.NewGuid();
                    CurrentUser.Instance.SetUser(userAccount, role, claims, sessionId);
                    return true;
                }
                else
                {
                    throw new BusinessException("Thông tin đăng nhập không chính xác");
                }
            }
            catch
            {
                throw new BusinessException("Có lỗi trong quá trình đăng nhập");
            }
        }

        public async Task<bool> RegisterAsync(UserAccount userAccount)
        {
            return await ExecuteAsync(async () =>
            {
                await _dbContext.UserAccounts.AddAsync(userAccount);
                await _dbContext.SaveChangesAsync();
                return true;
            });
        }

        public async Task<bool> ChangePasswordAsync(ChangePasswordDto input)
        {
            try
            {
                return await ExecuteAsync(async () =>
                {
                    var userAccount = await _dbContext.UserAccounts.Where(x => x.UserName.ToLower() == UserSetting.Instance.LastUserName).FirstOrDefaultAsync();
                    if(userAccount == null)
                    {
                        throw new BusinessException("Người dùng không tồn tại");
                    }
                    if (!userAccount.PasswordHash.Equals(Encryption.Encrypt(input.CurrentPassword, true)))
                    {
                        throw new BusinessException("Mật khẩu cũ không đúng");
                    }
                    userAccount.PasswordHash = Encryption.Encrypt(input.NewPassword, true);
                    _dbContext.UserAccounts.Update(userAccount);
                    await _dbContext.SaveChangesAsync();
                    return true;
                });
            }
            catch(Exception ex)
            {
                throw ex;
            }
        }
    }
}
