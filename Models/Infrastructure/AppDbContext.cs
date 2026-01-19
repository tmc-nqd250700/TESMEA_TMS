using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Storage.ValueConversion;
using System.IO;
using System.Linq.Expressions;
using TESMEA_TMS.Configs;
using TESMEA_TMS.Models.Entities;

namespace TESMEA_TMS.Models.Infrastructure
{
    public class AppDbContext : DbContext
    {
        public AppDbContext(DbContextOptions<AppDbContext> options) : base(options)
        {

        }

        public virtual DbSet<Module> Modules { get; set; }
        public virtual DbSet<Permission> Permissions { get; set; }
        public virtual DbSet<Role> Roles { get; set; }
        public virtual DbSet<RolePermissionMapping> RolePermissions { get; set; }
        public virtual DbSet<UserAccount> UserAccounts { get; set; }
        public virtual DbSet<Scenario> Scenarios { get; set; }
        public virtual DbSet<ScenarioParam> ScenarioParams { get; set; }
        public virtual DbSet<Library> Libraries { get; set; }
        public virtual DbSet<BienTan> BienTans { get; set; }
        public virtual DbSet<CamBien> CamBiens { get; set; }
        public virtual DbSet<OngGio> OngGios { get; set; }

        public override Task<int> SaveChangesAsync(CancellationToken cancellationToken = default)
        {
            var fullName = CurrentUser.Instance.UserAccount.FullName;
            var dateNow = DateTime.Now;
            foreach (var entry in ChangeTracker.Entries())
            {
                if (entry.Entity is BaseEntity baseEntity)
                {
                    if (entry.State == EntityState.Added)
                    {
                        baseEntity.CreatedUser = fullName;
                        baseEntity.CreatedDate = dateNow;
                    }
                    else if (entry.State == EntityState.Modified)
                    {
                        baseEntity.ModifiedUser = fullName;
                        baseEntity.ModifiedDate = dateNow;
                    }
                    else if (entry.State == EntityState.Deleted)
                    {
                        entry.State = EntityState.Modified;
                        baseEntity.IsDeleted = true;
                        baseEntity.DeletedUser = fullName;
                        baseEntity.DeletedDate = dateNow;
                    }
                }
            }
            return base.SaveChangesAsync(cancellationToken);
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            foreach (var entityType in modelBuilder.Model.GetEntityTypes())
            {
                // Chỉ áp dụng filter cho các entity kế thừa BaseEntity
                if (typeof(BaseEntity).IsAssignableFrom(entityType.ClrType))
                {
                    var parameter = Expression.Parameter(entityType.ClrType, "e");
                    var isDeletedProperty = Expression.Property(parameter, nameof(BaseEntity.IsDeleted));
                    var isDeletedCondition = Expression.Equal(isDeletedProperty, Expression.Constant(false));
                    var lambda = Expression.Lambda(isDeletedCondition, parameter);

                    modelBuilder.Entity(entityType.ClrType).HasQueryFilter(lambda);
                }
            }
            var boolToIntConverter = new BoolToZeroOneConverter<int>();

            base.OnModelCreating(modelBuilder); 
            modelBuilder.Entity<UserAccount>().ToTable("tbl_UserAccount");
            modelBuilder.Entity<Role>().ToTable("tbl_Role");
            modelBuilder.Entity<Module>().ToTable("tbl_Module");
            modelBuilder.Entity<Permission>().ToTable("tbl_Permission");
            modelBuilder.Entity<RolePermissionMapping>().ToTable("tbl_RolePermissionMapping");

            modelBuilder.Entity<Scenario>().ToTable("tbl_Scenario");
            modelBuilder.Entity<ScenarioParam>().ToTable("tbl_ScenarioParam");
            modelBuilder.Entity<Library>().ToTable("tbl_Library");
            modelBuilder.Entity<BienTan>().ToTable("tbl_BienTan");
            modelBuilder.Entity<OngGio>().ToTable("tbl_OngGio");

            modelBuilder.Entity<CamBien>().ToTable("tbl_CamBien");
            modelBuilder.Entity<CamBien>(entity =>
            {
                entity.Property(e => e.IsImportNhietDoMoiTruong).HasConversion(boolToIntConverter);
                entity.Property(e => e.IsImportDoAmMoiTruong).HasConversion(boolToIntConverter);
                entity.Property(e => e.IsImportApSuatKhiQuyen).HasConversion(boolToIntConverter);
                entity.Property(e => e.IsImportChenhLechApSuat).HasConversion(boolToIntConverter);
                entity.Property(e => e.IsImportApSuatTinh).HasConversion(boolToIntConverter);
                entity.Property(e => e.IsImportDoRung).HasConversion(boolToIntConverter);
                entity.Property(e => e.IsImportDoOn).HasConversion(boolToIntConverter);
                entity.Property(e => e.IsImportSoVongQuay).HasConversion(boolToIntConverter);
                entity.Property(e => e.IsImportMomen).HasConversion(boolToIntConverter);
                entity.Property(e => e.IsImportPhanHoiDongDien).HasConversion(boolToIntConverter);
                entity.Property(e => e.IsImportPhanHoiCongSuat).HasConversion(boolToIntConverter);
                entity.Property(e => e.IsImportPhanHoiViTriVan).HasConversion(boolToIntConverter);
                entity.Property(e => e.IsImportPhanHoiDienAp).HasConversion(boolToIntConverter);
                entity.Property(e => e.IsImportNhietDoGoiTruc).HasConversion(boolToIntConverter);
                entity.Property(e => e.IsImportNhietDoBauKho).HasConversion(boolToIntConverter);
                entity.Property(e => e.IsImportCamBienLuuLuong).HasConversion(boolToIntConverter);
                entity.Property(e => e.IsImportPhanHoiTanSo).HasConversion(boolToIntConverter);
            });
        }

        //protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        //{
        //    if (!optionsBuilder.IsConfigured)
        //    {
        //        var dbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tesmea_tms.db");
        //        var connectionString = $"Data Source={dbPath}";
        //        optionsBuilder.UseSqlite(connectionString);
        //    }

        //    // Kiểm tra kết nối và schema
        //    try
        //    {
        //        using var connection = new Microsoft.Data.Sqlite.SqliteConnection(optionsBuilder.Options.FindExtension<Microsoft.EntityFrameworkCore.Sqlite.Infrastructure.Internal.SqliteOptionsExtension>()?.ConnectionString);
        //        connection.Open();

        //        // Kiểm tra một bảng quan trọng, ví dụ: tbl_UserAccount
        //        using var cmd = connection.CreateCommand();
        //        cmd.CommandText = "SELECT name FROM sqlite_master WHERE type='table' AND name='tbl_UserAccount';";
        //        var result = cmd.ExecuteScalar();
        //        if (result == null)
        //        {
        //            throw new InvalidOperationException("Database schema không hợp lệ");
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        throw new InvalidOperationException("Không thể kết nối hoặc kiểm tra schema database.", ex);
        //    }
        //}
    }
}
