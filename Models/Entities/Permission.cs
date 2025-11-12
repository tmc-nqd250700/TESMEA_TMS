using System.ComponentModel.DataAnnotations;

namespace TESMEA_TMS.Models.Entities
{
    public class Permission
    {
        [Key]
        public Guid PermissionId { get; set; }
        public Guid ModuleId { get; set; }
        public string PermissionCode { get; set; }
        public string PermissionName { get; set; }
        public string PermissionDesc { get; set; }
        public int IndexOrder { get; set; }
        public Permission()
        {

        }
    }
}
