using System.ComponentModel.DataAnnotations;

namespace TESMEA_TMS.Models.Entities
{
    public class RolePermissionMapping
    {
        [Key]
        public Guid MappingId { get; set; }
        public Guid RoleId { get; set; }
        public Guid PermissionId { get; set; }
        public bool IsGranted { get; set; } = false;
    }
}
