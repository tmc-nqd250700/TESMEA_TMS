using System.ComponentModel.DataAnnotations;

namespace TESMEA_TMS.Models.Entities
{
    public class Role
    {
        [Key]
        public Guid RoleId { get; set; }
        public string RoleCode { get; set; }
        public string RoleName { get; set; }
        public string RoleDesc { get; set; }
        public int IndexOrder { get; set; } = 1;

        public Role()
        {

        }
    }
}
