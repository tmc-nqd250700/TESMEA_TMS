using System.ComponentModel.DataAnnotations;

namespace TESMEA_TMS.Models.Entities
{
    public class Module
    {
        [Key]
        public Guid ModuleId { get; set; }
        public string ModuleName { get; set; }
        public string Description { get; set; }
        public int IndexOrder { get; set; }
    }
}
