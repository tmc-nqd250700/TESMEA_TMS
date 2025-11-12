using System.ComponentModel.DataAnnotations.Schema;
namespace TESMEA_TMS.Models.Entities
{
    public abstract class BaseEntity
    {
        [Column(TypeName = "datetime")]
        public DateTime CreatedDate { get; set; }
        public string CreatedUser { get; set; }
        [Column(TypeName = "datetime")]
        public DateTime? ModifiedDate { get; set; }
        public string? ModifiedUser { get; set; }
        public bool IsDeleted { get; set; } = false;
        [Column(TypeName = "datetime")]
        public DateTime? DeletedDate { get; set; }
        public string? DeletedUser { get; set; }
    }
}
