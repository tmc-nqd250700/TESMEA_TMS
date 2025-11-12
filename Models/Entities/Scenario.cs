using System.ComponentModel.DataAnnotations;

namespace TESMEA_TMS.Models.Entities
{
    public class Scenario : BaseEntity
    {
        [Key]
        public Guid ScenarioId { get; set; }
        public string ScenarioName { get; set; }
        public float StandardDeviation { get; set; } // độ lệch chuẩn (max/min-1)*100 < %STable của vùng để cho là ổn định
        public float TimeRange { get; set; } // khoảng thời gian (s) để tính ổn định
    }

    public class ScenarioParam
    {
        [Key]
        public Guid ParamId { get; set; }
        public int STT { get; set; }
        public Guid ScenarioId { get; set; }
        public float S { get; set; }
        public float CV { get; set; }
    }
}
