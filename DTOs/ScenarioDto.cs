using System.ComponentModel;
using System.Runtime.CompilerServices;
using TESMEA_TMS.Models.Entities;
namespace TESMEA_TMS.DTOs
{
    public class ScenarioDto : INotifyPropertyChanged
    {
        private bool _isNew;
        private bool _isEdited;
        private bool _isMarkedForDeletion;

        public Guid ScenarioId { get; set; }
        public string ScenarioName { get; set; }
        public float StandardDeviation { get; set; }
        public float TimeRange { get; set; }
        public DateTime CreatedDate { get; set; }
        public string CreatedUser { get; set; }
        public DateTime? ModifiedDate { get; set; }
        public string ModifiedUser { get; set; }

        /// <summary>
        /// Đánh dấu kịch bản mới tạo (chưa lưu vào DB)
        /// </summary>
        public bool IsNew
        {
            get => _isNew;
            set
            {
                _isNew = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Đánh dấu kịch bản đã chỉnh sửa (đã tồn tại nhưng có thay đổi)
        /// </summary>
        public bool IsEdited
        {
            get => _isEdited;
            set
            {
                _isEdited = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Đánh dấu kịch bản đang chờ xóa (có thể hoàn tác)
        /// </summary>
        public bool IsMarkedForDeletion
        {
            get => _isMarkedForDeletion;
            set
            {
                _isMarkedForDeletion = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Kiểm tra có thể chỉnh sửa hay không
        /// </summary>
        public bool CanEdit => IsNew;

        /// <summary>
        /// Chuyển từ Entity sang DTO
        /// </summary>
        public static ScenarioDto FromEntity(Scenario entity, bool isNew = false)
        {
            return new ScenarioDto
            {
                ScenarioId = entity.ScenarioId,
                ScenarioName = entity.ScenarioName,
                StandardDeviation = entity.StandardDeviation,
                TimeRange = entity.TimeRange,
                CreatedDate = entity.CreatedDate,
                CreatedUser = entity.CreatedUser,
                ModifiedDate = entity.ModifiedDate,
                ModifiedUser = entity.ModifiedUser,
                IsNew = isNew,
                IsEdited = false,
                IsMarkedForDeletion = false
            };
        }

        /// <summary>
        /// Chuyển từ DTO sang Entity
        /// </summary>
        public Scenario ToEntity()
        {
            return new Scenario
            {
                ScenarioId = ScenarioId,
                ScenarioName = ScenarioName,
                StandardDeviation = StandardDeviation,
                TimeRange = TimeRange,
                CreatedDate = CreatedDate,
                CreatedUser = CreatedUser,
                ModifiedDate = ModifiedDate,
                ModifiedUser = ModifiedUser
            };
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    /// <summary>
    /// DTO cho ScenarioParam với STT tự động
    /// </summary>
    public class ScenarioParamDTO : INotifyPropertyChanged
    {
        private int _stt;
        private float _s;
        private float _cv;
        private bool _isNew;
        private bool _isEdited;

        public Guid ScenarioId { get; set; }

        /// <summary>
        /// STT tự động tăng
        /// </summary>
        public int STT
        {
            get => _stt;
            set
            {
                _stt = value;
                OnPropertyChanged();
            }
        }

        public float S
        {
            get => _s;
            set
            {
                if (_s != value)
                {
                    _s = value;
                    MarkAsEdited();
                    OnPropertyChanged();
                }
            }
        }

        public float CV
        {
            get => _cv;
            set
            {
                if (_cv != value)
                {
                    _cv = value;
                    MarkAsEdited();
                    OnPropertyChanged();
                }
            }
        }

        public bool IsNew
        {
            get => _isNew;
            set
            {
                _isNew = value;
                OnPropertyChanged();
            }
        }

        public bool IsEdited
        {
            get => _isEdited;
            set
            {
                _isEdited = value;
                OnPropertyChanged();
            }
        }

        private void MarkAsEdited()
        {
            if (!IsNew)
            {
                IsEdited = true;
            }
        }

        public static ScenarioParamDTO FromEntity(ScenarioParam entity, int stt, bool isNew = false)
        {
            return new ScenarioParamDTO
            {
                ScenarioId = entity.ScenarioId,
                STT = stt,
                S = entity.S,
                CV = entity.CV,
                IsNew = isNew,
                IsEdited = false
            };
        }

        public ScenarioParam ToEntity()
        {
            return new ScenarioParam
            {
                ScenarioId = ScenarioId,
                STT = STT,
                S = S,
                CV = CV
            };
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class ScenarioUpdateDto
    {
        public ScenarioDto Scenario { get; set; }
        public List<ScenarioParamDTO> Params { get; set; }
    }
}
