using System.ComponentModel;
using System.Runtime.CompilerServices;
using TESMEA_TMS.Models.Entities;

namespace TESMEA_TMS.DTOs
{
    public class LibraryDto : INotifyPropertyChanged
    {
        private bool _isNew;
        private bool _isEdited;
        private bool _isMarkedForDeletion;
        private string _paramName;

        public Guid LibId { get; set; }

        public string LibName
        {
            get => _paramName;
            set
            {
                if (_paramName != value)
                {
                    _paramName = value;
                    if (!IsNew)
                    {
                        IsEdited = true;
                    }
                    OnPropertyChanged();
                }
            }
        }

        public DateTime CreatedDate { get; set; }
        public string CreatedUser { get; set; }
        public DateTime? ModifiedDate { get; set; }
        public string ModifiedUser { get; set; }

        /// <summary>
        /// Đánh dấu library mới tạo (chưa lưu vào DB)
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
        /// Đánh dấu library đã chỉnh sửa (đã tồn tại nhưng có thay đổi)
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
        /// Đánh dấu library đang chờ xóa (có thể hoàn tác)
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

        public static LibraryDto FromEntity(Library entity, bool isNew = false)
        {
            return new LibraryDto
            {
                LibId = entity.LibId,
                LibName = entity.LibName,
                CreatedDate = entity.CreatedDate,
                CreatedUser = entity.CreatedUser,
                ModifiedDate = entity.ModifiedDate,
                ModifiedUser = entity.ModifiedUser,
                IsNew = isNew,
                IsEdited = false,
                IsMarkedForDeletion = false
            };
        }

        public Library ToEntity()
        {
            return new Library
            {
                LibId = LibId,
                LibName = LibName,
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
    /// DTO cho BienTan với tracking thay đổi
    /// </summary>
    public class BienTanDto : INotifyPropertyChanged
    {
        private float _dienApVao;
        private float _dongDienVao;
        private float _tanSoVao;
        private float _congSuatVao;
        private float _dienApRa;
        private float _dongDienRa;
        private float _tanSoRa;
        private float _congSuatTongRa;
        private float _congSuatHieuDungRa;
        private float _hieuSuatBoTruyen;
        private float _hieuSuatNoiTruc;
        private float _hieuSuatGoiTruc;

        public Guid LibId { get; set; }

        public float DienApVao
        {
            get => _dienApVao;
            set { _dienApVao = value; OnPropertyChanged(); NotifyChanged(); }
        }

        public float DongDienVao
        {
            get => _dongDienVao;
            set { _dongDienVao = value; OnPropertyChanged(); NotifyChanged(); }
        }

        public float TanSoVao
        {
            get => _tanSoVao;
            set { _tanSoVao = value; OnPropertyChanged(); NotifyChanged(); }
        }

        public float CongSuatVao
        {
            get => _congSuatVao;
            set { _congSuatVao = value; OnPropertyChanged(); NotifyChanged(); }
        }

        public float DienApRa
        {
            get => _dienApRa;
            set { _dienApRa = value; OnPropertyChanged(); NotifyChanged(); }
        }

        public float DongDienRa
        {
            get => _dongDienRa;
            set { _dongDienRa = value; OnPropertyChanged(); NotifyChanged(); }
        }

        public float TanSoRa
        {
            get => _tanSoRa;
            set { _tanSoRa = value; OnPropertyChanged(); NotifyChanged(); }
        }

        public float CongSuatTongRa
        {
            get => _congSuatTongRa;
            set { _congSuatTongRa = value; OnPropertyChanged(); NotifyChanged(); }
        }

        public float CongSuatHieuDungRa
        {
            get => _congSuatHieuDungRa;
            set { _congSuatHieuDungRa = value; OnPropertyChanged(); NotifyChanged(); }
        }

        public float HieuSuatBoTruyen
        {
            get => _hieuSuatBoTruyen;
            set { _hieuSuatBoTruyen = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float HieuSuatNoiTruc
        {
            get => _hieuSuatNoiTruc;
            set { _hieuSuatNoiTruc = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float HieuSuatGoiTruc
        {
            get => _hieuSuatGoiTruc;
            set { _hieuSuatGoiTruc = value; OnPropertyChanged(); NotifyChanged(); }
        }

        public event EventHandler DataChanged;
        private void NotifyChanged()
        {
            DataChanged?.Invoke(this, EventArgs.Empty);
        }

        public static BienTanDto FromEntity(BienTan entity)
        {
            return new BienTanDto
            {
                LibId = entity.LibId,
                DienApVao = entity.DienApVao,
                DongDienVao = entity.DongDienVao,
                TanSoVao = entity.TanSoVao,
                CongSuatVao = entity.CongSuatVao,
                DienApRa = entity.DienApRa,
                DongDienRa = entity.DongDienRa,
                TanSoRa = entity.TanSoRa,
                CongSuatTongRa = entity.CongSuatTongRa,
                CongSuatHieuDungRa = entity.CongSuatHieuDungRa,
                HieuSuatBoTruyen = entity.HieuSuatBoTruyen,
                HieuSuatNoiTruc = entity.HieuSuatNoiTruc,
                HieuSuatGoiTruc = entity.HieuSuatGoiTruc
            };
        }

        public BienTan ToEntity()
        {
            return new BienTan
            {
                LibId = LibId,
                DienApVao = DienApVao,
                DongDienVao = DongDienVao,
                TanSoVao = TanSoVao,
                CongSuatVao = CongSuatVao,
                DienApRa = DienApRa,
                DongDienRa = DongDienRa,
                TanSoRa = TanSoRa,
                CongSuatTongRa = CongSuatTongRa,
                CongSuatHieuDungRa = CongSuatHieuDungRa,
                HieuSuatBoTruyen = HieuSuatBoTruyen,
                HieuSuatNoiTruc = HieuSuatNoiTruc,
                HieuSuatGoiTruc = HieuSuatGoiTruc
            };
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
    /// <summary>
    /// DTO cho CamBien với tracking thay đổi
    /// </summary>
    public class CamBienDto : INotifyPropertyChanged
    {
        private float _nhietDoMoiTruongMin;
        private float _nhietDoMoiTruongMax;
        private float _doAmMoiTruongMin;
        private float _doAmMoiTruongMax;
        private float _apSuatKhiQuyenMin;
        private float _apSuatKhiQuyenMax;
        private float _chenhLechApSuatMin;
        private float _chenhLechApSuatMax;
        private float _apSuatTinhMin;
        private float _apSuatTinhMax;
        private float _doRungMin;
        private float _doRungMax;
        private float _doOnMin;
        private float _doOnMax;
        private float _soVongQuayMin;
        private float _soVongQuayMax;
        private float _momenMin;
        private float _momenMax;
        private float _phanHoiDongDienMin;
        private float _phanHoiDongDienMax;
        private float _phanHoiCongSuatMin;
        private float _phanHoiCongSuatMax;
        private float _phanHoiViTriVanMin;
        private float _phanHoiViTriVanMax;
        private float _phanHoiDienApMin;
        private float _phanHoiDienApMax;
        private float _nhietDoGoiTrucMin;
        private float _nhietDoGoiTrucMax;
        private float _nhietDoBauKhoMin;
        private float _nhietDoBauKhoMax;
        private float _camBienLuuLuongMin;
        private float _camBienLuuLuongMax;
        private float _phanHoiTanSoMin;
        private float _phanHoiTanSoMax;


        public Guid LibId { get; set; }

        public float NhietDoMoiTruongMin
        {
            get => _nhietDoMoiTruongMin;
            set { _nhietDoMoiTruongMin = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float NhietDoMoiTruongMax
        {
            get => _nhietDoMoiTruongMax;
            set { _nhietDoMoiTruongMax = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float DoAmMoiTruongMin
        {
            get => _doAmMoiTruongMin;
            set { _doAmMoiTruongMin = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float DoAmMoiTruongMax
        {
            get => _doAmMoiTruongMax;
            set { _doAmMoiTruongMax = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float ApSuatKhiQuyenMin
        {
            get => _apSuatKhiQuyenMin;
            set { _apSuatKhiQuyenMin = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float ApSuatKhiQuyenMax
        {
            get => _apSuatKhiQuyenMax;
            set { _apSuatKhiQuyenMax = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float ChenhLechApSuatMin
        {
            get => _chenhLechApSuatMin;
            set { _chenhLechApSuatMin = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float ChenhLechApSuatMax
        {
            get => _chenhLechApSuatMax;
            set { _chenhLechApSuatMax = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float ApSuatTinhMin
        {
            get => _apSuatTinhMin;
            set { _apSuatTinhMin = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float ApSuatTinhMax
        {
            get => _apSuatTinhMax;
            set { _apSuatTinhMax = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float DoRungMin
        {
            get => _doRungMin;
            set { _doRungMin = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float DoRungMax
        {
            get => _doRungMax;
            set { _doRungMax = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float DoOnMin
        {
            get => _doOnMin;
            set { _doOnMin = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float DoOnMax
        {
            get => _doOnMax;
            set { _doOnMax = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float SoVongQuayMin
        {
            get => _soVongQuayMin;
            set { _soVongQuayMin = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float SoVongQuayMax
        {
            get => _soVongQuayMax;
            set { _soVongQuayMax = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float MomenMin
        {
            get => _momenMin;
            set { _momenMin = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float MomenMax
        {
            get => _momenMax;
            set { _momenMax = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float PhanHoiDongDienMin
        {
            get => _phanHoiDongDienMin;
            set { _phanHoiDongDienMin = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float PhanHoiDongDienMax
        {
            get => _phanHoiDongDienMax;
            set { _phanHoiDongDienMax = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float PhanHoiCongSuatMin
        {
            get => _phanHoiCongSuatMin;
            set { _phanHoiCongSuatMin = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float PhanHoiCongSuatMax
        {
            get => _phanHoiCongSuatMax;
            set { _phanHoiCongSuatMax = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float PhanHoiViTriVanMin
        {
            get => _phanHoiViTriVanMin;
            set { _phanHoiViTriVanMin = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float PhanHoiViTriVanMax
        {
            get => _phanHoiViTriVanMax;
            set { _phanHoiViTriVanMax = value; OnPropertyChanged(); NotifyChanged(); }
        }

        public float PhanHoiDienApMin
        {
            get => _phanHoiDienApMin;
            set { _phanHoiDienApMin = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float PhanHoiDienApMax
        {
            get => _phanHoiDienApMax;
            set { _phanHoiDienApMax = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float NhietDoGoiTrucMin
        {
            get => _nhietDoGoiTrucMin;
            set { _nhietDoGoiTrucMin = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float NhietDoGoiTrucMax
        {
            get => _nhietDoGoiTrucMax;
            set { _nhietDoGoiTrucMax = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float NhietDoBauKhoMin
        {
            get => _nhietDoBauKhoMin;
            set { _nhietDoBauKhoMin = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float NhietDoBauKhoMax
        {
            get => _nhietDoBauKhoMax;
            set { _nhietDoBauKhoMax = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float CamBienLuuLuongMin
        {
            get => _camBienLuuLuongMin;
            set { _camBienLuuLuongMin = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float CamBienLuuLuongMax
        {
            get => _camBienLuuLuongMax;
            set { _camBienLuuLuongMax = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float PhanHoiTanSoMin
        {
            get => _phanHoiTanSoMin;
            set { _phanHoiTanSoMin = value; OnPropertyChanged(); NotifyChanged(); }
        }

        public float PhanHoiTanSoMax
        {
            get => _phanHoiTanSoMax;
            set { _phanHoiTanSoMax = value; OnPropertyChanged(); NotifyChanged(); }
        }

        public event EventHandler DataChanged;
        private void NotifyChanged()
        {
            DataChanged?.Invoke(this, EventArgs.Empty);
        }

        public static CamBienDto FromEntity(CamBien entity)
        {
            return new CamBienDto
            {
                LibId = entity.LibId,
                NhietDoMoiTruongMin = entity.NhietDoMoiTruongMin,
                NhietDoMoiTruongMax = entity.NhietDoMoiTruongMax,
                DoAmMoiTruongMin = entity.DoAmMoiTruongMin,
                DoAmMoiTruongMax = entity.DoAmMoiTruongMax,
                ApSuatKhiQuyenMin = entity.ApSuatKhiQuyenMin,
                ApSuatKhiQuyenMax = entity.ApSuatKhiQuyenMax,
                ChenhLechApSuatMin = entity.ChenhLechApSuatMin,
                ChenhLechApSuatMax = entity.ChenhLechApSuatMax,
                ApSuatTinhMin = entity.ApSuatTinhMin,
                ApSuatTinhMax = entity.ApSuatTinhMax,
                DoRungMin = entity.DoRungMin,
                DoRungMax = entity.DoRungMax,
                DoOnMin = entity.DoOnMin,
                DoOnMax = entity.DoOnMax,
                SoVongQuayMin = entity.SoVongQuayMin,
                SoVongQuayMax = entity.SoVongQuayMax,
                MomenMin = entity.MomenMin,
                MomenMax = entity.MomenMax,
                PhanHoiDongDienMin = entity.PhanHoiDongDienMin,
                PhanHoiDongDienMax = entity.PhanHoiDongDienMax,
                PhanHoiCongSuatMin = entity.PhanHoiCongSuatMin,
                PhanHoiCongSuatMax = entity.PhanHoiCongSuatMax,
                PhanHoiViTriVanMin = entity.PhanHoiViTriVanMin,
                PhanHoiViTriVanMax = entity.PhanHoiViTriVanMax,
                PhanHoiDienApMin = entity.PhanHoiDienApMin,
                PhanHoiDienApMax = entity.PhanHoiDienApMax,
                NhietDoGoiTrucMin = entity.NhietDoGoiTrucMin,
                NhietDoGoiTrucMax = entity.NhietDoGoiTrucMax,
                NhietDoBauKhoMin = entity.NhietDoBauKhoMin,
                NhietDoBauKhoMax = entity.NhietDoBauKhoMax,
                CamBienLuuLuongMin = entity.CamBienLuuLuongMin,
                CamBienLuuLuongMax = entity.CamBienLuuLuongMax,
                PhanHoiTanSoMin = entity.PhanHoiTanSoMin,
                PhanHoiTanSoMax = entity.PhanHoiTanSoMax
            };
        }

        public CamBien ToEntity()
        {
            return new CamBien
            {
                LibId = LibId,
                NhietDoMoiTruongMin = NhietDoMoiTruongMin,
                NhietDoMoiTruongMax = NhietDoMoiTruongMax,
                DoAmMoiTruongMin = DoAmMoiTruongMin,
                DoAmMoiTruongMax = DoAmMoiTruongMax,
                ApSuatKhiQuyenMin = ApSuatKhiQuyenMin,
                ApSuatKhiQuyenMax = ApSuatKhiQuyenMax,
                ChenhLechApSuatMin = ChenhLechApSuatMin,
                ChenhLechApSuatMax = ChenhLechApSuatMax,
                ApSuatTinhMin = ApSuatTinhMin,
                ApSuatTinhMax = ApSuatTinhMax,
                DoRungMin = DoRungMin,
                DoRungMax = DoRungMax,
                DoOnMin = DoOnMin,
                DoOnMax = DoOnMax,
                SoVongQuayMin = SoVongQuayMin,
                SoVongQuayMax = SoVongQuayMax,
                MomenMin = MomenMin,
                MomenMax = MomenMax,
                PhanHoiDongDienMin = PhanHoiDongDienMin,
                PhanHoiDongDienMax = PhanHoiDongDienMax,
                PhanHoiCongSuatMin = PhanHoiCongSuatMin,
                PhanHoiCongSuatMax = PhanHoiCongSuatMax,
                PhanHoiViTriVanMin = PhanHoiViTriVanMin,
                PhanHoiViTriVanMax = PhanHoiViTriVanMax,
                PhanHoiDienApMin = PhanHoiDienApMin,
                PhanHoiDienApMax = PhanHoiDienApMax,
                NhietDoGoiTrucMin = NhietDoGoiTrucMin,
                NhietDoGoiTrucMax = NhietDoGoiTrucMax,
                NhietDoBauKhoMin = NhietDoBauKhoMin,
                NhietDoBauKhoMax = NhietDoBauKhoMax,
                CamBienLuuLuongMin = CamBienLuuLuongMin,
                CamBienLuuLuongMax = CamBienLuuLuongMax,
                PhanHoiTanSoMin = PhanHoiTanSoMin, 
                PhanHoiTanSoMax = PhanHoiTanSoMax
            };
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    /// <summary>
    /// DTO cho OngGio với tracking thay đổi
    /// </summary>
    public class OngGioDto : INotifyPropertyChanged
    {
        private float _duongKinhOngGio;
        private float _chieuDaiOngGioSauQuat;
        private float _chieuDaiOngGioTruocQuat;
        private float _duongKinhLoPhut;
        private float _duongKinhMiengQuat;
        private float _chieuDaiConQuat;

        public Guid LibId { get; set; }

        public float DuongKinhOngGio
        {
            get => _duongKinhOngGio;
            set { _duongKinhOngGio = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float ChieuDaiOngGioSauQuat
        {
            get => _chieuDaiOngGioSauQuat;
            set { _chieuDaiOngGioSauQuat = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float ChieuDaiOngGioTruocQuat
        {
            get => _chieuDaiOngGioTruocQuat;
            set { _chieuDaiOngGioTruocQuat = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float DuongKinhLoPhut
        {
            get => _duongKinhLoPhut;
            set { _duongKinhLoPhut = value; OnPropertyChanged(); NotifyChanged(); }
        }

        public float DuongKinhMiengQuat
        {
            get => _duongKinhMiengQuat;
            set { _duongKinhMiengQuat = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float ChieuDaiConQuat
        {
            get => _chieuDaiConQuat;
            set { _chieuDaiConQuat = value; OnPropertyChanged(); NotifyChanged(); }
        }

        public event EventHandler DataChanged;
        private void NotifyChanged()
        {
            DataChanged?.Invoke(this, EventArgs.Empty);
        }

        public static OngGioDto FromEntity(OngGio entity)
        {
            return new OngGioDto
            {
                LibId = entity.LibId,
                DuongKinhOngGio = entity.DuongKinhOngGio,
                ChieuDaiOngGioSauQuat = entity.ChieuDaiOngGioSauQuat,
                ChieuDaiOngGioTruocQuat = entity.ChieuDaiOngGioTruocQuat,
                DuongKinhLoPhut = entity.DuongKinhLoPhut,
                DuongKinhMiengQuat = entity.DuongKinhMiengQuat,
                ChieuDaiConQuat = entity.ChieuDaiConQuat
            };
        }

        public OngGio ToEntity()
        {
            return new OngGio
            {
                LibId = LibId,
                DuongKinhOngGio = DuongKinhOngGio,
                ChieuDaiOngGioSauQuat = ChieuDaiOngGioSauQuat,
                ChieuDaiOngGioTruocQuat = ChieuDaiOngGioTruocQuat,
                DuongKinhLoPhut = DuongKinhLoPhut,
                DuongKinhMiengQuat = DuongKinhMiengQuat,
                ChieuDaiConQuat = ChieuDaiConQuat
            };
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}