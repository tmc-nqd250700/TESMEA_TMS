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
    /// <summary>
    /// DTO cho CamBien với tracking thay đổi
    /// </summary>
    public class CamBienDto : INotifyPropertyChanged
    {
        public Guid LibId { get; set; }

        // Nhiệt độ môi trường C
        private bool _isImportNhietDoMoiTruong;
        private float _nhietDoMoiTruongValue;
        private float _nhietDoMoiTruongMin;
        private float _nhietDoMoiTruongMax;
        public bool IsImportNhietDoMoiTruong
        {
            get => _isImportNhietDoMoiTruong;
            set { _isImportNhietDoMoiTruong = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float NhietDoMoiTruongValue
        {
            get => _nhietDoMoiTruongValue;
            set { _nhietDoMoiTruongValue = value; OnPropertyChanged(); NotifyChanged(); }
        }
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

        // Độ ẩm môi trường %
        private bool _isImportDoAmMoiTruong;
        private float _doAmMoiTruongValue;
        private float _doAmMoiTruongMin;
        private float _doAmMoiTruongMax;
        public bool IsImportDoAmMoiTruong
        {
            get => _isImportDoAmMoiTruong;
            set { _isImportDoAmMoiTruong = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float DoAmMoiTruongValue
        {
            get => _doAmMoiTruongValue;
            set { _doAmMoiTruongValue = value; OnPropertyChanged(); NotifyChanged(); }
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

        // Áp suất khí quyển Pa
        private bool _isImportApSuatKhiQuyen;
        private float _apSuatKhiQuyenValue;
        private float _apSuatKhiQuyenMin;
        private float _apSuatKhiQuyenMax;
        public bool IsImportApSuatKhiQuyen
        {
            get => _isImportApSuatKhiQuyen;
            set { _isImportApSuatKhiQuyen = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float ApSuatKhiQuyenValue
        {
            get => _apSuatKhiQuyenValue;
            set { _apSuatKhiQuyenValue = value; OnPropertyChanged(); NotifyChanged(); }
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

        // Áp suất chênh lệch Pa
        private bool _isImportChenhLechApSuat;
        private float _chenhLechApSuatValue;
        private float _chenhLechApSuatMin;
        private float _chenhLechApSuatMax;
        public bool IsImportChenhLechApSuat
        {
            get => _isImportChenhLechApSuat;
            set { _isImportChenhLechApSuat = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float ChenhLechApSuatValue
        {
            get => _chenhLechApSuatValue;
            set { _chenhLechApSuatValue = value; OnPropertyChanged(); NotifyChanged(); }
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

        // Áp suất tĩnh Pa
        private bool _isImportApSuatTinh;
        private float _apSuatTinhValue;
        private float _apSuatTinhMin;
        private float _apSuatTinhMax;
        public bool IsImportApSuatTinh
        {
            get => _isImportApSuatTinh;
            set { _isImportApSuatTinh = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float ApSuatTinhValue
        {
            get => _apSuatTinhValue;
            set { _apSuatTinhValue = value; OnPropertyChanged(); NotifyChanged(); }
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

        // Độ rung
        private bool _isImportDoRung;
        private float _doRungValue;
        private float _doRungMin;
        private float _doRungMax;
        public bool IsImportDoRung
        {
            get => _isImportDoRung;
            set { _isImportDoRung = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float DoRungValue
        {
            get => _doRungValue;
            set { _doRungValue = value; OnPropertyChanged(); NotifyChanged(); }
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

        // Độ ồn
        private bool _isImportDoOn;
        private float _doOnValue;
        private float _doOnMin;
        private float _doOnMax;
        public bool IsImportDoOn
        {
            get => _isImportDoOn;
            set { _isImportDoOn = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float DoOnValue
        {
            get => _doOnValue;
            set { _doOnValue = value; OnPropertyChanged(); NotifyChanged(); }
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

        // Số vòng quay
        private bool _isImportSoVongQuay;
        private float _soVongQuayValue;
        private float _soVongQuayMin;
        private float _soVongQuayMax;
        public bool IsImportSoVongQuay
        {
            get => _isImportSoVongQuay;
            set { _isImportSoVongQuay = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float SoVongQuayValue
        {
            get => _soVongQuayValue;
            set { _soVongQuayValue = value; OnPropertyChanged(); NotifyChanged(); }
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

        // Momen xoắn
        private bool _isImportMomen;
        private float _momenValue;
        private float _momenMin;
        private float _momenMax;
        public bool IsImportMomen
        {
            get => _isImportMomen;
            set { _isImportMomen = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float MomenValue
        {
            get => _momenValue;
            set { _momenValue = value; OnPropertyChanged(); NotifyChanged(); }
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

        // Phản hồi dòng điện
        private bool _isImportPhanHoiDongDien;
        private float _phanHoiDongDienValue;
        private float _phanHoiDongDienMin;
        private float _phanHoiDongDienMax;
        public bool IsImportPhanHoiDongDien
        {
            get => _isImportPhanHoiDongDien;
            set { _isImportPhanHoiDongDien = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float PhanHoiDongDienValue
        {
            get => _phanHoiDongDienValue;
            set { _phanHoiDongDienValue = value; OnPropertyChanged(); NotifyChanged(); }
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

        // Phản hồi công suất
        private bool _isImportPhanHoiCongSuat;
        private float _phanHoiCongSuatValue;
        private float _phanHoiCongSuatMin;
        private float _phanHoiCongSuatMax;
        public bool IsImportPhanHoiCongSuat
        {
            get => _isImportPhanHoiCongSuat;
            set { _isImportPhanHoiCongSuat = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float PhanHoiCongSuatValue
        {
            get => _phanHoiCongSuatValue;
            set { _phanHoiCongSuatValue = value; OnPropertyChanged(); NotifyChanged(); }
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

        // Phản hồi vị trí van
        private bool _isImportPhanHoiViTriVan;
        private float _phanHoiViTriVanValue;
        private float _phanHoiViTriVanMin;
        private float _phanHoiViTriVanMax;
        public bool IsImportPhanHoiViTriVan
        {
            get => _isImportPhanHoiViTriVan;
            set { _isImportPhanHoiViTriVan = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float PhanHoiViTriVanValue
        {
            get => _phanHoiViTriVanValue;
            set { _phanHoiViTriVanValue = value; OnPropertyChanged(); NotifyChanged(); }
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

        // Phản hồi điện áp
        private bool _isImportPhanHoiDienAp;
        private float _phanHoiDienApValue;
        private float _phanHoiDienApMin;
        private float _phanHoiDienApMax;
        public bool IsImportPhanHoiDienAp
        {
            get => _isImportPhanHoiDienAp;
            set { _isImportPhanHoiDienAp = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float PhanHoiDienApValue
        {
            get => _phanHoiDienApValue;
            set { _phanHoiDienApValue = value; OnPropertyChanged(); NotifyChanged(); }
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

        // Nhiệt độ gối trục
        private bool _isImportNhietDoGoiTruc;
        private float _nhietDoGoiTrucValue;
        private float _nhietDoGoiTrucMin;
        private float _nhietDoGoiTrucMax;
        public bool IsImportNhietDoGoiTruc
        {
            get => _isImportNhietDoGoiTruc;
            set { _isImportNhietDoGoiTruc = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float NhietDoGoiTrucValue
        {
            get => _nhietDoGoiTrucValue;
            set { _nhietDoGoiTrucValue = value; OnPropertyChanged(); NotifyChanged(); }
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

        // Nhiệt độ bầu khô
        private bool _isImportNhietDoBauKho;
        private float _nhietDoBauKhoValue;
        private float _nhietDoBauKhoMin;
        private float _nhietDoBauKhoMax;
        public bool IsImportNhietDoBauKho
        {
            get => _isImportNhietDoBauKho;
            set { _isImportNhietDoBauKho = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float NhietDoBauKhoValue
        {
            get => _nhietDoBauKhoValue;
            set { _nhietDoBauKhoValue = value; OnPropertyChanged(); NotifyChanged(); }
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

        // Cảm biến lưu lượng
        private bool _isImportCamBienLuuLuong;
        private float _camBienLuuLuongValue;
        private float _camBienLuuLuongMin;
        private float _camBienLuuLuongMax;
        public bool IsImportCamBienLuuLuong
        {
            get => _isImportCamBienLuuLuong;
            set { _isImportCamBienLuuLuong = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float CamBienLuuLuongValue
        {
            get => _camBienLuuLuongValue;
            set { _camBienLuuLuongValue = value; OnPropertyChanged(); NotifyChanged(); }
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

        // Phản hồi tần số
        private bool _isImportPhanHoiTanSo;
        private float _phanHoiTanSoValue;
        private float _phanHoiTanSoMin;
        private float _phanHoiTanSoMax;
        public bool IsImportPhanHoiTanSo
        {
            get => _isImportPhanHoiTanSo;
            set { _isImportPhanHoiTanSo = value; OnPropertyChanged(); NotifyChanged(); }
        }
        public float PhanHoiTanSoValue
        {
            get => _phanHoiTanSoValue;
            set { _phanHoiTanSoValue = value; OnPropertyChanged(); NotifyChanged(); }
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
                IsImportNhietDoMoiTruong = entity.IsImportNhietDoMoiTruong,
                NhietDoMoiTruongValue = entity.NhietDoMoiTruongValue,
                NhietDoMoiTruongMin = entity.NhietDoMoiTruongMin,
                NhietDoMoiTruongMax = entity.NhietDoMoiTruongMax,

                IsImportDoAmMoiTruong = entity.IsImportDoAmMoiTruong,
                DoAmMoiTruongValue = entity.DoAmMoiTruongValue,
                DoAmMoiTruongMin = entity.DoAmMoiTruongMin,
                DoAmMoiTruongMax = entity.DoAmMoiTruongMax,

                IsImportApSuatKhiQuyen = entity.IsImportApSuatKhiQuyen,
                ApSuatKhiQuyenValue = entity.ApSuatKhiQuyenValue,
                ApSuatKhiQuyenMin = entity.ApSuatKhiQuyenMin,
                ApSuatKhiQuyenMax = entity.ApSuatKhiQuyenMax,

                IsImportChenhLechApSuat = entity.IsImportChenhLechApSuat,
                ChenhLechApSuatValue = entity.ChenhLechApSuatValue,
                ChenhLechApSuatMin = entity.ChenhLechApSuatMin,
                ChenhLechApSuatMax = entity.ChenhLechApSuatMax,

                IsImportApSuatTinh = entity.IsImportApSuatTinh,
                ApSuatTinhValue = entity.ApSuatTinhValue,
                ApSuatTinhMin = entity.ApSuatTinhMin,
                ApSuatTinhMax = entity.ApSuatTinhMax,

                IsImportDoRung = entity.IsImportDoRung,
                DoRungValue = entity.DoRungValue,
                DoRungMin = entity.DoRungMin,
                DoRungMax = entity.DoRungMax,

                IsImportDoOn = entity.IsImportDoOn,
                DoOnValue = entity.DoOnValue,
                DoOnMin = entity.DoOnMin,
                DoOnMax = entity.DoOnMax,

                IsImportSoVongQuay = entity.IsImportSoVongQuay,
                SoVongQuayValue = entity.SoVongQuayValue,
                SoVongQuayMin = entity.SoVongQuayMin,
                SoVongQuayMax = entity.SoVongQuayMax,

                IsImportMomen = entity.IsImportMomen,
                MomenValue = entity.MomenValue,
                MomenMin = entity.MomenMin,
                MomenMax = entity.MomenMax,

                IsImportPhanHoiDongDien = entity.IsImportPhanHoiDongDien,
                PhanHoiDongDienValue = entity.PhanHoiDongDienValue,
                PhanHoiDongDienMin = entity.PhanHoiDongDienMin,
                PhanHoiDongDienMax = entity.PhanHoiDongDienMax,

                IsImportPhanHoiCongSuat = entity.IsImportPhanHoiCongSuat,
                PhanHoiCongSuatValue = entity.PhanHoiCongSuatValue,
                PhanHoiCongSuatMin = entity.PhanHoiCongSuatMin,
                PhanHoiCongSuatMax = entity.PhanHoiCongSuatMax,

                IsImportPhanHoiViTriVan = entity.IsImportPhanHoiViTriVan,
                PhanHoiViTriVanValue = entity.PhanHoiViTriVanValue,
                PhanHoiViTriVanMin = entity.PhanHoiViTriVanMin,
                PhanHoiViTriVanMax = entity.PhanHoiViTriVanMax,

                IsImportPhanHoiDienAp = entity.IsImportPhanHoiDienAp,
                PhanHoiDienApValue = entity.PhanHoiDienApValue,
                PhanHoiDienApMin = entity.PhanHoiDienApMin,
                PhanHoiDienApMax = entity.PhanHoiDienApMax,

                IsImportNhietDoGoiTruc = entity.IsImportNhietDoGoiTruc,
                NhietDoGoiTrucValue = entity.NhietDoGoiTrucValue,
                NhietDoGoiTrucMin = entity.NhietDoGoiTrucMin,
                NhietDoGoiTrucMax = entity.NhietDoGoiTrucMax,

                IsImportNhietDoBauKho = entity.IsImportNhietDoBauKho,
                NhietDoBauKhoValue = entity.NhietDoBauKhoValue,
                NhietDoBauKhoMin = entity.NhietDoBauKhoMin,
                NhietDoBauKhoMax = entity.NhietDoBauKhoMax,

                IsImportCamBienLuuLuong = entity.IsImportCamBienLuuLuong,
                CamBienLuuLuongValue = entity.CamBienLuuLuongValue,
                CamBienLuuLuongMin = entity.CamBienLuuLuongMin,
                CamBienLuuLuongMax = entity.CamBienLuuLuongMax,

                IsImportPhanHoiTanSo = entity.IsImportPhanHoiTanSo,
                PhanHoiTanSoValue = entity.PhanHoiTanSoValue,
                PhanHoiTanSoMin = entity.PhanHoiTanSoMin,
                PhanHoiTanSoMax = entity.PhanHoiTanSoMax
            };
        }

        public CamBien ToEntity()
        {
            return new CamBien
            {
                LibId = LibId,
                IsImportNhietDoMoiTruong = IsImportNhietDoMoiTruong,
                NhietDoMoiTruongValue = NhietDoMoiTruongValue,
                NhietDoMoiTruongMin = NhietDoMoiTruongMin,
                NhietDoMoiTruongMax = NhietDoMoiTruongMax,

                IsImportDoAmMoiTruong = IsImportDoAmMoiTruong,
                DoAmMoiTruongValue = DoAmMoiTruongValue,
                DoAmMoiTruongMin = DoAmMoiTruongMin,
                DoAmMoiTruongMax = DoAmMoiTruongMax,

                IsImportApSuatKhiQuyen = IsImportApSuatKhiQuyen,
                ApSuatKhiQuyenValue = ApSuatKhiQuyenValue,
                ApSuatKhiQuyenMin = ApSuatKhiQuyenMin,
                ApSuatKhiQuyenMax = ApSuatKhiQuyenMax,

                IsImportChenhLechApSuat = IsImportChenhLechApSuat,
                ChenhLechApSuatValue = ChenhLechApSuatValue,
                ChenhLechApSuatMin = ChenhLechApSuatMin,
                ChenhLechApSuatMax = ChenhLechApSuatMax,

                IsImportApSuatTinh = IsImportApSuatTinh,
                ApSuatTinhValue = ApSuatTinhValue,
                ApSuatTinhMin = ApSuatTinhMin,
                ApSuatTinhMax = ApSuatTinhMax,

                IsImportDoRung = IsImportDoRung,
                DoRungValue = DoRungValue,
                DoRungMin = DoRungMin,
                DoRungMax = DoRungMax,

                IsImportDoOn = IsImportDoOn,
                DoOnValue = DoOnValue,
                DoOnMin = DoOnMin,
                DoOnMax = DoOnMax,

                IsImportSoVongQuay = IsImportSoVongQuay,
                SoVongQuayValue = SoVongQuayValue,
                SoVongQuayMin = SoVongQuayMin,
                SoVongQuayMax = SoVongQuayMax,

                IsImportMomen = IsImportMomen,
                MomenValue = MomenValue,
                MomenMin = MomenMin,
                MomenMax = MomenMax,

                IsImportPhanHoiDongDien = IsImportPhanHoiDongDien,
                PhanHoiDongDienValue = PhanHoiDongDienValue,
                PhanHoiDongDienMin = PhanHoiDongDienMin,
                PhanHoiDongDienMax = PhanHoiDongDienMax,

                IsImportPhanHoiCongSuat = IsImportPhanHoiCongSuat,
                PhanHoiCongSuatValue = PhanHoiCongSuatValue,
                PhanHoiCongSuatMin = PhanHoiCongSuatMin,
                PhanHoiCongSuatMax = PhanHoiCongSuatMax,

                IsImportPhanHoiViTriVan = IsImportPhanHoiViTriVan,
                PhanHoiViTriVanValue = PhanHoiViTriVanValue,
                PhanHoiViTriVanMin = PhanHoiViTriVanMin,
                PhanHoiViTriVanMax = PhanHoiViTriVanMax,

                IsImportPhanHoiDienAp = IsImportPhanHoiDienAp,
                PhanHoiDienApValue = PhanHoiDienApValue,
                PhanHoiDienApMin = PhanHoiDienApMin,
                PhanHoiDienApMax = PhanHoiDienApMax,

                IsImportNhietDoGoiTruc = IsImportNhietDoGoiTruc,
                NhietDoGoiTrucValue = NhietDoGoiTrucValue,
                NhietDoGoiTrucMin = NhietDoGoiTrucMin,
                NhietDoGoiTrucMax = NhietDoGoiTrucMax,

                IsImportNhietDoBauKho = IsImportNhietDoBauKho,
                NhietDoBauKhoValue = NhietDoBauKhoValue,
                NhietDoBauKhoMin = NhietDoBauKhoMin,
                NhietDoBauKhoMax = NhietDoBauKhoMax,

                IsImportCamBienLuuLuong = IsImportCamBienLuuLuong,
                CamBienLuuLuongValue = CamBienLuuLuongValue,
                CamBienLuuLuongMin = CamBienLuuLuongMin,
                CamBienLuuLuongMax = CamBienLuuLuongMax,

                IsImportPhanHoiTanSo = IsImportPhanHoiTanSo,
                PhanHoiTanSoValue = PhanHoiTanSoValue,
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