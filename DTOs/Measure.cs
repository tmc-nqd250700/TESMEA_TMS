using System.ComponentModel;

namespace TESMEA_TMS.DTOs
{
    public class Measure : INotifyPropertyChanged
    {
        // sen (sensor), fb(feedback)
        public int k { get; set; }
        public float S { get; set; }
        public float CV { get; set; }
        public float NhietDoMoiTruong_sen { get; set; } // nhiệt độ bầu khô
        public float DoAm_sen { get; set; }
        public float ApSuatkhiQuyen_sen { get; set; }
        public float ChenhLechApSuat_sen { get; set; }
        public float ApSuatTinh_sen { get; set; }
        public float DoRung_sen { get; set; }
        public float DoOn_sen { get; set; }
        public float SoVongQuay_sen { get; set; } // tốc độ thực của guồn cánh
        public float Momen_sen { get; set; }
        public float DongDien_fb { get; set; }
        public float DienAp_fb { get; set; }
        public float CongSuat_fb { get; set; }
        public float ViTriVan_fb { get; set; }
        public float TanSo_fb { get; set; }

        private MeasureStatus _f;
        public MeasureStatus F
        {
            get => _f;
            set { _f = value; OnPropertyChanged(nameof(F)); }
        }

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        public event PropertyChangedEventHandler PropertyChanged;
    }
    public enum MeasureStatus
    {
        Completed, // F - xanh
        Pending,   // T - đen
        Error      // X - đỏ
    }

    public class MeasureConnect
    {
        public int Step { get; set; }
        public string Property { get; set; }
        public float Value { get; set; }
        private MeasureStatus _f;
        public MeasureStatus F
        {
            get => _f;
            set { _f = value; OnPropertyChanged(nameof(F)); }
        }

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        public event PropertyChangedEventHandler PropertyChanged;
    }

    // Class hiển thị kết quả tính toán thông số đo kiểm trên datagrid
    public class MeasureResponse
    {
        public int STT { get; set; }
        public float Airflow { get; set; }
        public float Ps { get; set; }
        public float Pt { get; set; }
        public float SEff { get; set; } // Hiệu suất tĩnh
        public float TEff { get; set; } // Hiệu suất tổng
        public float Power { get; set; }
        public float Prt { get; set; } // Công suất tính bằng momen xoắn
        public float Est { get; set; } // hs tĩnh theo momen xoắn
        public float Ett { get; set; } // hs tổng theo momen xoắn

    }

    // class hiển thị kết quả trên màn hình đo kiểm
    public class ParameterShow
    {
        public float Freq_show { get; set; }         // Tần số 
        public float Current_show { get; set; }      // Dòng điện
        public float Pw_show { get; set; }           // Công suất
        public float Speed_show { get; set; }        // Tốc độ
        public float TempB_show { get; set; }        // Nhiệt độ gối
        public float T_Show { get; set; }            // Momen xoắn trên trục quạt
        public float Ps_show { get; set; }           // Áp suất tĩnh
        public float Pt_show { get; set; }           // Áp suất tổng
        public float Flow_show { get; set; }         // Lưu lượng
        public float Ta_show { get; set; }           // Nhiệt độ T3
        public float Td_show { get; set; }           // Nhiệt độ bầu khô
        public float ViaB_show { get; set; }         // Nhiệt độ bầu khô
        public float Prt_show { get; set; }          // Công suất quạt tính bằng momen xoắn
        public float deltap { get; set; }
        public float Pe3 { get; set; }
        public float Ta { get; set; }
    }
    public class MeasureFittingFC
    {
        public double[] FlowPoint_ft { get; set; }
        public double[] PrPoint_ft { get; set; }
        public double[] PsPoint_ft { get; set; }
        public double[] PtPoint_ft { get; set; }
        public double[] EsPoint_ft { get; set; }
        public double[] EtPoint_ft { get; set; }
        public double[] PrtPoint_ft { get; set; }
        public double[] EstPoint_ft { get; set; }
        public double[] EttPoint_ft { get; set; }


        public double[] Ope_FlowPoint { get; set; }
        public double[] Ope_PrPoint { get; set; }
        public double[] Ope_PsPoint { get; set; }
        public double[] Ope_PtPoint { get; set; }
        public double[] Ope_EsPoint { get; set; }
        public double[] Ope_EtPoint { get; set; }
        public double[] Ope_EstPoint { get; set; }
        public double[] Ope_EttPoint { get; set; }
        public double[] PrtPoint { get; set; }

        public MeasureFittingFC() : this(0)
        {
        }

        public MeasureFittingFC(int range)
        {
            FlowPoint_ft = new double[range];
            PrPoint_ft = new double[range];
            PsPoint_ft = new double[range];
            PtPoint_ft = new double[range];
            EsPoint_ft = new double[range];
            EtPoint_ft = new double[range];
            PrtPoint_ft = new double[range];
            EstPoint_ft = new double[range];
            EttPoint_ft = new double[range];


            Ope_FlowPoint = new double[range];
            Ope_PrPoint = new double[range];
            Ope_PsPoint = new double[range];
            Ope_PtPoint = new double[range];
            Ope_EsPoint = new double[range];
            Ope_EtPoint = new double[range];
            Ope_EstPoint = new double[range];
            Ope_EttPoint = new double[range];
            PrtPoint = new double[range];
        }
    }
}
