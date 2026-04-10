using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using TESMEA_TMS.Helpers;

namespace TESMEA_TMS.Views
{
    /// <summary>
    /// Interaction logic for ConfirmAddScenarioDialog.xaml
    /// </summary>
    public partial class ConfirmAddScenarioDialog : Window, INotifyPropertyChanged
    {

        public class PreviewParamRow : INotifyPropertyChanged
        {
            private float _s;
            private float _cv;

            public int STT { get; set; }

            public float S
            {
                get => _s;
                set { _s = value; OnPropertyChanged(); }
            }
            public float CV
            {
                get => _cv;
                set { _cv = value; OnPropertyChanged(); }
            }

            public event PropertyChangedEventHandler PropertyChanged;
            protected void OnPropertyChanged([CallerMemberName] string name = null)
                => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        private float _inverterHz = 50f;
        private int _freqNum = 4;
        private int _valveNumPerFreq = 5;
        private bool _hasPreview;
        public float InverterHz
        {
            get => _inverterHz;
            set { _inverterHz = value; OnPropertyChanged(); }
        }
        public int FreqNum
        {
            get => _freqNum;
            set { _freqNum = value; OnPropertyChanged(); }
        }
        public int ValveNumPerFreq
        {
            get => _valveNumPerFreq;
            set { _valveNumPerFreq = value; OnPropertyChanged(); }
        }

        /// <summary>Điều khiển hiển thị GroupBox preview.</summary>
        public bool HasPreview
        {
            get => _hasPreview;
            set { _hasPreview = value; OnPropertyChanged(); }
        }
        public ObservableCollection<PreviewParamRow> PreviewParams { get; }
            = new ObservableCollection<PreviewParamRow>();


        private string _inputText;
        private float _standardDeviation = 105f;
        private float _timeRange = 20;
        public string InputText
        {
            get => _inputText;
            set
            {
                _inputText = value;
                OnPropertyChanged();
            }
        }

        public float StandardDeviation
        {
            get => _standardDeviation;
            set
            {
                _standardDeviation = value;
                OnPropertyChanged();
            }
        }

        public float TimeRange
        {
            get => _timeRange;
            set
            {
                _timeRange = value;
                OnPropertyChanged();
            }
        }

        public ConfirmAddScenarioDialog()
        {
            InitializeComponent();
            DataContext = this;
            Loaded += (s, e) => InputTextBox.Focus();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(InputText))
            {
                MessageBoxHelper.ShowWarning("Vui lòng nhập thông tin tên kịch bản!");
                return;
            }

            if (!HasPreview || PreviewParams.Count == 0)
            {
                if(MessageBoxHelper.ShowQuestion("Bạn có muốn thực hiện tạo danh sách điểm đo mẫu không?"))
                {
                    this.GenerateButton_Click(null, null);
                    return;
                }    
            }

            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void GenerateButton_Click(object sender, RoutedEventArgs e)
        {
            if (InverterHz <= 0)
            {
                MessageBoxHelper.ShowWarning("Tần số đo kiểm phải lớn hơn 0");
                return;
            }

            if(InverterHz > 60)
            {
                MessageBoxHelper.ShowWarning("Tần số đo kiểm thường không vượt quá 60Hz. Vui lòng kiểm tra lại!");
                return;
            }

            if (FreqNum <= 0)
            {
                MessageBoxHelper.ShowWarning("Số lượng tần số phải lớn hơn 0");
                return;
            }
            if (ValveNumPerFreq <= 0)
            {
                MessageBoxHelper.ShowWarning("Số lượng điểm đo phải lớn hơn 0");
                return;
            }
            if (ValveNumPerFreq > 10)
            {
                MessageBoxHelper.ShowWarning("Chỉ được phép đo tối đa 10 điểm mỗi tần số để tránh quá tải dữ liệu!");
                return;
            }

            // Tần số: từ 50% → 100% của InverterHz, chia đều FreqNum bước
            // Ví dụ: 50Hz, 4 tần số → [25.00, 33.33, 41.67, 50.00]
            var freqList = GenerateLinear(InverterHz * 0.5f, InverterHz, FreqNum);

            // Điểm đo: từ 10% → 100%, chia đều ValveNumPerFreq bước
            // Ví dụ: 5 điểm → [10, 25, 50, 75, 100]
            var cvList = GenerateLinear(10f, 100f, ValveNumPerFreq);

            PreviewParams.Clear();
            int stt = 1;
            foreach (var freq in freqList)
                foreach (var cv in cvList)
                    PreviewParams.Add(new PreviewParamRow
                    {
                        STT = stt++,
                        S = freq,
                        CV = cv
                    });

            HasPreview = true;
        }

        /// <summary>
        /// Sinh n giá trị chia đều từ <paramref name="from"/> đến <paramref name="to"/> (inclusive).
        /// Nếu n == 1 thì trả về đúng [to].
        /// </summary>
        private static List<float> GenerateLinear(float from, float to, int n)
        {
            var list = new List<float>(n);
            if (n == 1) { list.Add(to); return list; }
            float step = (to - from) / (n - 1);
            for (int i = 0; i < n; i++)
                list.Add(from + step * i);
            return list;
        }

        private void PreviewDataGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            // Chỉ xử lý khi người dùng commit (không phải Cancel)
            if (e.EditAction != DataGridEditAction.Commit) return;

            var editedRow = e.Row.Item as PreviewParamRow;
            if (editedRow == null) return;

            // Lấy giá trị mới người dùng vừa gõ
            var textBox = e.EditingElement as System.Windows.Controls.TextBox;
            if (textBox == null) return;

            string newText = textBox.Text.Trim();
            string columnHeader = (e.Column.Header as string) ?? string.Empty;

            if (!float.TryParse(newText, System.Globalization.NumberStyles.Float,
                                System.Globalization.CultureInfo.CurrentCulture, out float newValue))
            {
                MessageBoxHelper.ShowWarning("Giá trị không hợp lệ, vui lòng nhập số");
                e.Cancel = true;
                return;
            }

            // ── Chỉnh sửa cột Tần số (S) ──────────────────────────────
            if (columnHeader == "Tần số (Hz)")
            {
                float oldFreq = editedRow.S;

                // Không thay đổi → bỏ qua
                if (Math.Abs(newValue - oldFreq) < 0.001f) return;

                // Kiểm tra tần số mới đã tồn tại chưa
                bool freqExists = PreviewParams
                    .Any(r => r != editedRow && Math.Abs(r.S - newValue) < 0.001f);

                if (freqExists)
                {
                    MessageBoxHelper.ShowWarning(
                        $"Tần số {newValue:F2} Hz đã tồn tại trong danh sách.\nThay đổi sẽ bị hủy");
                    e.Cancel = true;

                    // Reset lại text trong ô về giá trị cũ
                    textBox.Text = oldFreq.ToString("F0");
                    return;
                }

                // Hợp lệ → cập nhật toàn bộ rows cùng tần số cũ
                // Dùng Dispatcher để tránh xung đột với commit cycle của DataGrid
                float capturedOld = oldFreq;
                float capturedNew = newValue;
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    foreach (var row in PreviewParams.Where(r => Math.Abs(r.S - capturedOld) < 0.001f))
                        row.S = capturedNew;
                }), System.Windows.Threading.DispatcherPriority.Background);
            }

            // ── Chỉnh sửa cột Điểm đo (CV) ────────────────────────────
            else if (columnHeader == "Điểm đo (%)")
            {
                float oldCv = editedRow.CV;

                if (Math.Abs(newValue - oldCv) < 0.001f) return;

                // Kiểm tra cặp (S, CV mới) đã tồn tại chưa
                bool pairExists = PreviewParams
                    .Any(r => r != editedRow
                           && Math.Abs(r.S - editedRow.S) < 0.001f
                           && Math.Abs(r.CV - newValue) < 0.001f);

                if (pairExists)
                {
                    MessageBoxHelper.ShowWarning(
                        $"Điểm đo {newValue:F1}% tại tần số {editedRow.S:F2} Hz đã tồn tại.\nThay đổi sẽ bị hủy");
                    e.Cancel = true;
                    textBox.Text = oldCv.ToString("F0");
                    return;
                }

                // Hợp lệ → DataGrid tự commit, không cần làm thêm
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
