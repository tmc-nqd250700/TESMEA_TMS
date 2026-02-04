using System.Windows.Controls;
using System.Windows.Input;

namespace TESMEA_TMS.Views
{
    /// <summary>
    /// Interaction logic for ScenarioView.xaml
    /// </summary>
    public partial class ScenarioView : System.Windows.Controls.UserControl
    {
        public ScenarioView()
        {
            InitializeComponent();
        }

        private void ScenarioParamsDataGrid_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            var dataGrid = sender as DataGrid;
            if (e.Key == Key.Tab && dataGrid.CurrentCell != null)
            {
                // Xác định cột "CV" (cột cuối cùng)
                int lastColumnIndex = dataGrid.Columns.Count - 1;
                int sColumnIndex = 1; // Giả sử cột S là cột thứ 1 (sau STT)

                var currentCell = dataGrid.CurrentCell;
                int currentColumnIndex = dataGrid.Columns.IndexOf(currentCell.Column);
                int currentRowIndex = dataGrid.Items.IndexOf(dataGrid.CurrentItem);

                // Chỉ xử lý khi đang ở cột cuối cùng (CV)
                if (currentColumnIndex == lastColumnIndex)
                {
                    // Nếu đang ở dòng cuối cùng (dòng mới)
                    if (currentRowIndex == dataGrid.Items.Count - 2)
                    {
                        // Đợi dòng mới được tạo, rồi focus vào cột S của dòng mới
                        dataGrid.Dispatcher.BeginInvoke(new Action(() =>
                        {
                            dataGrid.CurrentCell = new DataGridCellInfo(
                                dataGrid.Items[dataGrid.Items.Count - 1],
                                dataGrid.Columns[sColumnIndex]
                            );
                            dataGrid.BeginEdit();
                        }), System.Windows.Threading.DispatcherPriority.Background);
                    }
                    else
                    {
                        // Focus vào cột S của dòng tiếp theo
                        dataGrid.Dispatcher.BeginInvoke(new Action(() =>
                        {
                            dataGrid.CurrentCell = new DataGridCellInfo(
                                dataGrid.Items[currentRowIndex + 1],
                                dataGrid.Columns[sColumnIndex]
                            );
                            dataGrid.BeginEdit();
                        }), System.Windows.Threading.DispatcherPriority.Background);
                    }
                    e.Handled = true; // Ngăn Tab mặc định
                }
            }
        }
    }
}
