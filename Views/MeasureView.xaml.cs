using System.ComponentModel;
using System.Windows;
using TESMEA_TMS.ViewModels;

namespace TESMEA_TMS.Views
{
    /// <summary>
    /// Interaction logic for MeasureView.xaml
    /// </summary>
    public partial class MeasureView : System.Windows.Controls.UserControl
    {
        float Torque_trend = 0;
        public MeasureView()
        {
            InitializeComponent();
            var vm = (MeasureViewModel)DataContext;
            vm.PropertyChanged += DataGrid_SelectedMeasureChanged;
            this.Loaded += MeasureView_Loaded;
            this.Unloaded += MeasureView_Unloaded;
        }

        private void DataGrid_SelectedMeasureChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(MeasureViewModel.SelectedMeasure))
            {
                var vm = (MeasureViewModel)DataContext;
                if (vm.SelectedMeasure != null)
                    dgvMeasure.ScrollIntoView(vm.SelectedMeasure);
            }
        }

        private void MeasureView_Loaded(object sender, System.Windows.RoutedEventArgs e)
        {

        }

        private void MeasureView_Unloaded(object sender, RoutedEventArgs e)
        {

            if (pvPower != null && pvPower.Model != null)
            {
                pvPower.Model = null;
            }
            if (pvPr_Ef != null && pvPr_Ef.Model != null)
            {
                pvPr_Ef.Model = null;
            }

            this.Unloaded -= MeasureView_Unloaded;
        }
    }
}
