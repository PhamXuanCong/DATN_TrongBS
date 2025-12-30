using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Autodesk.Revit.DB;

namespace ProjectApp.ModelFromCad
{
    /// <summary>
    /// Interaction logic for ScheduleVolumeView.xaml
    /// </summary>
    public partial class ScheduleVolumeView : Window
    {
        public List<ScheduleInfo> ScheduleInfos { get; set; } = [];

        public ScheduleVolumeView(Document doc)
        {
            InitializeComponent();

            // Lấy khối lượng cột
            var columns = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_StructuralColumns)
                .WhereElementIsNotElementType()
                .Cast<FamilyInstance>()
                .ToList();

            var sumVolCol = 0.0;

            foreach (var col in columns)
            {
                double v = GetElementVolumeFt3(col);
                if (v > 0) sumVolCol += v;
            }

            ScheduleInfos.Add(new ScheduleInfo()
            {
                Name = "Tổng khối lượng cột",
                Volume = Math.Round(sumVolCol , 2) 
            });


            // Lấy khối lượng dầm
            var beams = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_StructuralFraming)
                .WhereElementIsNotElementType()
                .Cast<FamilyInstance>()
                .ToList();
            var sumVolBeam = 0.0;
            foreach (var beam in beams)
            {
                double v = GetElementVolumeFt3(beam);
                if (v > 0) sumVolBeam += v;
            }

            ScheduleInfos.Add(new ScheduleInfo()
            {
                Name = "Tổng khối lượng dầm",
                Volume = Math.Round(sumVolBeam , 2) 
            });

            // Lấy khối lượng sàn
            var floors = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Floors)
                .WhereElementIsNotElementType()
                .Cast<Floor>()
                .ToList();
            var sumVolFloor = 0.0;
            foreach (var floor in floors)
            {
                double v = GetElementVolumeFt3(floor);
                if (v > 0) sumVolFloor += v;
            }

            ScheduleInfos.Add(new ScheduleInfo()
            {
                Name = "Tổng khối lượng sàn",
                Volume = Math.Round(sumVolFloor , 2) 
            });

            // Khối lượng tổng = khối lượng cột + dầm + sàn
            ScheduleInfos.Add(new ScheduleInfo()
            {
                Name = "Tổng khối lượng (cột + dầm + sàn)",
                Volume = Math.Round(sumVolCol , 2)  + Math.Round(sumVolBeam ,2) + Math.Round(sumVolFloor, 2)
            });

            dgSchedule.ItemsSource = ScheduleInfos;
        }

        private static double GetElementVolumeFt3(Element e)
        {
            var p =
                e.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED);

            if (p == null)
                p = e.LookupParameter("Volume");

            if (p == null || p.StorageType != StorageType.Double || !p.HasValue)
                return 0.0;

            double v = p.AsDouble(); // ft^3
            if (double.IsNaN(v) || double.IsInfinity(v) || v < 0) return 0.0;

            return UnitUtils.ConvertFromInternalUnits(v, UnitTypeId.CubicMeters) ;
        }

        private void ButtonBase_OnClick(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }

    public class ScheduleInfo
    {
        public string TypeName { get; set; }
        public string Name { get; set; }
        public double Volume { get; set; }
    }
}