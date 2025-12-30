using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ProjectApp.ModelFromCad;
using ProjectApp.Utils;

namespace ProjectApp.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class ScheduleVolumeCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            AC.GetInformation(commandData, "Schedule Volume");
            var view = new ScheduleVolumeView(AC.Document);
            view.ShowDialog();

            return Result.Succeeded;
        }
    }
}
