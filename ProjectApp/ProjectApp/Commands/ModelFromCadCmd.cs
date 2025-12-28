using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ProjectApp.ModelFromCad;
using ProjectApp.Utils;

namespace ProjectApp.Commands
{
    [Transaction(TransactionMode.Manual)]
   public class ModelFromCadCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            AC.GetInformation(commandData, "DATN");
            var view = new ModelFromCadView(commandData.Application.ActiveUIDocument.Document);
            view.ShowDialog();

            return Result.Succeeded;
        }
    }
}
