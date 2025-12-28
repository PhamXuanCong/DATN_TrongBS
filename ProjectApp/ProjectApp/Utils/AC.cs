using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace ProjectApp.Utils;

public class AC
{
    public static UIDocument UiDoc;
    public static Document Document;
    public static UIApplication UiApplication;
    public static Autodesk.Revit.UI.Selection.Selection Selection;
    public static Autodesk.Revit.DB.View ActiveView;
    private static ExternalEventHandler externalEventHandler;
    private static ExternalEventHandlers externalEventHandlers;
    private static ExternalEvent externalEvent;
    
    public static void GetInformation(ExternalCommandData data, string currentCommand)
    {
        var uidoc = data.Application.ActiveUIDocument;
        UiDoc = uidoc;
        Document = uidoc.Document;
        UiApplication = uidoc.Application;
        Selection = uidoc.Selection;
        ActiveView = Document.ActiveView;
    }
    
    public static ExternalEvent ExternalEvent
    {
        get
        {
            if (externalEvent == null)
            {
                externalEvent = ExternalEvent.Create(ExternalEventHandler);
            }
            return externalEvent;
        }
        set => externalEvent = value;
    }

    public static ExternalEventHandler ExternalEventHandler
    {
        get
        {
            if (externalEventHandler == null)
            {
                externalEventHandler = new ExternalEventHandler();
            }
            return externalEventHandler;
        }
        set => externalEventHandler = value;
    }

    public static ExternalEventHandlers ExternalEventHandlers
    {
        get
        {
            if (externalEventHandlers == null)
            {
                externalEventHandlers = new ExternalEventHandlers();
            }
            return externalEventHandlers;
        }
        set => externalEventHandlers = value;
    }
}