using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using SeperatorAddin.Forms;
using System.Reflection;

namespace SeperatorAddin
{
    [Transaction(TransactionMode.Manual)]
    public class cmdHorizontalSplit : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Create an instance of the event handler and the external event
            SplitterEventHandler handler = new SplitterEventHandler();
            ExternalEvent externalEvent = ExternalEvent.Create(handler);

            // Pass the handler and event to the form
            var form = new frmSplitter(commandData, handler, externalEvent);
            form.ShowDialog();

            return Result.Succeeded;
        }

        internal static PushButtonData GetButtonData()
        {
            string buttonInternalName = "cmdHorizontalSplit";
            string buttonTitle = "Horizontal\nSplit";

            Common.ButtonDataClass myButtonData = new Common.ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Yellow_32,
                Properties.Resources.Yellow_16,
                "Split horizontal elements like Floors, Roofs, Ceilings, and more using model lines.");

            return myButtonData.Data;
        }
    }
}