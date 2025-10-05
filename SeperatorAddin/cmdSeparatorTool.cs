using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using SeperatorAddin.Forms;
using System.Reflection;

namespace SeperatorAddin
{
    [Transaction(TransactionMode.Manual)]
    public class cmdSeparator : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Create an instance of the event handler and the external event
            SeparatorEventHandler handler = new SeparatorEventHandler();
            ExternalEvent externalEvent = ExternalEvent.Create(handler);

            // Pass the handler and event to the form
            var form = new frmSeparator(commandData, handler, externalEvent);
            form.ShowDialog();

            return Result.Succeeded;
        }

        internal static PushButtonData GetButtonData()
        {
            string buttonInternalName = "cmdSeparator";
            string buttonTitle = "Separator";

            Common.ButtonDataClass myButtonData = new Common.ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "Separate multi-layer elements into individual layers.");

            return myButtonData.Data;
        }
    }
}