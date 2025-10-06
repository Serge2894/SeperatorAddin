using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using SeperatorAddin.Forms;
using System.Reflection;

namespace SeperatorAddin
{
    [Transaction(TransactionMode.Manual)]
    public class cmdVerticalSplit : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Create an instance of the event handler and the external event
            VerticalSplitEventHandler handler = new VerticalSplitEventHandler();
            ExternalEvent externalEvent = ExternalEvent.Create(handler);

            // Pass the handler and event to the form
            var form = new frmVerticalSplit(commandData, handler, externalEvent);
            form.ShowDialog();

            return Result.Succeeded;
        }

        internal static PushButtonData GetButtonData()
        {
            string buttonInternalName = "cmdVerticalSplit";
            string buttonTitle = "Vertical\nSplit";

            Common.ButtonDataClass myButtonData = new Common.ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Green_32, 
                Properties.Resources.Green_16,
                "Split vertical elements like Walls, Columns, Ducts, and Pipes by selected levels.");

            return myButtonData.Data;
        }
    }
}