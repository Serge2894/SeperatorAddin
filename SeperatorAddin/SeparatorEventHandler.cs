using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using SeperatorAddin.Forms;
using System.Collections.Generic;
using System.Linq;

namespace SeperatorAddin
{
    public class SeparatorEventHandler : IExternalEventHandler
    {
        private List<Reference> _floorRefs = new List<Reference>();
        private List<Reference> _roofRefs = new List<Reference>();
        private List<Reference> _wallRefs = new List<Reference>();
        private List<Reference> _ceilingRefs = new List<Reference>();

        /// <summary>
        /// Sets the data (selected element references) for the handler to process.
        /// </summary>
        public void SetData(List<Reference> floors, List<Reference> roofs, List<Reference> walls, List<Reference> ceilings)
        {
            _floorRefs = floors ?? new List<Reference>();
            _roofRefs = roofs ?? new List<Reference>();
            _wallRefs = walls ?? new List<Reference>();
            _ceilingRefs = ceilings ?? new List<Reference>();
        }

        /// <summary>
        /// This method is executed by Revit in a valid API context when the ExternalEvent is raised.
        /// </summary>
        public void Execute(UIApplication app)
        {
            UIDocument uidoc = app.ActiveUIDocument;
            Document doc = uidoc.Document;

            if (!_floorRefs.Any() && !_roofRefs.Any() && !_wallRefs.Any() && !_ceilingRefs.Any())
            {
                TaskDialog.Show("Separator", "No elements were selected for separation.");
                return;
            }

            using (Transaction t = new Transaction(doc, "Separate Element Layers"))
            {
                t.Start();

                // Process Floors if any were selected
                if (_floorRefs.Any())
                {
                    var processor = new cmdFloorSeperator();
                    foreach (var reference in _floorRefs)
                    {
                        if (doc.GetElement(reference) is Floor element)
                        {
                            processor.ProcessFloor(doc, element);
                        }
                    }
                }

                // Process Roofs if any were selected
                if (_roofRefs.Any())
                {
                    var processor = new cmdRoofSeperator();
                    foreach (var reference in _roofRefs)
                    {
                        if (doc.GetElement(reference) is RoofBase element)
                        {
                            processor.ProcessRoof(doc, element);
                        }
                    }
                }

                // Process Walls if any were selected
                if (_wallRefs.Any())
                {
                    var processor = new cmdWallSeperator();
                    processor.ProcessWallSeparation(doc, _wallRefs);
                }

                // Process Ceilings if any were selected
                if (_ceilingRefs.Any())
                {
                    var processor = new cmdCeilingSeperator();
                    foreach (var reference in _ceilingRefs)
                    {
                        if (doc.GetElement(reference) is Ceiling element)
                        {
                            processor.ProcessCeiling(doc, element);
                        }
                    }
                }

                t.Commit();
            }

            // Show the success dialog after the transaction is complete
            var successDialog = new frmSuccessDialog("Elements modified successfully.");
            successDialog.ShowDialog();
        }

        public string GetName() => "Element Separator Event Handler";
    }
}