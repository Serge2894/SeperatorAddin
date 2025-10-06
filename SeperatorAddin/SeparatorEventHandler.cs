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

        public void SetData(List<Reference> floors, List<Reference> roofs, List<Reference> walls, List<Reference> ceilings)
        {
            _floorRefs = floors ?? new List<Reference>();
            _roofRefs = roofs ?? new List<Reference>();
            _wallRefs = walls ?? new List<Reference>();
            _ceilingRefs = ceilings ?? new List<Reference>();
        }

        public void Execute(UIApplication app)
        {
            UIDocument uidoc = app.ActiveUIDocument;
            Document doc = uidoc.Document;

            // REMOVED: Redundant check, as it's now handled in the form.

            using (Transaction t = new Transaction(doc, "Separate Element Layers"))
            {
                t.Start();

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

                if (_wallRefs.Any())
                {
                    var processor = new cmdWallSeperator();
                    processor.ProcessWallSeparation(doc, _wallRefs);
                }

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

            var successDialog = new frmInfoDialog("Elements separated successfully.", "Separator");
            successDialog.ShowDialog();
        }

        public string GetName() => "Element Separator Event Handler";
    }
}