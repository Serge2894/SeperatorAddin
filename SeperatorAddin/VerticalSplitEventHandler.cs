using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using SeperatorAddin.Forms;
using System.Collections.Generic;
using System.Linq;

namespace SeperatorAddin
{
    public class VerticalSplitEventHandler : IExternalEventHandler
    {
        private List<Reference> _columnRefs = new List<Reference>();
        private List<Reference> _wallRefs = new List<Reference>();
        private List<Reference> _ductRefs = new List<Reference>();
        private List<Reference> _pipeRefs = new List<Reference>();
        private List<Level> _selectedLevels = new List<Level>();

        public void SetData(List<Reference> columns, List<Reference> walls, List<Reference> ducts, List<Reference> pipes, List<Level> levels)
        {
            _columnRefs = columns ?? new List<Reference>();
            _wallRefs = walls ?? new List<Reference>();
            _ductRefs = ducts ?? new List<Reference>();
            _pipeRefs = pipes ?? new List<Reference>();
            _selectedLevels = levels ?? new List<Level>();
        }

        public void Execute(UIApplication app)
        {
            UIDocument uidoc = app.ActiveUIDocument;
            Document doc = uidoc.Document;

            if (!_columnRefs.Any() && !_wallRefs.Any() && !_ductRefs.Any() && !_pipeRefs.Any())
            {
                return; // Validation is now in the form
            }

            int totalElementsProcessed = 0;

            using (Transaction t = new Transaction(doc, "Vertically Split Elements"))
            {
                t.Start();

                // Process Columns
                if (_columnRefs.Any())
                {
                    var processor = new cmdColumnSplit();
                    int successCount = 0;
                    foreach (var reference in _columnRefs)
                    {
                        if (doc.GetElement(reference) is Element element)
                        {
                            if (processor.SplitColumn(doc, element, _selectedLevels))
                            {
                                successCount++;
                            }
                        }
                    }
                    totalElementsProcessed += successCount;
                }

                // Process Walls
                if (_wallRefs.Any())
                {
                    var processor = new cmdWallSplit();
                    int successCount = 0;
                    List<ElementId> originalWallsToDelete = new List<ElementId>();
                    foreach (var reference in _wallRefs)
                    {
                        if (doc.GetElement(reference) is Wall wall)
                        {
                            var newWalls = processor.SplitWallByLevels(doc, wall, _selectedLevels);
                            if (newWalls.Any())
                            {
                                originalWallsToDelete.Add(wall.Id);
                                successCount++;
                            }
                        }
                    }
                    if (originalWallsToDelete.Any())
                    {
                        doc.Delete(originalWallsToDelete);
                    }
                    totalElementsProcessed += successCount;
                }

                // Process Ducts
                if (_ductRefs.Any())
                {
                    var processor = new cmdDuctSplit();
                    int successCount = 0;
                    foreach (var reference in _ductRefs)
                    {
                        if (doc.GetElement(reference) is Duct element)
                        {
                            if (processor.SplitDuctByLevels(doc, element, _selectedLevels))
                            {
                                successCount++;
                            }
                        }
                    }
                    totalElementsProcessed += successCount;
                }

                // Process Pipes
                if (_pipeRefs.Any())
                {
                    var processor = new cmdPipeSplit();
                    int successCount = 0;
                    foreach (var reference in _pipeRefs)
                    {
                        if (doc.GetElement(reference) is Pipe element)
                        {
                            if (processor.SplitPipeByLevels(doc, element, _selectedLevels))
                            {
                                successCount++;
                            }
                        }
                    }
                    totalElementsProcessed += successCount;
                }

                t.Commit();
            }

            if (totalElementsProcessed > 0)
            {
                var successDialog = new frmInfoDialog("Elements were split successfully.", "Split Results");
                successDialog.ShowDialog();
            }
        }

        public string GetName() => "Vertical Element Split Event Handler";
    }
}