using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using SeperatorAddin.Forms;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace SeperatorAddin
{
    [Transaction(TransactionMode.Manual)]
    public class cmdPipeSplit : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            TaskDialog.Show("Command Info", "This function is now part of the 'Vertical Split' tool.");
            return Result.Succeeded;
        }

        public bool SplitPipeByLevels(Document doc, Pipe originalPipe, List<Level> selectedLevels)
        {
            var locCurve = originalPipe.Location as LocationCurve;
            if (locCurve == null || !(locCurve.Curve is Line pipeLine)) return false;

            if (Math.Abs(pipeLine.Direction.Z) < 0.9) return false;

            XYZ startPoint = pipeLine.GetEndPoint(0);
            XYZ endPoint = pipeLine.GetEndPoint(1);

            if (startPoint.Z > endPoint.Z) (startPoint, endPoint) = (endPoint, startPoint);

            InsulationData originalInsulationData = GetPipeInsulation(doc, originalPipe);

            var pipeType = originalPipe.PipeType;
            var systemTypeId = originalPipe.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId();
            var diameter = originalPipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble();

            var elevations = new List<double> { startPoint.Z };
            elevations.AddRange(selectedLevels
                .Select(l => l.Elevation)
                .Where(e => e > startPoint.Z + 0.001 && e < endPoint.Z - 0.001));
            elevations.Add(endPoint.Z);
            elevations = elevations.Distinct().OrderBy(e => e).ToList();

            if (elevations.Count < 2) return false;

            Level originalLevel = doc.GetElement(originalPipe.ReferenceLevel.Id) as Level;
            if (originalLevel == null)
            {
                originalLevel = new FilteredElementCollector(doc)
                   .OfClass(typeof(Level)).Cast<Level>()
                   .OrderBy(l => Math.Abs(l.Elevation - startPoint.Z)).FirstOrDefault();
            }

            var newPipes = new List<Pipe>();
            for (int i = 0; i < elevations.Count - 1; i++)
            {
                double z1 = elevations[i];
                double z2 = elevations[i + 1];
                if (Math.Abs(z2 - z1) < 0.01) continue;

                try
                {
                    XYZ p1 = new XYZ(startPoint.X, startPoint.Y, z1);
                    XYZ p2 = new XYZ(startPoint.X, startPoint.Y, z2);
                    Pipe newPipe = Pipe.Create(doc, systemTypeId, pipeType.Id, originalLevel.Id, p1, p2);

                    if (newPipe != null)
                    {
                        newPipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).Set(diameter);
                        CopyPipeParameters(originalPipe, newPipe);
                        newPipes.Add(newPipe);
                    }
                }
                catch { continue; }
            }

            if (newPipes.Count == 0) return false;

            if (originalInsulationData != null && originalInsulationData.HasInsulation)
            {
                foreach (var newPipe in newPipes) AddInsulationToPipe(doc, newPipe, originalInsulationData);
            }

            doc.Delete(originalPipe.Id);
            return true;
        }

        // --- PASTE ALL HELPER METHODS FROM THE ORIGINAL cmdPipeSplit.cs HERE ---
        // (CopyPipeParameters, InsulationData class, GetPipeInsulation, AddInsulationToPipe, CopyInsulationParameters)
        #region Helper Classes and Methods
        private void CopyPipeParameters(Pipe source, Pipe target) { /* Original implementation */ }
        private class InsulationData { public bool HasInsulation { get; set; } public ElementId InsulationTypeId { get; set; } public double Thickness { get; set; } public Dictionary<string, object> Parameters { get; set; } = new Dictionary<string, object>(); }
        private InsulationData GetPipeInsulation(Document doc, Pipe pipe) { /* Original implementation */ return new InsulationData(); }
        private void AddInsulationToPipe(Document doc, Pipe pipe, InsulationData insulationData) { /* Original implementation */ }
        private void CopyInsulationParameters(PipeInsulation targetInsulation, InsulationData sourceData) { /* Original implementation */ }
        #endregion
    }
}