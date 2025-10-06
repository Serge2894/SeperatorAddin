using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
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
    public class cmdDuctSplit : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            TaskDialog.Show("Command Info", "This function is now part of the 'Vertical Split' tool.");
            return Result.Succeeded;
        }

        public bool SplitDuctByLevels(Document doc, Duct originalDuct, List<Level> selectedLevels)
        {
            var locCurve = originalDuct.Location as LocationCurve;
            if (locCurve == null || !(locCurve.Curve is Line ductLine)) return false;

            if (Math.Abs(ductLine.Direction.Z) < 0.9) return false;

            XYZ startPoint = ductLine.GetEndPoint(0);
            XYZ endPoint = ductLine.GetEndPoint(1);

            if (startPoint.Z > endPoint.Z) (startPoint, endPoint) = (endPoint, startPoint);

            InsulationData originalInsulationData = GetDuctInsulation(doc, originalDuct);
            LiningData originalLiningData = GetDuctLining(doc, originalDuct);

            var ductType = originalDuct.DuctType;
            var systemTypeId = originalDuct.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM).AsElementId();
            var width = originalDuct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM).AsDouble();
            var height = originalDuct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM).AsDouble();

            var elevations = new List<double> { startPoint.Z };
            elevations.AddRange(selectedLevels
                .Select(l => l.Elevation)
                .Where(e => e > startPoint.Z + 0.001 && e < endPoint.Z - 0.001));
            elevations.Add(endPoint.Z);
            elevations = elevations.Distinct().OrderBy(e => e).ToList();

            if (elevations.Count < 2) return false;

            Level originalLevel = doc.GetElement(originalDuct.ReferenceLevel.Id) as Level;
            if (originalLevel == null)
            {
                originalLevel = new FilteredElementCollector(doc)
                    .OfClass(typeof(Level)).Cast<Level>()
                    .OrderBy(l => Math.Abs(l.Elevation - startPoint.Z)).FirstOrDefault();
            }

            var newDucts = new List<Duct>();
            for (int i = 0; i < elevations.Count - 1; i++)
            {
                double z1 = elevations[i];
                double z2 = elevations[i + 1];
                if (Math.Abs(z2 - z1) < 0.01) continue;

                try
                {
                    XYZ p1 = new XYZ(startPoint.X, startPoint.Y, z1);
                    XYZ p2 = new XYZ(startPoint.X, startPoint.Y, z2);

                    Duct newDuct = Duct.Create(doc, systemTypeId, ductType.Id, originalLevel.Id, p1, p2);

                    if (newDuct != null)
                    {
                        newDuct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM).Set(width);
                        newDuct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM).Set(height);
                        CopyDuctParameters(originalDuct, newDuct);
                        newDucts.Add(newDuct);
                    }
                }
                catch { continue; }
            }

            if (newDucts.Count == 0) return false;

            if (originalInsulationData != null && originalInsulationData.HasInsulation)
            {
                foreach (var newDuct in newDucts) AddInsulationToDuct(doc, newDuct, originalInsulationData);
            }
            if (originalLiningData != null && originalLiningData.HasLining)
            {
                foreach (var newDuct in newDucts) AddLiningToDuct(doc, newDuct, originalLiningData);
            }

            doc.Delete(originalDuct.Id);
            return true;
        }

        // --- PASTE ALL HELPER METHODS FROM THE ORIGINAL cmdDuctSplit.cs HERE ---
        // (CopyDuctParameters, InsulationData class, GetDuctInsulation, AddInsulationToDuct, CopyInsulationParameters, LiningData class, GetDuctLining, AddLiningToDuct, CopyLiningParameters)
        #region Helper Classes and Methods
        private void CopyDuctParameters(Duct source, Duct target) { /* Original implementation */ }
        private class InsulationData { public bool HasInsulation { get; set; } public ElementId InsulationTypeId { get; set; } public double Thickness { get; set; } public Dictionary<string, object> Parameters { get; set; } = new Dictionary<string, object>(); }
        private InsulationData GetDuctInsulation(Document doc, Duct duct) { /* Original implementation */ return new InsulationData(); }
        private void AddInsulationToDuct(Document doc, Duct newDuct, InsulationData originalInsulationData) { /* Original implementation */ }
        private class LiningData { public bool HasLining { get; set; } public ElementId LiningTypeId { get; set; } public double Thickness { get; set; } public Dictionary<string, object> Parameters { get; set; } = new Dictionary<string, object>(); }
        private LiningData GetDuctLining(Document doc, Duct duct) { /* Original implementation */ return new LiningData(); }
        private void AddLiningToDuct(Document doc, Duct newDuct, LiningData originalLiningData) { /* Original implementation */ }
        #endregion
    }
}