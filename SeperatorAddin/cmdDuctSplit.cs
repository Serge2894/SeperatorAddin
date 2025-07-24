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
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                IList<Reference> ductRefs;
                try
                {
                    ISelectionFilter ductFilter = new DuctSelectionFilter();
                    ductRefs = uidoc.Selection.PickObjects(ObjectType.Element, ductFilter, "Select vertical ducts to split");
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException) { return Result.Cancelled; }

                if (ductRefs == null || ductRefs.Count == 0) return Result.Cancelled;
                List<Duct> selectedDucts = ductRefs.Select(r => doc.GetElement(r)).Cast<Duct>().ToList();

                var allLevels = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().OrderBy(l => l.Elevation).ToList();
                if (allLevels.Count < 2)
                {
                    TaskDialog.Show("Error", "At least 2 levels are required in the project.");
                    return Result.Cancelled;
                }

                var levelDisplayNames = allLevels.Select(l => $"{l.Name} (Elev: {l.Elevation:F2}')").ToList();
                var form = new DuctSelectionForm(levelDisplayNames);
                if (form.ShowDialog() != true)
                {
                    return Result.Cancelled;
                }

                var selectedLevelDisplayNames = form.SelectedLevelNames;
                if (selectedLevelDisplayNames.Count < 2)
                {
                    TaskDialog.Show("Error", "Please select at least 2 levels.");
                    return Result.Cancelled;
                }
                List<Level> selectedLevels = allLevels.Where(l => selectedLevelDisplayNames.Contains($"{l.Name} (Elev: {l.Elevation:F2}')")).ToList();

                using (var trans = new Transaction(doc, "Split Ducts by Level"))
                {
                    trans.Start();
                    int successCount = 0;
                    int failCount = 0;
                    foreach (Duct duct in selectedDucts)
                    {
                        if (SplitDuctByLevels(doc, duct, selectedLevels))
                            successCount++;
                        else
                            failCount++;
                    }
                    trans.Commit();
                    TaskDialog.Show("Success", $"Operation complete.\nSuccessfully split: {successCount}\nFailed to split: {failCount}");
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private bool SplitDuctByLevels(Document doc, Duct originalDuct, List<Level> selectedLevels)
        {
            var locCurve = originalDuct.Location as LocationCurve;
            if (locCurve == null || !(locCurve.Curve is Line ductLine)) return false;

            XYZ startPoint = ductLine.GetEndPoint(0).Z < ductLine.GetEndPoint(1).Z ? ductLine.GetEndPoint(0) : ductLine.GetEndPoint(1);
            XYZ endPoint = ductLine.GetEndPoint(0).Z < ductLine.GetEndPoint(1).Z ? ductLine.GetEndPoint(1) : ductLine.GetEndPoint(0);

            if (Math.Abs(ductLine.Direction.Z) < 0.9) return false;

            var elevations = new List<double> { startPoint.Z };
            elevations.AddRange(selectedLevels.Select(l => l.Elevation)
                .Where(e => e > startPoint.Z + 0.001 && e < endPoint.Z - 0.001));
            elevations.Add(endPoint.Z);
            elevations = elevations.Distinct().OrderBy(e => e).ToList();

            if (elevations.Count < 2) return false;

            var ductType = originalDuct.DuctType;
            var systemTypeId = originalDuct.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM).AsElementId();
            var levelId = originalDuct.ReferenceLevel.Id;
            var width = originalDuct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM).AsDouble();
            var height = originalDuct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM).AsDouble();

            var newDucts = new List<Duct>();
            for (int i = 0; i < elevations.Count - 1; i++)
            {
                double z1 = elevations[i];
                double z2 = elevations[i + 1];

                if (Math.Abs(z2 - z1) < 0.01) continue;

                XYZ p1 = new XYZ(startPoint.X, startPoint.Y, z1);
                XYZ p2 = new XYZ(startPoint.X, startPoint.Y, z2);

                Duct newDuct = Duct.Create(doc, systemTypeId, ductType.Id, levelId, p1, p2);
                newDuct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM).Set(width);
                newDuct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM).Set(height);
                CopyAllDuctParameters(originalDuct, newDuct);
                newDucts.Add(newDuct);
            }

            if (newDucts.Any())
            {
                doc.Delete(originalDuct.Id);
                return true;
            }
            return false;
        }

        private void CopyAllDuctParameters(Duct source, Duct target)
        {
            var ignoredParams = new List<BuiltInParameter> { BuiltInParameter.CURVE_ELEM_LENGTH, BuiltInParameter.RBS_CURVE_WIDTH_PARAM, BuiltInParameter.RBS_CURVE_HEIGHT_PARAM };

            foreach (Parameter sourceParam in source.Parameters)
            {
                if (sourceParam.IsReadOnly || !sourceParam.HasValue) continue;
                if (Enum.IsDefined(typeof(BuiltInParameter), sourceParam.Id.IntegerValue))
                {
                    if (ignoredParams.Contains((BuiltInParameter)sourceParam.Id.IntegerValue)) continue;
                }

                Parameter targetParam = target.get_Parameter(sourceParam.Definition);
                if (targetParam != null && !targetParam.IsReadOnly)
                {
                    try
                    {
                        // **FIXED CODE BLOCK**
                        switch (sourceParam.StorageType)
                        {
                            case StorageType.Double:
                                targetParam.Set(sourceParam.AsDouble());
                                break;
                            case StorageType.Integer:
                                targetParam.Set(sourceParam.AsInteger());
                                break;
                            case StorageType.String:
                                targetParam.Set(sourceParam.AsString());
                                break;
                            case StorageType.ElementId:
                                targetParam.Set(sourceParam.AsElementId());
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Could not set parameter {sourceParam.Definition.Name}: {ex.Message}");
                    }
                }
            }
        }
    }
}