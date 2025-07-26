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
                // Select ducts to split
                IList<Reference> ductRefs;
                try
                {
                    ISelectionFilter ductFilter = new DuctSelectionFilter();
                    ductRefs = uidoc.Selection.PickObjects(ObjectType.Element, ductFilter, "Select vertical ducts to split");
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    return Result.Cancelled;
                }

                if (ductRefs == null || ductRefs.Count == 0) return Result.Cancelled;
                List<Duct> selectedDucts = ductRefs.Select(r => doc.GetElement(r)).Cast<Duct>().ToList();

                // Get all levels in the project
                var allLevels = new FilteredElementCollector(doc)
                    .OfClass(typeof(Level))
                    .Cast<Level>()
                    .OrderBy(l => l.Elevation)
                    .ToList();

                if (allLevels.Count < 2)
                {
                    TaskDialog.Show("Error", "At least 2 levels are required in the project.");
                    return Result.Cancelled;
                }

                // Create level display names with elevation info
                var levelDisplayNames = allLevels.Select(l => $"{l.Name} (Elev: {l.Elevation:F2}')").ToList();

                // Show form to select levels
                var form = new DuctSelectionForm(levelDisplayNames);
                if (form.ShowDialog() != true)
                {
                    return Result.Cancelled;
                }

                // Get selected levels
                var selectedLevelDisplayNames = form.SelectedLevelNames;
                if (selectedLevelDisplayNames.Count < 1)
                {
                    TaskDialog.Show("Error", "Please select at least 1 level.");
                    return Result.Cancelled;
                }

                // Convert display names back to Level objects
                List<Level> selectedLevels = allLevels
                    .Where(l => selectedLevelDisplayNames.Contains($"{l.Name} (Elev: {l.Elevation:F2}')"))
                    .ToList();

                // Execute splitting
                using (var trans = new Transaction(doc, "Split Ducts by Level"))
                {
                    trans.Start();

                    int successCount = 0;
                    int failCount = 0;
                    List<string> errors = new List<string>();

                    foreach (Duct duct in selectedDucts)
                    {
                        try
                        {
                            if (SplitDuctByLevels(doc, duct, selectedLevels))
                            {
                                successCount++;
                            }
                            else
                            {
                                failCount++;
                                errors.Add($"Failed to split duct ID: {duct.Id}");
                            }
                        }
                        catch (Exception ex)
                        {
                            failCount++;
                            errors.Add($"Duct ID {duct.Id}: {ex.Message}");
                        }
                    }

                    trans.Commit();

                    // Show results
                    string resultMessage = $"Operation complete.\nSuccessfully split: {successCount} ducts\nFailed: {failCount} ducts";

                    if (errors.Count > 0)
                    {
                        resultMessage += "\n\nErrors:\n" + string.Join("\n", errors.Take(10));
                        if (errors.Count > 10)
                            resultMessage += $"\n... and {errors.Count - 10} more errors";
                    }

                    TaskDialog.Show("Split Ducts Result", resultMessage);
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
            // Get duct location
            var locCurve = originalDuct.Location as LocationCurve;
            if (locCurve == null || !(locCurve.Curve is Line ductLine))
                return false;

            // Get start and end points
            XYZ startPoint = ductLine.GetEndPoint(0);
            XYZ endPoint = ductLine.GetEndPoint(1);

            // Ensure start is lower than end
            if (startPoint.Z > endPoint.Z)
            {
                XYZ temp = startPoint;
                startPoint = endPoint;
                endPoint = temp;
            }

            // Check if duct is vertical (within tolerance)
            if (Math.Abs(ductLine.Direction.Z) < 0.9)
                return false;

            // Get duct insulation and lining data before splitting
            InsulationData originalInsulationData = GetDuctInsulation(doc, originalDuct);
            LiningData originalLiningData = GetDuctLining(doc, originalDuct);

            // Store original duct parameters
            var ductType = originalDuct.DuctType;
            var systemTypeId = originalDuct.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM).AsElementId();
            var width = originalDuct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM).AsDouble();
            var height = originalDuct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM).AsDouble();

            // Build elevation list for splitting
            var elevations = new List<double> { startPoint.Z };

            // Add selected level elevations that fall within duct range
            elevations.AddRange(selectedLevels
                .Select(l => l.Elevation)
                .Where(e => e > startPoint.Z + 0.001 && e < endPoint.Z - 0.001));

            elevations.Add(endPoint.Z);
            elevations = elevations.Distinct().OrderBy(e => e).ToList();

            if (elevations.Count < 2) return false;

            // Get the original duct's reference level
            Level originalLevel = doc.GetElement(originalDuct.ReferenceLevel.Id) as Level;
            if (originalLevel == null)
            {
                // If no reference level, get the nearest level
                var allLevels = new FilteredElementCollector(doc)
                    .OfClass(typeof(Level))
                    .Cast<Level>()
                    .OrderBy(l => Math.Abs(l.Elevation - startPoint.Z))
                    .FirstOrDefault();
                originalLevel = allLevels;
            }

            // Create new duct segments
            var newDucts = new List<Duct>();

            for (int i = 0; i < elevations.Count - 1; i++)
            {
                double z1 = elevations[i];
                double z2 = elevations[i + 1];

                if (Math.Abs(z2 - z1) < 0.01) continue; // Skip very small segments

                try
                {
                    // Create duct using exact coordinates
                    XYZ p1 = new XYZ(startPoint.X, startPoint.Y, z1);
                    XYZ p2 = new XYZ(startPoint.X, startPoint.Y, z2);

                    // Create a line for the new duct segment
                    Line segmentLine = Line.CreateBound(p1, p2);

                    // Create the duct
                    Duct newDuct = null;

                    // Use transaction sub-steps to ensure proper creation
                    using (SubTransaction subTrans = new SubTransaction(doc))
                    {
                        subTrans.Start();

                        // Create duct at original level first
                        newDuct = Duct.Create(doc, systemTypeId, ductType.Id, originalLevel.Id, p1, p2);

                        if (newDuct != null)
                        {
                            // Set dimensions immediately
                            newDuct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM).Set(width);
                            newDuct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM).Set(height);

                            // Force the location curve to be exactly what we want
                            LocationCurve newLoc = newDuct.Location as LocationCurve;
                            if (newLoc != null)
                            {
                                newLoc.Curve = segmentLine;
                            }
                        }

                        subTrans.Commit();
                    }

                    if (newDuct != null)
                    {
                        // Copy all parameters from original duct
                        CopyDuctParameters(originalDuct, newDuct);
                        newDucts.Add(newDuct);

                        Debug.WriteLine($"Created duct segment {i + 1}: Bottom={z1:F2}, Top={z2:F2}");
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Failed to create duct segment {i}: {ex.Message}");
                    continue;
                }
            }

            if (newDucts.Count == 0) return false;

            // Handle insulation for split ducts
            if (originalInsulationData != null && originalInsulationData.HasInsulation)
            {
                foreach (var newDuct in newDucts)
                {
                    AddInsulationToDuct(doc, newDuct, originalInsulationData);
                }
            }

            // Handle lining for split ducts
            if (originalLiningData != null && originalLiningData.HasLining)
            {
                foreach (var newDuct in newDucts)
                {
                    AddLiningToDuct(doc, newDuct, originalLiningData);
                }
            }

            // Delete original duct
            doc.Delete(originalDuct.Id);

            return true;
        }

        private void CopyDuctParameters(Duct source, Duct target)
        {
            // Copy important parameters but skip ones that would conflict
            try
            {
                // Copy flow parameters
                var flowParam = source.get_Parameter(BuiltInParameter.RBS_DUCT_FLOW_PARAM);
                if (flowParam != null && flowParam.HasValue)
                {
                    var targetFlowParam = target.get_Parameter(BuiltInParameter.RBS_DUCT_FLOW_PARAM);
                    if (targetFlowParam != null && !targetFlowParam.IsReadOnly)
                    {
                        targetFlowParam.Set(flowParam.AsDouble());
                    }
                }

                // Copy custom parameters
                foreach (Parameter param in source.Parameters)
                {
                    if (param.IsReadOnly || !param.HasValue) continue;

                    // Skip built-in parameters we've already handled or shouldn't copy
                    if (param.Id.IntegerValue < 0) continue;

                    Parameter targetParam = target.get_Parameter(param.Definition);
                    if (targetParam != null && !targetParam.IsReadOnly)
                    {
                        try
                        {
                            switch (param.StorageType)
                            {
                                case StorageType.Double:
                                    targetParam.Set(param.AsDouble());
                                    break;
                                case StorageType.Integer:
                                    targetParam.Set(param.AsInteger());
                                    break;
                                case StorageType.String:
                                    targetParam.Set(param.AsString());
                                    break;
                                case StorageType.ElementId:
                                    targetParam.Set(param.AsElementId());
                                    break;
                            }
                        }
                        catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error copying parameters: {ex.Message}");
            }
        }

        #region Insulation Handling

        private class InsulationData
        {
            public bool HasInsulation { get; set; }
            public ElementId InsulationTypeId { get; set; }
            public double Thickness { get; set; }
            public Dictionary<string, object> Parameters { get; set; }

            public InsulationData()
            {
                Parameters = new Dictionary<string, object>();
            }
        }

        private InsulationData GetDuctInsulation(Document doc, Duct duct)
        {
            InsulationData data = new InsulationData();

            // Find insulation associated with the duct
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            var insulations = collector
                .OfClass(typeof(DuctInsulation))
                .Cast<DuctInsulation>()
                .Where(ins => ins.HostElementId == duct.Id)
                .ToList();

            if (insulations.Count == 0)
            {
                data.HasInsulation = false;
                return data;
            }

            DuctInsulation insulation = insulations.First();
            data.HasInsulation = true;
            data.InsulationTypeId = insulation.GetTypeId();
            data.Thickness = insulation.Thickness;

            // Store all parameters
            foreach (Parameter param in insulation.Parameters)
            {
                if (param == null || param.IsReadOnly) continue;

                string paramName = param.Definition.Name;

                try
                {
                    switch (param.StorageType)
                    {
                        case StorageType.Double:
                            data.Parameters[paramName] = param.AsDouble();
                            break;
                        case StorageType.Integer:
                            data.Parameters[paramName] = param.AsInteger();
                            break;
                        case StorageType.String:
                            data.Parameters[paramName] = param.AsString();
                            break;
                        case StorageType.ElementId:
                            data.Parameters[paramName] = param.AsElementId();
                            break;
                    }
                }
                catch { }
            }

            return data;
        }

        private DuctInsulation AddInsulationToDuct(Document doc, Duct duct, InsulationData insulationData)
        {
            try
            {
                // Check if duct already has insulation
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                var existingInsulation = collector
                    .OfClass(typeof(DuctInsulation))
                    .Cast<DuctInsulation>()
                    .FirstOrDefault(ins => ins.HostElementId == duct.Id);

                if (existingInsulation != null)
                {
                    // Update existing insulation
                    CopyInsulationParameters(existingInsulation, insulationData);
                    return existingInsulation;
                }

                // Create new insulation
                DuctInsulation newInsulation = DuctInsulation.Create(
                    doc, duct.Id, insulationData.InsulationTypeId, insulationData.Thickness);

                if (newInsulation != null)
                {
                    CopyInsulationParameters(newInsulation, insulationData);
                }

                return newInsulation;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to add insulation to duct: {ex.Message}");
                return null;
            }
        }

        private void CopyInsulationParameters(DuctInsulation targetInsulation, InsulationData sourceData)
        {
            foreach (var kvp in sourceData.Parameters)
            {
                try
                {
                    Parameter param = targetInsulation.LookupParameter(kvp.Key);
                    if (param == null || param.IsReadOnly) continue;

                    switch (param.StorageType)
                    {
                        case StorageType.Double:
                            if (kvp.Value is double)
                                param.Set((double)kvp.Value);
                            break;
                        case StorageType.Integer:
                            if (kvp.Value is int)
                                param.Set((int)kvp.Value);
                            break;
                        case StorageType.String:
                            if (kvp.Value is string)
                                param.Set((string)kvp.Value);
                            break;
                        case StorageType.ElementId:
                            if (kvp.Value is ElementId)
                                param.Set((ElementId)kvp.Value);
                            break;
                    }
                }
                catch { }
            }
        }

        #endregion

        #region Lining Handling

        private class LiningData
        {
            public bool HasLining { get; set; }
            public ElementId LiningTypeId { get; set; }
            public double Thickness { get; set; }
            public Dictionary<string, object> Parameters { get; set; }

            public LiningData()
            {
                Parameters = new Dictionary<string, object>();
            }
        }

        private LiningData GetDuctLining(Document doc, Duct duct)
        {
            LiningData data = new LiningData();

            // Find lining associated with the duct
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            var linings = collector
                .OfClass(typeof(DuctLining))
                .Cast<DuctLining>()
                .Where(lining => lining.HostElementId == duct.Id)
                .ToList();

            if (linings.Count == 0)
            {
                data.HasLining = false;
                return data;
            }

            DuctLining lining = linings.First();
            data.HasLining = true;
            data.LiningTypeId = lining.GetTypeId();
            data.Thickness = lining.Thickness;

            // Store all parameters
            foreach (Parameter param in lining.Parameters)
            {
                if (param == null || param.IsReadOnly) continue;

                string paramName = param.Definition.Name;

                try
                {
                    switch (param.StorageType)
                    {
                        case StorageType.Double:
                            data.Parameters[paramName] = param.AsDouble();
                            break;
                        case StorageType.Integer:
                            data.Parameters[paramName] = param.AsInteger();
                            break;
                        case StorageType.String:
                            data.Parameters[paramName] = param.AsString();
                            break;
                        case StorageType.ElementId:
                            data.Parameters[paramName] = param.AsElementId();
                            break;
                    }
                }
                catch { }
            }

            return data;
        }

        private DuctLining AddLiningToDuct(Document doc, Duct duct, LiningData liningData)
        {
            try
            {
                // Check if duct already has lining
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                var existingLining = collector
                    .OfClass(typeof(DuctLining))
                    .Cast<DuctLining>()
                    .FirstOrDefault(lining => lining.HostElementId == duct.Id);

                if (existingLining != null)
                {
                    // Update existing lining
                    CopyLiningParameters(existingLining, liningData);
                    return existingLining;
                }

                // Create new lining
                DuctLining newLining = DuctLining.Create(
                    doc, duct.Id, liningData.LiningTypeId, liningData.Thickness);

                if (newLining != null)
                {
                    CopyLiningParameters(newLining, liningData);
                }

                return newLining;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to add lining to duct: {ex.Message}");
                return null;
            }
        }

        private void CopyLiningParameters(DuctLining targetLining, LiningData sourceData)
        {
            foreach (var kvp in sourceData.Parameters)
            {
                try
                {
                    Parameter param = targetLining.LookupParameter(kvp.Key);
                    if (param == null || param.IsReadOnly) continue;

                    switch (param.StorageType)
                    {
                        case StorageType.Double:
                            if (kvp.Value is double)
                                param.Set((double)kvp.Value);
                            break;
                        case StorageType.Integer:
                            if (kvp.Value is int)
                                param.Set((int)kvp.Value);
                            break;
                        case StorageType.String:
                            if (kvp.Value is string)
                                param.Set((string)kvp.Value);
                            break;
                        case StorageType.ElementId:
                            if (kvp.Value is ElementId)
                                param.Set((ElementId)kvp.Value);
                            break;
                    }
                }
                catch { }
            }
        }

        #endregion
    }
}