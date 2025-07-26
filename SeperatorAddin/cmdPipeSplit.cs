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
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Select pipes to split
                IList<Reference> pipeRefs;
                try
                {
                    ISelectionFilter pipeFilter = new PipeSelectionFilter();
                    pipeRefs = uidoc.Selection.PickObjects(ObjectType.Element, pipeFilter, "Select vertical pipes to split");
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    return Result.Cancelled;
                }

                if (pipeRefs == null || pipeRefs.Count == 0) return Result.Cancelled;
                List<Pipe> selectedPipes = pipeRefs.Select(r => doc.GetElement(r)).Cast<Pipe>().ToList();

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
                var form = new PipeSelectionForm(levelDisplayNames);
                if (form.ShowDialog() != true)
                {
                    return Result.Cancelled;
                }

                // Get selected levels
                var selectedLevelDisplayNames = form.SelectedLevelNames;
                if (selectedLevelDisplayNames.Count < 2)
                {
                    TaskDialog.Show("Error", "Please select at least 2 levels.");
                    return Result.Cancelled;
                }

                // Convert display names back to Level objects
                List<Level> selectedLevels = allLevels
                    .Where(l => selectedLevelDisplayNames.Contains($"{l.Name} (Elev: {l.Elevation:F2}')"))
                    .ToList();

                // Execute splitting
                using (var trans = new Transaction(doc, "Split Pipes by Level"))
                {
                    trans.Start();

                    int successCount = 0;
                    int failCount = 0;
                    List<string> errors = new List<string>();

                    foreach (Pipe pipe in selectedPipes)
                    {
                        try
                        {
                            if (SplitPipeByLevels(doc, pipe, selectedLevels))
                            {
                                successCount++;
                            }
                            else
                            {
                                failCount++;
                                errors.Add($"Failed to split pipe ID: {pipe.Id}");
                            }
                        }
                        catch (Exception ex)
                        {
                            failCount++;
                            errors.Add($"Pipe ID {pipe.Id}: {ex.Message}");
                        }
                    }

                    trans.Commit();

                    // Show results
                    string resultMessage = $"Operation complete.\nSuccessfully split: {successCount} pipes\nFailed: {failCount} pipes";

                    if (errors.Count > 0)
                    {
                        resultMessage += "\n\nErrors:\n" + string.Join("\n", errors.Take(10));
                        if (errors.Count > 10)
                            resultMessage += $"\n... and {errors.Count - 10} more errors";
                    }

                    TaskDialog.Show("Split Pipes Result", resultMessage);
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private bool SplitPipeByLevels(Document doc, Pipe originalPipe, List<Level> selectedLevels)
        {
            // Get pipe location
            var locCurve = originalPipe.Location as LocationCurve;
            if (locCurve == null || !(locCurve.Curve is Line pipeLine))
                return false;

            // Get start and end points
            XYZ startPoint = pipeLine.GetEndPoint(0);
            XYZ endPoint = pipeLine.GetEndPoint(1);

            // Ensure start is lower than end
            if (startPoint.Z > endPoint.Z)
            {
                XYZ temp = startPoint;
                startPoint = endPoint;
                endPoint = temp;
            }

            // Check if pipe is vertical (within tolerance)
            if (Math.Abs(pipeLine.Direction.Z) < 0.9)
                return false;

            // Get pipe insulation data before splitting
            InsulationData originalInsulationData = GetPipeInsulation(doc, originalPipe);

            // Store original pipe parameters
            var pipeType = originalPipe.PipeType;
            var systemTypeId = originalPipe.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId();
            var diameter = originalPipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble();

            // Build elevation list for splitting
            var elevations = new List<double> { startPoint.Z };

            // Add selected level elevations that fall within pipe range
            elevations.AddRange(selectedLevels
                .Select(l => l.Elevation)
                .Where(e => e > startPoint.Z + 0.001 && e < endPoint.Z - 0.001));

            elevations.Add(endPoint.Z);
            elevations = elevations.Distinct().OrderBy(e => e).ToList();

            if (elevations.Count < 2) return false;

            // Get the original pipe's reference level
            Level originalLevel = doc.GetElement(originalPipe.ReferenceLevel.Id) as Level;
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

            // Create new pipe segments
            var newPipes = new List<Pipe>();

            for (int i = 0; i < elevations.Count - 1; i++)
            {
                double z1 = elevations[i];
                double z2 = elevations[i + 1];

                if (Math.Abs(z2 - z1) < 0.01) continue; // Skip very small segments

                try
                {
                    // Create pipe using PlumbingUtils.BreakCurve approach
                    // First create a temporary pipe with the exact coordinates we need
                    XYZ p1 = new XYZ(startPoint.X, startPoint.Y, z1);
                    XYZ p2 = new XYZ(startPoint.X, startPoint.Y, z2);

                    // Create a line for the new pipe segment
                    Line segmentLine = Line.CreateBound(p1, p2);

                    // Create the pipe using the line
                    Pipe newPipe = null;

                    // Try to create pipe with exact coordinates
                    // Use transaction sub-steps to ensure proper creation
                    using (SubTransaction subTrans = new SubTransaction(doc))
                    {
                        subTrans.Start();

                        // Create pipe at original level first
                        newPipe = Pipe.Create(doc, systemTypeId, pipeType.Id, originalLevel.Id, p1, p2);

                        if (newPipe != null)
                        {
                            // Set diameter immediately
                            newPipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).Set(diameter);

                            // Force the location curve to be exactly what we want
                            LocationCurve newLoc = newPipe.Location as LocationCurve;
                            if (newLoc != null)
                            {
                                newLoc.Curve = segmentLine;
                            }
                        }

                        subTrans.Commit();
                    }

                    if (newPipe != null)
                    {
                        // Copy all parameters from original pipe
                        CopyPipeParameters(originalPipe, newPipe);
                        newPipes.Add(newPipe);

                        Debug.WriteLine($"Created pipe segment {i + 1}: Bottom={z1:F2}, Top={z2:F2}");
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Failed to create pipe segment {i}: {ex.Message}");
                    continue;
                }
            }

            if (newPipes.Count == 0) return false;

            // Handle insulation for split pipes
            if (originalInsulationData != null && originalInsulationData.HasInsulation)
            {
                foreach (var newPipe in newPipes)
                {
                    AddInsulationToPipe(doc, newPipe, originalInsulationData);
                }
            }

            // Delete original pipe
            doc.Delete(originalPipe.Id);

            return true;
        }

        private void CopyPipeParameters(Pipe source, Pipe target)
        {
            // Copy important parameters but skip ones that would conflict
            try
            {
                // Copy slope if applicable
                var slopeParam = source.get_Parameter(BuiltInParameter.RBS_PIPE_SLOPE);
                if (slopeParam != null && slopeParam.HasValue)
                {
                    var targetSlopeParam = target.get_Parameter(BuiltInParameter.RBS_PIPE_SLOPE);
                    if (targetSlopeParam != null && !targetSlopeParam.IsReadOnly)
                    {
                        targetSlopeParam.Set(slopeParam.AsDouble());
                    }
                }

                // Copy insulation thickness if present
                var insulationParam = source.get_Parameter(BuiltInParameter.RBS_PIPE_INSULATION_THICKNESS);
                if (insulationParam != null && insulationParam.HasValue)
                {
                    var targetInsulationParam = target.get_Parameter(BuiltInParameter.RBS_PIPE_INSULATION_THICKNESS);
                    if (targetInsulationParam != null && !targetInsulationParam.IsReadOnly)
                    {
                        targetInsulationParam.Set(insulationParam.AsDouble());
                    }
                }

                // Copy flow parameters
                var flowParam = source.get_Parameter(BuiltInParameter.RBS_PIPE_FLOW_PARAM);
                if (flowParam != null && flowParam.HasValue)
                {
                    var targetFlowParam = target.get_Parameter(BuiltInParameter.RBS_PIPE_FLOW_PARAM);
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

        private InsulationData GetPipeInsulation(Document doc, Pipe pipe)
        {
            InsulationData data = new InsulationData();

            // Find insulation associated with the pipe
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            var insulations = collector
                .OfClass(typeof(PipeInsulation))
                .Cast<PipeInsulation>()
                .Where(ins => ins.HostElementId == pipe.Id)
                .ToList();

            if (insulations.Count == 0)
            {
                data.HasInsulation = false;
                return data;
            }

            PipeInsulation insulation = insulations.First();
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

        private PipeInsulation AddInsulationToPipe(Document doc, Pipe pipe, InsulationData insulationData)
        {
            try
            {
                // Check if pipe already has insulation
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                var existingInsulation = collector
                    .OfClass(typeof(PipeInsulation))
                    .Cast<PipeInsulation>()
                    .FirstOrDefault(ins => ins.HostElementId == pipe.Id);

                if (existingInsulation != null)
                {
                    // Update existing insulation
                    CopyInsulationParameters(existingInsulation, insulationData);
                    return existingInsulation;
                }

                // Create new insulation
                PipeInsulation newInsulation = PipeInsulation.Create(
                    doc, pipe.Id, insulationData.InsulationTypeId, insulationData.Thickness);

                if (newInsulation != null)
                {
                    CopyInsulationParameters(newInsulation, insulationData);
                }

                return newInsulation;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to add insulation to pipe: {ex.Message}");
                return null;
            }
        }

        private void CopyInsulationParameters(PipeInsulation targetInsulation, InsulationData sourceData)
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

        #region Parameter Handling

        private Dictionary<string, object> StorePipeParameters(Pipe pipe)
        {
            Dictionary<string, object> parameters = new Dictionary<string, object>();

            // Store important pipe parameters
            Parameter systemType = pipe.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM);
            if (systemType != null)
                parameters["SystemType"] = systemType.AsElementId();

            Parameter diameter = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
            if (diameter != null)
                parameters["Diameter"] = diameter.AsDouble();

            // Store custom parameters
            foreach (Parameter param in pipe.Parameters)
            {
                if (param == null || param.IsReadOnly || !param.HasValue) continue;

                string paramName = param.Definition.Name;
                if (parameters.ContainsKey(paramName)) continue;

                try
                {
                    switch (param.StorageType)
                    {
                        case StorageType.Double:
                            parameters[paramName] = param.AsDouble();
                            break;
                        case StorageType.Integer:
                            parameters[paramName] = param.AsInteger();
                            break;
                        case StorageType.String:
                            parameters[paramName] = param.AsString();
                            break;
                        case StorageType.ElementId:
                            parameters[paramName] = param.AsElementId();
                            break;
                    }
                }
                catch { }
            }

            return parameters;
        }

        private void RestorePipeParameters(Pipe pipe, Dictionary<string, object> parameters)
        {
            foreach (var kvp in parameters)
            {
                try
                {
                    Parameter param = pipe.LookupParameter(kvp.Key);
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