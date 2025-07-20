using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace SeperatorAddin
{
    [Transaction(TransactionMode.Manual)]
    public class cmdPipeSplitter : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Select a pipe
                Reference pipeRef = uidoc.Selection.PickObject(
                    ObjectType.Element,
                    new PipeSelectionFilter(),
                    "Select a pipe to split");

                Pipe originalPipe = doc.GetElement(pipeRef) as Pipe;
                if (originalPipe == null)
                {
                    TaskDialog.Show("Error", "Selected element is not a pipe.");
                    return Result.Failed;
                }

                // Get split point
                XYZ splitPoint = uidoc.Selection.PickPoint("Pick a point on the pipe to split");

                // Project the point onto the pipe
                LocationCurve locCurve = originalPipe.Location as LocationCurve;
                if (locCurve == null)
                {
                    TaskDialog.Show("Error", "Could not get pipe location curve.");
                    return Result.Failed;
                }

                Line pipeLine = locCurve.Curve as Line;
                if (pipeLine == null)
                {
                    TaskDialog.Show("Error", "This tool only works with straight pipes.");
                    return Result.Failed;
                }

                // Project split point onto pipe line
                IntersectionResult projection = pipeLine.Project(splitPoint);
                if (projection == null)
                {
                    TaskDialog.Show("Error", "Could not project point onto pipe.");
                    return Result.Failed;
                }

                XYZ projectedPoint = projection.XYZPoint;

                // Check if split point is valid (not too close to ends)
                double minDistance = 0.5; // 6 inches minimum
                if (projectedPoint.DistanceTo(pipeLine.GetEndPoint(0)) < minDistance ||
                    projectedPoint.DistanceTo(pipeLine.GetEndPoint(1)) < minDistance)
                {
                    TaskDialog.Show("Error", "Split point is too close to pipe ends.");
                    return Result.Failed;
                }

                using (Transaction trans = new Transaction(doc, "Split Pipe with Insulation"))
                {
                    trans.Start();

                    try
                    {
                        // Store original pipe insulation data before splitting
                        InsulationData originalInsulationData = GetPipeInsulation(doc, originalPipe);

                        // Store original pipe parameters
                        Dictionary<string, object> pipeParameters = StorePipeParameters(originalPipe);

                        // Get original pipe system
                        ElementId systemTypeId = originalPipe.get_Parameter(
                            BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId();

                        // Split the pipe using Revit's built-in method
                        ElementId newPipeId = PlumbingUtils.BreakCurve(
                            doc, originalPipe.Id, projectedPoint);

                        if (newPipeId == null || newPipeId == ElementId.InvalidElementId)
                        {
                            trans.RollBack();
                            TaskDialog.Show("Error", "Failed to split pipe.");
                            return Result.Failed;
                        }

                        // Get the new pipe created by the split
                        Pipe newPipe = doc.GetElement(newPipeId) as Pipe;

                        // Restore parameters to both pipes
                        RestorePipeParameters(originalPipe, pipeParameters);
                        RestorePipeParameters(newPipe, pipeParameters);

                        // Handle insulation for split pipes
                        if (originalInsulationData != null && originalInsulationData.HasInsulation)
                        {
                            // Add insulation to the new pipe
                            PipeInsulation newPipeInsulation = AddInsulationToPipe(
                                doc, newPipe, originalInsulationData);

                            // Add insulation to the original pipe (it may have been removed during split)
                            PipeInsulation originalPipeNewInsulation = AddInsulationToPipe(
                                doc, originalPipe, originalInsulationData);
                        }

                        trans.Commit();

                        // Report results
                        string resultMessage = "Pipe split successfully.";
                        if (originalInsulationData != null && originalInsulationData.HasInsulation)
                        {
                            resultMessage += "\nInsulation parameters copied to both pipe segments.";
                        }

                        TaskDialog.Show("Success", resultMessage);
                        return Result.Succeeded;
                    }
                    catch (Exception ex)
                    {
                        trans.RollBack();
                        TaskDialog.Show("Error", $"Failed to split pipe: {ex.Message}");
                        return Result.Failed;
                    }
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        #region Insulation Data Storage and Retrieval

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

        #endregion

        #region Pipe Parameter Handling

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

        #region Insulation Creation

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
                TaskDialog.Show("Warning", $"Failed to add insulation to pipe: {ex.Message}");
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

        #region Utility Methods

        internal static PushButtonData GetButtonData()
        {
            string buttonInternalName = "btnPipeSplitter";
            string buttonTitle = "Split Pipe";

            Common.ButtonDataClass myButtonData = new Common.ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "Splits a pipe at a selected point and copies insulation parameters to both segments");

            return myButtonData.Data;
        }

        #endregion
    }

    public class PipeSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            return elem is Pipe;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }
}
