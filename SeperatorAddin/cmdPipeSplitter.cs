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

                using (Transaction trans = new Transaction(doc, "Split Pipe"))
                {
                    trans.Start();
                    bool success = SplitPipe(doc, originalPipe, splitPoint);
                    if (success)
                    {
                        trans.Commit();
                        TaskDialog.Show("Success", "Pipe split successfully.");
                        return Result.Succeeded;
                    }
                    else
                    {
                        trans.RollBack();
                        message = "Failed to split pipe. The split point may be too close to an end.";
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

        public bool SplitPipe(Document doc, Pipe originalPipe, XYZ splitPoint)
        {
            if (originalPipe == null) return false;

            LocationCurve locCurve = originalPipe.Location as LocationCurve;
            if (locCurve == null || !(locCurve.Curve is Line pipeLine)) return false;

            IntersectionResult projection = pipeLine.Project(splitPoint);
            if (projection == null) return false;
            XYZ projectedPoint = projection.XYZPoint;

            // Check if split point is valid (not too close to ends)
            double minDistance = 0.5; // 6 inches minimum
            if (projectedPoint.DistanceTo(pipeLine.GetEndPoint(0)) < minDistance ||
                projectedPoint.DistanceTo(pipeLine.GetEndPoint(1)) < minDistance)
            {
                return false;
            }

            try
            {
                // Store original pipe insulation data before splitting
                InsulationData originalInsulationData = GetPipeInsulation(doc, originalPipe);
                Dictionary<string, object> pipeParameters = StorePipeParameters(originalPipe);

                // Split the pipe using Revit's built-in method
                ElementId newPipeId = PlumbingUtils.BreakCurve(doc, originalPipe.Id, projectedPoint);
                if (newPipeId == null || newPipeId == ElementId.InvalidElementId) return false;

                Pipe newPipe = doc.GetElement(newPipeId) as Pipe;
                if (newPipe == null) return false;

                // Restore parameters to both pipes
                RestorePipeParameters(originalPipe, pipeParameters);
                RestorePipeParameters(newPipe, pipeParameters);

                // Handle insulation for split pipes
                if (originalInsulationData != null && originalInsulationData.HasInsulation)
                {
                    AddInsulationToPipe(doc, newPipe, originalInsulationData);
                    AddInsulationToPipe(doc, originalPipe, originalInsulationData);
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        #region Insulation Data Storage and Retrieval
        private class InsulationData
        {
            public bool HasInsulation { get; set; }
            public ElementId InsulationTypeId { get; set; }
            public double Thickness { get; set; }
            public Dictionary<string, object> Parameters { get; set; } = new Dictionary<string, object>();
        }

        private InsulationData GetPipeInsulation(Document doc, Pipe pipe)
        {
            var data = new InsulationData();
            // Find insulation associated with the pipe
            var insulation = new FilteredElementCollector(doc)
                .OfClass(typeof(PipeInsulation))
                .Cast<PipeInsulation>()
                .FirstOrDefault(ins => ins.HostElementId == pipe.Id);

            if (insulation == null)
            {
                data.HasInsulation = false;
                return data;
            }

            data.HasInsulation = true;
            data.InsulationTypeId = insulation.GetTypeId();
            data.Thickness = insulation.Thickness;
            StoreParameters(insulation, data.Parameters);
            return data;
        }

        private void AddInsulationToPipe(Document doc, Pipe pipe, InsulationData insulationData)
        {
            try
            {
                // Check if pipe already has insulation
                var existingInsulation = new FilteredElementCollector(doc).OfClass(typeof(PipeInsulation)).Cast<PipeInsulation>().FirstOrDefault(ins => ins.HostElementId == pipe.Id);
                if (existingInsulation != null)
                {
                    // Update existing insulation
                    RestoreParameters(existingInsulation, insulationData.Parameters);
                    return;
                }

                // Create new insulation
                PipeInsulation newInsulation = PipeInsulation.Create(doc, pipe.Id, insulationData.InsulationTypeId, insulationData.Thickness);
                if (newInsulation != null)
                {
                    RestoreParameters(newInsulation, insulationData.Parameters);
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Warning", $"Failed to add insulation to pipe: {ex.Message}");
            }
        }
        #endregion

        #region Pipe Parameter Handling
        private Dictionary<string, object> StorePipeParameters(Pipe pipe)
        {
            var parameters = new Dictionary<string, object>();
            StoreParameters(pipe, parameters);
            return parameters;
        }

        private void RestorePipeParameters(Pipe pipe, Dictionary<string, object> parameters) => RestoreParameters(pipe, parameters);

        private void StoreParameters(Element element, Dictionary<string, object> parameters)
        {
            foreach (Parameter param in element.Parameters)
            {
                if (param == null || param.IsReadOnly || !param.HasValue) continue;
                try
                {
                    string name = param.Definition.Name;
                    if (parameters.ContainsKey(name)) continue;

                    switch (param.StorageType)
                    {
                        case StorageType.Double:
                            parameters[name] = param.AsDouble();
                            break;
                        case StorageType.Integer:
                            parameters[name] = param.AsInteger();
                            break;
                        case StorageType.String:
                            parameters[name] = param.AsString();
                            break;
                        case StorageType.ElementId:
                            parameters[name] = param.AsElementId();
                            break;
                    }
                }
                catch { }
            }
        }

        private void RestoreParameters(Element element, Dictionary<string, object> parameters)
        {
            foreach (var kvp in parameters)
            {
                try
                {
                    Parameter param = element.LookupParameter(kvp.Key);
                    if (param == null || param.IsReadOnly) continue;

                    switch (param.StorageType)
                    {
                        case StorageType.Double:
                            if (kvp.Value is double) param.Set((double)kvp.Value);
                            break;
                        case StorageType.Integer:
                            if (kvp.Value is int) param.Set((int)kvp.Value);
                            break;
                        case StorageType.String:
                            if (kvp.Value is string) param.Set((string)kvp.Value);
                            break;
                        case StorageType.ElementId:
                            if (kvp.Value is ElementId) param.Set((ElementId)kvp.Value);
                            break;
                    }
                }
                catch { }
            }
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