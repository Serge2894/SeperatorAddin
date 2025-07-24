using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using SeperatorAddin.Forms;
using System;
using System.Collections.Generic;
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
                // 1. Select pipes from the model first
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

                // 2. Get levels to populate the form
                var allLevels = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().OrderBy(l => l.Elevation).ToList();
                if (allLevels.Count < 2)
                {
                    TaskDialog.Show("Error", "At least 2 levels are required in the project.");
                    return Result.Cancelled;
                }

                // 3. Show form to select levels
                List<string> levelNames = allLevels.Select(l => l.Name).ToList();
                // NOTE: This assumes a PipeSelectionForm that takes List<string>
                PipeSelectionForm form = new PipeSelectionForm(levelNames);
                if (form.ShowDialog() != true)
                {
                    return Result.Cancelled;
                }

                // 4. Get selected levels from the form
                List<string> selectedLevelNames = form.SelectedLevelNames;
                if (selectedLevelNames.Count < 2)
                {
                    TaskDialog.Show("Error", "Please select at least 2 levels.");
                    return Result.Cancelled;
                }
                List<Level> selectedLevels = allLevels.Where(l => selectedLevelNames.Contains(l.Name)).ToList();

                // 5. Perform the split operation
                using (var trans = new Transaction(doc, "Split Pipes by Level"))
                {
                    trans.Start();

                    int successCount = 0;
                    int failCount = 0;

                    foreach (Pipe pipe in selectedPipes)
                    {
                        if (SplitPipeByLevels(doc, pipe, selectedLevels))
                        {
                            successCount++;
                        }
                        else
                        {
                            failCount++;
                        }
                    }

                    trans.Commit();
                    TaskDialog.Show("Success", $"Operation complete.\nSuccessfully split: {successCount}\nFailed: {failCount}");
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        /// <summary>
        /// Splits a single vertical pipe at the specified levels.
        /// </summary>
        private bool SplitPipeByLevels(Document doc, Pipe pipe, List<Level> levels)
        {
            LocationCurve locCurve = pipe.Location as LocationCurve;
            if (locCurve == null || !(locCurve.Curve is Line pipeLine)) return false;

            XYZ startPoint = pipeLine.GetEndPoint(0);
            XYZ endPoint = pipeLine.GetEndPoint(1);

            // Only process pipes that are mostly vertical
            if (Math.Abs(pipeLine.Direction.Z) < 0.9) return false;

            // Find level elevations that are within the pipe's bounds
            List<double> splitHeights = levels
                .Select(l => l.Elevation)
                .Where(e => e > Math.Min(startPoint.Z, endPoint.Z) + 0.001 && e < Math.Max(startPoint.Z, endPoint.Z) - 0.001)
                .Distinct()
                .OrderBy(h => h)
                .ToList();

            if (splitHeights.Count == 0) return false;

            // Store original pipe properties
            PipeType pipeType = pipe.PipeType;
            ElementId systemTypeId = pipe.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId();
            ElementId levelId = pipe.ReferenceLevel.Id;
            double diameter = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble();

            // Generate all points for new segments (start, intersections, end)
            List<XYZ> splitPoints = new List<XYZ> { startPoint };
            splitPoints.AddRange(splitHeights.Select(h => startPoint + ((h - startPoint.Z) / pipeLine.Direction.Z) * pipeLine.Direction));
            splitPoints.Add(endPoint);
            splitPoints = splitPoints.OrderBy(p => p.Z).Distinct().ToList();

            List<Pipe> newPipes = new List<Pipe>();
            for (int i = 0; i < splitPoints.Count - 1; i++)
            {
                XYZ p1 = splitPoints[i];
                XYZ p2 = splitPoints[i + 1];

                if (p1.DistanceTo(p2) < 0.01) continue; // Skip tiny segments

                Pipe newPipe = Pipe.Create(doc, systemTypeId, pipeType.Id, levelId, p1, p2);
                newPipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).Set(diameter);
                CopyPipeParameters(pipe, newPipe);
                newPipes.Add(newPipe);
            }

            if (newPipes.Count == 0) return false;

            // Delete the original pipe
            doc.Delete(pipe.Id);

            return true;
        }

        private void CopyPipeParameters(Pipe source, Pipe target)
        {
            CopyParameter(source, target, BuiltInParameter.RBS_PIPE_SLOPE);
            CopyParameter(source, target, BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
            CopyParameter(source, target, BuiltInParameter.ALL_MODEL_MARK);
        }

        private void CopyParameter(Element source, Element target, BuiltInParameter paramEnum)
        {
            Parameter sourceParam = source.get_Parameter(paramEnum);
            Parameter targetParam = target.get_Parameter(paramEnum);

            if (sourceParam != null && targetParam != null && !targetParam.IsReadOnly && sourceParam.HasValue)
            {
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
        }
    }

    // NOTE: This class should only be defined ONCE in your project.
    // If it's already in another file (like cmdPipeSplitter.cs), you can remove it from here.
    public class PipeSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem) => elem is Pipe;
        public bool AllowReference(Reference reference, XYZ position) => false;
    }
}