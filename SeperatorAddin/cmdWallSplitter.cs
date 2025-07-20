using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI.Selection;
using SeperatorAddin.Common;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace SeperatorAddin
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class cmdWallSplitter : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            try
            {
                // 1. Select a wall
                Reference wallRef = uiDoc.Selection.PickObject(ObjectType.Element, new Utils.WallSelectionFilter(), "Select a wall to split");
                Wall selectedWall = doc.GetElement(wallRef) as Wall;

                // 2. Select a model line
                Reference lineRef = uiDoc.Selection.PickObject(ObjectType.Element, new Utils.ModelCurveSelectionFilter(), "Select a model line to split the wall");
                ModelCurve modelLine = doc.GetElement(lineRef) as ModelCurve;
                Line splitLine = modelLine.GeometryCurve as Line;

                if (splitLine == null)
                {
                    TaskDialog.Show("Error", "Please select a straight model line.");
                    return Result.Failed;
                }

                using (Transaction trans = new Transaction(doc, "Split Wall"))
                {
                    trans.Start();

                    try
                    {
                        LocationCurve wallLocationCurve = selectedWall.Location as LocationCurve;
                        if (wallLocationCurve == null)
                        {
                            TaskDialog.Show("Error", "Could not get wall location.");
                            return Result.Failed;
                        }
                        Line wallLine = wallLocationCurve.Curve as Line;
                        if (wallLine == null)
                        {
                            TaskDialog.Show("Error", "Wall splitting only works on straight walls.");
                            return Result.Failed;
                        }

                        // Find intersection
                        IntersectionResultArray results;
                        if (wallLine.Intersect(splitLine, out results) != SetComparisonResult.Overlap || results.Size == 0)
                        {
                            TaskDialog.Show("Error", "The selected line does not intersect the wall.");
                            return Result.Failed;
                        }
                        XYZ intersectionPoint = results.get_Item(0).XYZPoint;

                        // Get wall properties
                        WallType wallType = selectedWall.WallType;
                        Level level = doc.GetElement(selectedWall.LevelId) as Level;
                        double baseOffset = selectedWall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET).AsDouble();
                        double unconnectedHeight = selectedWall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).AsDouble();
                        bool isStructural = selectedWall.get_Parameter(BuiltInParameter.WALL_STRUCTURAL_SIGNIFICANT).AsInteger() == 1;

                        // Collect hosted elements
                        List<FamilyInstance> hostedInstances = GetHostedInstances(doc, selectedWall);

                        // Create two new wall segments
                        Line wallSegment1 = Line.CreateBound(wallLine.GetEndPoint(0), intersectionPoint);
                        Line wallSegment2 = Line.CreateBound(intersectionPoint, wallLine.GetEndPoint(1));

                        double minLength = 1.0 / 12.0; // 1 inch
                        if (wallSegment1.Length < minLength || wallSegment2.Length < minLength)
                        {
                            TaskDialog.Show("Error", "Split results in a wall segment that is too short. Please move the split line farther from the wall end.");
                            trans.RollBack();
                            return Result.Cancelled;
                        }

                        // Create new walls
                        Wall newWall1 = Wall.Create(doc, wallSegment1, wallType.Id, level.Id, unconnectedHeight, baseOffset, selectedWall.Flipped, isStructural);
                        Wall newWall2 = Wall.Create(doc, wallSegment2, wallType.Id, level.Id, unconnectedHeight, baseOffset, selectedWall.Flipped, isStructural);

                        // Copy parameters from original wall
                        CopyWallParameters(selectedWall, newWall1);
                        CopyWallParameters(selectedWall, newWall2);

                        // Re-host elements with correct orientation
                        RehostInstances(doc, newWall1, newWall2, hostedInstances, intersectionPoint, wallLine.Direction);

                        // Delete the original wall
                        doc.Delete(selectedWall.Id);

                        trans.Commit();
                        return Result.Succeeded;
                    }
                    catch (Exception ex)
                    {
                        trans.RollBack();
                        message = ex.Message;
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

        private List<FamilyInstance> GetHostedInstances(Document doc, Wall wall)
        {
            var hostedInstances = new List<FamilyInstance>();
            var dependentIds = wall.GetDependentElements(new ElementClassFilter(typeof(FamilyInstance)));
            if (dependentIds == null || !dependentIds.Any()) return hostedInstances;

            foreach (ElementId id in dependentIds)
            {
                hostedInstances.Add(doc.GetElement(id) as FamilyInstance);
            }
            return hostedInstances;
        }

        private void RehostInstances(Document doc, Wall wall1, Wall wall2, List<FamilyInstance> instances, XYZ splitPoint, XYZ wallDirection)
        {
            if (!instances.Any()) return;

            foreach (var fi in instances)
            {
                var locPoint = fi.Location as LocationPoint;
                if (locPoint == null) continue;

                // Get original orientation
                bool wasHandFlipped = fi.HandFlipped;
                bool wasFacingFlipped = fi.FacingFlipped;

                var instanceLocation = locPoint.Point;
                var dotProduct = (instanceLocation - splitPoint).DotProduct(wallDirection);

                var targetWall = dotProduct < 0 ? wall1 : wall2;
                Level level = doc.GetElement(targetWall.LevelId) as Level;

                var newFi = doc.Create.NewFamilyInstance(instanceLocation, fi.Symbol, targetWall, level, StructuralType.NonStructural);
                if (newFi != null)
                {
                    // Copy all instance parameters
                    CopyInstanceParameters(fi, newFi);

                    // Correct the orientation of the new instance to match the original
                    if (newFi.HandFlipped != wasHandFlipped)
                    {
                        newFi.flipHand();
                    }
                    if (newFi.FacingFlipped != wasFacingFlipped)
                    {
                        newFi.flipFacing();
                    }
                }
            }
        }

        private void CopyWallParameters(Wall source, Wall target)
        {
            // List of built-in parameters to explicitly ignore
            var ignoredParams = new List<BuiltInParameter>
            {
                BuiltInParameter.CURVE_ELEM_LENGTH,
                BuiltInParameter.HOST_AREA_COMPUTED,
                BuiltInParameter.HOST_VOLUME_COMPUTED,
                BuiltInParameter.WALL_USER_HEIGHT_PARAM,
                BuiltInParameter.WALL_BASE_CONSTRAINT,
                BuiltInParameter.WALL_BASE_OFFSET,
                BuiltInParameter.WALL_KEY_REF_PARAM // Location Line
            };

            foreach (Parameter sourceParam in source.Parameters)
            {
                // Skip read-only parameters or those in our ignore list
                if (sourceParam.IsReadOnly || ignoredParams.Contains((BuiltInParameter)sourceParam.Id.IntegerValue))
                {
                    continue;
                }

                Parameter targetParam = target.get_Parameter(sourceParam.Definition);
                if (targetParam != null && !targetParam.IsReadOnly)
                {
                    try
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
                    catch
                    {
                        // Fails silently if a parameter can't be set for any reason
                    }
                }
            }
        }


        private void CopyInstanceParameters(FamilyInstance source, FamilyInstance target)
        {
            foreach (Parameter param in source.Parameters)
            {
                if (param.IsReadOnly || !param.HasValue) continue;

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
                    catch
                    {
                        // Ignore if this specific parameter can't be set
                    }
                }
            }
        }

        internal static PushButtonData GetButtonData()
        {
            string buttonInternalName = "btnSplitWall";
            string buttonTitle = "Split Wall";

            ButtonDataClass myButtonData = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Red_32,
                Properties.Resources.Red_16,
                "Splits a wall along a model line.");

            return myButtonData.Data;
        }
    }
}