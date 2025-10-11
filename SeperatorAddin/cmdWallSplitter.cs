using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using SeperatorAddin.Common;
using System;
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

                using (Transaction trans = new Transaction(doc, "Split Wall"))
                {
                    trans.Start();
                    bool success = SplitWall(doc, selectedWall, modelLine);

                    if (success)
                    {
                        trans.Commit();
                        return Result.Succeeded;
                    }
                    else
                    {
                        trans.RollBack();
                        message = "Failed to split the wall. Ensure the model line intersects the wall.";
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

        public bool SplitWall(Document doc, Wall selectedWall, ModelCurve modelLine)
        {
            if (selectedWall == null || modelLine == null) return false;

            Line splitLine = modelLine.GeometryCurve as Line;
            if (splitLine == null) return false;

            try
            {
                LocationCurve wallLocationCurve = selectedWall.Location as LocationCurve;
                if (wallLocationCurve == null) return false;

                Line wallLine = wallLocationCurve.Curve as Line;
                if (wallLine == null) return false;

                // Find intersection
                IntersectionResultArray results;
                if (wallLine.Intersect(splitLine, out results) != SetComparisonResult.Overlap || results.Size == 0)
                {
                    return false;
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
                    return false; // Resulting segment is too short
                }

                // Create new walls
                Wall newWall1 = Wall.Create(doc, wallSegment1, wallType.Id, level.Id, unconnectedHeight, baseOffset, selectedWall.Flipped, isStructural);
                Wall newWall2 = Wall.Create(doc, wallSegment2, wallType.Id, level.Id, unconnectedHeight, baseOffset, selectedWall.Flipped, isStructural);

                if (newWall1 == null || newWall2 == null) return false;

                // Copy parameters from original wall
                CopyWallParameters(selectedWall, newWall1);
                CopyWallParameters(selectedWall, newWall2);

                // Re-host elements with correct orientation
                RehostInstances(doc, newWall1, newWall2, hostedInstances, intersectionPoint, wallLine.Direction);

                // Delete the original wall
                doc.Delete(selectedWall.Id);

                return true;
            }
            catch
            {
                return false;
            }
        }

        private List<FamilyInstance> GetHostedInstances(Document doc, Wall wall)
        {
            var hostedInstances = new List<FamilyInstance>();
            var dependentIds = wall.GetDependentElements(new ElementClassFilter(typeof(FamilyInstance)));
            if (dependentIds == null || !dependentIds.Any()) return hostedInstances;

            foreach (ElementId id in dependentIds)
            {
                if (doc.GetElement(id) is FamilyInstance fi)
                {
                    hostedInstances.Add(fi);
                }
            }
            return hostedInstances;
        }

        private void RehostInstances(Document doc, Wall wall1, Wall wall2, List<FamilyInstance> instances, XYZ splitPoint, XYZ wallDirection)
        {
            if (!instances.Any()) return;

            foreach (var fi in instances)
            {
                if (fi == null || !(fi.Location is LocationPoint locPoint)) continue;

                // Get original orientation
                bool wasHandFlipped = fi.HandFlipped;
                bool wasFacingFlipped = fi.FacingFlipped;

                var instanceLocation = locPoint.Point;
                var dotProduct = (instanceLocation - splitPoint).DotProduct(wallDirection);

                var targetWall = dotProduct < 0 ? wall1 : wall2;
                Level level = doc.GetElement(targetWall.LevelId) as Level;

                try
                {
                    var newFi = doc.Create.NewFamilyInstance(instanceLocation, fi.Symbol, targetWall, level, StructuralType.NonStructural);
                    if (newFi != null)
                    {
                        // Copy all instance parameters
                        CopyInstanceParameters(fi, newFi);

                        // Correct the orientation of the new instance to match the original
                        if (newFi.HandFlipped != wasHandFlipped) newFi.flipHand();
                        if (newFi.FacingFlipped != wasFacingFlipped) newFi.flipFacing();
                    }
                }
                catch { }
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
                if (sourceParam.IsReadOnly) continue;

                if (sourceParam.Definition is InternalDefinition internalDef && ignoredParams.Contains(internalDef.BuiltInParameter))
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
                            case StorageType.Double: targetParam.Set(sourceParam.AsDouble()); break;
                            case StorageType.Integer: targetParam.Set(sourceParam.AsInteger()); break;
                            case StorageType.String: targetParam.Set(sourceParam.AsString()); break;
                            case StorageType.ElementId: targetParam.Set(sourceParam.AsElementId()); break;
                        }
                    }
                    catch { }
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
                            case StorageType.Double: targetParam.Set(param.AsDouble()); break;
                            case StorageType.Integer: targetParam.Set(param.AsInteger()); break;
                            case StorageType.String: targetParam.Set(param.AsString()); break;
                            case StorageType.ElementId: targetParam.Set(param.AsElementId()); break;
                        }
                    }
                    catch { }
                }
            }
        }
    }
}