using Autodesk.Revit.UI.Selection;
using SeperatorAddin.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using static SeperatorAddin.Common.Utils;

namespace SeperatorAddin
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class cmdSplitTool : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            try
            {
                // Ensure we're in a plan view
                View activeView = doc.ActiveView;
                if (!(activeView is ViewPlan))
                {
                    TaskDialog.Show("Error", "Please run this command in a plan view.");
                    return Result.Failed;
                }

                // Select a floor
                Reference floorRef = uiDoc.Selection.PickObject(
                    ObjectType.Element,
                    new FloorSelectionFilter(),
                    "Select a floor to split");

                Floor selectedFloor = doc.GetElement(floorRef) as Floor;

                // Select ONE model line
                Reference lineRef = uiDoc.Selection.PickObject(
                    ObjectType.Element,
                    new ModelCurveSelectionFilter(),
                    "Select ONE straight model line that crosses the floor");

                ModelCurve modelLine = doc.GetElement(lineRef) as ModelCurve;
                Line splitLine = modelLine.GeometryCurve as Line;

                if (splitLine == null)
                {
                    TaskDialog.Show("Error", "Please select a straight line.");
                    return Result.Failed;
                }

                using (Transaction trans = new Transaction(doc, "Split Floor by Duplicating"))
                {
                    trans.Start();

                    try
                    {
                        // Get floor properties before duplication
                        Level level = doc.GetElement(selectedFloor.LevelId) as Level;

                        // Store all face-based families hosted on this floor
                        List<FamilyInstance> hostedFamilies = GetHostedFamilies(doc, selectedFloor);

                        // Get floor boundary curves including openings
                        var floorLoops = GetFloorBoundaryLoops(selectedFloor);

                        if (floorLoops.Item1.Count < 3)
                        {
                            trans.RollBack();
                            TaskDialog.Show("Error", "Could not get valid floor boundary.");
                            return Result.Failed;
                        }

                        // Extend split line
                        XYZ lineDir = splitLine.Direction;
                        XYZ lineStart = splitLine.GetEndPoint(0) - lineDir * 1000;
                        XYZ lineEnd = splitLine.GetEndPoint(1) + lineDir * 1000;
                        Line extendedLine = Line.CreateBound(lineStart, lineEnd);

                        // Find intersections with outer boundary
                        List<XYZ> intersectionPoints = new List<XYZ>();
                        foreach (Curve boundaryCurve in floorLoops.Item1)
                        {
                            IntersectionResultArray results;
                            if (boundaryCurve.Intersect(extendedLine, out results) == SetComparisonResult.Overlap)
                            {
                                foreach (IntersectionResult result in results)
                                {
                                    intersectionPoints.Add(result.XYZPoint);
                                }
                            }
                        }

                        if (intersectionPoints.Count < 2)
                        {
                            trans.RollBack();
                            TaskDialog.Show("Error",
                                "Split line must cross the floor boundary at two points.\n" +
                                "Make sure the line extends across the entire floor.");
                            return Result.Failed;
                        }

                        // Sort intersection points
                        intersectionPoints = intersectionPoints
                            .OrderBy(p => lineStart.DistanceTo(p))
                            .ToList();

                        XYZ splitStart = intersectionPoints.First();
                        XYZ splitEnd = intersectionPoints.Last();

                        // Create profiles for two floors
                        StringBuilder debugInfo = new StringBuilder();
                        var profiles = CreateSplitProfilesWithOpenings(
                            floorLoops.Item1, floorLoops.Item2, splitStart, splitEnd, lineDir, debugInfo);

                        if (profiles.Item1.Count == 0 || profiles.Item2.Count == 0)
                        {
                            trans.RollBack();
                            TaskDialog.Show("Error", "Failed to create valid split profiles.");
                            return Result.Failed;
                        }

                        // Duplicate the original floor
                        ICollection<ElementId> copiedIds = ElementTransformUtils.CopyElement(
                            doc, selectedFloor.Id, XYZ.Zero);

                        if (copiedIds.Count == 0)
                        {
                            trans.RollBack();
                            TaskDialog.Show("Error", "Failed to duplicate floor.");
                            return Result.Failed;
                        }

                        Floor duplicatedFloor = doc.GetElement(copiedIds.First()) as Floor;

                        // Now we have two floors: selectedFloor and duplicatedFloor
                        // Modify their boundaries using the split profiles

                        bool floor1Updated = UpdateFloorBoundary(doc, selectedFloor, profiles.Item1, profiles.Item3, debugInfo);
                        bool floor2Updated = UpdateFloorBoundary(doc, duplicatedFloor, profiles.Item2, profiles.Item4, debugInfo);

                        if (!floor1Updated || !floor2Updated)
                        {
                            trans.RollBack();
                            TaskDialog.Show("Error",
                                $"Failed to update floor boundaries.\n" +
                                $"Floor 1: {(floor1Updated ? "Success" : "Failed")}\n" +
                                $"Floor 2: {(floor2Updated ? "Success" : "Failed")}\n\n" +
                                debugInfo.ToString());
                            return Result.Failed;
                        }

                        // Handle hosted families
                        HandleHostedFamilies(doc, selectedFloor, duplicatedFloor,
                            extendedLine, splitStart, lineDir, hostedFamilies);

                        trans.Commit();
                        TaskDialog.Show("Success", "Floor split into 2 parts successfully.");
                        return Result.Succeeded;
                    }
                    catch (Exception ex)
                    {
                        trans.RollBack();
                        TaskDialog.Show("Error", $"Failed: {ex.Message}\n\n{ex.StackTrace}");
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

        #region Floor Boundary Update

        private bool UpdateFloorBoundary(Document doc, Floor floor, List<Curve> outerBoundary,
            List<List<Curve>> innerBoundaries, StringBuilder debugInfo)
        {
            try
            {
                // Get the floor's sketch
                if (!(floor is Floor))
                {
                    debugInfo.AppendLine("Element is not a floor.");
                    return false;
                }

                // Create new curve loops
                List<CurveLoop> newLoops = new List<CurveLoop>();

                // Validate and create outer loop
                if (!ValidateCurveLoop(outerBoundary, out string error))
                {
                    debugInfo.AppendLine($"Invalid outer boundary: {error}");
                    return false;
                }

                CurveLoop outerLoop = new CurveLoop();
                foreach (Curve c in outerBoundary)
                {
                    outerLoop.Append(c);
                }
                newLoops.Add(outerLoop);

                // Add inner loops (openings)
                foreach (var innerLoop in innerBoundaries)
                {
                    if (ValidateCurveLoop(innerLoop, out string openingError))
                    {
                        CurveLoop opening = new CurveLoop();
                        foreach (Curve c in innerLoop)
                        {
                            opening.Append(c);
                        }
                        newLoops.Add(opening);
                    }
                    else
                    {
                        debugInfo.AppendLine($"Skipping invalid opening: {openingError}");
                    }
                }

                // Use SlabShapeEditor to modify the floor boundary
                SlabShapeEditor shapeEditor = floor.SlabShapeEditor;
                if (shapeEditor != null)
                {
                    shapeEditor.ResetSlabShape();
                }

                // Try to modify the floor's boundary sketch
                // This is done by getting the floor's sketch and modifying it
                // Note: This approach works for sketch-based floors

                // First, we need to delete the existing floor and recreate it with new boundaries
                // Store floor properties
                FloorType floorType = floor.FloorType;
                Level level = doc.GetElement(floor.LevelId) as Level;
                double offset = floor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM).AsDouble();

                // Get top offset if it exists
                double topOffset = 0;
                Parameter topOffsetParam = floor.get_Parameter(BuiltInParameter.FLOOR_PARAM_IS_STRUCTURAL);
                if (topOffsetParam != null)
                {
                    topOffset = topOffsetParam.AsDouble();
                }

                // Store the floor's ID to preserve any relationships
                ElementId originalFloorId = floor.Id;

                // Create a new floor with the updated boundary
                Floor newFloor = Floor.Create(doc, newLoops, floorType.Id, level.Id);

                if (newFloor == null)
                {
                    debugInfo.AppendLine("Failed to create new floor with updated boundary.");
                    return false;
                }

                // Copy parameters from old floor to new floor
                newFloor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM).Set(offset);

                // Copy other parameters
                CopyFloorParameters(floor, newFloor);

                // Delete the original floor
                doc.Delete(floor.Id);

                return true;
            }
            catch (Exception ex)
            {
                debugInfo.AppendLine($"UpdateFloorBoundary exception: {ex.Message}");
                return false;
            }
        }

        private void CopyFloorParameters(Floor sourceFloor, Floor targetFloor)
        {
            try
            {
                // Copy ALL parameters - including custom/shared parameters
                foreach (Parameter sourceParam in sourceFloor.Parameters)
                {
                    if (sourceParam == null || sourceParam.IsReadOnly) continue;

                    // Skip certain parameters that shouldn't be copied
                    if (sourceParam.Id.IntegerValue == (int)BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM ||
                        sourceParam.Id.IntegerValue == (int)BuiltInParameter.ELEM_FAMILY_PARAM ||
                        sourceParam.Id.IntegerValue == (int)BuiltInParameter.ELEM_TYPE_PARAM ||
                        sourceParam.Id.IntegerValue == (int)BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM)
                    {
                        continue;
                    }

                    try
                    {
                        Parameter targetParam = targetFloor.get_Parameter(sourceParam.Definition);
                        if (targetParam != null && !targetParam.IsReadOnly)
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
                                    string value = sourceParam.AsString();
                                    if (value != null)
                                        targetParam.Set(value);
                                    break;
                                case StorageType.ElementId:
                                    ElementId id = sourceParam.AsElementId();
                                    if (id != ElementId.InvalidElementId)
                                        targetParam.Set(id);
                                    break;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Failed to copy parameter {sourceParam.Definition.Name}: {ex.Message}");
                    }
                }

                // Ensure critical parameters are copied
                CopyPhaseParameters(sourceFloor, targetFloor);
            }
            catch (Exception ex)
            {
                // Log but don't fail the operation
                Debug.WriteLine($"Failed to copy some parameters: {ex.Message}");
            }
        }

        private void CopyPhaseParameters(Floor source, Floor target)
        {
            try
            {
                // Phase Created
                Parameter phaseCreated = source.get_Parameter(BuiltInParameter.PHASE_CREATED);
                if (phaseCreated != null)
                {
                    Parameter targetPhaseCreated = target.get_Parameter(BuiltInParameter.PHASE_CREATED);
                    if (targetPhaseCreated != null && !targetPhaseCreated.IsReadOnly)
                    {
                        targetPhaseCreated.Set(phaseCreated.AsElementId());
                    }
                }

                // Phase Demolished
                Parameter phaseDemolished = source.get_Parameter(BuiltInParameter.PHASE_DEMOLISHED);
                if (phaseDemolished != null)
                {
                    Parameter targetPhaseDemolished = target.get_Parameter(BuiltInParameter.PHASE_DEMOLISHED);
                    if (targetPhaseDemolished != null && !targetPhaseDemolished.IsReadOnly)
                    {
                        targetPhaseDemolished.Set(phaseDemolished.AsElementId());
                    }
                }
            }
            catch { }
        }

        #endregion

        #region Hosted Families Handling

        private List<FamilyInstance> GetHostedFamilies(Document doc, Floor floor)
        {
            List<FamilyInstance> hostedFamilies = new List<FamilyInstance>();

            try
            {
                // Get all dependent elements
                ICollection<ElementId> dependentIds = floor.GetDependentElements(
                    new ElementClassFilter(typeof(FamilyInstance)));

                foreach (ElementId id in dependentIds)
                {
                    FamilyInstance fi = doc.GetElement(id) as FamilyInstance;
                    if (fi != null && fi.Host != null && fi.Host.Id == floor.Id)
                    {
                        hostedFamilies.Add(fi);
                    }
                }

                // Also check using FilteredElementCollector as backup
                FilteredElementCollector collector = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilyInstance))
                    .WhereElementIsNotElementType();

                foreach (FamilyInstance fi in collector)
                {
                    if (fi.Host != null && fi.Host.Id == floor.Id && !hostedFamilies.Contains(fi))
                    {
                        hostedFamilies.Add(fi);
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error getting hosted families: {ex.Message}");
            }

            return hostedFamilies;
        }

        private void HandleHostedFamilies(Document doc, Floor floor1, Floor floor2,
            Line splitLine, XYZ splitStart, XYZ lineDir, List<FamilyInstance> hostedFamilies)
        {
            // Perpendicular to split line
            XYZ perpendicular = XYZ.BasisZ.CrossProduct(lineDir).Normalize();

            foreach (FamilyInstance family in hostedFamilies)
            {
                try
                {
                    LocationPoint locPoint = family.Location as LocationPoint;
                    if (locPoint == null) continue;

                    XYZ familyPoint = locPoint.Point;

                    // Determine which side of the split line the family is on
                    double side = (familyPoint - splitStart).DotProduct(perpendicular);

                    // Check if family crosses the split line
                    BoundingBoxXYZ bbox = family.get_BoundingBox(null);
                    if (bbox != null)
                    {
                        bool crossesSplit = DoesBoundingBoxCrossSplitLine(bbox, splitLine, splitStart, perpendicular);

                        if (crossesSplit)
                        {
                            // Family crosses split line - needs to be on both floors
                            // The original is already on floor1, create a copy for floor2

                            // First, we need to ensure the family is properly hosted
                            FamilyInstance copiedFamily = null;

                            try
                            {
                                // Get face reference from the floor
                                Options geomOptions = new Options();
                                geomOptions.ComputeReferences = true;
                                GeometryElement geomElem = floor2.get_Geometry(geomOptions);

                                Reference faceRef = null;
                                foreach (GeometryObject geomObj in geomElem)
                                {
                                    if (geomObj is Solid solid)
                                    {
                                        foreach (Face face in solid.Faces)
                                        {
                                            if (face is PlanarFace planarFace &&
                                                Math.Abs(planarFace.FaceNormal.Z - 1) < 0.01)
                                            {
                                                faceRef = face.Reference;
                                                break;
                                            }
                                        }
                                        if (faceRef != null) break;
                                    }
                                }

                                if (faceRef != null)
                                {
                                    // Create instance on the face
                                    copiedFamily = doc.Create.NewFamilyInstance(
                                        faceRef, familyPoint, XYZ.BasisX, family.Symbol);

                                    // Copy parameters
                                    CopyFamilyInstanceParameters(family, copiedFamily);
                                }
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"Failed to duplicate family across split: {ex.Message}");
                            }
                        }
                        else
                        {
                            // Family doesn't cross split - reassign to correct floor if needed
                            if (side > 0)
                            {
                                // Should be on floor1 (already is)
                            }
                            else
                            {
                                // Should be on floor2 - need to move it
                                try
                                {
                                    // This is tricky because we can't directly change the host
                                    // We need to recreate the family on the new floor
                                    Options geomOptions = new Options();
                                    geomOptions.ComputeReferences = true;
                                    GeometryElement geomElem = floor2.get_Geometry(geomOptions);

                                    Reference faceRef = null;
                                    foreach (GeometryObject geomObj in geomElem)
                                    {
                                        if (geomObj is Solid solid)
                                        {
                                            foreach (Face face in solid.Faces)
                                            {
                                                if (face is PlanarFace planarFace &&
                                                    Math.Abs(planarFace.FaceNormal.Z - 1) < 0.01)
                                                {
                                                    faceRef = face.Reference;
                                                    break;
                                                }
                                            }
                                            if (faceRef != null) break;
                                        }
                                    }

                                    if (faceRef != null)
                                    {
                                        // Create new instance on floor2
                                        FamilyInstance newFamily = doc.Create.NewFamilyInstance(
                                            faceRef, familyPoint, XYZ.BasisX, family.Symbol);

                                        // Copy parameters
                                        CopyFamilyInstanceParameters(family, newFamily);

                                        // Delete original
                                        doc.Delete(family.Id);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Debug.WriteLine($"Failed to move family to other floor: {ex.Message}");
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error handling hosted family: {ex.Message}");
                }
            }
        }

        private bool DoesBoundingBoxCrossSplitLine(BoundingBoxXYZ bbox, Line splitLine,
            XYZ splitStart, XYZ perpendicular)
        {
            // Get the min and max points of the bounding box
            XYZ min = bbox.Min;
            XYZ max = bbox.Max;

            // Check all 4 corners of the bounding box in plan
            XYZ[] corners = new XYZ[]
            {
                new XYZ(min.X, min.Y, min.Z),
                new XYZ(max.X, min.Y, min.Z),
                new XYZ(max.X, max.Y, min.Z),
                new XYZ(min.X, max.Y, min.Z)
            };

            // Check if corners are on different sides of the split line
            bool hasPositive = false;
            bool hasNegative = false;

            foreach (XYZ corner in corners)
            {
                double side = (corner - splitStart).DotProduct(perpendicular);
                if (side > 0) hasPositive = true;
                if (side < 0) hasNegative = true;
            }

            // If corners are on both sides, the box crosses the split line
            return hasPositive && hasNegative;
        }

        private void CopyFamilyInstanceParameters(FamilyInstance source, FamilyInstance target)
        {
            try
            {
                foreach (Parameter param in source.Parameters)
                {
                    if (param.IsReadOnly) continue;

                    Parameter targetParam = target.get_Parameter(param.Definition);
                    if (targetParam != null && !targetParam.IsReadOnly)
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
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error copying family parameters: {ex.Message}");
            }
        }

        #endregion

        #region Geometry Processing (Reused from original)

        private Tuple<List<Curve>, List<List<Curve>>> GetFloorBoundaryLoops(Floor floor)
        {
            List<Curve> outerBoundary = new List<Curve>();
            List<List<Curve>> innerBoundaries = new List<List<Curve>>();

            Options options = new Options();
            options.ComputeReferences = true;
            options.IncludeNonVisibleObjects = true;

            GeometryElement geomElem = floor.get_Geometry(options);

            foreach (GeometryObject geomObj in geomElem)
            {
                if (geomObj is Solid solid && solid.Volume > 0)
                {
                    foreach (Face face in solid.Faces)
                    {
                        if (face is PlanarFace planarFace &&
                            Math.Abs(planarFace.FaceNormal.Z - 1) < 0.01)
                        {
                            EdgeArray largestLoop = null;
                            double maxArea = 0;
                            List<EdgeArray> innerLoops = new List<EdgeArray>();

                            foreach (EdgeArray edgeArray in face.EdgeLoops)
                            {
                                double area = GetLoopArea(edgeArray);
                                if (area > maxArea)
                                {
                                    if (largestLoop != null)
                                        innerLoops.Add(largestLoop);
                                    maxArea = area;
                                    largestLoop = edgeArray;
                                }
                                else
                                {
                                    innerLoops.Add(edgeArray);
                                }
                            }

                            if (largestLoop != null)
                            {
                                foreach (Edge edge in largestLoop)
                                {
                                    outerBoundary.Add(edge.AsCurve());
                                }
                            }

                            foreach (EdgeArray innerLoop in innerLoops)
                            {
                                List<Curve> innerCurves = new List<Curve>();
                                foreach (Edge edge in innerLoop)
                                {
                                    innerCurves.Add(edge.AsCurve());
                                }
                                if (innerCurves.Count > 0)
                                    innerBoundaries.Add(innerCurves);
                            }

                            return new Tuple<List<Curve>, List<List<Curve>>>(
                                outerBoundary, innerBoundaries);
                        }
                    }
                }
            }

            return new Tuple<List<Curve>, List<List<Curve>>>(
                outerBoundary, innerBoundaries);
        }

        private double GetLoopArea(EdgeArray edges)
        {
            double area = 0;
            List<XYZ> points = new List<XYZ>();

            foreach (Edge edge in edges)
            {
                points.Add(edge.AsCurve().GetEndPoint(0));
            }

            for (int i = 0; i < points.Count; i++)
            {
                int j = (i + 1) % points.Count;
                area += points[i].X * points[j].Y;
                area -= points[j].X * points[i].Y;
            }

            return Math.Abs(area) / 2.0;
        }

        private Tuple<List<Curve>, List<Curve>, List<List<Curve>>, List<List<Curve>>>
            CreateSplitProfilesWithOpenings(
                List<Curve> outerBoundary,
                List<List<Curve>> innerBoundaries,
                XYZ splitStart,
                XYZ splitEnd,
                XYZ splitDirection,
                StringBuilder debugInfo)
        {
            List<Curve> profile1 = new List<Curve>();
            List<Curve> profile2 = new List<Curve>();
            List<List<Curve>> innerLoops1 = new List<List<Curve>>();
            List<List<Curve>> innerLoops2 = new List<List<Curve>>();

            // Create the split line
            Line splitLine = Line.CreateBound(splitStart, splitEnd);

            // Determine which side each boundary segment belongs to
            XYZ normal = XYZ.BasisZ;
            XYZ perpendicular = normal.CrossProduct(splitDirection).Normalize();

            // Process outer boundary curves
            foreach (Curve curve in outerBoundary)
            {
                ProcessCurveForSplit(curve, splitLine, perpendicular, splitStart,
                    profile1, profile2, debugInfo);
            }

            // Add the split line to both profiles
            profile1.Add(Line.CreateBound(splitStart, splitEnd));
            profile2.Add(Line.CreateBound(splitEnd, splitStart));

            // Sort curves to form continuous loops
            profile1 = SortAndFixCurveLoop(profile1, debugInfo);
            profile2 = SortAndFixCurveLoop(profile2, debugInfo);

            // Process inner boundaries (openings)
            int openingIndex = 0;
            foreach (var innerLoop in innerBoundaries)
            {
                debugInfo.AppendLine($"\nProcessing opening {openingIndex++}:");

                try
                {
                    // Check if opening intersects split line
                    List<XYZ> intersections = new List<XYZ>();
                    foreach (Curve curve in innerLoop)
                    {
                        IntersectionResultArray results;
                        if (curve.Intersect(splitLine, out results) == SetComparisonResult.Overlap)
                        {
                            foreach (IntersectionResult result in results)
                            {
                                intersections.Add(result.XYZPoint);
                            }
                        }
                    }

                    if (intersections.Count == 0)
                    {
                        // Opening doesn't intersect - assign to one side
                        XYZ center = GetLoopCenter(innerLoop);
                        double side = (center - splitStart).DotProduct(perpendicular);

                        if (side > 0)
                        {
                            innerLoops1.Add(innerLoop);
                            debugInfo.AppendLine("  Assigned to side 1 (no intersection)");
                        }
                        else
                        {
                            innerLoops2.Add(innerLoop);
                            debugInfo.AppendLine("  Assigned to side 2 (no intersection)");
                        }
                    }
                    else
                    {
                        // Opening intersects split line
                        debugInfo.AppendLine($"  Opening intersects at {intersections.Count} points");

                        // Try to split the opening
                        var splitOpeningResult = SplitOpeningLoop(innerLoop, splitLine,
                            perpendicular, splitStart, debugInfo);

                        if (splitOpeningResult.Item1 != null && splitOpeningResult.Item1.Count >= 3)
                        {
                            innerLoops1.Add(splitOpeningResult.Item1);
                            debugInfo.AppendLine($"  Added split opening to side 1 with {splitOpeningResult.Item1.Count} curves");
                        }

                        if (splitOpeningResult.Item2 != null && splitOpeningResult.Item2.Count >= 3)
                        {
                            innerLoops2.Add(splitOpeningResult.Item2);
                            debugInfo.AppendLine($"  Added split opening to side 2 with {splitOpeningResult.Item2.Count} curves");
                        }
                    }
                }
                catch (Exception ex)
                {
                    debugInfo.AppendLine($"  ERROR processing opening: {ex.Message}");
                }
            }

            return new Tuple<List<Curve>, List<Curve>, List<List<Curve>>, List<List<Curve>>>(
                profile1, profile2, innerLoops1, innerLoops2);
        }

        private Tuple<List<Curve>, List<Curve>> SplitOpeningLoop(
            List<Curve> openingLoop, Line splitLine, XYZ perpendicular,
            XYZ splitStart, StringBuilder debugInfo)
        {
            List<Curve> side1Curves = new List<Curve>();
            List<Curve> side2Curves = new List<Curve>();

            try
            {
                // Process each curve in the opening
                foreach (Curve curve in openingLoop)
                {
                    IntersectionResultArray results;
                    if (curve.Intersect(splitLine, out results) == SetComparisonResult.Overlap)
                    {
                        // Curve intersects split line
                        XYZ intersectionPoint = results.get_Item(0).XYZPoint;
                        XYZ curveStart = curve.GetEndPoint(0);
                        XYZ curveEnd = curve.GetEndPoint(1);

                        // Create segments on each side
                        if (curveStart.DistanceTo(intersectionPoint) > 0.01)
                        {
                            Curve seg1 = CreateSegment(curve, curveStart, intersectionPoint);
                            XYZ midPoint = EvaluateMidpoint(seg1);
                            double side = (midPoint - splitStart).DotProduct(perpendicular);

                            if (side > 0)
                                side1Curves.Add(seg1);
                            else
                                side2Curves.Add(seg1);
                        }

                        if (intersectionPoint.DistanceTo(curveEnd) > 0.01)
                        {
                            Curve seg2 = CreateSegment(curve, intersectionPoint, curveEnd);
                            XYZ midPoint = EvaluateMidpoint(seg2);
                            double side = (midPoint - splitStart).DotProduct(perpendicular);

                            if (side > 0)
                                side1Curves.Add(seg2);
                            else
                                side2Curves.Add(seg2);
                        }
                    }
                    else
                    {
                        // Curve doesn't intersect - assign to one side
                        XYZ midPoint = EvaluateMidpoint(curve);
                        double side = (midPoint - splitStart).DotProduct(perpendicular);

                        if (side > 0)
                            side1Curves.Add(curve);
                        else
                            side2Curves.Add(curve);
                    }
                }

                // Now we need to close each opening loop with a segment along the split line
                if (side1Curves.Count > 0 && side2Curves.Count > 0)
                {
                    // Find the intersection points on the split line
                    List<XYZ> side1Points = new List<XYZ>();
                    List<XYZ> side2Points = new List<XYZ>();

                    // Collect endpoints that are on the split line
                    foreach (var c in side1Curves)
                    {
                        XYZ start = c.GetEndPoint(0);
                        XYZ end = c.GetEndPoint(1);

                        if (IsPointOnLine(start, splitLine)) side1Points.Add(start);
                        if (IsPointOnLine(end, splitLine)) side1Points.Add(end);
                    }

                    foreach (var c in side2Curves)
                    {
                        XYZ start = c.GetEndPoint(0);
                        XYZ end = c.GetEndPoint(1);

                        if (IsPointOnLine(start, splitLine)) side2Points.Add(start);
                        if (IsPointOnLine(end, splitLine)) side2Points.Add(end);
                    }

                    // Remove duplicates and sort
                    side1Points = side1Points.Distinct(new XYZComparer()).ToList();
                    side2Points = side2Points.Distinct(new XYZComparer()).ToList();

                    // Add closing segments along split line
                    if (side1Points.Count >= 2)
                    {
                        side1Points = side1Points.OrderBy(p => splitLine.Project(p).Parameter).ToList();
                        side1Curves.Add(Line.CreateBound(side1Points.First(), side1Points.Last()));
                    }

                    if (side2Points.Count >= 2)
                    {
                        side2Points = side2Points.OrderBy(p => splitLine.Project(p).Parameter).ToList();
                        side2Curves.Add(Line.CreateBound(side2Points.Last(), side2Points.First()));
                    }

                    // Sort curves into loops
                    side1Curves = SortCurvesIntoLoop(side1Curves);
                    side2Curves = SortCurvesIntoLoop(side2Curves);
                }
            }
            catch (Exception ex)
            {
                debugInfo.AppendLine($"  Error splitting opening: {ex.Message}");
            }

            return new Tuple<List<Curve>, List<Curve>>(side1Curves, side2Curves);
        }

        private bool IsPointOnLine(XYZ point, Line line)
        {
            XYZ projection = line.Project(point).XYZPoint;
            return point.DistanceTo(projection) < 0.01;
        }

        private class XYZComparer : IEqualityComparer<XYZ>
        {
            public bool Equals(XYZ x, XYZ y)
            {
                if (ReferenceEquals(x, y)) return true;
                if (x == null || y == null) return false;
                return x.DistanceTo(y) < 0.01;
            }

            public int GetHashCode(XYZ obj)
            {
                if (obj == null) return 0;
                return obj.X.GetHashCode() ^ obj.Y.GetHashCode() ^ obj.Z.GetHashCode();
            }
        }

        private void ProcessCurveForSplit(Curve curve, Line splitLine, XYZ perpendicular,
            XYZ splitStart, List<Curve> profile1, List<Curve> profile2, StringBuilder debugInfo)
        {
            IntersectionResultArray results;
            if (curve.Intersect(splitLine, out results) == SetComparisonResult.Overlap)
            {
                // Curve intersects split line
                XYZ intersectionPoint = results.get_Item(0).XYZPoint;
                XYZ curveStart = curve.GetEndPoint(0);
                XYZ curveEnd = curve.GetEndPoint(1);

                // Create segments
                if (curveStart.DistanceTo(intersectionPoint) > 0.01)
                {
                    try
                    {
                        Curve seg1 = CreateSegment(curve, curveStart, intersectionPoint);
                        XYZ midPoint = EvaluateMidpoint(seg1);
                        double side = (midPoint - splitStart).DotProduct(perpendicular);

                        if (side > 0)
                            profile1.Add(seg1);
                        else
                            profile2.Add(seg1);
                    }
                    catch (Exception ex)
                    {
                        debugInfo.AppendLine($"  Failed to create segment 1: {ex.Message}");
                    }
                }

                if (intersectionPoint.DistanceTo(curveEnd) > 0.01)
                {
                    try
                    {
                        Curve seg2 = CreateSegment(curve, intersectionPoint, curveEnd);
                        XYZ midPoint = EvaluateMidpoint(seg2);
                        double side = (midPoint - splitStart).DotProduct(perpendicular);

                        if (side > 0)
                            profile1.Add(seg2);
                        else
                            profile2.Add(seg2);
                    }
                    catch (Exception ex)
                    {
                        debugInfo.AppendLine($"  Failed to create segment 2: {ex.Message}");
                    }
                }
            }
            else
            {
                // Curve doesn't intersect - assign to one side
                XYZ midPoint = EvaluateMidpoint(curve);
                double side = (midPoint - splitStart).DotProduct(perpendicular);

                if (side > 0)
                    profile1.Add(curve);
                else
                    profile2.Add(curve);
            }
        }

        private Curve CreateSegment(Curve originalCurve, XYZ start, XYZ end)
        {
            if (originalCurve is Line)
            {
                return Line.CreateBound(start, end);
            }
            else if (originalCurve is Arc arc)
            {
                try
                {
                    // For arcs, try to maintain the arc
                    double startParam = originalCurve.Project(start).Parameter;
                    double endParam = originalCurve.Project(end).Parameter;

                    Curve newArc = arc.Clone();
                    newArc.MakeBound(startParam, endParam);
                    return newArc;
                }
                catch
                {
                    // If arc creation fails, use line
                    return Line.CreateBound(start, end);
                }
            }
            else
            {
                // For other curves, use line approximation
                return Line.CreateBound(start, end);
            }
        }

        private XYZ EvaluateMidpoint(Curve curve)
        {
            try
            {
                double midParam = (curve.GetEndParameter(0) + curve.GetEndParameter(1)) / 2.0;
                return curve.Evaluate(midParam, false);
            }
            catch
            {
                return (curve.GetEndPoint(0) + curve.GetEndPoint(1)) / 2.0;
            }
        }

        private List<Curve> SortAndFixCurveLoop(List<Curve> curves, StringBuilder debugInfo)
        {
            if (curves.Count < 3)
            {
                debugInfo.AppendLine($"  Loop has only {curves.Count} curves");
                return curves;
            }

            List<Curve> sorted = new List<Curve>();
            List<Curve> remaining = new List<Curve>(curves);

            // Start with first curve
            sorted.Add(remaining[0]);
            remaining.RemoveAt(0);

            int iterations = 0;
            while (remaining.Count > 0 && iterations < curves.Count * 2)
            {
                XYZ currentEnd = sorted.Last().GetEndPoint(1);
                bool found = false;

                for (int i = 0; i < remaining.Count; i++)
                {
                    Curve curve = remaining[i];
                    XYZ curveStart = curve.GetEndPoint(0);
                    XYZ curveEnd = curve.GetEndPoint(1);

                    if (currentEnd.DistanceTo(curveStart) < 0.01)
                    {
                        sorted.Add(curve);
                        remaining.RemoveAt(i);
                        found = true;
                        break;
                    }
                    else if (currentEnd.DistanceTo(curveEnd) < 0.01)
                    {
                        sorted.Add(curve.CreateReversed());
                        remaining.RemoveAt(i);
                        found = true;
                        break;
                    }
                }

                if (!found)
                {
                    // Check if loop closes
                    XYZ loopStart = sorted.First().GetEndPoint(0);
                    if (currentEnd.DistanceTo(loopStart) < 0.01)
                    {
                        debugInfo.AppendLine($"  Loop closed with {sorted.Count} curves");
                        break;
                    }
                    else
                    {
                        // Try to close gap
                        debugInfo.AppendLine($"  Gap in loop: {currentEnd.DistanceTo(loopStart):F3} feet");
                        try
                        {
                            Line closingLine = Line.CreateBound(currentEnd, loopStart);
                            sorted.Add(closingLine);
                            debugInfo.AppendLine("  Added closing line");
                        }
                        catch
                        {
                            debugInfo.AppendLine("  Failed to add closing line");
                        }
                        break;
                    }
                }

                iterations++;
            }

            if (remaining.Count > 0)
            {
                debugInfo.AppendLine($"  WARNING: {remaining.Count} curves not connected");
            }

            return sorted;
        }

        private bool ValidateCurveLoop(List<Curve> curves, out string error)
        {
            error = "";

            if (curves.Count < 3)
            {
                error = $"Too few curves ({curves.Count})";
                return false;
            }

            // Check if curves form a closed loop
            for (int i = 0; i < curves.Count; i++)
            {
                XYZ currentEnd = curves[i].GetEndPoint(1);
                XYZ nextStart = curves[(i + 1) % curves.Count].GetEndPoint(0);

                double distance = currentEnd.DistanceTo(nextStart);
                if (distance > 0.01)
                {
                    error = $"Gap at curve {i}: {distance:F3} feet";
                    return false;
                }
            }

            // Check for zero-length curves
            for (int i = 0; i < curves.Count; i++)
            {
                if (curves[i].Length < 0.001)
                {
                    error = $"Zero-length curve at index {i}";
                    return false;
                }
            }

            return true;
        }

        private XYZ GetLoopCenter(List<Curve> loop)
        {
            double sumX = 0, sumY = 0, sumZ = 0;
            int count = 0;

            foreach (Curve curve in loop)
            {
                XYZ start = curve.GetEndPoint(0);
                XYZ end = curve.GetEndPoint(1);
                sumX += start.X + end.X;
                sumY += start.Y + end.Y;
                sumZ += start.Z + end.Z;
                count += 2;
            }

            return new XYZ(sumX / count, sumY / count, sumZ / count);
        }

        private List<Curve> SortCurvesIntoLoop(List<Curve> curves)
        {
            if (curves.Count == 0) return curves;

            List<Curve> sorted = new List<Curve>();
            List<Curve> remaining = new List<Curve>(curves);

            sorted.Add(remaining[0]);
            remaining.RemoveAt(0);

            while (remaining.Count > 0)
            {
                XYZ currentEnd = sorted.Last().GetEndPoint(1);
                bool found = false;

                for (int i = 0; i < remaining.Count; i++)
                {
                    Curve curve = remaining[i];
                    XYZ curveStart = curve.GetEndPoint(0);
                    XYZ curveEnd = curve.GetEndPoint(1);

                    if (currentEnd.DistanceTo(curveStart) < 0.001)
                    {
                        sorted.Add(curve);
                        remaining.RemoveAt(i);
                        found = true;
                        break;
                    }
                    else if (currentEnd.DistanceTo(curveEnd) < 0.001)
                    {
                        sorted.Add(curve.CreateReversed());
                        remaining.RemoveAt(i);
                        found = true;
                        break;
                    }
                }

                if (!found)
                {
                    XYZ loopStart = sorted.First().GetEndPoint(0);
                    if (currentEnd.DistanceTo(loopStart) < 0.001)
                    {
                        break;
                    }
                    else
                    {
                        break;
                    }
                }
            }

            return sorted;
        }

        internal static PushButtonData GetButtonData()
        {
            string buttonInternalName = "btnSplitFloor";
            string buttonTitle = "Split Floor";

            Common.ButtonDataClass myButtonData = new Common.ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Yellow_32,
                Properties.Resources.Yellow_16,
                "Splits a floor into two parts using a model line while preserving hosted families and openings");

            return myButtonData.Data;
        }

        #endregion
    }

    public class FloorSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            return elem is Floor;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }

    public class ModelCurveSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            return elem is ModelCurve;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }
}