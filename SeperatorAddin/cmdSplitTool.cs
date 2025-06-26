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

                using (Transaction trans = new Transaction(doc, "Split Floor"))
                {
                    trans.Start();

                    try
                    {
                        // Get floor properties
                        FloorType floorType = selectedFloor.FloorType;
                        Level level = doc.GetElement(selectedFloor.LevelId) as Level;
                        double offset = selectedFloor.get_Parameter(
                            BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM).AsDouble();

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

                        // Sort intersection points by distance from line start
                        intersectionPoints = intersectionPoints
                            .OrderBy(p => lineStart.DistanceTo(p))
                            .ToList();

                        // Take first and last intersection
                        XYZ splitStart = intersectionPoints.First();
                        XYZ splitEnd = intersectionPoints.Last();

                        // DEBUG: Log opening information
                        StringBuilder debugInfo = new StringBuilder();
                        debugInfo.AppendLine($"Found {floorLoops.Item2.Count} openings in floor");

                        // Check which openings intersect the split line
                        int intersectingOpenings = 0;
                        foreach (var opening in floorLoops.Item2)
                        {
                            bool intersects = false;
                            foreach (var curve in opening)
                            {
                                IntersectionResultArray results;
                                if (curve.Intersect(extendedLine, out results) == SetComparisonResult.Overlap)
                                {
                                    intersects = true;
                                    break;
                                }
                            }
                            if (intersects) intersectingOpenings++;
                        }
                        debugInfo.AppendLine($"{intersectingOpenings} openings intersect the split line");

                        // Create profile loops for two floors - with alternative approach
                        var profiles = CreateSplitProfilesWithOpeningsAlternative(
                            floorLoops.Item1, floorLoops.Item2, splitStart, splitEnd, lineDir, debugInfo);

                        if (profiles.Item1.Count == 0 || profiles.Item2.Count == 0)
                        {
                            trans.RollBack();
                            TaskDialog.Show("Error", "Failed to create valid split profiles.\n\n" + debugInfo.ToString());
                            return Result.Failed;
                        }

                        // Create new floors
                        int floorsCreated = 0;
                        string creationErrors = "";

                        // First floor
                        try
                        {
                            List<CurveLoop> loops1 = new List<CurveLoop>();

                            // Validate and add outer loop
                            if (ValidateCurveLoop(profiles.Item1, out string error1))
                            {
                                CurveLoop outerLoop1 = new CurveLoop();
                                foreach (Curve c in profiles.Item1)
                                    outerLoop1.Append(c);
                                loops1.Add(outerLoop1);

                                // Add inner loops (openings) - only if they're valid
                                foreach (var innerLoop in profiles.Item3)
                                {
                                    if (ValidateCurveLoop(innerLoop, out string openingError))
                                    {
                                        CurveLoop opening = new CurveLoop();
                                        foreach (Curve c in innerLoop)
                                            opening.Append(c);
                                        loops1.Add(opening);
                                    }
                                    else
                                    {
                                        debugInfo.AppendLine($"Skipping invalid opening: {openingError}");
                                    }
                                }

                                Floor floor1 = Floor.Create(doc, loops1, floorType.Id, level.Id);

                                if (floor1 != null)
                                {
                                    floor1.get_Parameter(
                                        BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM).Set(offset);
                                    floorsCreated++;
                                }
                            }
                            else
                            {
                                creationErrors += $"Floor 1: {error1}\n";
                            }
                        }
                        catch (Exception ex)
                        {
                            creationErrors += $"Floor 1 Exception: {ex.Message}\n";
                        }

                        // Second floor
                        try
                        {
                            List<CurveLoop> loops2 = new List<CurveLoop>();

                            // Validate and add outer loop
                            if (ValidateCurveLoop(profiles.Item2, out string error2))
                            {
                                CurveLoop outerLoop2 = new CurveLoop();
                                foreach (Curve c in profiles.Item2)
                                    outerLoop2.Append(c);
                                loops2.Add(outerLoop2);

                                // Add inner loops (openings) - only if they're valid
                                foreach (var innerLoop in profiles.Item4)
                                {
                                    if (ValidateCurveLoop(innerLoop, out string openingError))
                                    {
                                        CurveLoop opening = new CurveLoop();
                                        foreach (Curve c in innerLoop)
                                            opening.Append(c);
                                        loops2.Add(opening);
                                    }
                                    else
                                    {
                                        debugInfo.AppendLine($"Skipping invalid opening: {openingError}");
                                    }
                                }

                                Floor floor2 = Floor.Create(doc, loops2, floorType.Id, level.Id);

                                if (floor2 != null)
                                {
                                    floor2.get_Parameter(
                                        BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM).Set(offset);
                                    floorsCreated++;
                                }
                            }
                            else
                            {
                                creationErrors += $"Floor 2: {error2}\n";
                            }
                        }
                        catch (Exception ex)
                        {
                            creationErrors += $"Floor 2 Exception: {ex.Message}\n";
                        }

                        if (floorsCreated == 2)
                        {
                            // Delete original floor
                            doc.Delete(selectedFloor.Id);
                            trans.Commit();
                            TaskDialog.Show("Success", "Floor split into 2 parts.");
                            return Result.Succeeded;
                        }
                        else if (floorsCreated == 1)
                        {
                            trans.RollBack();
                            TaskDialog.Show("Partial Success",
                                $"Only {floorsCreated} floor was created.\n\nErrors:\n{creationErrors}\n\nDebug Info:\n{debugInfo}");
                            return Result.Failed;
                        }
                        else
                        {
                            trans.RollBack();
                            TaskDialog.Show("Error",
                                $"Failed to create both floor parts.\n\nErrors:\n{creationErrors}\n\nDebug Info:\n{debugInfo}");
                            return Result.Failed;
                        }
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

        // Alternative approach that's more conservative with opening splits
        private Tuple<List<Curve>, List<Curve>, List<List<Curve>>, List<List<Curve>>>
            CreateSplitProfilesWithOpeningsAlternative(
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

            // Process inner boundaries (openings) with alternative approach
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
                        // Opening intersects split line - try conservative approach
                        debugInfo.AppendLine($"  Opening intersects at {intersections.Count} points");

                        // For now, skip openings that intersect the split line
                        // This is more conservative but should work
                        debugInfo.AppendLine("  WARNING: Skipping opening that intersects split line");

                        // Alternative: Try to split the opening
                        /*
                        var splitResult = TrySplitOpeningConservative(
                            innerLoop, intersections, splitLine, perpendicular, splitStart, debugInfo);
                        
                        if (splitResult.Item1 != null && splitResult.Item1.Count >= 3)
                            innerLoops1.Add(splitResult.Item1);
                        
                        if (splitResult.Item2 != null && splitResult.Item2.Count >= 3)
                            innerLoops2.Add(splitResult.Item2);
                        */
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

        // Process a single curve for splitting
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

        // Create a segment of a curve
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

        // Evaluate midpoint of a curve
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

        // Sort and fix curve loop with debugging
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

        // Validate curve loop
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

        // Get floor boundary loops (keeping existing implementation)
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