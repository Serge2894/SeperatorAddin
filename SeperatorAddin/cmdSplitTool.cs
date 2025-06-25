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

                        // Create profile loops for two floors
                        var profiles = CreateSplitProfilesWithOpenings(
                            floorLoops.Item1, floorLoops.Item2, splitStart, splitEnd, lineDir);

                        if (profiles.Item1.Count == 0 || profiles.Item2.Count == 0)
                        {
                            trans.RollBack();
                            TaskDialog.Show("Error", "Failed to create valid split profiles.");
                            return Result.Failed;
                        }

                        // Create new floors
                        int floorsCreated = 0;

                        // First floor
                        try
                        {
                            List<CurveLoop> loops1 = new List<CurveLoop>();

                            // Add outer loop
                            CurveLoop outerLoop1 = new CurveLoop();
                            foreach (Curve c in profiles.Item1)
                                outerLoop1.Append(c);
                            loops1.Add(outerLoop1);

                            // Add inner loops (openings)
                            foreach (var innerLoop in profiles.Item3)
                            {
                                CurveLoop opening = new CurveLoop();
                                foreach (Curve c in innerLoop)
                                    opening.Append(c);
                                loops1.Add(opening);
                            }

                            Floor floor1 = Floor.Create(doc, loops1, floorType.Id, level.Id);

                            if (floor1 != null)
                            {
                                floor1.get_Parameter(
                                    BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM).Set(offset);
                                floorsCreated++;
                            }
                        }
                        catch { }

                        // Second floor
                        try
                        {
                            List<CurveLoop> loops2 = new List<CurveLoop>();

                            // Add outer loop
                            CurveLoop outerLoop2 = new CurveLoop();
                            foreach (Curve c in profiles.Item2)
                                outerLoop2.Append(c);
                            loops2.Add(outerLoop2);

                            // Add inner loops (openings)
                            foreach (var innerLoop in profiles.Item4)
                            {
                                CurveLoop opening = new CurveLoop();
                                foreach (Curve c in innerLoop)
                                    opening.Append(c);
                                loops2.Add(opening);
                            }

                            Floor floor2 = Floor.Create(doc, loops2, floorType.Id, level.Id);

                            if (floor2 != null)
                            {
                                floor2.get_Parameter(
                                    BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM).Set(offset);
                                floorsCreated++;
                            }
                        }
                        catch { }

                        if (floorsCreated == 2)
                        {
                            // Delete original floor
                            doc.Delete(selectedFloor.Id);
                            trans.Commit();
                            TaskDialog.Show("Success", "Floor split into 2 parts.");
                            return Result.Succeeded;
                        }
                        else
                        {
                            trans.RollBack();
                            TaskDialog.Show("Error", "Failed to create both floor parts.");
                            return Result.Failed;
                        }
                    }
                    catch (Exception ex)
                    {
                        trans.RollBack();
                        TaskDialog.Show("Error", $"Failed: {ex.Message}");
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
                    // Find the top horizontal face
                    foreach (Face face in solid.Faces)
                    {
                        if (face is PlanarFace planarFace &&
                            Math.Abs(planarFace.FaceNormal.Z - 1) < 0.01) // Top face
                        {
                            // Get all edge loops
                            EdgeArray largestLoop = null;
                            double maxArea = 0;
                            List<EdgeArray> innerLoops = new List<EdgeArray>();

                            foreach (EdgeArray edgeArray in face.EdgeLoops)
                            {
                                double area = GetLoopArea(edgeArray);
                                if (area > maxArea)
                                {
                                    // Previous largest becomes inner loop
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

                            // Extract outer boundary curves
                            if (largestLoop != null)
                            {
                                foreach (Edge edge in largestLoop)
                                {
                                    outerBoundary.Add(edge.AsCurve());
                                }
                            }

                            // Extract inner boundary curves
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

            // Shoelace formula
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
                XYZ splitDirection)
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
            for (int i = 0; i < outerBoundary.Count; i++)
            {
                Curve curve = outerBoundary[i];
                XYZ curveStart = curve.GetEndPoint(0);
                XYZ curveEnd = curve.GetEndPoint(1);

                // Check if curve intersects split line
                IntersectionResultArray results;
                if (curve.Intersect(splitLine, out results) == SetComparisonResult.Overlap)
                {
                    // Curve intersects - need to split it
                    XYZ intersectionPoint = results.get_Item(0).XYZPoint;

                    // Create two segments
                    if (curveStart.DistanceTo(intersectionPoint) > 0.001)
                    {
                        Line seg1 = Line.CreateBound(curveStart, intersectionPoint);

                        // Determine which side
                        XYZ midPoint = (curveStart + intersectionPoint) / 2;
                        XYZ toMid = midPoint - splitStart;
                        double side = toMid.DotProduct(perpendicular);

                        if (side > 0)
                            profile1.Add(seg1);
                        else
                            profile2.Add(seg1);
                    }

                    if (intersectionPoint.DistanceTo(curveEnd) > 0.001)
                    {
                        Line seg2 = Line.CreateBound(intersectionPoint, curveEnd);

                        // Determine which side
                        XYZ midPoint = (intersectionPoint + curveEnd) / 2;
                        XYZ toMid = midPoint - splitStart;
                        double side = toMid.DotProduct(perpendicular);

                        if (side > 0)
                            profile1.Add(seg2);
                        else
                            profile2.Add(seg2);
                    }
                }
                else
                {
                    // Curve doesn't intersect - assign to one side
                    XYZ midPoint = (curveStart + curveEnd) / 2;
                    XYZ toMid = midPoint - splitStart;
                    double side = toMid.DotProduct(perpendicular);

                    if (side > 0)
                        profile1.Add(curve);
                    else
                        profile2.Add(curve);
                }
            }

            // Add the split line to both profiles (in opposite directions)
            profile1.Add(Line.CreateBound(splitStart, splitEnd));
            profile2.Add(Line.CreateBound(splitEnd, splitStart));

            // Sort curves to form continuous loops
            profile1 = SortCurvesIntoLoop(profile1);
            profile2 = SortCurvesIntoLoop(profile2);

            // Process inner boundaries (openings)
            foreach (var innerLoop in innerBoundaries)
            {
                // Check which side the opening belongs to
                XYZ loopCenter = GetLoopCenter(innerLoop);
                XYZ toCenter = loopCenter - splitStart;
                double side = toCenter.DotProduct(perpendicular);

                // Check if opening intersects split line
                bool intersectsSplit = false;
                foreach (Curve curve in innerLoop)
                {
                    IntersectionResultArray results;
                    if (curve.Intersect(splitLine, out results) == SetComparisonResult.Overlap)
                    {
                        intersectsSplit = true;
                        break;
                    }
                }

                if (!intersectsSplit)
                {
                    // Opening doesn't intersect split - assign to appropriate side
                    if (side > 0)
                        innerLoops1.Add(innerLoop);
                    else
                        innerLoops2.Add(innerLoop);
                }
                // If opening intersects split line, it's cut and we ignore it
                // (more complex handling would be needed to split the opening)
            }

            return new Tuple<List<Curve>, List<Curve>, List<List<Curve>>, List<List<Curve>>>(
                profile1, profile2, innerLoops1, innerLoops2);
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

            // Start with first curve
            sorted.Add(remaining[0]);
            remaining.RemoveAt(0);

            // Build the loop
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
                        // Need to reverse the curve
                        sorted.Add(curve.CreateReversed());
                        remaining.RemoveAt(i);
                        found = true;
                        break;
                    }
                }

                if (!found)
                {
                    // Try to close the loop
                    XYZ loopStart = sorted.First().GetEndPoint(0);
                    if (currentEnd.DistanceTo(loopStart) < 0.001)
                    {
                        // Loop is closed
                        break;
                    }
                    else
                    {
                        // Can't continue building loop
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