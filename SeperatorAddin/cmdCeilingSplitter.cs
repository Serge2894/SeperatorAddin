using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using SeperatorAddin.Common;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;

namespace SeperatorAddin
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class cmdCeilingSplitter : IExternalCommand
    {
        // Tolerance for geometric comparisons
        private const double TOLERANCE = 0.001; // ~1/32" in feet
        private const double OPENING_MERGE_TOLERANCE = 0.1; // Distance to merge split openings

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

                // Select a ceiling
                Reference ceilingRef = uiDoc.Selection.PickObject(
                    ObjectType.Element,
                    new CeilingSelectionFilter(),
                    "Select a ceiling to split");

                Ceiling selectedCeiling = doc.GetElement(ceilingRef) as Ceiling;

                // Select ONE model line
                Reference lineRef = uiDoc.Selection.PickObject(
                    ObjectType.Element,
                    new Utils.ModelCurveSelectionFilter(),
                    "Select ONE straight model line that crosses the ceiling");

                ModelCurve modelLine = doc.GetElement(lineRef) as ModelCurve;
                Line splitLine = modelLine.GeometryCurve as Line;

                if (splitLine == null)
                {
                    TaskDialog.Show("Error", "Please select a straight line.");
                    return Result.Failed;
                }

                using (Transaction trans = new Transaction(doc, "Split Ceiling with Openings"))
                {
                    trans.Start();

                    try
                    {
                        // Get ceiling properties
                        CeilingType ceilingType = doc.GetElement(selectedCeiling.GetTypeId()) as CeilingType;
                        Level level = doc.GetElement(selectedCeiling.LevelId) as Level;
                        double offset = selectedCeiling.get_Parameter(
                            BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM).AsDouble();

                        // Collect face-based families before splitting
                        List<FamilyInstance> hostedFamilies = GetHostedFamilies(doc, selectedCeiling);
                        Dictionary<FamilyInstance, XYZ> familyLocations = new Dictionary<FamilyInstance, XYZ>();

                        foreach (var family in hostedFamilies)
                        {
                            LocationPoint locPoint = family.Location as LocationPoint;
                            if (locPoint != null)
                                familyLocations[family] = locPoint.Point;
                        }

                        // Get ceiling boundary curves including openings
                        var ceilingLoops = GetCeilingBoundaryLoops(selectedCeiling);

                        if (ceilingLoops.Item1.Count < 3)
                        {
                            trans.RollBack();
                            TaskDialog.Show("Error", "Could not get valid ceiling boundary.");
                            return Result.Failed;
                        }

                        // Extend split line
                        XYZ lineDir = splitLine.Direction;
                        XYZ lineStart = splitLine.GetEndPoint(0) - lineDir * 1000;
                        XYZ lineEnd = splitLine.GetEndPoint(1) + lineDir * 1000;
                        Line extendedLine = Line.CreateBound(lineStart, lineEnd);

                        // Find intersections with outer boundary
                        List<XYZ> intersectionPoints = new List<XYZ>();

                        foreach (Curve boundaryCurve in ceilingLoops.Item1)
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
                                "Split line must cross the ceiling boundary at two points.\n" +
                                "Make sure the line extends across the entire ceiling.");
                            return Result.Failed;
                        }

                        // Sort intersection points by distance from line start
                        intersectionPoints = intersectionPoints
                            .OrderBy(p => lineStart.DistanceTo(p))
                            .ToList();

                        // Take first and last intersection
                        XYZ splitStart = intersectionPoints.First();
                        XYZ splitEnd = intersectionPoints.Last();

                        // Create profile loops for two ceilings
                        StringBuilder debugInfo = new StringBuilder();
                        var profiles = CreateSplitProfilesImproved(
                            ceilingLoops.Item1, ceilingLoops.Item2, splitStart, splitEnd, lineDir, debugInfo);

                        if (profiles == null || profiles.Item1.Count == 0 || profiles.Item2.Count == 0)
                        {
                            trans.RollBack();
                            TaskDialog.Show("Error", "Failed to create valid split profiles.\n\n" + debugInfo.ToString());
                            return Result.Failed;
                        }

                        // Create new ceilings
                        Ceiling ceiling1 = CreateCeilingWithOpenings(doc, profiles.Item1, profiles.Item3,
                            ceilingType, level, offset);
                        Ceiling ceiling2 = CreateCeilingWithOpenings(doc, profiles.Item2, profiles.Item4,
                            ceilingType, level, offset);

                        if (ceiling1 != null && ceiling2 != null)
                        {
                            // Copy parameters from original ceiling
                            CopyParameters(selectedCeiling, ceiling1);
                            CopyParameters(selectedCeiling, ceiling2);

                            // Delete original
                            doc.Delete(selectedCeiling.Id);
                        }


                        if (ceiling1 == null || ceiling2 == null)
                        {
                            trans.RollBack();
                            TaskDialog.Show("Error", "Failed to create split ceilings.\n\n" + debugInfo.ToString());
                            return Result.Failed;
                        }

                        // Reassign hosted families
                        ReassignHostedFamilies(doc, ceiling1, ceiling2, hostedFamilies,
                            familyLocations, splitLine, lineDir);

                        trans.Commit();
                        TaskDialog.Show("Success",
                            $"Ceiling split into 2 parts.\n" +
                            $"{hostedFamilies.Count} hosted families processed.");
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

        #region Profile Creation and Modification

        private Tuple<List<Curve>, List<Curve>, List<List<Curve>>, List<List<Curve>>>
            CreateSplitProfilesImproved(
                List<Curve> outerBoundary,
                List<List<Curve>> innerBoundaries,
                XYZ splitStart,
                XYZ splitEnd,
                XYZ splitDirection,
                StringBuilder debugInfo)
        {
            try
            {
                List<Curve> profile1 = new List<Curve>();
                List<Curve> profile2 = new List<Curve>();
                List<List<Curve>> innerLoops1 = new List<List<Curve>>();
                List<List<Curve>> innerLoops2 = new List<List<Curve>>();

                // Create the split line
                Line splitLine = Line.CreateBound(splitStart, splitEnd);
                XYZ normal = XYZ.BasisZ;
                XYZ perpendicular = normal.CrossProduct(splitDirection).Normalize();

                // Process outer boundary
                ProcessBoundaryForSplit(outerBoundary, splitLine, perpendicular, splitStart,
                    profile1, profile2, debugInfo);

                // Add the split line to both profiles
                profile1.Add(Line.CreateBound(splitStart, splitEnd));
                profile2.Add(Line.CreateBound(splitEnd, splitStart));

                // Sort curves to form continuous loops
                profile1 = SortCurvesIntoLoop(profile1, debugInfo);
                profile2 = SortCurvesIntoLoop(profile2, debugInfo);

                // Process openings with improved logic
                int openingIndex = 0;
                foreach (var innerLoop in innerBoundaries)
                {
                    debugInfo.AppendLine($"\nProcessing opening {openingIndex++}:");

                    var openingResult = ProcessOpeningForSplit(
                        innerLoop, splitLine, perpendicular, splitStart, debugInfo);

                    if (openingResult.Item1 != null && ValidateOpeningLoop(openingResult.Item1))
                        innerLoops1.Add(openingResult.Item1);

                    if (openingResult.Item2 != null && ValidateOpeningLoop(openingResult.Item2))
                        innerLoops2.Add(openingResult.Item2);
                }

                return new Tuple<List<Curve>, List<Curve>, List<List<Curve>>, List<List<Curve>>>(
                    profile1, profile2, innerLoops1, innerLoops2);
            }
            catch (Exception ex)
            {
                debugInfo.AppendLine($"Error in CreateSplitProfilesImproved: {ex.Message}");
                return null;
            }
        }

        private void ProcessBoundaryForSplit(List<Curve> boundary, Line splitLine,
            XYZ perpendicular, XYZ splitStart, List<Curve> profile1, List<Curve> profile2,
            StringBuilder debugInfo)
        {
            foreach (Curve curve in boundary)
            {
                IntersectionResultArray results;
                if (curve.Intersect(splitLine, out results) == SetComparisonResult.Overlap)
                {
                    // Curve intersects split line - split it
                    XYZ intersectionPoint = results.get_Item(0).XYZPoint;

                    List<Curve> segments = SplitCurveAtPoint(curve, intersectionPoint);

                    foreach (var segment in segments)
                    {
                        XYZ midPoint = EvaluateMidpoint(segment);
                        double side = (midPoint - splitStart).DotProduct(perpendicular);

                        if (side > TOLERANCE)
                            profile1.Add(segment);
                        else if (side < -TOLERANCE)
                            profile2.Add(segment);
                        // If very close to split line, add to both
                        else
                        {
                            profile1.Add(segment);
                            profile2.Add(segment);
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
        }

        private Tuple<List<Curve>, List<Curve>> ProcessOpeningForSplit(
            List<Curve> opening, Line splitLine, XYZ perpendicular, XYZ splitStart,
            StringBuilder debugInfo)
        {
            List<Curve> opening1 = null;
            List<Curve> opening2 = null;

            // Find all intersection points
            List<XYZ> intersections = new List<XYZ>();
            Dictionary<XYZ, Curve> intersectionCurves = new Dictionary<XYZ, Curve>();

            foreach (Curve curve in opening)
            {
                IntersectionResultArray results;
                if (curve.Intersect(splitLine, out results) == SetComparisonResult.Overlap)
                {
                    foreach (IntersectionResult result in results)
                    {
                        intersections.Add(result.XYZPoint);
                        intersectionCurves[result.XYZPoint] = curve;
                    }
                }
            }

            if (intersections.Count == 0)
            {
                // Opening doesn't intersect - assign to one side
                XYZ center = GetLoopCenter(opening);
                double side = (center - splitStart).DotProduct(perpendicular);

                if (side > 0)
                    opening1 = new List<Curve>(opening);
                else
                    opening2 = new List<Curve>(opening);
            }
            else if (intersections.Count == 2)
            {
                // Opening crosses split line - split it
                debugInfo.AppendLine($"  Opening crosses at 2 points - splitting");

                opening1 = new List<Curve>();
                opening2 = new List<Curve>();

                // Sort intersections along split line
                intersections = intersections.OrderBy(p => splitLine.Project(p).Parameter).ToList();
                XYZ int1 = intersections[0];
                XYZ int2 = intersections[1];

                // Process each curve in the opening
                foreach (Curve curve in opening)
                {
                    bool hasInt1 = intersectionCurves.ContainsKey(int1) &&
                                  intersectionCurves[int1].Id == curve.Id;
                    bool hasInt2 = intersectionCurves.ContainsKey(int2) &&
                                  intersectionCurves[int2].Id == curve.Id;

                    if (hasInt1 || hasInt2)
                    {
                        // This curve intersects the split line
                        XYZ intPoint = hasInt1 ? int1 : int2;
                        List<Curve> segments = SplitCurveAtPoint(curve, intPoint);

                        foreach (var segment in segments)
                        {
                            XYZ midPoint = EvaluateMidpoint(segment);
                            double side = (midPoint - splitStart).DotProduct(perpendicular);

                            if (side > TOLERANCE)
                                opening1.Add(segment);
                            else if (side < -TOLERANCE)
                                opening2.Add(segment);
                        }
                    }
                    else
                    {
                        // Curve doesn't intersect - assign based on position
                        XYZ midPoint = EvaluateMidpoint(curve);
                        double side = (midPoint - splitStart).DotProduct(perpendicular);

                        if (side > 0)
                            opening1.Add(curve);
                        else
                            opening2.Add(curve);
                    }
                }

                // Add connecting line between intersection points for each side
                if (opening1.Count > 0 && opening2.Count > 0)
                {
                    Line connector = Line.CreateBound(int1, int2);
                    opening1.Add(connector);
                    opening2.Add(connector.CreateReversed() as Line);
                }

                // Sort the curves into proper loops
                opening1 = SortCurvesIntoLoop(opening1, debugInfo);
                opening2 = SortCurvesIntoLoop(opening2, debugInfo);
            }
            else
            {
                // Complex intersection - skip for safety
                debugInfo.AppendLine($"  WARNING: Opening has {intersections.Count} intersections - skipping");
            }

            return new Tuple<List<Curve>, List<Curve>>(opening1, opening2);
        }

        private List<Curve> SplitCurveAtPoint(Curve curve, XYZ point)
        {
            List<Curve> segments = new List<Curve>();

            try
            {
                double param = curve.Project(point).Parameter;
                XYZ projectedPoint = curve.Evaluate(param, false);

                // Check if point is at curve endpoints
                if (projectedPoint.DistanceTo(curve.GetEndPoint(0)) < TOLERANCE ||
                    projectedPoint.DistanceTo(curve.GetEndPoint(1)) < TOLERANCE)
                {
                    segments.Add(curve);
                    return segments;
                }

                if (curve is Line)
                {
                    segments.Add(Line.CreateBound(curve.GetEndPoint(0), projectedPoint));
                    segments.Add(Line.CreateBound(projectedPoint, curve.GetEndPoint(1)));
                }
                else if (curve is Arc arc)
                {
                    try
                    {
                        // Split arc at parameter
                        Curve arc1 = arc.Clone();
                        Curve arc2 = arc.Clone();

                        arc1.MakeBound(arc.GetEndParameter(0), param);
                        arc2.MakeBound(param, arc.GetEndParameter(1));

                        segments.Add(arc1);
                        segments.Add(arc2);
                    }
                    catch
                    {
                        // If arc splitting fails, use line segments
                        segments.Add(Line.CreateBound(curve.GetEndPoint(0), projectedPoint));
                        segments.Add(Line.CreateBound(projectedPoint, curve.GetEndPoint(1)));
                    }
                }
                else
                {
                    // For other curve types, use line approximation
                    segments.Add(Line.CreateBound(curve.GetEndPoint(0), projectedPoint));
                    segments.Add(Line.CreateBound(projectedPoint, curve.GetEndPoint(1)));
                }
            }
            catch
            {
                // If splitting fails, return original curve
                segments.Add(curve);
            }

            return segments;
        }

        private Ceiling CreateCeilingWithOpenings(Document doc, List<Curve> outerBoundary,
            List<List<Curve>> innerBoundaries, CeilingType ceilingType, Level level,
            double offset)
        {
            try
            {
                var allLoops = new List<CurveLoop>();

                // Create outer loop
                var outerLoop = CurveLoop.Create(outerBoundary);
                allLoops.Add(outerLoop);

                // Create inner loops (openings)
                foreach (var innerBoundary in innerBoundaries)
                {
                    if (innerBoundary.Count >= 3)
                    {
                        var innerLoop = CurveLoop.Create(innerBoundary);
                        allLoops.Add(innerLoop);
                    }
                }

                // Create ceiling with all loops at once
                Ceiling newCeiling = Ceiling.Create(doc, allLoops, ceilingType.Id, level.Id);

                if (newCeiling != null)
                {
                    // Set offset
                    newCeiling.get_Parameter(BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM).Set(offset);
                }

                return newCeiling;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to create ceiling: {ex.Message}");
                return null;
            }
        }

        private void CopyParameters(Element source, Element target)
        {
            try
            {
                // Copy instance parameters
                foreach (Parameter param in source.Parameters)
                {
                    if (param.IsReadOnly || param.Id.IntegerValue < 0) continue;

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
            catch { }
        }

        #endregion

        #region Hosted Family Handling

        private List<FamilyInstance> GetHostedFamilies(Document doc, Ceiling ceiling)
        {
            List<FamilyInstance> hostedFamilies = new List<FamilyInstance>();

            // Use dependency filter to find elements hosted on this ceiling
            ElementClassFilter familyFilter = new ElementClassFilter(typeof(FamilyInstance));
            ElementId ceilingId = ceiling.Id;

            FilteredElementCollector collector = new FilteredElementCollector(doc)
                .WherePasses(familyFilter);

            foreach (Element elem in collector)
            {
                FamilyInstance fi = elem as FamilyInstance;
                if (fi != null && fi.Host != null && fi.Host.Id == ceilingId)
                {
                    hostedFamilies.Add(fi);
                }
            }

            return hostedFamilies;
        }

        private void ReassignHostedFamilies(Document doc, Ceiling ceiling1, Ceiling ceiling2,
            List<FamilyInstance> originalFamilies, Dictionary<FamilyInstance, XYZ> familyLocations,
            Line splitLine, XYZ splitDirection)
        {
            XYZ normal = XYZ.BasisZ;
            XYZ perpendicular = normal.CrossProduct(splitDirection).Normalize();
            XYZ splitStart = splitLine.GetEndPoint(0);

            List<ElementId> familiesToDelete = new List<ElementId>();

            foreach (var family in originalFamilies)
            {
                if (!familyLocations.ContainsKey(family)) continue;

                XYZ location = familyLocations[family];

                // Determine which side of split line
                double side = (location - splitStart).DotProduct(perpendicular);

                // Check if family crosses the split line (based on bounding box)
                BoundingBoxXYZ bbox = family.get_BoundingBox(null);
                bool crossesSplit = false;

                if (bbox != null)
                {
                    XYZ min = bbox.Min;
                    XYZ max = bbox.Max;

                    // Check if bounding box straddles the split line
                    double minSide = (min - splitStart).DotProduct(perpendicular);
                    double maxSide = (max - splitStart).DotProduct(perpendicular);

                    crossesSplit = (minSide * maxSide) < 0;
                }

                try
                {
                    if (crossesSplit)
                    {
                        // Family crosses split line - duplicate it for both ceilings

                        // Place on ceiling1
                        FamilyInstance copy1 = PlaceFamilyOnCeiling(doc, family, ceiling1, location);

                        // Place on ceiling2
                        FamilyInstance copy2 = PlaceFamilyOnCeiling(doc, family, ceiling2, location);

                        // Mark original for deletion
                        familiesToDelete.Add(family.Id);
                    }
                    else
                    {
                        // Family doesn't cross - assign to appropriate ceiling
                        Ceiling targetCeiling = side > 0 ? ceiling1 : ceiling2;

                        // Try to reassign host
                        Parameter hostParam = family.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_PARAM);
                        if (hostParam != null && !hostParam.IsReadOnly)
                        {
                            hostParam.Set(targetCeiling.Id);
                        }
                        else
                        {
                            // If can't reassign, recreate
                            FamilyInstance newInstance = PlaceFamilyOnCeiling(doc, family, targetCeiling, location);
                            if (newInstance != null)
                                familiesToDelete.Add(family.Id);
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Log error but continue processing other families
                    Debug.WriteLine($"Failed to process family {family.Id}: {ex.Message}");
                }
            }

            // Delete families that were recreated
            if (familiesToDelete.Count > 0)
            {
                doc.Delete(familiesToDelete);
            }
        }

        private FamilyInstance PlaceFamilyOnCeiling(Document doc, FamilyInstance originalFamily,
            Ceiling targetCeiling, XYZ location)
        {
            try
            {
                FamilySymbol symbol = originalFamily.Symbol;
                Level level = doc.GetElement(targetCeiling.LevelId) as Level;

                // Get the face of the ceiling at the location
                Reference faceRef = GetCeilingBottomFaceReference(targetCeiling, location);
                if (faceRef == null) return null;

                // For face-based families, we need to use NewFamilyInstance with Reference
                FamilyInstance newInstance = null;

                // Get original family placement info
                LocationPoint origLoc = originalFamily.Location as LocationPoint;
                double rotation = 0;
                if (origLoc != null)
                {
                    rotation = origLoc.Rotation;
                }

                //Try as face-based family
                newInstance = doc.Create.NewFamilyInstance(
                        faceRef,
                        location,
                        originalFamily.HandOrientation,
                        symbol);

                if (newInstance != null)
                {
                    // Copy parameters
                    CopyParameters(originalFamily, newInstance);

                    // Apply rotation if needed
                    if (Math.Abs(rotation) > 0.001)
                    {
                        LocationPoint newLoc = newInstance.Location as LocationPoint;
                        if (newLoc != null)
                        {
                            Line axis = Line.CreateBound(newLoc.Point,
                                newLoc.Point + XYZ.BasisZ);
                            ElementTransformUtils.RotateElement(doc, newInstance.Id,
                                axis, rotation);
                        }
                    }
                }

                return newInstance;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to place family on ceiling: {ex.Message}");
                return null;
            }
        }

        private Reference GetCeilingBottomFaceReference(Ceiling ceiling, XYZ location)
        {
            Options options = new Options();
            options.ComputeReferences = true;
            options.DetailLevel = ViewDetailLevel.Fine;

            GeometryElement geomElem = ceiling.get_Geometry(options);

            foreach (GeometryObject geomObj in geomElem)
            {
                if (geomObj is Solid solid)
                {
                    foreach (Face face in solid.Faces)
                    {
                        if (face is PlanarFace planarFace)
                        {
                            // Check if face normal points down (bottom face)
                            if (planarFace.FaceNormal.DotProduct(XYZ.BasisZ) < -0.9)
                            {
                                // Check if location projects onto this face
                                IntersectionResult result = face.Project(location);
                                if (result != null && result.Distance < 10.0) // Within 10 feet
                                {
                                    return face.Reference;
                                }
                            }
                        }
                    }
                }
            }
            return null;
        }

        #endregion

        #region Utility Methods

        private bool ValidateOpeningLoop(List<Curve> curves)
        {
            if (curves.Count < 3) return false;

            // Check if curves form a closed loop
            for (int i = 0; i < curves.Count; i++)
            {
                XYZ currentEnd = curves[i].GetEndPoint(1);
                XYZ nextStart = curves[(i + 1) % curves.Count].GetEndPoint(0);

                if (currentEnd.DistanceTo(nextStart) > TOLERANCE)
                {
                    return false;
                }
            }

            return true;
        }

        private List<Curve> SortCurvesIntoLoop(List<Curve> curves, StringBuilder debugInfo)
        {
            if (curves.Count < 3)
            {
                debugInfo?.AppendLine($"  Too few curves to form loop: {curves.Count}");
                return curves;
            }

            List<Curve> sorted = new List<Curve>();
            List<Curve> remaining = new List<Curve>(curves);

            // Start with first curve
            sorted.Add(remaining[0]);
            remaining.RemoveAt(0);

            int iterations = 0;
            int maxIterations = curves.Count * 2;

            while (remaining.Count > 0 && iterations < maxIterations)
            {
                XYZ currentEnd = sorted.Last().GetEndPoint(1);
                bool found = false;

                for (int i = 0; i < remaining.Count; i++)
                {
                    Curve curve = remaining[i];
                    XYZ curveStart = curve.GetEndPoint(0);
                    XYZ curveEnd = curve.GetEndPoint(1);

                    if (currentEnd.DistanceTo(curveStart) < TOLERANCE)
                    {
                        sorted.Add(curve);
                        remaining.RemoveAt(i);
                        found = true;
                        break;
                    }
                    else if (currentEnd.DistanceTo(curveEnd) < TOLERANCE)
                    {
                        // Reverse the curve
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
                    if (currentEnd.DistanceTo(loopStart) < TOLERANCE)
                    {
                        debugInfo?.AppendLine($"  Loop closed with {sorted.Count} curves");
                        break;
                    }
                    else
                    {
                        // Try to find nearest curve
                        double minDist = double.MaxValue;
                        int nearestIndex = -1;
                        bool reverseNearest = false;

                        for (int i = 0; i < remaining.Count; i++)
                        {
                            double distToStart = currentEnd.DistanceTo(remaining[i].GetEndPoint(0));
                            double distToEnd = currentEnd.DistanceTo(remaining[i].GetEndPoint(1));

                            if (distToStart < minDist)
                            {
                                minDist = distToStart;
                                nearestIndex = i;
                                reverseNearest = false;
                            }
                            if (distToEnd < minDist)
                            {
                                minDist = distToEnd;
                                nearestIndex = i;
                                reverseNearest = true;
                            }
                        }

                        if (nearestIndex >= 0 && minDist < OPENING_MERGE_TOLERANCE)
                        {
                            // Add bridge if needed
                            if (minDist > TOLERANCE)
                            {
                                XYZ bridgeEnd = reverseNearest ?
                                    remaining[nearestIndex].GetEndPoint(1) :
                                    remaining[nearestIndex].GetEndPoint(0);
                                sorted.Add(Line.CreateBound(currentEnd, bridgeEnd));
                                debugInfo?.AppendLine($"  Added bridge: {minDist:F4} feet");
                            }

                            // Add the nearest curve
                            if (reverseNearest)
                                sorted.Add(remaining[nearestIndex].CreateReversed());
                            else
                                sorted.Add(remaining[nearestIndex]);

                            remaining.RemoveAt(nearestIndex);
                            found = true;
                        }
                    }

                    if (!found)
                    {
                        debugInfo?.AppendLine($"  Gap in loop: {remaining.Count} curves unconnected");
                        break;
                    }
                }

                iterations++;
            }

            // Close loop if needed
            if (sorted.Count > 2)
            {
                XYZ loopEnd = sorted.Last().GetEndPoint(1);
                XYZ loopStart = sorted.First().GetEndPoint(0);
                double gap = loopEnd.DistanceTo(loopStart);

                if (gap > TOLERANCE && gap < OPENING_MERGE_TOLERANCE)
                {
                    sorted.Add(Line.CreateBound(loopEnd, loopStart));
                    debugInfo?.AppendLine($"  Closed loop with bridge: {gap:F4} feet");
                }
            }

            return sorted;
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
                // Fallback to average of endpoints
                return (curve.GetEndPoint(0) + curve.GetEndPoint(1)) / 2.0;
            }
        }

        private XYZ GetLoopCenter(List<Curve> loop)
        {
            double sumX = 0, sumY = 0, sumZ = 0;
            int count = 0;

            foreach (Curve curve in loop)
            {
                XYZ midPoint = EvaluateMidpoint(curve);
                sumX += midPoint.X;
                sumY += midPoint.Y;
                sumZ += midPoint.Z;
                count++;
            }

            if (count > 0)
                return new XYZ(sumX / count, sumY / count, sumZ / count);
            else
                return XYZ.Zero;
        }

        private Tuple<List<Curve>, List<List<Curve>>> GetCeilingBoundaryLoops(Ceiling ceiling)
        {
            List<Curve> outerBoundary = new List<Curve>();
            List<List<Curve>> innerBoundaries = new List<List<Curve>>();

            Options options = new Options();
            options.ComputeReferences = true;
            options.IncludeNonVisibleObjects = true;
            options.DetailLevel = ViewDetailLevel.Fine;

            GeometryElement geomElem = ceiling.get_Geometry(options);

            foreach (GeometryObject geomObj in geomElem)
            {
                if (geomObj is Solid solid && solid.Volume > 0)
                {
                    // Find the bottom face
                    foreach (Face face in solid.Faces)
                    {
                        if (face is PlanarFace planarFace &&
                            planarFace.FaceNormal.DotProduct(XYZ.BasisZ) < -0.9) // Bottom face
                        {
                            // Find the largest loop (outer boundary)
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

                            // Extract curves from outer loop
                            if (largestLoop != null)
                            {
                                foreach (Edge edge in largestLoop)
                                {
                                    outerBoundary.Add(edge.AsCurve());
                                }
                            }

                            // Extract curves from inner loops
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

        internal static PushButtonData GetButtonData()
        {
            string buttonInternalName = "btnSplitCeiling";
            string buttonTitle = "Split Ceiling";

            ButtonDataClass myButtonData = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Green_32,
                Properties.Resources.Green_16,
                "Splits a ceiling along a model line while preserving openings and hosted families");

            return myButtonData.Data;
        }

        #endregion
    }

    public class CeilingSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            return elem is Ceiling;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }
}