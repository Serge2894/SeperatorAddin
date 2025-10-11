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
    public class cmdRoofSplitter : IExternalCommand
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
                // Select a roof
                Reference roofRef = uiDoc.Selection.PickObject(ObjectType.Element, new RoofSelectionFilter(), "Select a roof to split");
                RoofBase selectedRoof = doc.GetElement(roofRef) as RoofBase;

                // Select ONE model line
                Reference lineRef = uiDoc.Selection.PickObject(ObjectType.Element, new Utils.ModelCurveSelectionFilter(), "Select ONE straight model line that crosses the roof");
                ModelCurve modelLine = doc.GetElement(lineRef) as ModelCurve;

                using (Transaction trans = new Transaction(doc, "Split Roof with Openings"))
                {
                    trans.Start();
                    bool success = SplitRoof(doc, selectedRoof, modelLine);
                    if (success)
                    {
                        trans.Commit();
                        TaskDialog.Show("Success", "Roof split successfully.");
                        return Result.Succeeded;
                    }
                    else
                    {
                        trans.RollBack();
                        message = "Failed to split the roof. Ensure the model line fully intersects the roof boundary.";
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

        public bool SplitRoof(Document doc, RoofBase selectedRoof, ModelCurve modelLine)
        {
            if (selectedRoof == null || modelLine == null) return false;

            Line splitLine = modelLine.GeometryCurve as Line;
            if (splitLine == null) return false;

            try
            {
                // Get roof properties
                RoofType roofType = selectedRoof.RoofType;
                Level level = doc.GetElement(selectedRoof.LevelId) as Level;
                double offset = selectedRoof.get_Parameter(BuiltInParameter.ROOF_BASE_LEVEL_PARAM)?.AsDouble() ?? 0;
                bool isStructural = selectedRoof.get_Parameter(BuiltInParameter.FLOOR_PARAM_IS_STRUCTURAL)?.AsInteger() == 1;

                // Collect hosted families
                List<FamilyInstance> hostedFamilies = GetHostedFamilies(doc, selectedRoof);
                Dictionary<FamilyInstance, XYZ> familyLocations = new Dictionary<FamilyInstance, XYZ>();
                foreach (var family in hostedFamilies)
                {
                    if (family.Location is LocationPoint locPoint)
                        familyLocations[family] = locPoint.Point;
                }

                // Get roof boundary
                var roofLoops = GetRoofBoundaryLoops(selectedRoof);
                if (roofLoops.Item1.Count < 3) return false;

                // Extend split line
                XYZ lineDir = splitLine.Direction;
                Line extendedLine = Line.CreateBound(splitLine.GetEndPoint(0) - lineDir * 1000, splitLine.GetEndPoint(1) + lineDir * 1000);

                List<XYZ> intersectionPoints = new List<XYZ>();
                foreach (Curve boundaryCurve in roofLoops.Item1)
                {
                    if (boundaryCurve.Intersect(extendedLine, out IntersectionResultArray results) == SetComparisonResult.Overlap)
                    {
                        foreach (IntersectionResult result in results) intersectionPoints.Add(result.XYZPoint);
                    }
                }

                if (intersectionPoints.Count < 2) return false;

                intersectionPoints = intersectionPoints.OrderBy(p => extendedLine.GetEndPoint(0).DistanceTo(p)).ToList();
                XYZ splitStart = intersectionPoints.First();
                XYZ splitEnd = intersectionPoints.Last();

                var profiles = CreateSplitProfilesImproved(roofLoops.Item1, roofLoops.Item2, splitStart, splitEnd, lineDir, new StringBuilder());
                if (profiles == null || !profiles.Item1.Any() || !profiles.Item2.Any()) return false;

                RoofBase roof1 = CreateRoofWithOpenings(doc, profiles.Item1, profiles.Item3, roofType, level, offset, isStructural);
                RoofBase roof2 = CreateRoofWithOpenings(doc, profiles.Item2, profiles.Item4, roofType, level, offset, isStructural);

                if (roof1 != null && roof2 != null)
                {
                    CopyRoofParameters(selectedRoof, roof1);
                    CopyRoofParameters(selectedRoof, roof2);
                    ReassignHostedFamilies(doc, roof1, roof2, hostedFamilies, familyLocations, splitLine, lineDir);
                    doc.Delete(selectedRoof.Id);
                    return true;
                }

                return false;
            }
            catch
            {
                return false;
            }
        }


        #region Improved Profile Creation

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

                Line splitLine = Line.CreateBound(splitStart, splitEnd);
                XYZ perpendicular = XYZ.BasisZ.CrossProduct(splitDirection).Normalize();

                ProcessBoundaryForSplit(outerBoundary, splitLine, perpendicular, splitStart, profile1, profile2, debugInfo);

                profile1.Add(Line.CreateBound(splitStart, splitEnd));
                profile2.Add(Line.CreateBound(splitEnd, splitStart));

                profile1 = SortCurvesIntoLoop(profile1, debugInfo);
                profile2 = SortCurvesIntoLoop(profile2, debugInfo);

                foreach (var innerLoop in innerBoundaries)
                {
                    var openingResult = ProcessOpeningForSplit(innerLoop, splitLine, perpendicular, splitStart, debugInfo);
                    if (openingResult.Item1 != null && ValidateOpeningLoop(openingResult.Item1)) innerLoops1.Add(openingResult.Item1);
                    if (openingResult.Item2 != null && ValidateOpeningLoop(openingResult.Item2)) innerLoops2.Add(openingResult.Item2);
                }

                return new Tuple<List<Curve>, List<Curve>, List<List<Curve>>, List<List<Curve>>>(profile1, profile2, innerLoops1, innerLoops2);
            }
            catch (Exception ex)
            {
                debugInfo.AppendLine($"Error in CreateSplitProfilesImproved: {ex.Message}");
                return null;
            }
        }

        private void ProcessBoundaryForSplit(List<Curve> boundary, Line splitLine, XYZ perpendicular, XYZ splitStart, List<Curve> profile1, List<Curve> profile2, StringBuilder debugInfo)
        {
            foreach (Curve curve in boundary)
            {
                if (curve.Intersect(splitLine, out IntersectionResultArray results) == SetComparisonResult.Overlap)
                {
                    XYZ intersectionPoint = results.get_Item(0).XYZPoint;
                    List<Curve> segments = SplitCurveAtPoint(curve, intersectionPoint);
                    foreach (var segment in segments)
                    {
                        double side = (EvaluateMidpoint(segment) - splitStart).DotProduct(perpendicular);
                        if (side > TOLERANCE) profile1.Add(segment);
                        else if (side < -TOLERANCE) profile2.Add(segment);
                        else { profile1.Add(segment); profile2.Add(segment); }
                    }
                }
                else
                {
                    double side = (EvaluateMidpoint(curve) - splitStart).DotProduct(perpendicular);
                    if (side > 0) profile1.Add(curve);
                    else profile2.Add(curve);
                }
            }
        }

        private Tuple<List<Curve>, List<Curve>> ProcessOpeningForSplit(List<Curve> opening, Line splitLine, XYZ perpendicular, XYZ splitStart, StringBuilder debugInfo)
        {
            List<Curve> opening1 = null;
            List<Curve> opening2 = null;
            List<XYZ> intersections = new List<XYZ>();
            Dictionary<XYZ, Curve> intersectionCurves = new Dictionary<XYZ, Curve>();

            foreach (Curve curve in opening)
            {
                if (curve.Intersect(splitLine, out IntersectionResultArray results) == SetComparisonResult.Overlap)
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
                double side = (GetLoopCenter(opening) - splitStart).DotProduct(perpendicular);
                if (side > 0) opening1 = new List<Curve>(opening);
                else opening2 = new List<Curve>(opening);
            }
            else if (intersections.Count == 2)
            {
                opening1 = new List<Curve>();
                opening2 = new List<Curve>();
                intersections = intersections.OrderBy(p => splitLine.Project(p).Parameter).ToList();
                XYZ int1 = intersections[0];
                XYZ int2 = intersections[1];

                foreach (Curve curve in opening)
                {
                    bool hasInt1 = intersectionCurves.ContainsKey(int1) && intersectionCurves[int1] == curve;
                    bool hasInt2 = intersectionCurves.ContainsKey(int2) && intersectionCurves[int2] == curve;

                    if (hasInt1 || hasInt2)
                    {
                        XYZ intPoint = hasInt1 ? int1 : int2;
                        foreach (var segment in SplitCurveAtPoint(curve, intPoint))
                        {
                            double side = (EvaluateMidpoint(segment) - splitStart).DotProduct(perpendicular);
                            if (side > TOLERANCE) opening1.Add(segment);
                            else if (side < -TOLERANCE) opening2.Add(segment);
                        }
                    }
                    else
                    {
                        double side = (EvaluateMidpoint(curve) - splitStart).DotProduct(perpendicular);
                        if (side > 0) opening1.Add(curve);
                        else opening2.Add(curve);
                    }
                }

                if (opening1.Any() && opening2.Any())
                {
                    Line connector = Line.CreateBound(int1, int2);
                    opening1.Add(connector);
                    opening2.Add(connector.CreateReversed() as Line);
                }

                opening1 = SortCurvesIntoLoop(opening1, debugInfo);
                opening2 = SortCurvesIntoLoop(opening2, debugInfo);
            }
            return new Tuple<List<Curve>, List<Curve>>(opening1, opening2);
        }

        private List<Curve> SplitCurveAtPoint(Curve curve, XYZ point)
        {
            List<Curve> segments = new List<Curve>();
            try
            {
                double param = curve.Project(point).Parameter;
                if (point.DistanceTo(curve.GetEndPoint(0)) < TOLERANCE || point.DistanceTo(curve.GetEndPoint(1)) < TOLERANCE)
                {
                    segments.Add(curve);
                    return segments;
                }

                segments.Add(Line.CreateBound(curve.GetEndPoint(0), point));
                segments.Add(Line.CreateBound(point, curve.GetEndPoint(1)));
            }
            catch { segments.Add(curve); }
            return segments;
        }

        #endregion

        #region Roof Modification

        private RoofBase CreateRoofWithOpenings(Document doc, List<Curve> outerBoundary, List<List<Curve>> innerBoundaries, RoofType roofType, Level level, double offset, bool isStructural)
        {
            try
            {
                CurveArray outerCurves = new CurveArray();
                foreach (Curve c in outerBoundary) outerCurves.Append(c);

                FootPrintRoof newRoof = doc.Create.NewFootPrintRoof(outerCurves, level, roofType, out ModelCurveArray modelCurves);

                if (newRoof != null)
                {
                    newRoof.get_Parameter(BuiltInParameter.ROOF_CONSTRAINT_OFFSET_PARAM)?.Set(offset);
                    newRoof.get_Parameter(BuiltInParameter.FLOOR_PARAM_IS_STRUCTURAL)?.Set(isStructural ? 1 : 0);

                    foreach (var innerBoundary in innerBoundaries)
                    {
                        if (innerBoundary.Count >= 3)
                        {
                            try
                            {
                                CurveArray openingCurves = new CurveArray();
                                foreach (Curve c in innerBoundary) openingCurves.Append(c);
                                doc.Create.NewOpening(newRoof, openingCurves, true);
                            }
                            catch { continue; }
                        }
                    }
                }
                return newRoof;
            }
            catch
            {
                return null;
            }
        }

        private void CopyRoofParameters(RoofBase sourceRoof, RoofBase targetRoof)
        {
            foreach (Parameter param in sourceRoof.Parameters)
            {
                if (param.IsReadOnly || param.Id.IntegerValue < 0) continue;
                Parameter targetParam = targetRoof.get_Parameter(param.Definition);
                if (targetParam != null && !targetParam.IsReadOnly)
                {
                    try
                    {
                        if (param.StorageType == StorageType.Double) targetParam.Set(param.AsDouble());
                        else if (param.StorageType == StorageType.Integer) targetParam.Set(param.AsInteger());
                        else if (param.StorageType == StorageType.String) targetParam.Set(param.AsString());
                        else if (param.StorageType == StorageType.ElementId) targetParam.Set(param.AsElementId());
                    }
                    catch { }
                }
            }
        }

        #endregion

        #region Hosted Family Handling

        private List<FamilyInstance> GetHostedFamilies(Document doc, RoofBase roof)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => fi.Host != null && fi.Host.Id == roof.Id)
                .ToList();
        }

        private void ReassignHostedFamilies(Document doc, RoofBase roof1, RoofBase roof2, List<FamilyInstance> originalFamilies, Dictionary<FamilyInstance, XYZ> familyLocations, Line splitLine, XYZ splitDirection)
        {
            XYZ perpendicular = XYZ.BasisZ.CrossProduct(splitDirection).Normalize();
            XYZ splitStart = splitLine.GetEndPoint(0);
            List<ElementId> familiesToDelete = new List<ElementId>();

            foreach (var family in originalFamilies)
            {
                if (!familyLocations.ContainsKey(family) || familyLocations[family] == null) continue;
                XYZ location = familyLocations[family];
                double side = (location - splitStart).DotProduct(perpendicular);

                RoofBase targetRoof = side > 0 ? roof1 : roof2;
                FamilyInstance newInstance = PlaceFamilyOnRoof(doc, family, targetRoof, location);
                if (newInstance != null) familiesToDelete.Add(family.Id);
            }

            if (familiesToDelete.Any()) doc.Delete(familiesToDelete);
        }

        private FamilyInstance PlaceFamilyOnRoof(Document doc, FamilyInstance originalFamily, RoofBase targetRoof, XYZ location)
        {
            try
            {
                FamilySymbol symbol = originalFamily.Symbol;
                Reference faceRef = GetRoofTopFaceReference(targetRoof, location);
                if (faceRef == null) return null;

                FamilyInstance newInstance = doc.Create.NewFamilyInstance(faceRef, location, originalFamily.HandOrientation, symbol);

                if (newInstance != null)
                {
                    CopyFamilyParameters(originalFamily, newInstance);
                }
                return newInstance;
            }
            catch
            {
                return null;
            }
        }

        private Reference GetRoofTopFaceReference(RoofBase roof, XYZ location)
        {
            Options options = new Options { ComputeReferences = true, DetailLevel = ViewDetailLevel.Fine };
            GeometryElement geomElem = roof.get_Geometry(options);
            foreach (GeometryObject geomObj in geomElem)
            {
                if (geomObj is Solid solid)
                {
                    foreach (Face face in solid.Faces)
                    {
                        if (face is PlanarFace planarFace && planarFace.FaceNormal.DotProduct(XYZ.BasisZ) > 0.9)
                        {
                            if (face.Project(location) != null) return face.Reference;
                        }
                    }
                }
            }
            return null;
        }

        private void CopyFamilyParameters(FamilyInstance source, FamilyInstance target)
        {
            foreach (Parameter param in source.Parameters)
            {
                if (param.IsReadOnly || !param.HasValue) continue;
                Parameter targetParam = target.get_Parameter(param.Definition);
                if (targetParam != null && !targetParam.IsReadOnly)
                {
                    try
                    {
                        if (param.StorageType == StorageType.Double) targetParam.Set(param.AsDouble());
                        else if (param.StorageType == StorageType.Integer) targetParam.Set(param.AsInteger());
                        else if (param.StorageType == StorageType.String) targetParam.Set(param.AsString());
                        else if (param.StorageType == StorageType.ElementId) targetParam.Set(param.AsElementId());
                    }
                    catch { }
                }
            }
        }

        #endregion

        #region Utility Methods

        private bool ValidateOpeningLoop(List<Curve> curves)
        {
            if (curves.Count < 3) return false;
            for (int i = 0; i < curves.Count; i++)
            {
                if (curves[i].GetEndPoint(1).DistanceTo(curves[(i + 1) % curves.Count].GetEndPoint(0)) > TOLERANCE) return false;
            }
            return true;
        }

        private List<Curve> SortCurvesIntoLoop(List<Curve> curves, StringBuilder debugInfo)
        {
            if (curves.Count < 2) return curves;
            List<Curve> sorted = new List<Curve> { curves[0] };
            curves.RemoveAt(0);
            while (curves.Any())
            {
                XYZ currentEnd = sorted.Last().GetEndPoint(1);
                bool found = false;
                for (int i = 0; i < curves.Count; i++)
                {
                    if (currentEnd.DistanceTo(curves[i].GetEndPoint(0)) < TOLERANCE)
                    {
                        sorted.Add(curves[i]);
                        curves.RemoveAt(i);
                        found = true;
                        break;
                    }
                    if (currentEnd.DistanceTo(curves[i].GetEndPoint(1)) < TOLERANCE)
                    {
                        sorted.Add(curves[i].CreateReversed());
                        curves.RemoveAt(i);
                        found = true;
                        break;
                    }
                }
                if (!found) break; // Gap in loop
            }
            return sorted;
        }

        private XYZ EvaluateMidpoint(Curve curve)
        {
            try { return curve.Evaluate(0.5, true); }
            catch { return (curve.GetEndPoint(0) + curve.GetEndPoint(1)) / 2.0; }
        }

        private XYZ GetLoopCenter(List<Curve> loop)
        {
            XYZ center = XYZ.Zero;
            foreach (Curve curve in loop) center += EvaluateMidpoint(curve);
            return center / loop.Count;
        }

        private Tuple<List<Curve>, List<List<Curve>>> GetRoofBoundaryLoops(RoofBase roof)
        {
            List<Curve> outerBoundary = new List<Curve>();
            List<List<Curve>> innerBoundaries = new List<List<Curve>>();
            Options options = new Options { ComputeReferences = true, DetailLevel = ViewDetailLevel.Fine };
            GeometryElement geomElem = roof.get_Geometry(options);

            foreach (GeometryObject geomObj in geomElem)
            {
                if (geomObj is Solid solid && solid.Volume > 0)
                {
                    foreach (Face face in solid.Faces)
                    {
                        if (face is PlanarFace planarFace && planarFace.FaceNormal.DotProduct(XYZ.BasisZ) > 0.9) // Top face
                        {
                            EdgeArray largestLoop = null;
                            double maxArea = 0;
                            List<EdgeArray> innerLoops = new List<EdgeArray>();

                            foreach (EdgeArray edgeArray in face.EdgeLoops)
                            {
                                double area = GetLoopArea(edgeArray);
                                if (area > maxArea)
                                {
                                    if (largestLoop != null) innerLoops.Add(largestLoop);
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
                                foreach (Edge edge in largestLoop) outerBoundary.Add(edge.AsCurve());
                            }
                            foreach (EdgeArray innerLoop in innerLoops)
                            {
                                List<Curve> innerCurves = new List<Curve>();
                                foreach (Edge edge in innerLoop) innerCurves.Add(edge.AsCurve());
                                if (innerCurves.Any()) innerBoundaries.Add(innerCurves);
                            }
                            return new Tuple<List<Curve>, List<List<Curve>>>(outerBoundary, innerBoundaries);
                        }
                    }
                }
            }
            return new Tuple<List<Curve>, List<List<Curve>>>(outerBoundary, innerBoundaries);
        }

        private double GetLoopArea(EdgeArray edges)
        {
            List<XYZ> points = new List<XYZ>();
            foreach (Edge edge in edges) points.Add(edge.AsCurve().GetEndPoint(0));
            double area = 0;
            for (int i = 0; i < points.Count; i++)
            {
                int j = (i + 1) % points.Count;
                area += points[i].X * points[j].Y;
                area -= points[j].X * points[i].Y;
            }
            return Math.Abs(area) / 2.0;
        }

        #endregion
    }

    public class RoofSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem) => elem is RoofBase;
        public bool AllowReference(Reference reference, XYZ position) => false;
    }
}