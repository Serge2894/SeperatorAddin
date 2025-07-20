using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using SeperatorAddin.Common;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using static SeperatorAddin.Common.Utils;

namespace SeperatorAddin
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class cmdCeilingSplitter : IExternalCommand
    {
        // Track failed hosted elements
        private List<ElementId> failedHostedElements = new List<ElementId>();

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            try
            {
                // Clear failed elements list
                failedHostedElements.Clear();

                // Ensure we're in a reflected ceiling plan
                View activeView = doc.ActiveView;
                if (!(activeView is ViewPlan viewPlan) || viewPlan.ViewType != ViewType.CeilingPlan)
                {
                    TaskDialog.Show("Error", "Please run this command in a reflected ceiling plan view.");
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
                    new ModelCurveSelectionFilter(),
                    "Select ONE straight model line that crosses the ceiling");

                ModelCurve modelLine = doc.GetElement(lineRef) as ModelCurve;
                Line splitLine = modelLine.GeometryCurve as Line;

                if (splitLine == null)
                {
                    TaskDialog.Show("Error", "Please select a straight line.");
                    return Result.Failed;
                }

                using (Transaction trans = new Transaction(doc, "Split Ceiling"))
                {
                    trans.Start();

                    try
                    {
                        // Get ceiling properties
                        CeilingType ceilingType = doc.GetElement(selectedCeiling.GetTypeId()) as CeilingType;
                        Level level = doc.GetElement(selectedCeiling.LevelId) as Level;
                        double offset = selectedCeiling.get_Parameter(
                            BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM).AsDouble();

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

                        // DEBUG: Log opening information
                        StringBuilder debugInfo = new StringBuilder();
                        debugInfo.AppendLine($"Found {ceilingLoops.Item2.Count} openings in ceiling");

                        // Create profile loops for two ceilings
                        var profiles = CreateSplitProfilesWithOpeningsAlternative(
                            ceilingLoops.Item1, ceilingLoops.Item2, splitStart, splitEnd, lineDir, debugInfo);

                        if (profiles.Item1.Count == 0 || profiles.Item2.Count == 0)
                        {
                            trans.RollBack();
                            TaskDialog.Show("Error", "Failed to create valid split profiles.\n\n" + debugInfo.ToString());
                            return Result.Failed;
                        }

                        // Get hosted elements before deleting original ceiling
                        List<FamilyInstance> hostedFamilies = GetHostedFamilies(doc, selectedCeiling);
                        debugInfo.AppendLine($"\nFound {hostedFamilies.Count} hosted families");

                        // Create new ceilings
                        int ceilingsCreated = 0;
                        string creationErrors = "";
                        Ceiling newCeiling1 = null;
                        Ceiling newCeiling2 = null;

                        // First ceiling
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

                                // Add inner loops (openings)
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

                                newCeiling1 = Ceiling.Create(doc, loops1, ceilingType.Id, level.Id);

                                if (newCeiling1 != null)
                                {
                                    newCeiling1.get_Parameter(
                                        BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM).Set(offset);
                                    ceilingsCreated++;
                                }
                            }
                            else
                            {
                                creationErrors += $"Ceiling 1: {error1}\n";
                            }
                        }
                        catch (Exception ex)
                        {
                            creationErrors += $"Ceiling 1 Exception: {ex.Message}\n";
                        }

                        // Second ceiling
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

                                // Add inner loops (openings)
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

                                newCeiling2 = Ceiling.Create(doc, loops2, ceilingType.Id, level.Id);

                                if (newCeiling2 != null)
                                {
                                    newCeiling2.get_Parameter(
                                        BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM).Set(offset);
                                    ceilingsCreated++;
                                }
                            }
                            else
                            {
                                creationErrors += $"Ceiling 2: {error2}\n";
                            }
                        }
                        catch (Exception ex)
                        {
                            creationErrors += $"Ceiling 2 Exception: {ex.Message}\n";
                        }

                        if (ceilingsCreated == 2)
                        {
                            // Copy parameters from original ceiling to new ceilings
                            CopyCeilingParameters(selectedCeiling, newCeiling1);
                            CopyCeilingParameters(selectedCeiling, newCeiling2);

                            // Delete original ceiling
                            doc.Delete(selectedCeiling.Id);

                            // Re-host families on new ceilings
                            RehostFamiliesOnNewCeilings(doc, hostedFamilies, newCeiling1, newCeiling2,
                                splitStart, splitEnd, lineDir, debugInfo);

                            trans.Commit();

                            // Show success message with any warnings
                            string successMessage = "Ceiling split into 2 parts.";
                            if (failedHostedElements.Count > 0)
                            {
                                successMessage += $"\n\nWarning: {failedHostedElements.Count} hosted elements could not be transferred to the new ceilings.";
                            }
                            if (hostedFamilies.Count > 0)
                            {
                                successMessage += $"\n{hostedFamilies.Count - failedHostedElements.Count} hosted families processed.";
                            }

                            TaskDialog.Show("Success", successMessage);
                            return Result.Succeeded;
                        }
                        else
                        {
                            trans.RollBack();
                            TaskDialog.Show("Error",
                                $"Failed to create both ceiling parts.\n\nErrors:\n{creationErrors}\n\nDebug Info:\n{debugInfo}");
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

        // Get all families hosted to the ceiling
        private List<FamilyInstance> GetHostedFamilies(Document doc, Ceiling ceiling)
        {
            List<FamilyInstance> hostedFamilies = new List<FamilyInstance>();

            try
            {
                // Get dependent elements
                ICollection<ElementId> dependentIds = ceiling.GetDependentElements(null);

                foreach (ElementId id in dependentIds)
                {
                    Element elem = doc.GetElement(id);
                    if (elem is FamilyInstance fi && fi.Host?.Id == ceiling.Id)
                    {
                        hostedFamilies.Add(fi);
                    }
                }

                // Alternative method: use filtered collector
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                var allFamilyInstances = collector.OfClass(typeof(FamilyInstance)).ToElements();

                foreach (Element elem in allFamilyInstances)
                {
                    if (elem is FamilyInstance fi && fi.Host?.Id == ceiling.Id)
                    {
                        if (!hostedFamilies.Any(f => f.Id == fi.Id))
                        {
                            hostedFamilies.Add(fi);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error getting hosted families: {ex.Message}");
            }

            return hostedFamilies;
        }

        // Copy parameters from source ceiling to target ceiling
        private void CopyCeilingParameters(Ceiling source, Ceiling target)
        {
            try
            {
                foreach (Parameter sourceParam in source.Parameters)
                {
                    if (!sourceParam.IsReadOnly && sourceParam.HasValue)
                    {
                        Parameter targetParam = target.LookupParameter(sourceParam.Definition.Name);
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
                            catch { }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error copying ceiling parameters: {ex.Message}");
            }
        }

        // Re-host families on new ceilings
        private void RehostFamiliesOnNewCeilings(Document doc, List<FamilyInstance> hostedFamilies,
            Ceiling ceiling1, Ceiling ceiling2, XYZ splitStart, XYZ splitEnd, XYZ splitDirection,
            StringBuilder debugInfo)
        {
            if (hostedFamilies.Count == 0) return;

            XYZ normal = XYZ.BasisZ;
            XYZ perpendicular = normal.CrossProduct(splitDirection).Normalize();

            foreach (FamilyInstance fi in hostedFamilies)
            {
                try
                {
                    // Get family location
                    LocationPoint locPoint = fi.Location as LocationPoint;
                    if (locPoint == null) continue;

                    XYZ fiLocation = locPoint.Point;
                    double fiRotation = locPoint.Rotation;

                    // Store all necessary information before deleting
                    FamilySymbol fiSymbol = fi.Symbol;
                    ElementId originalId = fi.Id;
                    bool isFaceBased = fi.HostFace != null;

                    // Get the exact Z coordinate (elevation)
                    double originalElevation = fiLocation.Z;

                    // Get offset parameter if it exists
                    double hostOffset = 0;
                    double elevationFromLevel = 0;
                    try
                    {
                        Parameter offsetParam = fi.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM);
                        if (offsetParam != null && offsetParam.HasValue)
                        {
                            hostOffset = offsetParam.AsDouble();
                        }

                        // Also get elevation from level parameter
                        Parameter elevParam = fi.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM);
                        if (elevParam != null && elevParam.HasValue)
                        {
                            elevationFromLevel = elevParam.AsDouble();
                        }

                        // Check for schedule level parameter
                        Parameter scheduleLevelParam = fi.get_Parameter(BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM);
                        if (scheduleLevelParam != null && scheduleLevelParam.HasValue)
                        {
                            parameterValues["ScheduleLevel"] = scheduleLevelParam.AsElementId();
                        }
                    }
                    catch { }

                    // Store parameter values before deletion
                    Dictionary<string, object> parameterValues = new Dictionary<string, object>();
                    parameterValues["HostOffset"] = hostOffset; // Store the offset separately
                    parameterValues["ElevationFromLevel"] = elevationFromLevel; // Store elevation separately
                    foreach (Parameter param in fi.Parameters)
                    {
                        if (!param.IsReadOnly && param.HasValue)
                        {
                            try
                            {
                                switch (param.StorageType)
                                {
                                    case StorageType.Double:
                                        parameterValues[param.Definition.Name] = param.AsDouble();
                                        break;
                                    case StorageType.Integer:
                                        parameterValues[param.Definition.Name] = param.AsInteger();
                                        break;
                                    case StorageType.String:
                                        parameterValues[param.Definition.Name] = param.AsString();
                                        break;
                                    case StorageType.ElementId:
                                        parameterValues[param.Definition.Name] = param.AsElementId();
                                        break;
                                }
                            }
                            catch { }
                        }
                    }

                    // Determine which side of split line
                    double side = (fiLocation - splitStart).DotProduct(perpendicular);

                    // Choose appropriate ceiling
                    Ceiling newCeiling = side > 0 ? ceiling1 : ceiling2;

                    // Delete old instance
                    doc.Delete(fi.Id);

                    // Create new hosted instance with exact same location
                    CreateHostedFamilyOnNewCeiling(doc, fiSymbol, newCeiling, fiLocation,
                        parameterValues, originalId, fiRotation, isFaceBased, originalElevation, debugInfo);
                }
                catch (Exception ex)
                {
                    debugInfo.AppendLine($"Failed to re-host family: {ex.Message}");
                    if (fi != null && fi.IsValidObject)
                    {
                        failedHostedElements.Add(fi.Id);
                    }
                }
            }
        }

        // Create hosted family on new ceiling (updated method signature)
        private void CreateHostedFamilyOnNewCeiling(Document doc, FamilySymbol fiSymbol,
            Ceiling newCeiling, XYZ location, Dictionary<string, object> parameterValues,
            ElementId originalId, double rotation, bool wasFaceBased, double originalElevation,
            StringBuilder debugInfo)
        {
            try
            {
                FamilyInstance newInstance = null;

                if (wasFaceBased)
                {
                    // Get reference to new ceiling's bottom face
                    Reference newCeilingRef = GetCeilingBottomFaceReference(newCeiling, location);

                    if (newCeilingRef != null)
                    {
                        try
                        {
                            // For face-based families, we need to project the location onto the ceiling face
                            // to get the correct placement point
                            Face ceilingFace = doc.GetElement(newCeilingRef).GetGeometryObjectFromReference(newCeilingRef) as Face;
                            if (ceilingFace != null)
                            {
                                // Project the original location onto the ceiling face
                                IntersectionResult projResult = ceilingFace.Project(location);
                                if (projResult != null)
                                {
                                    // Use the projected point but maintain the original Z coordinate
                                    XYZ projectedPoint = projResult.XYZPoint;
                                    XYZ adjustedLocation = new XYZ(location.X, location.Y, originalElevation);

                                    newInstance = doc.Create.NewFamilyInstance(
                                        newCeilingRef, adjustedLocation, XYZ.BasisZ, fiSymbol);
                                }
                                else
                                {
                                    // If projection fails, use original location
                                    newInstance = doc.Create.NewFamilyInstance(
                                        newCeilingRef, location, XYZ.BasisZ, fiSymbol);
                                }
                            }
                            else
                            {
                                // Fallback to original location
                                newInstance = doc.Create.NewFamilyInstance(
                                    newCeilingRef, location, XYZ.BasisZ, fiSymbol);
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Failed to create face-based instance: {ex.Message}");

                            // Try alternative method with ceiling host
                            try
                            {
                                Level level = doc.GetElement(newCeiling.LevelId) as Level;
                                newInstance = doc.Create.NewFamilyInstance(
                                    location, fiSymbol, newCeiling, level,
                                    Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                            }
                            catch { }
                        }
                    }
                }

                // If not face-based or face-based creation failed
                if (newInstance == null)
                {
                    try
                    {
                        Level level = doc.GetElement(newCeiling.LevelId) as Level;

                        // First try with ceiling as host
                        try
                        {
                            newInstance = doc.Create.NewFamilyInstance(
                                location, fiSymbol, newCeiling, level,
                                Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                        }
                        catch
                        {
                            // If that fails, try level-based
                            newInstance = doc.Create.NewFamilyInstance(
                                location, fiSymbol, level,
                                Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"All placement methods failed: {ex.Message}");
                        failedHostedElements.Add(originalId);
                        return;
                    }
                }

                if (newInstance != null)
                {
                    Debug.WriteLine($"Successfully created hosted element on new ceiling");

                    // First, restore the elevation from level parameter immediately
                    if (parameterValues.ContainsKey("ElevationFromLevel"))
                    {
                        try
                        {
                            Parameter elevParam = newInstance.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM);
                            if (elevParam != null && !elevParam.IsReadOnly)
                            {
                                double elevValue = (double)parameterValues["ElevationFromLevel"];
                                elevParam.Set(elevValue);
                                Debug.WriteLine($"Set elevation from level to: {elevValue:F4}");
                            }
                        }
                        catch { }
                    }

                    // Check if we need to adjust the elevation further
                    LocationPoint newLoc = newInstance.Location as LocationPoint;
                    if (newLoc != null)
                    {
                        XYZ currentPos = newLoc.Point;
                        double elevationDiff = originalElevation - currentPos.Z;

                        // If there's still a significant elevation difference, try to correct it
                        if (Math.Abs(elevationDiff) > 0.001)
                        {
                            try
                            {
                                // Try using offset parameter instead of moving
                                Parameter offsetParam = newInstance.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM);
                                if (offsetParam != null && !offsetParam.IsReadOnly)
                                {
                                    offsetParam.Set(elevationDiff);
                                    Debug.WriteLine($"Set offset parameter to {elevationDiff:F4}");
                                }
                                else
                                {
                                    // Try to move the element to the correct elevation
                                    XYZ translation = new XYZ(0, 0, elevationDiff);
                                    ElementTransformUtils.MoveElement(doc, newInstance.Id, translation);
                                    Debug.WriteLine($"Moved element by {elevationDiff:F4} feet");
                                }
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"Could not adjust elevation: {ex.Message}");
                            }
                        }
                    }

                    // Restore rotation
                    if (Math.Abs(rotation) > 0.001)
                    {
                        try
                        {
                            newLoc = newInstance.Location as LocationPoint;
                            if (newLoc != null)
                            {
                                Line axis = Line.CreateBound(newLoc.Point, newLoc.Point + XYZ.BasisZ);
                                newLoc.Rotate(axis, rotation);
                            }
                        }
                        catch { }
                    }

                    // Restore all other parameters
                    RestoreParameters(newInstance, parameterValues);

                    // Final verification
                    newLoc = newInstance.Location as LocationPoint;
                    if (newLoc != null)
                    {
                        XYZ finalPos = newLoc.Point;
                        double posDiff = location.DistanceTo(finalPos);
                        double zDiff = Math.Abs(originalElevation - finalPos.Z);

                        if (posDiff > 0.001 || zDiff > 0.001)
                        {
                            debugInfo.AppendLine($"Warning: Position difference - XY: {posDiff:F4} ft, Z: {zDiff:F4} ft");

                            // Log parameter values for debugging
                            Parameter finalElevParam = newInstance.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM);
                            if (finalElevParam != null)
                            {
                                debugInfo.AppendLine($"Final elevation from level: {finalElevParam.AsDouble():F4}");
                            }
                        }
                    }
                }
                else
                {
                    failedHostedElements.Add(originalId);
                }
            }
            catch (Exception ex)
            {
                debugInfo.AppendLine($"Error creating hosted family: {ex.Message}");
                failedHostedElements.Add(originalId);
            }
        }

        // Restore parameters to new instance
        private void RestoreParameters(FamilyInstance target, Dictionary<string, object> parameterValues)
        {
            try
            {
                foreach (var kvp in parameterValues)
                {
                    // Handle special case for host offset
                    if (kvp.Key == "HostOffset" && kvp.Value is double offsetValue)
                    {
                        try
                        {
                            Parameter offsetParam = target.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM);
                            if (offsetParam != null && !offsetParam.IsReadOnly)
                            {
                                offsetParam.Set(offsetValue);
                                Debug.WriteLine($"Restored host offset: {offsetValue:F4}");
                                continue;
                            }
                        }
                        catch { }
                    }

                    // Handle elevation from level
                    if (kvp.Key == "ElevationFromLevel" && kvp.Value is double elevValue)
                    {
                        try
                        {
                            Parameter elevParam = target.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM);
                            if (elevParam != null && !elevParam.IsReadOnly)
                            {
                                elevParam.Set(elevValue);
                                Debug.WriteLine($"Restored elevation from level: {elevValue:F4}");
                                continue;
                            }
                        }
                        catch { }
                    }

                    // Handle schedule level
                    if (kvp.Key == "ScheduleLevel" && kvp.Value is ElementId levelId)
                    {
                        try
                        {
                            Parameter scheduleLevelParam = target.get_Parameter(BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM);
                            if (scheduleLevelParam != null && !scheduleLevelParam.IsReadOnly)
                            {
                                scheduleLevelParam.Set(levelId);
                                Debug.WriteLine($"Restored schedule level");
                                continue;
                            }
                        }
                        catch { }
                    }

                    Parameter targetParam = target.LookupParameter(kvp.Key);
                    if (targetParam != null && !targetParam.IsReadOnly)
                    {
                        try
                        {
                            if (kvp.Value is double)
                                targetParam.Set((double)kvp.Value);
                            else if (kvp.Value is int)
                                targetParam.Set((int)kvp.Value);
                            else if (kvp.Value is string)
                                targetParam.Set((string)kvp.Value);
                            else if (kvp.Value is ElementId)
                                targetParam.Set((ElementId)kvp.Value);
                        }
                        catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error restoring parameters: {ex.Message}");
            }
        }

        // Get reference to ceiling's bottom face
        private Reference GetCeilingBottomFaceReference(Ceiling ceiling, XYZ location)
        {
            try
            {
                Options options = new Options();
                options.ComputeReferences = true;
                options.IncludeNonVisibleObjects = true;
                options.View = ceiling.Document.ActiveView; // Add view context

                GeometryElement geomElem = ceiling.get_Geometry(options);

                // Check if geometry is null
                if (geomElem == null)
                {
                    Debug.WriteLine($"Warning: Could not get geometry for ceiling {ceiling.Id}");
                    return null;
                }

                foreach (GeometryObject geomObj in geomElem)
                {
                    if (geomObj is Solid solid && solid.Volume > 0) // Check for valid solid
                    {
                        foreach (Face face in solid.Faces)
                        {
                            if (face is PlanarFace planarFace)
                            {
                                // Check if face normal points down (bottom face)
                                if (planarFace.FaceNormal.DotProduct(XYZ.BasisZ) < -0.9) // More precise check
                                {
                                    try
                                    {
                                        IntersectionResult result = face.Project(location);
                                        if (result != null && result.Distance < 10.0) // Reasonable distance check
                                        {
                                            return face.Reference;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Debug.WriteLine($"Failed to project point to face: {ex.Message}");
                                    }
                                }
                            }
                        }
                    }
                    else if (geomObj is GeometryInstance instance)
                    {
                        // Handle geometry instances (for complex ceilings)
                        GeometryElement instanceGeometry = instance.GetInstanceGeometry();
                        if (instanceGeometry != null)
                        {
                            foreach (GeometryObject instObj in instanceGeometry)
                            {
                                if (instObj is Solid instSolid && instSolid.Volume > 0)
                                {
                                    foreach (Face face in instSolid.Faces)
                                    {
                                        if (face is PlanarFace planarFace &&
                                            planarFace.FaceNormal.DotProduct(XYZ.BasisZ) < -0.9)
                                        {
                                            try
                                            {
                                                XYZ transformedLocation = instance.Transform.Inverse.OfPoint(location);
                                                IntersectionResult result = face.Project(transformedLocation);
                                                if (result != null && result.Distance < 10.0)
                                                {
                                                    return face.Reference;
                                                }
                                            }
                                            catch { }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                Debug.WriteLine($"Warning: Could not find suitable bottom face for ceiling {ceiling.Id}");
                return null;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error in GetCeilingBottomFaceReference: {ex.Message}");
                return null;
            }
        }

        // Copy parameters from one instance to another
        private void CopyInstanceParameters(FamilyInstance source, FamilyInstance target)
        {
            try
            {
                foreach (Parameter sourceParam in source.Parameters)
                {
                    if (!sourceParam.IsReadOnly && sourceParam.HasValue)
                    {
                        Parameter targetParam = target.LookupParameter(sourceParam.Definition.Name);
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
                                    targetParam.Set(sourceParam.AsString());
                                    break;
                                case StorageType.ElementId:
                                    targetParam.Set(sourceParam.AsElementId());
                                    break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error copying parameters: {ex.Message}");
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
                        debugInfo.AppendLine("  WARNING: Skipping opening that intersects split line");
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

        // Get ceiling boundary loops
        private Tuple<List<Curve>, List<List<Curve>>> GetCeilingBoundaryLoops(Ceiling ceiling)
        {
            List<Curve> outerBoundary = new List<Curve>();
            List<List<Curve>> innerBoundaries = new List<List<Curve>>();

            Options options = new Options();
            options.ComputeReferences = true;
            options.IncludeNonVisibleObjects = true;
            options.View = ceiling.Document.ActiveView;

            GeometryElement geomElem = ceiling.get_Geometry(options);

            if (geomElem == null)
            {
                Debug.WriteLine($"Warning: Could not get geometry for ceiling {ceiling.Id}");
                return new Tuple<List<Curve>, List<List<Curve>>>(outerBoundary, innerBoundaries);
            }

            foreach (GeometryObject geomObj in geomElem)
            {
                if (geomObj is Solid solid && solid.Volume > 0)
                {
                    foreach (Face face in solid.Faces)
                    {
                        if (face is PlanarFace planarFace &&
                            Math.Abs(planarFace.FaceNormal.Z + 1) < 0.01) // Bottom face (normal points down)
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
                else if (geomObj is GeometryInstance instance)
                {
                    // Handle geometry instances
                    GeometryElement instanceGeometry = instance.GetInstanceGeometry();
                    if (instanceGeometry != null)
                    {
                        foreach (GeometryObject instObj in instanceGeometry)
                        {
                            if (instObj is Solid instSolid && instSolid.Volume > 0)
                            {
                                foreach (Face face in instSolid.Faces)
                                {
                                    if (face is PlanarFace planarFace &&
                                        Math.Abs(planarFace.FaceNormal.Z + 1) < 0.01)
                                    {
                                        // Process similar to above
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
                                                Curve transformedCurve = edge.AsCurve();
                                                transformedCurve = transformedCurve.CreateTransformed(instance.Transform);
                                                outerBoundary.Add(transformedCurve);
                                            }
                                        }

                                        foreach (EdgeArray innerLoop in innerLoops)
                                        {
                                            List<Curve> innerCurves = new List<Curve>();
                                            foreach (Edge edge in innerLoop)
                                            {
                                                Curve transformedCurve = edge.AsCurve();
                                                transformedCurve = transformedCurve.CreateTransformed(instance.Transform);
                                                innerCurves.Add(transformedCurve);
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

        // Get button data for ribbon
        internal static PushButtonData GetButtonData()
        {
            string buttonInternalName = "btnCeilingSplitter";
            string buttonTitle = "Split Ceiling";

            ButtonDataClass myButtonData = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "Splits a ceiling into two parts along a model line while preserving hosted elements");

            return myButtonData.Data;
        }
    }

    // Selection filter for ceilings
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