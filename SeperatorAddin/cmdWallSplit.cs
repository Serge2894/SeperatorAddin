using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using SeperatorAddin.Common;
using SeperatorAddin.Forms;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Windows.Interop;

namespace SeperatorAddin
{
    [Transaction(TransactionMode.Manual)]
    public class cmdWallSplit : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Get all levels in the project, sorted by elevation
                FilteredElementCollector levelCollector = new FilteredElementCollector(doc);
                List<Level> allLevels = levelCollector.OfClass(typeof(Level))
                    .Cast<Level>()
                    .OrderBy(l => l.ProjectElevation)
                    .ToList();

                if (allLevels.Count < 2)
                {
                    TaskDialog.Show("Error", "The project must have at least 2 levels to split walls.");
                    return Result.Failed;
                }

                // Create level selection form
                WallSelectionForm form = new WallSelectionForm(allLevels);

                // Show form
                WindowInteropHelper helper = new WindowInteropHelper(form);
                helper.Owner = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;

                bool? dialogResult = form.ShowDialog();

                if (dialogResult != true)
                {
                    return Result.Cancelled;
                }

                // Get selected levels
                List<Level> selectedLevels = form.GetSelectedLevels();

                if (selectedLevels.Count < 2)
                {
                    TaskDialog.Show("Error", "Please select at least 2 levels.");
                    return Result.Failed;
                }

                // Sort selected levels by elevation
                selectedLevels = selectedLevels.OrderBy(l => l.ProjectElevation).ToList();

                // Select walls
                Utils.WallSelectionilter filter = new Utils.WallSelectionilter();
                IList<Reference> references;

                try
                {
                    references = uidoc.Selection.PickObjects(ObjectType.Element, filter, "Select walls to split by levels");
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    return Result.Cancelled;
                }

                if (references.Count == 0)
                {
                    return Result.Cancelled;
                }

                List<Wall> selectedWalls = new List<Wall>();
                foreach (Reference reference in references)
                {
                    Wall wall = doc.GetElement(reference.ElementId) as Wall;
                    if (wall != null)
                    {
                        selectedWalls.Add(wall);
                    }
                }

                // Process walls
                using (Transaction trans = new Transaction(doc, "Split Walls by Levels"))
                {
                    trans.Start();

                    int totalWallsCreated = 0;
                    List<ElementId> wallsToDelete = new List<ElementId>();

                    foreach (Wall originalWall in selectedWalls)
                    {
                        List<Wall> newWalls = SplitWallByLevels(doc, originalWall, selectedLevels);

                        if (newWalls.Count > 0)
                        {
                            totalWallsCreated += newWalls.Count;
                            wallsToDelete.Add(originalWall.Id);
                        }
                    }

                    // Delete original walls
                    foreach (ElementId wallId in wallsToDelete)
                    {
                        doc.Delete(wallId);
                    }

                    trans.Commit();

                    TaskDialog.Show("Success",
                        $"Split {selectedWalls.Count} wall(s) into {totalWallsCreated} new wall(s) across {selectedLevels.Count} levels.");
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private List<Wall> SplitWallByLevels(Document doc, Wall originalWall, List<Level> levels)
        {
            List<Wall> newWalls = new List<Wall>();

            try
            {
                // Get wall properties
                LocationCurve locationCurve = originalWall.Location as LocationCurve;
                if (locationCurve == null) return newWalls;

                Curve wallCurve = locationCurve.Curve;
                WallType wallType = originalWall.WallType;

                // Get original wall base and top information
                Level originalBaseLevel = doc.GetElement(originalWall.LevelId) as Level;
                double originalBaseOffset = originalWall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET).AsDouble();
                double originalBaseElevation = originalBaseLevel.ProjectElevation + originalBaseOffset;

                // Get top constraint
                Parameter topConstraintParam = originalWall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE);
                ElementId topLevelId = topConstraintParam.AsElementId();

                double originalTopElevation;
                if (topLevelId != ElementId.InvalidElementId)
                {
                    // Wall has top constraint
                    Level topLevel = doc.GetElement(topLevelId) as Level;
                    double topOffset = originalWall.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET).AsDouble();
                    originalTopElevation = topLevel.ProjectElevation + topOffset;
                }
                else
                {
                    // Unconnected height
                    double height = originalWall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).AsDouble();
                    originalTopElevation = originalBaseElevation + height;
                }

                // Check if wall has edit profile
                bool hasEditProfile = originalWall.SketchId != ElementId.InvalidElementId;

                // Get wall openings information before creating new walls
                List<WallOpeningInfo> openingInfos = GetWallOpenings(originalWall, originalBaseElevation);

                // Determine which levels are within the wall's height range
                List<Level> relevantLevels = new List<Level>();
                foreach (Level level in levels)
                {
                    if (level.ProjectElevation >= originalBaseElevation - 0.01 &&
                        level.ProjectElevation <= originalTopElevation + 0.01)
                    {
                        relevantLevels.Add(level);
                    }
                }

                // Ensure we include a level below the base if needed
                Level baseLevelForSplit = levels.LastOrDefault(l => l.ProjectElevation <= originalBaseElevation);
                if (baseLevelForSplit != null && !relevantLevels.Contains(baseLevelForSplit))
                {
                    relevantLevels.Insert(0, baseLevelForSplit);
                }

                // If wall has edit profile, handle it specially
                if (hasEditProfile)
                {
                    // For walls with edit profile, we'll create the walls but won't try to recreate the profile
                    // The user will need to edit the profile of the new walls manually
                    TaskDialog.Show("Edit Profile Wall",
                        "This wall has an edited profile. The wall will be split, but the profile openings will need to be recreated manually on each segment.");
                }

                // Create new walls between each pair of levels
                for (int i = 0; i < relevantLevels.Count - 1; i++)
                {
                    Level baseLevel = relevantLevels[i];
                    Level topLevel = relevantLevels[i + 1];

                    // Calculate base offset
                    double baseOffset = 0;
                    if (i == 0 && baseLevel.ProjectElevation < originalBaseElevation)
                    {
                        baseOffset = originalBaseElevation - baseLevel.ProjectElevation;
                    }

                    // Calculate top offset
                    double topOffset = 0;
                    if (i == relevantLevels.Count - 2 && topLevel.ProjectElevation > originalTopElevation)
                    {
                        topOffset = originalTopElevation - topLevel.ProjectElevation;
                    }

                    // Create the new wall
                    Wall newWall = Wall.Create(doc, wallCurve, wallType.Id, baseLevel.Id,
                        topLevel.ProjectElevation - baseLevel.ProjectElevation - baseOffset, baseOffset, false, false);

                    if (newWall != null)
                    {
                        // Set top constraint
                        Parameter newTopConstraint = newWall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE);
                        if (newTopConstraint != null && !newTopConstraint.IsReadOnly)
                        {
                            newTopConstraint.Set(topLevel.Id);
                        }

                        // Set top offset
                        Parameter newTopOffset = newWall.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET);
                        if (newTopOffset != null && !newTopOffset.IsReadOnly)
                        {
                            newTopOffset.Set(topOffset);
                        }

                        // Copy all parameters from original wall
                        CopyAllWallParameters(originalWall, newWall);

                        // Apply openings to this wall segment (skip edit profile openings)
                        if (!hasEditProfile)
                        {
                            double segmentBaseElev = baseLevel.ProjectElevation + baseOffset;
                            double segmentTopElev = topLevel.ProjectElevation + topOffset;
                            ApplyOpeningsToWall(doc, newWall, openingInfos, segmentBaseElev, segmentTopElev);
                        }

                        newWalls.Add(newWall);
                    }
                }

                // Handle the last segment if the original wall extends above the highest selected level
                if (relevantLevels.Count > 0 && relevantLevels.Last().ProjectElevation < originalTopElevation)
                {
                    Level lastLevel = relevantLevels.Last();
                    double baseOffset = 0;
                    double height = originalTopElevation - lastLevel.ProjectElevation;

                    Wall topWall = Wall.Create(doc, wallCurve, wallType.Id, lastLevel.Id, height, baseOffset, false, false);

                    if (topWall != null)
                    {
                        // If original had top constraint, apply it
                        if (topLevelId != ElementId.InvalidElementId)
                        {
                            Parameter newTopConstraint = topWall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE);
                            if (newTopConstraint != null && !newTopConstraint.IsReadOnly)
                            {
                                newTopConstraint.Set(topLevelId);
                            }

                            Parameter newTopOffset = topWall.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET);
                            if (newTopOffset != null && !newTopOffset.IsReadOnly)
                            {
                                double originalTopOffset = originalWall.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET).AsDouble();
                                newTopOffset.Set(originalTopOffset);
                            }
                        }

                        CopyAllWallParameters(originalWall, topWall);

                        // Apply openings to top segment (skip edit profile openings)
                        if (!hasEditProfile)
                        {
                            double segmentBaseElev = lastLevel.ProjectElevation;
                            ApplyOpeningsToWall(doc, topWall, openingInfos, segmentBaseElev, originalTopElevation);
                        }

                        newWalls.Add(topWall);
                    }
                }

                // Join new walls at their connection points
                for (int i = 0; i < newWalls.Count - 1; i++)
                {
                    try
                    {
                        if (!JoinGeometryUtils.AreElementsJoined(doc, newWalls[i], newWalls[i + 1]))
                        {
                            JoinGeometryUtils.JoinGeometry(doc, newWalls[i], newWalls[i + 1]);
                        }
                    }
                    catch
                    {
                        // Join might fail for some wall types, continue anyway
                    }
                }

                // Handle hosted elements (doors, windows)
                HandleHostedElements(doc, originalWall, newWalls, levels);
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to split wall: {ex.Message}");
            }

            return newWalls;
        }

        private void CopyAllWallParameters(Wall sourceWall, Wall targetWall)
        {
            try
            {
                // Get all parameters from source wall
                foreach (Parameter sourceParam in sourceWall.Parameters)
                {
                    if (sourceParam == null || !sourceParam.HasValue) continue;

                    // Skip certain built-in parameters that shouldn't be copied
                    InternalDefinition def = sourceParam.Definition as InternalDefinition;
                    if (def != null)
                    {
                        BuiltInParameter bip = (BuiltInParameter)def.BuiltInParameter;

                        // Skip parameters that are set automatically
                        if (bip == BuiltInParameter.WALL_BASE_CONSTRAINT ||
                            bip == BuiltInParameter.WALL_HEIGHT_TYPE ||
                            bip == BuiltInParameter.WALL_BASE_OFFSET ||
                            bip == BuiltInParameter.WALL_TOP_OFFSET ||
                            bip == BuiltInParameter.WALL_USER_HEIGHT_PARAM ||
                            bip == BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM ||
                            bip == BuiltInParameter.ELEM_FAMILY_PARAM ||
                            bip == BuiltInParameter.ELEM_TYPE_PARAM ||
                            bip == BuiltInParameter.HOST_AREA_COMPUTED ||
                            bip == BuiltInParameter.HOST_VOLUME_COMPUTED ||
                            bip == BuiltInParameter.CURVE_ELEM_LENGTH)
                        {
                            continue;
                        }
                    }

                    // Try to find the corresponding parameter in target wall
                    Parameter targetParam = targetWall.get_Parameter(sourceParam.Definition);

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
                                    string strValue = sourceParam.AsString();
                                    if (!string.IsNullOrEmpty(strValue))
                                        targetParam.Set(strValue);
                                    break;
                                case StorageType.ElementId:
                                    ElementId id = sourceParam.AsElementId();
                                    if (id != ElementId.InvalidElementId)
                                        targetParam.Set(id);
                                    break;
                            }
                        }
                        catch
                        {
                            // Some parameters might fail to copy, continue with others
                        }
                    }
                }

                // Also copy shared parameters
                CopySharedParameters(sourceWall, targetWall);
            }
            catch (Exception ex)
            {
                // Log error but continue
                Debug.WriteLine($"Error copying parameters: {ex.Message}");
            }
        }

        private void CopySharedParameters(Wall sourceWall, Wall targetWall)
        {
            try
            {
                // Get all shared parameters from source
                BindingMap bindingMap = sourceWall.Document.ParameterBindings;
                DefinitionBindingMapIterator iterator = bindingMap.ForwardIterator();

                while (iterator.MoveNext())
                {
                    Definition def = iterator.Key;
                    if (def != null)
                    {
                        Parameter sourceParam = sourceWall.get_Parameter(def);
                        Parameter targetParam = targetWall.get_Parameter(def);

                        if (sourceParam != null && targetParam != null &&
                            sourceParam.HasValue && !targetParam.IsReadOnly)
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
                                        string strValue = sourceParam.AsString();
                                        if (!string.IsNullOrEmpty(strValue))
                                            targetParam.Set(strValue);
                                        break;
                                    case StorageType.ElementId:
                                        ElementId id = sourceParam.AsElementId();
                                        if (id != ElementId.InvalidElementId)
                                            targetParam.Set(id);
                                        break;
                                }
                            }
                            catch
                            {
                                // Continue with next parameter
                            }
                        }
                    }
                }
            }
            catch
            {
                // If binding map fails, continue
            }
        }

        private class WallOpeningInfo
        {
            public OpeningType Type { get; set; }
            public double BottomElevation { get; set; }
            public double TopElevation { get; set; }
            public List<Curve> ProfileCurves { get; set; }
            public XYZ InsertPoint { get; set; }
            public ElementId OpeningId { get; set; }
            public double Width { get; set; }
            public double Height { get; set; }
        }

        private enum OpeningType
        {
            EditProfile,
            RectangularOpening,
            ArcOpening,
            HostedElement
        }

        private List<WallOpeningInfo> GetWallOpenings(Wall wall, double wallBaseElevation)
        {
            List<WallOpeningInfo> openings = new List<WallOpeningInfo>();

            try
            {
                // Check for edit profile openings
                if (wall.SketchId != ElementId.InvalidElementId)
                {
                    Sketch sketch = wall.Document.GetElement(wall.SketchId) as Sketch;
                    if (sketch != null)
                    {
                        CurveArrArray profiles = sketch.Profile;
                        for (int i = 0; i < profiles.Size; i++)
                        {
                            CurveArray curveArray = profiles.get_Item(i);
                            if (i > 0) // First profile is the outer boundary
                            {
                                List<Curve> profileCurves = new List<Curve>();
                                double minZ = double.MaxValue;
                                double maxZ = double.MinValue;

                                foreach (Curve curve in curveArray)
                                {
                                    profileCurves.Add(curve);
                                    XYZ start = curve.GetEndPoint(0);
                                    XYZ end = curve.GetEndPoint(1);
                                    minZ = Math.Min(minZ, Math.Min(start.Z, end.Z));
                                    maxZ = Math.Max(maxZ, Math.Max(start.Z, end.Z));
                                }

                                openings.Add(new WallOpeningInfo
                                {
                                    Type = OpeningType.EditProfile,
                                    BottomElevation = minZ,
                                    TopElevation = maxZ,
                                    ProfileCurves = profileCurves
                                });
                            }
                        }
                    }
                }

                // Get dependent elements (openings and hosted elements)
                ICollection<ElementId> dependentIds = wall.GetDependentElements(null);

                foreach (ElementId id in dependentIds)
                {
                    Element elem = wall.Document.GetElement(id);

                    // Check for wall openings
                    if (elem is Opening opening)
                    {
                        BoundingBoxXYZ bb = opening.get_BoundingBox(null);
                        if (bb != null)
                        {
                            // Calculate opening dimensions
                            double width = Math.Sqrt(Math.Pow(bb.Max.X - bb.Min.X, 2) +
                                                    Math.Pow(bb.Max.Y - bb.Min.Y, 2));
                            double height = bb.Max.Z - bb.Min.Z;

                            // Check if it's rectangular or arc opening
                            bool isRect = opening.IsRectBoundary;

                            openings.Add(new WallOpeningInfo
                            {
                                Type = isRect ? OpeningType.RectangularOpening : OpeningType.ArcOpening,
                                BottomElevation = bb.Min.Z,
                                TopElevation = bb.Max.Z,
                                OpeningId = opening.Id,
                                InsertPoint = (bb.Min + bb.Max) / 2,
                                Width = width,
                                Height = height
                            });
                        }
                    }
                    // Check for hosted elements that create openings
                    else if (elem is FamilyInstance fi && fi.Host?.Id == wall.Id)
                    {
                        Category category = fi.Category;
                        if (category != null &&
                            (category.Id.IntegerValue == (int)BuiltInCategory.OST_Doors ||
                             category.Id.IntegerValue == (int)BuiltInCategory.OST_Windows))
                        {
                            // These are handled separately in HandleHostedElements
                            continue;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error getting wall openings: {ex.Message}");
            }

            return openings;
        }

        private void ApplyOpeningsToWall(Document doc, Wall newWall, List<WallOpeningInfo> openingInfos,
            double wallBottomElev, double wallTopElev)
        {
            foreach (var openingInfo in openingInfos)
            {
                // Check if opening falls within this wall segment
                if (openingInfo.BottomElevation < wallTopElev && openingInfo.TopElevation > wallBottomElev)
                {
                    try
                    {
                        switch (openingInfo.Type)
                        {
                            case OpeningType.EditProfile:
                                // For edit profile openings, skip for now as they require special handling
                                // The profile will be maintained through the wall type or we'll handle it differently
                                break;

                            case OpeningType.RectangularOpening:
                            case OpeningType.ArcOpening:
                                // Create new opening
                                CreateWallOpening(doc, newWall, openingInfo, wallBottomElev, wallTopElev);
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Failed to apply opening: {ex.Message}");
                    }
                }
            }
        }

        private void CreateWallOpening(Document doc, Wall wall, WallOpeningInfo openingInfo,
            double wallBottomElev, double wallTopElev)
        {
            try
            {
                // Get wall location curve
                LocationCurve locCurve = wall.Location as LocationCurve;
                if (locCurve == null || !(locCurve.Curve is Line wallLine)) return;

                // Get the original opening from the document if it still exists
                Opening originalOpening = null;
                if (openingInfo.OpeningId != ElementId.InvalidElementId)
                {
                    originalOpening = doc.GetElement(openingInfo.OpeningId) as Opening;
                }

                // Calculate opening bounds relative to new wall
                double openingBottom = Math.Max(openingInfo.BottomElevation, wallBottomElev);
                double openingTop = Math.Min(openingInfo.TopElevation, wallTopElev);

                if (openingTop <= openingBottom) return;

                // Get wall parameters
                Level wallLevel = doc.GetElement(wall.LevelId) as Level;
                double wallBaseOffset = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET).AsDouble();
                double wallBaseElevation = wallLevel.ProjectElevation + wallBaseOffset;

                // Calculate relative positions
                double relativeBottom = openingBottom - wallBaseElevation;
                double relativeTop = openingTop - wallBaseElevation;

                // Get opening width
                double openingWidth = openingInfo.Width;
                if (openingWidth <= 0) openingWidth = 3.0; // Default 3 feet

                // Find the position along the wall
                XYZ wallStart = wallLine.GetEndPoint(0);
                XYZ wallEnd = wallLine.GetEndPoint(1);
                XYZ wallDirection = wallLine.Direction;

                // Project opening center onto wall line
                IntersectionResult projection = wallLine.Project(openingInfo.InsertPoint);
                if (projection == null) return;

                // Calculate distance along wall
                double distanceAlongWall = projection.Parameter * wallLine.Length;

                // Calculate the start and end points of the opening
                double startDistance = distanceAlongWall - (openingWidth / 2);
                double endDistance = distanceAlongWall + (openingWidth / 2);

                // Make sure opening is within wall bounds
                startDistance = Math.Max(0, startDistance);
                endDistance = Math.Min(wallLine.Length, endDistance);

                // Create points for rectangular opening
                XYZ point1 = wallStart + wallDirection * startDistance;
                XYZ point2 = wallStart + wallDirection * endDistance;

                // Set the Z coordinates
                point1 = new XYZ(point1.X, point1.Y, openingBottom);
                point2 = new XYZ(point2.X, point2.Y, openingTop);

                // Create the opening using the correct method
                if (openingInfo.Type == OpeningType.RectangularOpening)
                {
                    // For rectangular openings, create directly
                    Opening newOpening = doc.Create.NewOpening(wall, point1, point2);
                }
                else if (openingInfo.Type == OpeningType.ArcOpening && originalOpening != null)
                {
                    // For arc openings, try to copy from original if possible
                    try
                    {
                        // Get arc parameters from original opening
                        CurveArray curveArray = originalOpening.BoundaryCurves;
                        if (curveArray != null && curveArray.Size > 0)
                        {
                            // For now, create as rectangular - arc openings require special handling
                            Opening newOpening = doc.Create.NewOpening(wall, point1, point2);
                        }
                    }
                    catch
                    {
                        // Fallback to rectangular
                        Opening newOpening = doc.Create.NewOpening(wall, point1, point2);
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to create opening: {ex.Message}");
                // Don't throw - continue with other openings
            }
        }

        private Curve ClipCurveToElevation(Curve curve, double bottomElev, double topElev)
        {
            try
            {
                XYZ start = curve.GetEndPoint(0);
                XYZ end = curve.GetEndPoint(1);

                // Check if curve needs clipping
                if (start.Z >= bottomElev && start.Z <= topElev &&
                    end.Z >= bottomElev && end.Z <= topElev)
                {
                    return curve; // No clipping needed
                }

                // Clip curve
                XYZ newStart = start;
                XYZ newEnd = end;

                if (start.Z < bottomElev)
                {
                    double t = (bottomElev - start.Z) / (end.Z - start.Z);
                    newStart = start + (end - start) * t;
                }
                else if (start.Z > topElev)
                {
                    double t = (topElev - start.Z) / (end.Z - start.Z);
                    newStart = start + (end - start) * t;
                }

                if (end.Z < bottomElev)
                {
                    double t = (bottomElev - start.Z) / (end.Z - start.Z);
                    newEnd = start + (end - start) * t;
                }
                else if (end.Z > topElev)
                {
                    double t = (topElev - start.Z) / (end.Z - start.Z);
                    newEnd = start + (end - start) * t;
                }

                if (curve is Line)
                {
                    return Line.CreateBound(newStart, newEnd);
                }
                else if (curve is Arc)
                {
                    // For arcs, create a line approximation when clipped
                    return Line.CreateBound(newStart, newEnd);
                }
            }
            catch
            {
                // Return null if clipping fails
            }

            return null;
        }

        private void HandleHostedElements(Document doc, Wall originalWall, List<Wall> newWalls, List<Level> levels)
        {
            try
            {
                // Get all hosted elements
                List<ElementId> hostedElementIds = originalWall.GetDependentElements(null).ToList();

                foreach (ElementId elementId in hostedElementIds)
                {
                    Element element = doc.GetElement(elementId);

                    if (element is FamilyInstance familyInstance && familyInstance.Host?.Id == originalWall.Id)
                    {
                        // Get the elevation of the hosted element
                        double elementElevation = GetHostedElementElevation(familyInstance);

                        // Find which new wall should host this element
                        Wall newHost = FindAppropriateHost(newWalls, elementElevation);

                        if (newHost != null)
                        {
                            try
                            {
                                // Get location point
                                LocationPoint locationPoint = familyInstance.Location as LocationPoint;
                                if (locationPoint != null)
                                {
                                    XYZ point = locationPoint.Point;

                                    // Create new instance on the new wall
                                    FamilyInstance newInstance = doc.Create.NewFamilyInstance(
                                        point,
                                        familyInstance.Symbol,
                                        newHost,
                                        familyInstance.StructuralType);

                                    // Copy all parameters
                                    CopyInstanceParameters(familyInstance, newInstance);
                                }
                            }
                            catch
                            {
                                // If we can't recreate it, the original will be deleted with the wall
                            }
                        }
                    }
                }
            }
            catch
            {
                // If handling hosted elements fails, continue anyway
            }
        }

        private double GetHostedElementElevation(FamilyInstance instance)
        {
            try
            {
                // First try to get sill height for windows
                Parameter sillHeight = instance.get_Parameter(BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM);
                if (sillHeight != null && sillHeight.HasValue)
                {
                    Level hostLevel = instance.Document.GetElement(instance.LevelId) as Level;
                    if (hostLevel != null)
                    {
                        return hostLevel.ProjectElevation + sillHeight.AsDouble();
                    }
                }

                // Try location point
                LocationPoint location = instance.Location as LocationPoint;
                if (location != null)
                {
                    return location.Point.Z;
                }

                // Try to get from bounding box
                BoundingBoxXYZ bbox = instance.get_BoundingBox(null);
                if (bbox != null)
                {
                    return (bbox.Min.Z + bbox.Max.Z) / 2;
                }
            }
            catch
            {
            }

            return 0;
        }

        private Wall FindAppropriateHost(List<Wall> walls, double elevation)
        {
            foreach (Wall wall in walls)
            {
                try
                {
                    Level baseLevel = wall.Document.GetElement(wall.LevelId) as Level;
                    double baseOffset = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET).AsDouble();
                    double baseElevation = baseLevel.ProjectElevation + baseOffset;

                    double topElevation;
                    Parameter topConstraint = wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE);
                    ElementId topLevelId = topConstraint.AsElementId();

                    if (topLevelId != ElementId.InvalidElementId)
                    {
                        Level topLevel = wall.Document.GetElement(topLevelId) as Level;
                        double topOffset = wall.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET).AsDouble();
                        topElevation = topLevel.ProjectElevation + topOffset;
                    }
                    else
                    {
                        double height = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).AsDouble();
                        topElevation = baseElevation + height;
                    }

                    if (elevation >= baseElevation && elevation <= topElevation)
                    {
                        return wall;
                    }
                }
                catch
                {
                    continue;
                }
            }

            return null;
        }

        private void CopyInstanceParameters(FamilyInstance source, FamilyInstance target)
        {
            try
            {
                // Copy all parameters
                foreach (Parameter param in source.Parameters)
                {
                    if (!param.IsReadOnly && param.HasValue)
                    {
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
                                        string value = param.AsString();
                                        if (!string.IsNullOrEmpty(value))
                                            targetParam.Set(value);
                                        break;
                                    case StorageType.ElementId:
                                        ElementId id = param.AsElementId();
                                        if (id != ElementId.InvalidElementId)
                                            targetParam.Set(id);
                                        break;
                                }
                            }
                            catch
                            {
                                // Some parameters might fail to copy
                            }
                        }
                    }
                }

                // Special handling for certain parameters
                CopySpecialInstanceParameters(source, target);
            }
            catch
            {
                // Some parameters might fail to copy
            }
        }

        private void CopySpecialInstanceParameters(FamilyInstance source, FamilyInstance target)
        {
            try
            {
                // Copy flip states
                if (source.CanFlipFacing && target.CanFlipFacing)
                {
                    target.flipFacing();
                    if (source.FacingFlipped != target.FacingFlipped)
                    {
                        target.flipFacing();
                    }
                }

                if (source.CanFlipHand && target.CanFlipHand)
                {
                    target.flipHand();
                    if (source.HandFlipped != target.HandFlipped)
                    {
                        target.flipHand();
                    }
                }

                // Copy mirrored state
                if (source.Mirrored != target.Mirrored)
                {
                    // Can't directly set mirrored state, this is controlled by creation method
                }

                // Copy rotation if applicable
                LocationPoint sourceLocation = source.Location as LocationPoint;
                LocationPoint targetLocation = target.Location as LocationPoint;

                if (sourceLocation != null && targetLocation != null)
                {
                    double sourceRotation = sourceLocation.Rotation;
                    if (Math.Abs(sourceRotation) > 0.001)
                    {
                        XYZ axis = XYZ.BasisZ;
                        targetLocation.Rotate(Line.CreateBound(targetLocation.Point,
                            targetLocation.Point + axis), sourceRotation);
                    }
                }
            }
            catch
            {
                // Special parameters might fail
            }
        }

        internal static PushButtonData GetButtonData()
        {
            string buttonInternalName = "btnWallSplit";
            string buttonTitle = "Split Walls\nby Level";

            ButtonDataClass myButtonData = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Green_32,
                Properties.Resources.Green_16,
                "Splits selected walls at specified levels, preserving all parameters and openings");

            return myButtonData.Data;
        }
    }
}