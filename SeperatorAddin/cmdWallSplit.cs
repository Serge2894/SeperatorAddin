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

                // If the lowest relevant level is above the wall's base, create a segment for the lowest part
                if (relevantLevels.Any() && originalBaseElevation < relevantLevels.First().ProjectElevation)
                {
                    Level topLevelForBottomSegment = relevantLevels.First();
                    double height = topLevelForBottomSegment.ProjectElevation - originalBaseElevation;
                    Wall bottomWall = Wall.Create(doc, wallCurve, wallType.Id, originalBaseLevel.Id, height, originalBaseOffset, false, false);
                    if (bottomWall != null)
                    {
                        // Set top constraint to the first relevant level
                        Parameter newTopConstraint = bottomWall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE);
                        if (newTopConstraint != null && !newTopConstraint.IsReadOnly)
                        {
                            newTopConstraint.Set(topLevelForBottomSegment.Id);
                        }
                        CopyAllWallParameters(originalWall, bottomWall);

                        double segmentBaseElev = originalBaseElevation;
                        double segmentTopElev = topLevelForBottomSegment.ProjectElevation;
                        ApplyOpeningsToWall(doc, bottomWall, openingInfos, segmentBaseElev, segmentTopElev);

                        newWalls.Add(bottomWall);
                    }
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

                        double segmentBaseElev = baseLevel.ProjectElevation + baseOffset;
                        double segmentTopElev = topLevel.ProjectElevation + topOffset;
                        ApplyOpeningsToWall(doc, newWall, openingInfos, segmentBaseElev, segmentTopElev);

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

                        double segmentBaseElev = lastLevel.ProjectElevation;
                        ApplyOpeningsToWall(doc, topWall, openingInfos, segmentBaseElev, originalTopElevation);

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
                System.Diagnostics.Debug.WriteLine($"Error copying parameters: {ex.Message}");
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
                    if (sketch != null && (wall.Location as LocationCurve)?.Curve is Line line)
                    {
                        CurveArrArray profiles = sketch.Profile;
                        for (int i = 0; i < profiles.Size; i++)
                        {
                            CurveArray curveArray = profiles.get_Item(i);
                            if (i > 0) // First profile is the outer boundary
                            {
                                // For edited profiles, we will approximate them with their bounding box
                                // to create a rectangular opening.
                                XYZ min = new XYZ(double.MaxValue, double.MaxValue, double.MaxValue);
                                XYZ max = new XYZ(double.MinValue, double.MinValue, double.MinValue);

                                foreach (Curve curve in curveArray)
                                {
                                    XYZ p1 = curve.GetEndPoint(0);
                                    XYZ p2 = curve.GetEndPoint(1);
                                    min = new XYZ(Math.Min(min.X, p1.X), Math.Min(min.Y, p1.Y), Math.Min(min.Z, p1.Z));
                                    max = new XYZ(Math.Max(max.X, p1.X), Math.Max(max.Y, p1.Y), Math.Max(max.Z, p1.Z));
                                    min = new XYZ(Math.Min(min.X, p2.X), Math.Min(min.Y, p2.Y), Math.Min(min.Z, p2.Z));
                                    max = new XYZ(Math.Max(max.X, p2.X), Math.Max(max.Y, p2.Y), Math.Max(max.Z, p2.Z));
                                }

                                // Project the min and max points onto the wall's location line to find the true width
                                double paramMin = line.Project(min).Parameter;
                                double paramMax = line.Project(max).Parameter;

                                double openingWidth = Math.Abs(paramMax - paramMin);

                                openings.Add(new WallOpeningInfo
                                {
                                    Type = OpeningType.EditProfile,
                                    BottomElevation = min.Z,
                                    TopElevation = max.Z,
                                    InsertPoint = (min + max) / 2.0,
                                    Width = openingWidth,
                                    Height = max.Z - min.Z
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
                System.Diagnostics.Debug.WriteLine($"Error getting wall openings: {ex.Message}");
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
                            // MODIFICATION: Treat EditProfile openings as rectangular openings
                            // by creating a standard opening from their bounding box extents.
                            case OpeningType.EditProfile:
                            case OpeningType.RectangularOpening:
                            case OpeningType.ArcOpening:
                                // Create new opening
                                CreateWallOpening(doc, newWall, openingInfo, wallBottomElev, wallTopElev);
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Failed to apply opening: {ex.Message}");
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

                // Calculate opening bounds relative to new wall
                double openingBottom = Math.Max(openingInfo.BottomElevation, wallBottomElev);
                double openingTop = Math.Min(openingInfo.TopElevation, wallTopElev);

                if (openingTop <= openingBottom) return;

                // Get opening width
                double openingWidth = openingInfo.Width;
                if (openingWidth <= 0.01) return; // Skip openings that are too small

                // Find the position along the wall
                XYZ wallDirection = wallLine.Direction;

                // Project opening center onto wall line
                IntersectionResult projection = wallLine.Project(openingInfo.InsertPoint);
                if (projection == null) return;

                XYZ projectedCenter = projection.XYZPoint;

                // Calculate the start and end points of the opening
                XYZ openingStart = projectedCenter - wallDirection * (openingWidth / 2);
                XYZ openingEnd = projectedCenter + wallDirection * (openingWidth / 2);

                // Create points for rectangular opening
                XYZ point1 = new XYZ(openingStart.X, openingStart.Y, openingBottom);
                XYZ point2 = new XYZ(openingEnd.X, openingEnd.Y, openingTop);

                // Create the opening using the correct method
                doc.Create.NewOpening(wall, point1, point2);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to create opening: {ex.Message}");
                // Don't throw - continue with other openings
            }
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
            // A small tolerance is used to handle floating-point inaccuracies
            const double tolerance = 0.001;

            for (int i = 0; i < walls.Count; i++)
            {
                Wall wall = walls[i];
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

                    // For the last wall segment, the top boundary is inclusive.
                    if (i == walls.Count - 1)
                    {
                        if (elevation >= baseElevation - tolerance && elevation <= topElevation + tolerance)
                        {
                            return wall;
                        }
                    }
                    // For all other segments, the top boundary is exclusive.
                    // This ensures an element on a boundary is hosted by the wall segment ABOVE it.
                    else
                    {
                        if (elevation >= baseElevation - tolerance && elevation < topElevation - tolerance)
                        {
                            return wall;
                        }
                    }
                }
                catch
                {
                    // If parameters can't be read for a segment, skip to the next one.
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