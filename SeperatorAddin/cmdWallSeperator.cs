using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using SeperatorAddin.Common;
using System.Collections.Generic;
using System.Linq;
using static SeperatorAddin.Common.Utils;

namespace SeperatorAddin
{
    [Transaction(TransactionMode.Manual)]
    public class cmdWallSeperator : IExternalCommand
    {
        // Tolerance for geometric comparisons
        private const double TOLERANCE = 0.001; // 1/16" in feet
        private const double ANGLE_TOLERANCE = 0.01; // Radians
        private const double MAX_EXTENSION = 5.0; // Maximum extension distance in feet

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Get selected walls
                Utils.WallSelectionilter filter = new Utils.WallSelectionilter();
                List<Reference> references = uidoc.Selection.PickObjects(ObjectType.Element, filter).ToList();

                List<Wall> selectedWalls = new List<Wall>();
                foreach (Reference reference in references)
                {
                    Wall wall = doc.GetElement(reference.ElementId) as Wall;
                    selectedWalls.Add(wall);
                }

                // Store connection information before processing
                Dictionary<ElementId, WallConnectionInfo> wallConnections = AnalyzeAllConnections(doc, selectedWalls);

                // Store the new walls created from each original wall
                Dictionary<ElementId, List<Wall>> originalToNewWalls = new Dictionary<ElementId, List<Wall>>();

                using (Transaction t = new Transaction(doc, "Split and Join Walls"))
                {
                    t.Start();

                    // Unjoin all selected walls from each other
                    UnjoinAllWalls(doc, selectedWalls);

                    // Process each wall and create separated layers
                    foreach (Wall wall in selectedWalls)
                    {
                        List<Wall> newWalls = ProcessWall(doc, wall, wallConnections[wall.Id]);
                        if (newWalls.Count > 0)
                        {
                            originalToNewWalls[wall.Id] = newWalls;
                        }
                    }

                    // Delete original walls
                    foreach (Wall wall in selectedWalls)
                    {
                        doc.Delete(wall.Id);
                    }

                    // Handle connections based on angle type
                    HandleAllConnections(doc, originalToNewWalls, wallConnections);

                    t.Commit();
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private class WallConnectionInfo
        {
            public List<WallConnection> Connections { get; set; } = new List<WallConnection>();
            public XYZ StartPoint { get; set; }
            public XYZ EndPoint { get; set; }
            public Line LocationLine { get; set; }
            public XYZ Direction { get; set; }
            public double Width { get; set; }
        }

        private class WallConnection
        {
            public ElementId ConnectedWallId { get; set; }
            public ConnectionType Type { get; set; }
            public XYZ ConnectionPoint { get; set; }
            public double Angle { get; set; } // Angle between walls
            public bool IsAtStart { get; set; } // Connection at start or end of current wall
            public bool IsAtOtherStart { get; set; } // Connection at start or end of other wall
        }

        private enum ConnectionType
        {
            Corner,      // End-to-end connection
            TJunction,   // End-to-middle connection
            Cross        // Middle-to-middle connection
        }

        private Dictionary<ElementId, WallConnectionInfo> AnalyzeAllConnections(Document doc, List<Wall> walls)
        {
            var connections = new Dictionary<ElementId, WallConnectionInfo>();

            foreach (Wall wall in walls)
            {
                connections[wall.Id] = AnalyzeWallConnections(doc, wall, walls);
            }

            return connections;
        }

        private WallConnectionInfo AnalyzeWallConnections(Document doc, Wall wall, List<Wall> allWalls)
        {
            WallConnectionInfo info = new WallConnectionInfo();

            LocationCurve locCurve = wall.Location as LocationCurve;
            if (locCurve == null) return info;

            Line line = locCurve.Curve as Line;
            if (line == null) return info;

            info.LocationLine = line;
            info.StartPoint = line.GetEndPoint(0);
            info.EndPoint = line.GetEndPoint(1);
            info.Direction = line.Direction;
            info.Width = wall.WallType.Width;

            foreach (Wall otherWall in allWalls)
            {
                if (otherWall.Id == wall.Id) continue;

                LocationCurve otherLocCurve = otherWall.Location as LocationCurve;
                if (otherLocCurve == null) continue;

                Line otherLine = otherLocCurve.Curve as Line;
                if (otherLine == null) continue;

                // Check for various connection types
                WallConnection connection = DetectConnection(line, otherLine, wall.Id, otherWall.Id);
                if (connection != null)
                {
                    info.Connections.Add(connection);
                }
            }

            return info;
        }

        private WallConnection DetectConnection(Line line1, Line line2, ElementId wallId1, ElementId wallId2)
        {
            XYZ start1 = line1.GetEndPoint(0);
            XYZ end1 = line1.GetEndPoint(1);
            XYZ start2 = line2.GetEndPoint(0);
            XYZ end2 = line2.GetEndPoint(1);

            // Check end-to-end connections (corners)
            if (start1.DistanceTo(start2) < TOLERANCE)
            {
                return CreateConnection(wallId2, ConnectionType.Corner, start1, line1.Direction, line2.Direction, true, true);
            }
            if (start1.DistanceTo(end2) < TOLERANCE)
            {
                return CreateConnection(wallId2, ConnectionType.Corner, start1, line1.Direction, -line2.Direction, true, false);
            }
            if (end1.DistanceTo(start2) < TOLERANCE)
            {
                return CreateConnection(wallId2, ConnectionType.Corner, end1, -line1.Direction, line2.Direction, false, true);
            }
            if (end1.DistanceTo(end2) < TOLERANCE)
            {
                return CreateConnection(wallId2, ConnectionType.Corner, end1, -line1.Direction, -line2.Direction, false, false);
            }

            // Check T-junctions and crosses
            // Check if line1 endpoints are on line2
            double param;
            if (IsPointOnLine(start1, line2, out param) && param > TOLERANCE && param < 1.0 - TOLERANCE)
            {
                return CreateConnection(wallId2, ConnectionType.TJunction, start1, line1.Direction, line2.Direction, true, false);
            }
            if (IsPointOnLine(end1, line2, out param) && param > TOLERANCE && param < 1.0 - TOLERANCE)
            {
                return CreateConnection(wallId2, ConnectionType.TJunction, end1, -line1.Direction, line2.Direction, false, false);
            }

            // Check if line2 endpoints are on line1
            if (IsPointOnLine(start2, line1, out param) && param > TOLERANCE && param < 1.0 - TOLERANCE)
            {
                return CreateConnection(wallId2, ConnectionType.TJunction, start2, line1.Direction, line2.Direction, false, true);
            }
            if (IsPointOnLine(end2, line1, out param) && param > TOLERANCE && param < 1.0 - TOLERANCE)
            {
                return CreateConnection(wallId2, ConnectionType.TJunction, end2, line1.Direction, -line2.Direction, false, false);
            }

            // Check for crossing walls
            IntersectionResultArray results;
            if (line1.Intersect(line2, out results) == SetComparisonResult.Overlap)
            {
                if (results.Size > 0)
                {
                    XYZ intersectionPoint = results.get_Item(0).XYZPoint;
                    double param1 = line1.Project(intersectionPoint).Parameter;
                    double param2 = line2.Project(intersectionPoint).Parameter;

                    if (param1 > TOLERANCE && param1 < 1.0 - TOLERANCE &&
                        param2 > TOLERANCE && param2 < 1.0 - TOLERANCE)
                    {
                        return CreateConnection(wallId2, ConnectionType.Cross, intersectionPoint,
                            line1.Direction, line2.Direction, false, false);
                    }
                }
            }

            return null;
        }

        private WallConnection CreateConnection(ElementId connectedWallId, ConnectionType type,
            XYZ connectionPoint, XYZ dir1, XYZ dir2, bool isAtStart, bool isAtOtherStart)
        {
            double angle = Math.Acos(Math.Abs(dir1.DotProduct(dir2)));

            return new WallConnection
            {
                ConnectedWallId = connectedWallId,
                Type = type,
                ConnectionPoint = connectionPoint,
                Angle = angle,
                IsAtStart = isAtStart,
                IsAtOtherStart = isAtOtherStart
            };
        }

        private bool IsPointOnLine(XYZ point, Line line, out double parameter)
        {
            parameter = 0;
            XYZ projection = line.Project(point).XYZPoint;

            if (point.DistanceTo(projection) < TOLERANCE)
            {
                parameter = line.Project(point).Parameter;
                return true;
            }

            return false;
        }

        private void HandleAllConnections(Document doc, Dictionary<ElementId, List<Wall>> originalToNewWalls,
            Dictionary<ElementId, WallConnectionInfo> wallConnections)
        {
            // Process each connection
            foreach (var kvp in wallConnections)
            {
                ElementId originalWallId = kvp.Key;
                WallConnectionInfo connectionInfo = kvp.Value;

                if (!originalToNewWalls.ContainsKey(originalWallId)) continue;

                List<Wall> newWalls = originalToNewWalls[originalWallId];

                foreach (WallConnection connection in connectionInfo.Connections)
                {
                    if (!originalToNewWalls.ContainsKey(connection.ConnectedWallId)) continue;

                    List<Wall> connectedNewWalls = originalToNewWalls[connection.ConnectedWallId];

                    // Handle based on connection type and angle
                    if (connection.Type == ConnectionType.Corner)
                    {
                        HandleCornerConnection(doc, newWalls, connectedNewWalls, connection, connectionInfo);
                    }
                    else if (connection.Type == ConnectionType.TJunction)
                    {
                        HandleTJunctionConnection(doc, newWalls, connectedNewWalls, connection, connectionInfo);
                    }
                    else if (connection.Type == ConnectionType.Cross)
                    {
                        HandleCrossConnection(doc, newWalls, connectedNewWalls, connection, connectionInfo);
                    }
                }
            }

            // Join all walls that should be joined
            JoinWallsIntelligently(doc, originalToNewWalls);
        }

        private void HandleCornerConnection(Document doc, List<Wall> walls1, List<Wall> walls2,
            WallConnection connection, WallConnectionInfo info1)
        {
            // Determine join type based on angle
            double angle = connection.Angle;

            if (Math.Abs(angle - Math.PI / 2) < ANGLE_TOLERANCE)
            {
                // 90-degree corner
                ExtendWallsToCorner(doc, walls1, walls2, connection, false);
            }
            else if (angle < Math.PI / 2)
            {
                // Acute angle - use miter join
                ExtendWallsToCorner(doc, walls1, walls2, connection, true);
            }
            else
            {
                // Obtuse angle - extend walls
                ExtendWallsToCorner(doc, walls1, walls2, connection, false);
            }
        }

        private void ExtendWallsToCorner(Document doc, List<Wall> walls1, List<Wall> walls2,
            WallConnection connection, bool useMiter)
        {
            // Process each layer
            for (int i = 0; i < Math.Min(walls1.Count, walls2.Count); i++)
            {
                Wall wall1 = walls1[i];
                Wall wall2 = walls2[i];

                LocationCurve loc1 = wall1.Location as LocationCurve;
                LocationCurve loc2 = wall2.Location as LocationCurve;

                if (loc1 == null || loc2 == null) continue;

                Line line1 = loc1.Curve as Line;
                Line line2 = loc2.Curve as Line;

                if (line1 == null || line2 == null) continue;

                // Find intersection point
                XYZ intersectionPoint = FindExtendedIntersection(line1, line2);

                if (intersectionPoint == null) continue;

                // Update wall endpoints
                if (connection.IsAtStart)
                {
                    Line newLine1 = Line.CreateBound(intersectionPoint, line1.GetEndPoint(1));
                    loc1.Curve = newLine1;
                }
                else
                {
                    Line newLine1 = Line.CreateBound(line1.GetEndPoint(0), intersectionPoint);
                    loc1.Curve = newLine1;
                }

                if (connection.IsAtOtherStart)
                {
                    Line newLine2 = Line.CreateBound(intersectionPoint, line2.GetEndPoint(1));
                    loc2.Curve = newLine2;
                }
                else
                {
                    Line newLine2 = Line.CreateBound(line2.GetEndPoint(0), intersectionPoint);
                    loc2.Curve = newLine2;
                }
            }
        }

        private void HandleTJunctionConnection(Document doc, List<Wall> walls1, List<Wall> walls2,
            WallConnection connection, WallConnectionInfo info1)
        {
            // For T-junctions, we need to:
            // 1. Extend the terminating wall to meet the continuing wall
            // 2. Split the continuing wall at the junction point if needed

            for (int i = 0; i < Math.Min(walls1.Count, walls2.Count); i++)
            {
                Wall wall1 = walls1[i];
                Wall wall2 = walls2[i];

                LocationCurve loc1 = wall1.Location as LocationCurve;
                LocationCurve loc2 = wall2.Location as LocationCurve;

                if (loc1 == null || loc2 == null) continue;

                Line line1 = loc1.Curve as Line;
                Line line2 = loc2.Curve as Line;

                if (line1 == null || line2 == null) continue;

                // Find the exact intersection point on wall2
                XYZ intersectionPoint = line2.Project(connection.ConnectionPoint).XYZPoint;

                // Extend wall1 to meet wall2
                if (connection.IsAtStart)
                {
                    Line newLine1 = Line.CreateBound(intersectionPoint, line1.GetEndPoint(1));
                    loc1.Curve = newLine1;
                }
                else
                {
                    Line newLine1 = Line.CreateBound(line1.GetEndPoint(0), intersectionPoint);
                    loc1.Curve = newLine1;
                }
            }
        }

        private void HandleCrossConnection(Document doc, List<Wall> walls1, List<Wall> walls2,
            WallConnection connection, WallConnectionInfo info1)
        {
            // For crossing walls, we typically don't modify the geometry
            // but we need to ensure they join properly at the intersection
            // This is handled in the JoinWallsIntelligently method
        }

        private XYZ FindExtendedIntersection(Line line1, Line line2)
        {
            // Create unbounded lines for intersection
            Line unboundLine1 = line1.Clone() as Line;
            Line unboundLine2 = line2.Clone() as Line;

            unboundLine1.MakeUnbound();
            unboundLine2.MakeUnbound();

            IntersectionResultArray results;
            SetComparisonResult result = unboundLine1.Intersect(unboundLine2, out results);

            if (result == SetComparisonResult.Overlap && results.Size > 0)
            {
                return results.get_Item(0).XYZPoint;
            }

            return null;
        }

        private void JoinWallsIntelligently(Document doc, Dictionary<ElementId, List<Wall>> originalToNewWalls)
        {
            List<Wall> allNewWalls = originalToNewWalls.Values.SelectMany(list => list).ToList();

            // Join walls within each group (same original wall)
            foreach (var wallList in originalToNewWalls.Values)
            {
                for (int i = 0; i < wallList.Count - 1; i++)
                {
                    for (int j = i + 1; j < wallList.Count; j++)
                    {
                        try
                        {
                            if (!JoinGeometryUtils.AreElementsJoined(doc, wallList[i], wallList[j]))
                            {
                                JoinGeometryUtils.JoinGeometry(doc, wallList[i], wallList[j]);
                            }
                        }
                        catch { }
                    }
                }
            }

            // Join walls at intersections
            foreach (Wall wall1 in allNewWalls)
            {
                foreach (Wall wall2 in allNewWalls)
                {
                    if (wall1.Id == wall2.Id) continue;
                    if (JoinGeometryUtils.AreElementsJoined(doc, wall1, wall2)) continue;

                    // Check if walls should be joined
                    if (ShouldWallsBeJoined(wall1, wall2))
                    {
                        try
                        {
                            JoinGeometryUtils.JoinGeometry(doc, wall1, wall2);
                        }
                        catch { }
                    }
                }
            }
        }

        private bool ShouldWallsBeJoined(Wall wall1, Wall wall2)
        {
            LocationCurve loc1 = wall1.Location as LocationCurve;
            LocationCurve loc2 = wall2.Location as LocationCurve;

            if (loc1 == null || loc2 == null) return false;

            Line line1 = loc1.Curve as Line;
            Line line2 = loc2.Curve as Line;

            if (line1 == null || line2 == null) return false;

            // Check if walls intersect or touch
            IntersectionResultArray results;
            SetComparisonResult result = line1.Intersect(line2, out results);

            if (result == SetComparisonResult.Overlap)
            {
                return true;
            }

            // Check if endpoints are close
            XYZ[] endpoints1 = { line1.GetEndPoint(0), line1.GetEndPoint(1) };
            XYZ[] endpoints2 = { line2.GetEndPoint(0), line2.GetEndPoint(1) };

            foreach (XYZ pt1 in endpoints1)
            {
                foreach (XYZ pt2 in endpoints2)
                {
                    if (pt1.DistanceTo(pt2) < TOLERANCE)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private void UnjoinAllWalls(Document doc, List<Wall> walls)
        {
            foreach (Wall wall1 in walls)
            {
                foreach (Wall wall2 in walls)
                {
                    if (wall1.Id == wall2.Id) continue;
                    try
                    {
                        if (JoinGeometryUtils.AreElementsJoined(doc, wall1, wall2))
                        {
                            JoinGeometryUtils.UnjoinGeometry(doc, wall1, wall2);
                        }
                    }
                    catch { }
                }
            }
        }

        private List<Wall> ProcessWall(Document doc, Wall wall, WallConnectionInfo connectionInfo)
        {
            List<Wall> newWalls = new List<Wall>();

            try
            {
                WallType wallType = wall.WallType;
                CompoundStructure compoundStructure = wallType.GetCompoundStructure();
                if (compoundStructure == null) return newWalls;

                List<CompoundStructureLayer> layers = compoundStructure.GetLayers().ToList();

                Level baseLevel = doc.GetElement(wall.LevelId) as Level;
                double baseOffset = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET).AsDouble();
                double topOffset = wall.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET).AsDouble();
                double height = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).AsDouble();

                double totalWidth = wallType.Width;
                double currentPosition = -totalWidth / 2;

                LocationCurve locCurve = wall.Location as LocationCurve;
                Line originalLine = locCurve.Curve as Line;

                foreach (CompoundStructureLayer layer in layers)
                {
                    double layerWidth = layer.Width;
                    double layerCenterOffset = currentPosition + layerWidth / 2;

                    string materialName = GetLayerMaterialName(doc, layer);
                    WallType newWallType = GetOrCreateWallType(doc, wallType, materialName, layer);

                    XYZ offsetDir = originalLine.Direction.CrossProduct(XYZ.BasisZ).Normalize();
                    XYZ startPoint = originalLine.GetEndPoint(0) + offsetDir * layerCenterOffset;
                    XYZ endPoint = originalLine.GetEndPoint(1) + offsetDir * layerCenterOffset;
                    Line newLine = Line.CreateBound(startPoint, endPoint);

                    Wall newWall = Wall.Create(doc, newLine, newWallType.Id, baseLevel.Id, height, baseOffset, false, false);

                    Parameter topOffsetParam = newWall.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET);
                    if (topOffsetParam != null && !topOffsetParam.IsReadOnly)
                    {
                        topOffsetParam.Set(topOffset);
                    }

                    CopyWallParameters(wall, newWall);
                    newWalls.Add(newWall);
                    currentPosition += layerWidth;
                }

                // Remove hosted elements from all but the first wall
                for (int i = 1; i < newWalls.Count; i++)
                {
                    RemoveHostedElements(doc, newWalls[i]);
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to process wall: {ex.Message}");
            }

            return newWalls;
        }

        private string GetLayerMaterialName(Document doc, CompoundStructureLayer layer)
        {
            Element material = doc.GetElement(layer.MaterialId);
            return material != null ? material.Name : "Generic";
        }

        private WallType GetOrCreateWallType(Document doc, WallType originalWallType, string materialName, CompoundStructureLayer layer)
        {
            string newWallTypeName = $"A-Wal-{materialName}";

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            WallType existingType = collector.OfClass(typeof(WallType))
                .Cast<WallType>()
                .FirstOrDefault(wt => wt.Name == newWallTypeName);

            if (existingType != null)
                return existingType;

            WallType newWallType = originalWallType.Duplicate(newWallTypeName) as WallType;

            List<CompoundStructureLayer> listLayer = new List<CompoundStructureLayer>() { layer };
            CompoundStructure newStructure = CompoundStructure.CreateSimpleCompoundStructure(listLayer);
            newWallType.SetCompoundStructure(newStructure);

            return newWallType;
        }

        private void CopyWallParameters(Wall sourceWall, Wall targetWall)
        {
            try
            {
                Parameter structParam = sourceWall.get_Parameter(BuiltInParameter.WALL_STRUCTURAL_USAGE_PARAM);
                if (structParam != null && !structParam.IsReadOnly)
                {
                    Parameter targetStructParam = targetWall.get_Parameter(BuiltInParameter.WALL_STRUCTURAL_USAGE_PARAM);
                    if (targetStructParam != null && !targetStructParam.IsReadOnly)
                    {
                        targetStructParam.Set(structParam.AsInteger());
                    }
                }

                Parameter roomBoundingParam = sourceWall.get_Parameter(BuiltInParameter.WALL_ATTR_ROOM_BOUNDING);
                if (roomBoundingParam != null)
                {
                    Parameter targetRoomBoundingParam = targetWall.get_Parameter(BuiltInParameter.WALL_ATTR_ROOM_BOUNDING);
                    if (targetRoomBoundingParam != null && !targetRoomBoundingParam.IsReadOnly)
                    {
                        targetRoomBoundingParam.Set(roomBoundingParam.AsInteger());
                    }
                }
            }
            catch { }
        }

        private void RemoveHostedElements(Document doc, Wall wall)
        {
            try
            {
                List<ElementId> dependentIds = wall.GetDependentElements(null).ToList();

                foreach (ElementId elId in dependentIds)
                {
                    Element el = doc.GetElement(elId);

                    if (el is FamilyInstance fi && fi.Host?.Id == wall.Id)
                    {
                        doc.Delete(el.Id);
                    }
                }
            }
            catch { }
        }

        internal static PushButtonData GetButtonData()
        {
            string buttonInternalName = "btnCommand1";
            string buttonTitle = "Split Walls";

            Common.ButtonDataClass myButtonData = new Common.ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "Splits multi-layer walls into separate walls for each layer with proper joining");

            return myButtonData.Data;
        }
    }
}