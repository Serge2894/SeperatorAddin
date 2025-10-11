using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using SeperatorAddin.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

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

                using (Transaction t = new Transaction(doc, "Split and Join Walls"))
                {
                    t.Start();

                    ProcessWallSeparation(doc, references);

                    t.Commit();
                }

                return Result.Succeeded;
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

        public void ProcessWallSeparation(Document doc, List<Reference> references)
        {
            if (references == null || !references.Any()) return;

            List<Wall> selectedWalls = references.Select(r => doc.GetElement(r)).OfType<Wall>().ToList();
            if (!selectedWalls.Any()) return;

            // Store connection information before processing
            Dictionary<ElementId, WallConnectionInfo> wallConnections = AnalyzeAllConnections(doc, selectedWalls);

            // Store the new walls created from each original wall
            Dictionary<ElementId, List<Wall>> originalToNewWalls = new Dictionary<ElementId, List<Wall>>();

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

            // Handle connections based on angle type with improved logic
            HandleAllConnectionsImproved(doc, originalToNewWalls, wallConnections);
        }


        #region Helper Classes

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

        private class LayerInfo
        {
            public Wall Wall { get; set; }
            public CompoundStructureLayer Layer { get; set; }
            public double CenterOffset { get; set; }
            public MaterialFunctionAssignment Function { get; set; }
            public ElementId MaterialId { get; set; }
            public string MaterialName { get; set; }
            public int LayerIndex { get; set; }
            public double Width { get; set; }
        }

        private class WallTypeConfig
        {
            public string MaterialName { get; set; }
            public MaterialFunctionAssignment Function { get; set; }
            public int Priority { get; set; } // Join priority
            public bool CanJoinDifferentMaterials { get; set; }
        }

        #endregion

        #region Connection Analysis

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

        #endregion

        #region Wall Processing

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

                int layerIndex = 0;
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

                    // Store layer information in extensible storage or as a parameter for later use
                    StoreLayerInfo(newWall, layerIndex, layerCenterOffset, layer);

                    newWalls.Add(newWall);
                    currentPosition += layerWidth;
                    layerIndex++;
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

        private void StoreLayerInfo(Wall wall, int layerIndex, double centerOffset, CompoundStructureLayer layer)
        {
            // Store layer information for later use
            // This could be done through extensible storage or custom parameters
            // For now, we'll rely on the wall type name to identify the layer
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

        #endregion

        #region Improved Connection Handling

        private void HandleAllConnectionsImproved(Document doc, Dictionary<ElementId, List<Wall>> originalToNewWalls,
            Dictionary<ElementId, WallConnectionInfo> wallConnections)
        {
            // Process each connection with improved logic
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

                    // Use intelligent layer matching for different wall types
                    if (connection.Type == ConnectionType.Corner)
                    {
                        HandleCornerConnectionImproved(doc, newWalls, connectedNewWalls, connection, connectionInfo);
                    }
                    else if (connection.Type == ConnectionType.TJunction)
                    {
                        HandleTJunctionConnectionImproved(doc, newWalls, connectedNewWalls, connection, connectionInfo);
                    }
                    else if (connection.Type == ConnectionType.Cross)
                    {
                        HandleCrossConnectionImproved(doc, newWalls, connectedNewWalls, connection, connectionInfo);
                    }
                }
            }

            // Join walls using improved logic
            JoinWallsIntelligentlyImproved(doc, originalToNewWalls);
        }

        private void HandleCornerConnectionImproved(Document doc, List<Wall> walls1, List<Wall> walls2,
            WallConnection connection, WallConnectionInfo info1)
        {
            // Get layer information for both wall groups
            List<LayerInfo> layers1 = GetWallLayersInfo(walls1);
            List<LayerInfo> layers2 = GetWallLayersInfo(walls2);

            // Match layers based on function and material
            Dictionary<Wall, List<Wall>> wallPairs = MatchWallLayers(layers1, layers2);

            // Process each matched pair
            foreach (var pair in wallPairs)
            {
                Wall wall1 = pair.Key;
                List<Wall> matchedWalls = pair.Value;

                foreach (Wall wall2 in matchedWalls)
                {
                    ExtendWallsToCornerImproved(doc, wall1, wall2, connection);
                }
            }
        }

        private void HandleTJunctionConnectionImproved(Document doc, List<Wall> walls1, List<Wall> walls2,
            WallConnection connection, WallConnectionInfo info1)
        {
            // Get layer information
            List<LayerInfo> layers1 = GetWallLayersInfo(walls1);
            List<LayerInfo> layers2 = GetWallLayersInfo(walls2);

            // For T-junctions, we need special handling
            foreach (var layer1 in layers1)
            {
                Wall wall1 = layer1.Wall;
                LocationCurve loc1 = wall1.Location as LocationCurve;
                if (loc1 == null) continue;

                Line line1 = loc1.Curve as Line;
                if (line1 == null) continue;

                // Find the best matching wall in the continuing wall
                Wall bestMatch = null;
                double bestScore = 0;

                foreach (var layer2 in layers2)
                {
                    double score = CalculateJoinScore(layer1, layer2);
                    if (score > bestScore)
                    {
                        bestScore = score;
                        bestMatch = layer2.Wall;
                    }
                }

                if (bestMatch != null && bestScore > 0.5) // Threshold for joining
                {
                    LocationCurve loc2 = bestMatch.Location as LocationCurve;
                    if (loc2 == null) continue;

                    Line line2 = loc2.Curve as Line;
                    if (line2 == null) continue;

                    // Extend wall1 to meet wall2
                    XYZ intersectionPoint = line2.Project(connection.ConnectionPoint).XYZPoint;

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
        }

        private void HandleCrossConnectionImproved(Document doc, List<Wall> walls1, List<Wall> walls2,
            WallConnection connection, WallConnectionInfo info1)
        {
            // For crossing walls, we typically don't modify the geometry
            // but we need to ensure they join properly at the intersection
            // This is handled in the JoinWallsIntelligentlyImproved method
        }

        #endregion

        #region Layer Matching and Analysis

        private List<LayerInfo> GetWallLayersInfo(List<Wall> walls)
        {
            List<LayerInfo> layerInfos = new List<LayerInfo>();

            for (int i = 0; i < walls.Count; i++)
            {
                Wall wall = walls[i];
                WallType wallType = wall.WallType;

                // Extract material name from wall type name
                string wallTypeName = wallType.Name;
                string materialName = wallTypeName.Replace("A-Wal-", "");

                // Get the compound structure to find the layer function
                CompoundStructure cs = wallType.GetCompoundStructure();
                if (cs != null && cs.LayerCount > 0)
                {
                    var layer = cs.GetLayers()[0]; // Single layer wall

                    layerInfos.Add(new LayerInfo
                    {
                        Wall = wall,
                        Layer = layer,
                        Function = layer.Function,
                        MaterialId = layer.MaterialId,
                        MaterialName = materialName,
                        CenterOffset = GetWallCenterOffset(wall, i, walls.Count),
                        LayerIndex = i,
                        Width = layer.Width
                    });
                }
            }

            return layerInfos;
        }

        private double GetWallCenterOffset(Wall wall, int layerIndex, int totalLayers)
        {
            // Calculate the offset from the original wall centerline
            // This is simplified - in reality, you'd calculate based on the original wall's structure
            WallType wallType = wall.WallType;
            double wallWidth = wallType.Width;

            // Assuming walls are created from center outward
            double totalWidth = 0;
            for (int i = 0; i < totalLayers; i++)
            {
                totalWidth += wallWidth; // Simplified - assumes all layers same width
            }

            double centerOffset = -totalWidth / 2 + layerIndex * wallWidth + wallWidth / 2;
            return centerOffset;
        }

        private Dictionary<Wall, List<Wall>> MatchWallLayers(List<LayerInfo> layers1, List<LayerInfo> layers2)
        {
            Dictionary<Wall, List<Wall>> matches = new Dictionary<Wall, List<Wall>>();

            foreach (LayerInfo layer1 in layers1)
            {
                matches[layer1.Wall] = new List<Wall>();

                foreach (LayerInfo layer2 in layers2)
                {
                    if (ShouldLayersJoin(layer1, layer2))
                    {
                        matches[layer1.Wall].Add(layer2.Wall);
                    }
                }
            }

            return matches;
        }

        private bool ShouldLayersJoin(LayerInfo layer1, LayerInfo layer2)
        {
            // Priority 1: Same material
            if (layer1.MaterialId == layer2.MaterialId && layer1.MaterialId != ElementId.InvalidElementId)
            {
                return true;
            }

            // Check by material name if IDs don't match
            if (layer1.MaterialName == layer2.MaterialName)
            {
                return true;
            }

            // Priority 2: Same function (Structure, Substrate, Finish, etc.)
            if (layer1.Function == layer2.Function)
            {
                // For structural layers, always join
                if (layer1.Function == MaterialFunctionAssignment.Structure)
                {
                    return true;
                }

                // For finish layers on the same side, join
                if (layer1.Function == MaterialFunctionAssignment.Finish1 ||
                    layer1.Function == MaterialFunctionAssignment.Finish2)
                {
                    // Check if they're on the same side based on layer index
                    return Math.Abs(layer1.LayerIndex - layer2.LayerIndex) <= 1;
                }
            }

            // Priority 3: Check configuration rules
            var config = GetWallTypeConfigurations();
            if (config.ContainsKey(layer1.MaterialName) && config.ContainsKey(layer2.MaterialName))
            {
                var config1 = config[layer1.MaterialName];
                var config2 = config[layer2.MaterialName];

                if (config1.CanJoinDifferentMaterials && config2.CanJoinDifferentMaterials &&
                    config1.Function == config2.Function)
                {
                    return true;
                }
            }

            return false;
        }

        private double CalculateJoinScore(LayerInfo layer1, LayerInfo layer2)
        {
            double score = 0;

            // Same material: highest score
            if (layer1.MaterialId == layer2.MaterialId && layer1.MaterialId != ElementId.InvalidElementId)
            {
                score += 1.0;
            }
            else if (layer1.MaterialName == layer2.MaterialName)
            {
                score += 0.9;
            }

            // Same function: medium score
            if (layer1.Function == layer2.Function)
            {
                score += 0.5;
            }

            // Layer index proximity
            int indexDiff = Math.Abs(layer1.LayerIndex - layer2.LayerIndex);
            if (indexDiff == 0)
            {
                score += 0.3;
            }
            else if (indexDiff == 1)
            {
                score += 0.1;
            }

            return score;
        }

        #endregion

        #region Configuration

        private Dictionary<string, WallTypeConfig> GetWallTypeConfigurations()
        {
            return new Dictionary<string, WallTypeConfig>
            {
                {"Concrete", new WallTypeConfig
                    {
                        MaterialName = "Concrete",
                        Function = MaterialFunctionAssignment.Structure,
                        Priority = 1,
                        CanJoinDifferentMaterials = true
                    }},
                {"CMU", new WallTypeConfig
                    {
                        MaterialName = "CMU",
                        Function = MaterialFunctionAssignment.Structure,
                        Priority = 1,
                        CanJoinDifferentMaterials = true
                    }},
                {"Masonry", new WallTypeConfig
                    {
                        MaterialName = "Masonry",
                        Function = MaterialFunctionAssignment.Structure,
                        Priority = 1,
                        CanJoinDifferentMaterials = true
                    }},
                {"Gypsum Wall Board", new WallTypeConfig
                    {
                        MaterialName = "Gypsum Wall Board",
                        Function = MaterialFunctionAssignment.Finish1,
                        Priority = 3,
                        CanJoinDifferentMaterials = false
                    }},
                {"Insulation", new WallTypeConfig
                    {
                        MaterialName = "Insulation",
                        Function = MaterialFunctionAssignment.Insulation,
                        Priority = 2,
                        CanJoinDifferentMaterials = false
                    }},
                {"Metal - Stud Layer", new WallTypeConfig
                    {
                        MaterialName = "Metal - Stud Layer",
                        Function = MaterialFunctionAssignment.Structure,
                        Priority = 1,
                        CanJoinDifferentMaterials = true
                    }}
            };
        }

        #endregion

        #region Geometry Operations

        private void ExtendWallsToCornerImproved(Document doc, Wall wall1, Wall wall2, WallConnection connection)
        {
            LocationCurve loc1 = wall1.Location as LocationCurve;
            LocationCurve loc2 = wall2.Location as LocationCurve;

            if (loc1 == null || loc2 == null) return;

            Line line1 = loc1.Curve as Line;
            Line line2 = loc2.Curve as Line;

            if (line1 == null || line2 == null) return;

            // Calculate the intersection point considering wall thickness
            XYZ intersectionPoint = CalculateAdjustedIntersection(wall1, wall2, line1, line2, connection);

            if (intersectionPoint == null) return;

            // Update wall endpoints
            UpdateWallEndpoint(loc1, line1, intersectionPoint, connection.IsAtStart);
            UpdateWallEndpoint(loc2, line2, intersectionPoint, connection.IsAtOtherStart);
        }

        private XYZ CalculateAdjustedIntersection(Wall wall1, Wall wall2, Line line1, Line line2, WallConnection connection)
        {
            // Get wall thicknesses
            double thickness1 = wall1.WallType.Width;
            double thickness2 = wall2.WallType.Width;

            // Find base intersection
            XYZ baseIntersection = FindExtendedIntersection(line1, line2);
            if (baseIntersection == null) return null;

            // For different wall types, we might need to adjust the intersection point
            // based on how the layers should align

            // This is a simplified version - you might need more complex logic
            // based on your specific requirements
            return baseIntersection;
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

        private void UpdateWallEndpoint(LocationCurve loc, Line line, XYZ newPoint, bool updateStart)
        {
            if (updateStart)
            {
                Line newLine = Line.CreateBound(newPoint, line.GetEndPoint(1));
                loc.Curve = newLine;
            }
            else
            {
                Line newLine = Line.CreateBound(line.GetEndPoint(0), newPoint);
                loc.Curve = newLine;
            }
        }

        #endregion

        #region Wall Joining

        private void JoinWallsIntelligentlyImproved(Document doc, Dictionary<ElementId, List<Wall>> originalToNewWalls)
        {
            // Get all new walls
            List<Wall> allNewWalls = originalToNewWalls.Values.SelectMany(list => list).ToList();

            // Create layer info for all walls
            Dictionary<Wall, LayerInfo> wallLayerInfo = new Dictionary<Wall, LayerInfo>();
            foreach (var wallList in originalToNewWalls.Values)
            {
                var layerInfos = GetWallLayersInfo(wallList);
                foreach (var info in layerInfos)
                {
                    wallLayerInfo[info.Wall] = info;
                }
            }

            // Join walls that came from the same original wall first
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
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Failed to join walls from same original: {ex.Message}");
                        }
                    }
                }
            }

            // Join walls based on layer compatibility and geometry
            for (int i = 0; i < allNewWalls.Count - 1; i++)
            {
                for (int j = i + 1; j < allNewWalls.Count; j++)
                {
                    Wall wall1 = allNewWalls[i];
                    Wall wall2 = allNewWalls[j];

                    if (wall1.Id == wall2.Id) continue;
                    if (JoinGeometryUtils.AreElementsJoined(doc, wall1, wall2)) continue;

                    // Check if walls should be joined based on geometry and layer compatibility
                    if (ShouldWallsBeJoinedImproved(wall1, wall2, wallLayerInfo))
                    {
                        try
                        {
                            JoinGeometryUtils.JoinGeometry(doc, wall1, wall2);
                            Debug.WriteLine($"Successfully joined walls: {wall1.Id} and {wall2.Id}");
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Failed to join walls {wall1.Id} and {wall2.Id}: {ex.Message}");
                        }
                    }
                }
            }
        }

        private bool ShouldWallsBeJoinedImproved(Wall wall1, Wall wall2, Dictionary<Wall, LayerInfo> wallLayerInfo)
        {
            // First check geometric conditions
            if (!AreWallsGeometricallyAdjacent(wall1, wall2))
            {
                return false;
            }

            // Then check layer compatibility
            if (wallLayerInfo.ContainsKey(wall1) && wallLayerInfo.ContainsKey(wall2))
            {
                LayerInfo info1 = wallLayerInfo[wall1];
                LayerInfo info2 = wallLayerInfo[wall2];

                return ShouldLayersJoin(info1, info2);
            }

            // If we can't determine layer info, check by wall type configuration
            return ShouldWallsJoinByConfiguration(wall1, wall2);
        }

        private bool ShouldWallsJoinByConfiguration(Wall wall1, Wall wall2)
        {
            var configs = GetWallTypeConfigurations();

            string material1 = GetWallMaterialName(wall1);
            string material2 = GetWallMaterialName(wall2);

            if (!configs.ContainsKey(material1) || !configs.ContainsKey(material2))
                return false;

            var config1 = configs[material1];
            var config2 = configs[material2];

            // Same material always joins
            if (material1 == material2)
                return true;

            // Check if materials can join different types
            if (!config1.CanJoinDifferentMaterials || !config2.CanJoinDifferentMaterials)
                return false;

            // Check function compatibility
            if (config1.Function == config2.Function)
                return true;

            // Both structural elements can join
            if (config1.Function == MaterialFunctionAssignment.Structure &&
                config2.Function == MaterialFunctionAssignment.Structure)
                return true;

            return false;
        }

        private string GetWallMaterialName(Wall wall)
        {
            string typeName = wall.WallType.Name;
            if (typeName.StartsWith("A-Wal-"))
            {
                return typeName.Replace("A-Wal-", "");
            }

            // Try to get from compound structure
            CompoundStructure cs = wall.WallType.GetCompoundStructure();
            if (cs != null && cs.LayerCount > 0)
            {
                var layer = cs.GetLayers()[0];
                Element mat = wall.Document.GetElement(layer.MaterialId);
                if (mat != null)
                    return mat.Name;
            }

            return "Unknown";
        }

        private bool AreWallsGeometricallyAdjacent(Wall wall1, Wall wall2)
        {
            LocationCurve loc1 = wall1.Location as LocationCurve;
            LocationCurve loc2 = wall2.Location as LocationCurve;

            if (loc1 == null || loc2 == null) return false;

            Line line1 = loc1.Curve as Line;
            Line line2 = loc2.Curve as Line;

            if (line1 == null || line2 == null) return false;

            // Check if walls intersect
            IntersectionResultArray results;
            SetComparisonResult result = line1.Intersect(line2, out results);

            if (result == SetComparisonResult.Overlap)
            {
                return true;
            }

            // Check if endpoints are close
            double tolerance = 0.01; // 1/8" in feet
            XYZ[] endpoints1 = { line1.GetEndPoint(0), line1.GetEndPoint(1) };
            XYZ[] endpoints2 = { line2.GetEndPoint(0), line2.GetEndPoint(1) };

            foreach (XYZ pt1 in endpoints1)
            {
                foreach (XYZ pt2 in endpoints2)
                {
                    if (pt1.DistanceTo(pt2) < tolerance)
                    {
                        return true;
                    }
                }
            }

            // Check if walls are parallel and close
            if (AreWallsParallelAndClose(line1, line2, wall1.WallType.Width, wall2.WallType.Width))
            {
                return true;
            }

            return false;
        }

        private bool AreWallsParallelAndClose(Line line1, Line line2, double width1, double width2)
        {
            // Check if lines are parallel
            double dotProduct = Math.Abs(line1.Direction.DotProduct(line2.Direction));
            if (dotProduct < 0.99) return false; // Not parallel enough

            // Check distance between lines
            XYZ pointOnLine1 = line1.GetEndPoint(0);
            XYZ projection = line2.Project(pointOnLine1).XYZPoint;
            double distance = pointOnLine1.DistanceTo(projection);

            // Maximum distance should be sum of half-widths plus small tolerance
            double maxDistance = (width1 + width2) / 2 + 0.01;

            return distance < maxDistance;
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

        #endregion

        #region Utility Methods

        private void LogJoinAttempt(Wall wall1, Wall wall2, bool success, string reason)
        {
            string msg = $"Join {wall1.Id} to {wall2.Id}: {(success ? "Success" : "Failed")} - {reason}";
            Debug.WriteLine(msg);
            // Could also write to a file or show in UI
        }

        #endregion
    }
}