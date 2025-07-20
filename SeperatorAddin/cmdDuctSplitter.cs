using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI.Selection;
using SeperatorAddin.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace SeperatorAddin
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class cmdDuctSplitter : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            try
            {
                // Select a duct
                Reference ductRef = uiDoc.Selection.PickObject(
                    ObjectType.Element,
                    new DuctSelectionFilter(),
                    "Select a duct to split");

                Duct selectedDuct = doc.GetElement(ductRef) as Duct;

                // Select ONE model line
                Reference lineRef = uiDoc.Selection.PickObject(
                    ObjectType.Element,
                    new ModelCurveSelectionFilter(),
                    "Select ONE straight model line that crosses the duct");

                ModelCurve modelLine = doc.GetElement(lineRef) as ModelCurve;
                Line splitLine = modelLine.GeometryCurve as Line;

                if (splitLine == null)
                {
                    TaskDialog.Show("Error", "Please select a straight line.");
                    return Result.Failed;
                }

                using (Transaction trans = new Transaction(doc, "Split Duct"))
                {
                    trans.Start();

                    try
                    {
                        // Store duct properties before splitting
                        DuctProperties ductProps = GetDuctProperties(selectedDuct);

                        // Get the duct's location curve
                        LocationCurve locCurve = selectedDuct.Location as LocationCurve;
                        if (locCurve == null)
                        {
                            trans.RollBack();
                            TaskDialog.Show("Error", "Could not get duct location.");
                            return Result.Failed;
                        }

                        Line ductLine = locCurve.Curve as Line;
                        if (ductLine == null)
                        {
                            trans.RollBack();
                            TaskDialog.Show("Error", "This tool only works with straight ducts.");
                            return Result.Failed;
                        }

                        // Find intersection point
                        XYZ intersectionPoint = FindIntersectionPoint(ductLine, splitLine);
                        if (intersectionPoint == null)
                        {
                            trans.RollBack();
                            TaskDialog.Show("Error",
                                "The split line does not intersect with the duct in plan view.\n\n" +
                                "Make sure the model line crosses over the duct when viewed from above (plan view).");
                            return Result.Failed;
                        }

                        // Check if intersection point is valid (not too close to ends)
                        double distToStart = intersectionPoint.DistanceTo(ductLine.GetEndPoint(0));
                        double distToEnd = intersectionPoint.DistanceTo(ductLine.GetEndPoint(1));
                        double minDistance = 0.5; // Minimum 6 inches from ends

                        if (distToStart < minDistance || distToEnd < minDistance)
                        {
                            trans.RollBack();
                            TaskDialog.Show("Error", "Split point is too close to duct ends. Please select a point at least 6 inches from the ends.");
                            return Result.Failed;
                        }

                        // Get connected elements before splitting
                        ConnectorSet connectors = selectedDuct.ConnectorManager.Connectors;
                        Dictionary<int, ConnectedElement> connectedElements = GetConnectedElements(connectors);

                        // Check if duct has insulation BEFORE creating new ducts
                        List<ElementId> insulationIds = GetDuctInsulation(doc, selectedDuct);
                        DuctInsulation originalInsulation = null;
                        ElementId insulationTypeId = ElementId.InvalidElementId;
                        double insulationThickness = 0;

                        if (insulationIds.Count > 0)
                        {
                            originalInsulation = doc.GetElement(insulationIds[0]) as DuctInsulation;
                            if (originalInsulation != null)
                            {
                                insulationTypeId = originalInsulation.GetTypeId();
                                Parameter thicknessParam = originalInsulation.get_Parameter(BuiltInParameter.RBS_INSULATION_THICKNESS);
                                if (thicknessParam != null && thicknessParam.HasValue)
                                {
                                    insulationThickness = thicknessParam.AsDouble();
                                }

                                // Debug information
                                TaskDialog.Show("Debug", $"Found insulation:\nType ID: {insulationTypeId}\nThickness: {insulationThickness}");
                            }
                        }
                        else
                        {
                            TaskDialog.Show("Debug", "No insulation found on the original duct");
                        }

                        // Create two new ducts
                        Duct duct1 = CreateDuctSegment(doc, ductLine.GetEndPoint(0), intersectionPoint, ductProps);
                        Duct duct2 = CreateDuctSegment(doc, intersectionPoint, ductLine.GetEndPoint(1), ductProps);

                        if (duct1 == null || duct2 == null)
                        {
                            trans.RollBack();
                            TaskDialog.Show("Error", "Failed to create new duct segments.");
                            return Result.Failed;
                        }

                        // Copy parameters from original duct
                        CopyDuctParameters(selectedDuct, duct1);
                        CopyDuctParameters(selectedDuct, duct2);

                        // Reconnect to original connected elements
                        ReconnectElements(doc, duct1, duct2, connectedElements, intersectionPoint);

                        // Connect the two new ducts at the split point
                        ConnectDucts(doc, duct1, duct2, intersectionPoint);

                        // Delete original duct (this will also delete its insulation)
                        doc.Delete(selectedDuct.Id);

                        // Force regeneration to ensure ducts are fully created
                        doc.Regenerate();

                        // Add insulation to new ducts if the original had insulation
                        if (insulationTypeId != ElementId.InvalidElementId && insulationThickness > 0)
                        {
                            try
                            {
                                TaskDialog.Show("Debug", $"Attempting to create insulation:\nType ID: {insulationTypeId}\nThickness: {insulationThickness}");

                                // Create insulation for duct1
                                DuctInsulation insulation1 = DuctInsulation.Create(doc, duct1.Id, insulationTypeId, insulationThickness);

                                // Create insulation for duct2
                                DuctInsulation insulation2 = DuctInsulation.Create(doc, duct2.Id, insulationTypeId, insulationThickness);

                                if (insulation1 != null && insulation2 != null)
                                {
                                    TaskDialog.Show("Debug", "Insulation created successfully for both ducts");

                                    // Copy any custom parameters from original insulation
                                    if (originalInsulation != null)
                                    {
                                        CopyInsulationParameters(originalInsulation, insulation1);
                                        CopyInsulationParameters(originalInsulation, insulation2);
                                    }
                                }
                                else
                                {
                                    TaskDialog.Show("Warning", "Insulation creation returned null for one or both ducts");
                                }
                            }
                            catch (Exception ex)
                            {
                                TaskDialog.Show("Warning", $"Ducts were split successfully, but insulation could not be added:\n{ex.Message}\n\nStack trace:\n{ex.StackTrace}");
                            }
                        }
                        else
                        {
                            string reason = "";
                            if (insulationTypeId == ElementId.InvalidElementId)
                                reason += "Invalid insulation type ID\n";
                            if (insulationThickness <= 0)
                                reason += $"Invalid thickness: {insulationThickness}";

                            if (!string.IsNullOrEmpty(reason))
                                TaskDialog.Show("Debug", $"Not creating insulation because:\n{reason}");
                        }

                        trans.Commit();
                        TaskDialog.Show("Success", "Duct split into 2 segments successfully.");
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

        #region Helper Classes

        private class DuctProperties
        {
            public ElementId SystemTypeId { get; set; }
            public ElementId DuctTypeId { get; set; }
            public ElementId LevelId { get; set; }
            public double Width { get; set; }
            public double Height { get; set; }
            public double Diameter { get; set; }
            public double Offset { get; set; }
            public ConnectorProfileType Shape { get; set; }
            public MEPSystem System { get; set; }
        }

        private class ConnectedElement
        {
            public Element Element { get; set; }
            public Connector OtherConnector { get; set; }
            public bool IsStart { get; set; }
        }

        #endregion

        #region Property Extraction

        private DuctProperties GetDuctProperties(Duct duct)
        {
            DuctProperties props = new DuctProperties
            {
                DuctTypeId = duct.GetTypeId(),
                LevelId = duct.ReferenceLevel?.Id ?? ElementId.InvalidElementId,
                SystemTypeId = duct.MEPSystem?.GetTypeId() ?? ElementId.InvalidElementId
            };

            // Get shape and dimensions
            ConnectorSet connectors = duct.ConnectorManager.Connectors;
            foreach (Connector conn in connectors)
            {
                props.Shape = conn.Shape;
                if (conn.Shape == ConnectorProfileType.Round)
                {
                    props.Diameter = conn.Radius * 2;
                }
                else if (conn.Shape == ConnectorProfileType.Rectangular)
                {
                    props.Width = conn.Width;
                    props.Height = conn.Height;
                }
                break; // Just need one connector to get the shape
            }

            // Get offset
            Parameter offsetParam = duct.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM);
            if (offsetParam != null)
            {
                props.Offset = offsetParam.AsDouble();
            }

            // Get MEP System
            if (duct.MEPSystem != null)
            {
                props.System = duct.MEPSystem;
            }

            return props;
        }

        #endregion

        #region Geometry Operations

        private XYZ FindIntersectionPoint(Line ductLine, Line splitLine)
        {
            // Project both lines to XY plane (plan view) by setting Z to 0
            XYZ ductStart2D = new XYZ(ductLine.GetEndPoint(0).X, ductLine.GetEndPoint(0).Y, 0);
            XYZ ductEnd2D = new XYZ(ductLine.GetEndPoint(1).X, ductLine.GetEndPoint(1).Y, 0);
            XYZ splitStart2D = new XYZ(splitLine.GetEndPoint(0).X, splitLine.GetEndPoint(0).Y, 0);
            XYZ splitEnd2D = new XYZ(splitLine.GetEndPoint(1).X, splitLine.GetEndPoint(1).Y, 0);

            // Create 2D lines for intersection
            Line ductLine2D = Line.CreateBound(ductStart2D, ductEnd2D);
            Line splitLine2D = Line.CreateBound(splitStart2D, splitEnd2D);

            // Make them unbounded for intersection calculation
            ductLine2D.MakeUnbound();
            splitLine2D.MakeUnbound();

            IntersectionResultArray results;
            SetComparisonResult result = ductLine2D.Intersect(splitLine2D, out results);

            if (result == SetComparisonResult.Overlap && results.Size > 0)
            {
                // Get the 2D intersection point
                XYZ intersection2D = results.get_Item(0).XYZPoint;

                // Project back to the duct line to get the actual 3D point
                // First, find the parameter on the original duct line
                XYZ ductDir = (ductLine.GetEndPoint(1) - ductLine.GetEndPoint(0)).Normalize();
                XYZ ductDir2D = new XYZ(ductDir.X, ductDir.Y, 0).Normalize();

                // Calculate how far along the duct line this intersection occurs
                double distanceAlong2D = (intersection2D - ductStart2D).DotProduct(ductDir2D);

                // Calculate the parameter t for the original 3D duct line
                double ductLength = ductLine.Length;
                double ductLength2D = ductStart2D.DistanceTo(ductEnd2D);
                double t = distanceAlong2D / ductLength2D;

                // Ensure t is within valid range [0, 1]
                if (t < 0 || t > 1)
                    return null;

                // Calculate the actual 3D intersection point on the duct
                XYZ intersectionPoint3D = ductLine.GetEndPoint(0) + (ductLine.GetEndPoint(1) - ductLine.GetEndPoint(0)) * t;

                return intersectionPoint3D;
            }

            return null;
        }

        #endregion

        #region Duct Creation

        private Duct CreateDuctSegment(Document doc, XYZ startPoint, XYZ endPoint, DuctProperties props)
        {
            try
            {
                // Create the duct
                Duct newDuct = Duct.Create(doc, props.SystemTypeId, props.DuctTypeId,
                    props.LevelId, startPoint, endPoint);

                if (newDuct != null)
                {
                    // Set offset
                    Parameter offsetParam = newDuct.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM);
                    if (offsetParam != null && !offsetParam.IsReadOnly)
                    {
                        offsetParam.Set(props.Offset);
                    }

                    // Set dimensions based on shape
                    if (props.Shape == ConnectorProfileType.Round)
                    {
                        Parameter diamParam = newDuct.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM);
                        if (diamParam != null && !diamParam.IsReadOnly)
                        {
                            diamParam.Set(props.Diameter);
                        }
                    }
                    else if (props.Shape == ConnectorProfileType.Rectangular)
                    {
                        Parameter widthParam = newDuct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
                        Parameter heightParam = newDuct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM);

                        if (widthParam != null && !widthParam.IsReadOnly)
                        {
                            widthParam.Set(props.Width);
                        }
                        if (heightParam != null && !heightParam.IsReadOnly)
                        {
                            heightParam.Set(props.Height);
                        }
                    }
                }

                return newDuct;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to create duct segment: {ex.Message}");
                return null;
            }
        }

        private void CopyDuctParameters(Duct sourceDuct, Duct targetDuct)
        {
            try
            {
                // Copy system classification using GUID for compatibility across Revit versions
                Guid systemClassGuid = new Guid("f1a7b350-59ef-4062-a36d-6c48581cebe6");
                Parameter sourceSystemClass = sourceDuct.get_Parameter(systemClassGuid);
                Parameter targetSystemClass = targetDuct.get_Parameter(systemClassGuid);

                if (sourceSystemClass != null && targetSystemClass != null && !targetSystemClass.IsReadOnly)
                {
                    targetSystemClass.Set(sourceSystemClass.AsInteger());
                }

                // Copy flow parameters if they exist
                Parameter sourceFlow = sourceDuct.get_Parameter(BuiltInParameter.RBS_DUCT_FLOW_PARAM);
                Parameter targetFlow = targetDuct.get_Parameter(BuiltInParameter.RBS_DUCT_FLOW_PARAM);

                if (sourceFlow != null && targetFlow != null && !targetFlow.IsReadOnly && sourceFlow.HasValue)
                {
                    targetFlow.Set(sourceFlow.AsDouble());
                }

                // Copy other relevant parameters
                CopyParameter(sourceDuct, targetDuct, BuiltInParameter.RBS_REFERENCE_OVERALLSIZE);
                CopyParameter(sourceDuct, targetDuct, BuiltInParameter.RBS_REFERENCE_LINING_THICKNESS);
                CopyParameter(sourceDuct, targetDuct, BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS);

                // Copy all user-defined parameters
                foreach (Parameter sourceParam in sourceDuct.Parameters)
                {
                    if (!sourceParam.IsReadOnly && sourceParam.HasValue && sourceParam.Definition != null)
                    {
                        Parameter targetParam = targetDuct.get_Parameter(sourceParam.Definition);
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
                            catch
                            {
                                // Skip parameters that can't be copied
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Log but don't fail the operation
                Debug.WriteLine($"Failed to copy some parameters: {ex.Message}");
            }
        }

        private void CopyParameter(Element source, Element target, BuiltInParameter paramId)
        {
            try
            {
                Parameter sourceParam = source.get_Parameter(paramId);
                Parameter targetParam = target.get_Parameter(paramId);

                if (sourceParam != null && targetParam != null && !targetParam.IsReadOnly && sourceParam.HasValue)
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
            catch
            {
                // Ignore parameter copy failures
            }
        }

        #endregion

        #region Connection Handling

        private Dictionary<int, ConnectedElement> GetConnectedElements(ConnectorSet connectors)
        {
            Dictionary<int, ConnectedElement> connected = new Dictionary<int, ConnectedElement>();
            int index = 0;

            foreach (Connector conn in connectors)
            {
                if (conn.IsConnected)
                {
                    foreach (Connector other in conn.AllRefs)
                    {
                        if (other.Owner.Id != conn.Owner.Id)
                        {
                            connected[index] = new ConnectedElement
                            {
                                Element = other.Owner,
                                OtherConnector = other,
                                IsStart = (index == 0)
                            };
                            break;
                        }
                    }
                }
                index++;
            }

            return connected;
        }

        private void ReconnectElements(Document doc, Duct duct1, Duct duct2,
            Dictionary<int, ConnectedElement> originalConnections, XYZ splitPoint)
        {
            foreach (var kvp in originalConnections)
            {
                ConnectedElement connElem = kvp.Value;

                try
                {
                    // Determine which new duct to connect to based on position
                    Connector duct1Start = GetConnectorAtEnd(duct1, true);
                    Connector duct1End = GetConnectorAtEnd(duct1, false);
                    Connector duct2Start = GetConnectorAtEnd(duct2, true);
                    Connector duct2End = GetConnectorAtEnd(duct2, false);

                    // Find the connector closest to the original connected element
                    Connector targetConnector = null;
                    double minDist = double.MaxValue;

                    // Get location of the connected element's connector
                    XYZ connectedLocation = null;
                    try
                    {
                        connectedLocation = connElem.OtherConnector.Origin;
                    }
                    catch
                    {
                        // If Origin is not available, try to get location from the element
                        if (connElem.Element is MEPCurve mepCurve)
                        {
                            LocationCurve loc = mepCurve.Location as LocationCurve;
                            if (loc != null)
                            {
                                // Use the appropriate end based on the original connection
                                connectedLocation = connElem.IsStart ?
                                    loc.Curve.GetEndPoint(0) : loc.Curve.GetEndPoint(1);
                            }
                        }
                        else if (connElem.Element is FamilyInstance fi)
                        {
                            LocationPoint loc = fi.Location as LocationPoint;
                            if (loc != null)
                            {
                                connectedLocation = loc.Point;
                            }
                        }
                    }

                    if (connectedLocation == null)
                        continue;

                    // Check each connector
                    Connector[] connectors = { duct1Start, duct1End, duct2Start, duct2End };
                    XYZ[] ductEndpoints = {
                        (duct1.Location as LocationCurve)?.Curve.GetEndPoint(0),
                        (duct1.Location as LocationCurve)?.Curve.GetEndPoint(1),
                        (duct2.Location as LocationCurve)?.Curve.GetEndPoint(0),
                        (duct2.Location as LocationCurve)?.Curve.GetEndPoint(1)
                    };

                    for (int i = 0; i < connectors.Length; i++)
                    {
                        if (connectors[i] != null && ductEndpoints[i] != null)
                        {
                            double dist = ductEndpoints[i].DistanceTo(connectedLocation);
                            if (dist < minDist)
                            {
                                minDist = dist;
                                targetConnector = connectors[i];
                            }
                        }
                    }

                    if (targetConnector != null && minDist < 1.0) // Within 1 foot
                    {
                        try
                        {
                            targetConnector.ConnectTo(connElem.OtherConnector);
                        }
                        catch
                        {
                            // Connection might fail if geometry doesn't align perfectly
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Failed to reconnect element: {ex.Message}");
                    continue;
                }
            }
        }

        private void ConnectDucts(Document doc, Duct duct1, Duct duct2, XYZ connectionPoint)
        {
            try
            {
                // Find the connectors closest to the split point
                Connector conn1 = GetConnectorNearPoint(duct1, connectionPoint);
                Connector conn2 = GetConnectorNearPoint(duct2, connectionPoint);

                if (conn1 != null && conn2 != null)
                {
                    // Try direct connection first
                    try
                    {
                        conn1.ConnectTo(conn2);
                    }
                    catch
                    {
                        // If direct connection fails, try to create a fitting
                        try
                        {
                            // Get the actual endpoints for fitting creation
                            LocationCurve loc1 = duct1.Location as LocationCurve;
                            LocationCurve loc2 = duct2.Location as LocationCurve;

                            if (loc1 != null && loc2 != null)
                            {
                                XYZ end1 = loc1.Curve.GetEndPoint(1);
                                XYZ start2 = loc2.Curve.GetEndPoint(0);

                                // Check if the points are close enough
                                if (end1.DistanceTo(connectionPoint) < 0.1 &&
                                    start2.DistanceTo(connectionPoint) < 0.1)
                                {
                                    doc.Create.NewUnionFitting(conn1, conn2);
                                }
                            }
                        }
                        catch
                        {
                            // If fitting creation also fails, leave unconnected
                            Debug.WriteLine("Failed to create fitting between split ducts");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to connect ducts: {ex.Message}");
                // Don't fail the entire operation if connection fails
            }
        }

        private Connector GetConnectorAtEnd(Duct duct, bool start)
        {
            LocationCurve locCurve = duct.Location as LocationCurve;
            if (locCurve == null) return null;

            Curve curve = locCurve.Curve;
            XYZ endPoint = start ? curve.GetEndPoint(0) : curve.GetEndPoint(1);

            return GetConnectorNearPoint(duct, endPoint);
        }

        private Connector GetConnectorNearPoint(Duct duct, XYZ point)
        {
            ConnectorSet connectors = duct.ConnectorManager.Connectors;
            Connector closest = null;
            double minDist = double.MaxValue;

            foreach (Connector conn in connectors)
            {
                try
                {
                    // Check if the connector has a valid origin
                    if (conn.ConnectorType == ConnectorType.End ||
                        conn.ConnectorType == ConnectorType.Curve ||
                        conn.ConnectorType == ConnectorType.Physical)
                    {
                        XYZ connOrigin = conn.Origin;
                        if (connOrigin != null)
                        {
                            double dist = connOrigin.DistanceTo(point);
                            if (dist < minDist)
                            {
                                minDist = dist;
                                closest = conn;
                            }
                        }
                    }
                }
                catch
                {
                    // Skip connectors that don't support Origin
                    continue;
                }
            }

            // If no connector found with Origin, try using the connector's owner location
            if (closest == null)
            {
                LocationCurve locCurve = duct.Location as LocationCurve;
                if (locCurve != null)
                {
                    Curve curve = locCurve.Curve;
                    XYZ start = curve.GetEndPoint(0);
                    XYZ end = curve.GetEndPoint(1);

                    // Find which end is closer to the point
                    bool useStart = start.DistanceTo(point) < end.DistanceTo(point);

                    foreach (Connector conn in connectors)
                    {
                        try
                        {
                            // For connectors without Origin, determine by index
                            if (connectors.Size == 2)
                            {
                                int index = 0;
                                foreach (Connector c in connectors)
                                {
                                    if (c.Id == conn.Id)
                                    {
                                        if ((index == 0 && useStart) || (index == 1 && !useStart))
                                        {
                                            closest = conn;
                                            break;
                                        }
                                    }
                                    index++;
                                }
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }
                }
            }

            return closest;
        }

        #endregion

        #region Insulation Handling

        private List<ElementId> GetDuctInsulation(Document doc, Duct duct)
        {
            List<ElementId> insulationIds = new List<ElementId>();

            try
            {
                // Method 1: Get all dependent elements
                ICollection<ElementId> dependents = duct.GetDependentElements(
                    new ElementClassFilter(typeof(DuctInsulation)));

                if (dependents != null && dependents.Count > 0)
                {
                    insulationIds.AddRange(dependents);
                }
            }
            catch
            {
                // Method 1 failed, continue to method 2
            }

            // Method 2: If the above method fails or returns nothing, try alternative approach
            if (insulationIds.Count == 0)
            {
                try
                {
                    FilteredElementCollector collector = new FilteredElementCollector(doc);
                    ICollection<Element> allInsulation = collector.OfClass(typeof(DuctInsulation)).ToElements();

                    foreach (Element elem in allInsulation)
                    {
                        DuctInsulation insulation = elem as DuctInsulation;
                        if (insulation != null)
                        {
                            try
                            {
                                // Try to get the host element ID directly
                                if (insulation.HostElementId == duct.Id)
                                {
                                    insulationIds.Add(insulation.Id);
                                }
                            }
                            catch
                            {
                                // If HostElementId doesn't work, try another approach
                                // Get the insulation's host through its geometry
                                try
                                {
                                    // Check if the insulation references our duct
                                    ICollection<ElementId> refs = insulation.GetDependentElements(null);
                                    if (refs.Contains(duct.Id))
                                    {
                                        insulationIds.Add(insulation.Id);
                                    }
                                }
                                catch
                                {
                                    // Skip this insulation
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error finding insulation: {ex.Message}");
                }
            }

            return insulationIds;
        }

        private void CopyInsulationParameters(DuctInsulation source, DuctInsulation target)
        {
            try
            {
                foreach (Parameter sourceParam in source.Parameters)
                {
                    if (!sourceParam.IsReadOnly && sourceParam.HasValue && sourceParam.Definition != null)
                    {
                        // Skip system parameters that are automatically set
                        if (sourceParam.Definition.Name == "Thickness" ||
                            sourceParam.Definition.Name == "Type" ||
                            sourceParam.Definition.Name == "Host Element")
                            continue;

                        Parameter targetParam = target.get_Parameter(sourceParam.Definition);
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
                            catch
                            {
                                // Skip parameters that can't be copied
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to copy insulation parameters: {ex.Message}");
            }
        }

        #endregion

        #region Selection Filters

        public class DuctSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem)
            {
                return elem is Duct;
            }

            public bool AllowReference(Reference reference, XYZ position)
            {
                return false;
            }
        }

        #endregion

        #region Button Data

        internal static PushButtonData GetButtonData()
        {
            string buttonInternalName = "btnDuctSplitter";
            string buttonTitle = "Split Duct";

            ButtonDataClass myButtonData = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "Splits a duct into two segments at a model line intersection, preserving insulation and connections");

            return myButtonData.Data;
        }

        #endregion
    }
}