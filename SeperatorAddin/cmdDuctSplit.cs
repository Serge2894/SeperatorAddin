using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using SeperatorAddin.Forms;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace SeperatorAddin.Forms
{
    [Transaction(TransactionMode.Manual)]
    public class cmdDuctSplit : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Get all ducts in the model
                FilteredElementCollector ductCollector = new FilteredElementCollector(doc)
                    .OfClass(typeof(Duct))
                    .WhereElementIsNotElementType();

                List<Duct> allDucts = ductCollector.Cast<Duct>().ToList();

                if (allDucts.Count == 0)
                {
                    TaskDialog.Show("No Ducts", "No ducts found in the model.");
                    return Result.Cancelled;
                }

                // Get all levels in the model
                FilteredElementCollector levelCollector = new FilteredElementCollector(doc)
                    .OfClass(typeof(Level));

                List<Level> allLevels = levelCollector.Cast<Level>()
                    .OrderBy(l => l.Elevation)
                    .ToList();

                if (allLevels.Count < 2)
                {
                    TaskDialog.Show("Insufficient Levels", "At least 2 levels are required to split ducts.");
                    return Result.Cancelled;
                }

                // Show duct selection form
                DuctSelectionForm form = new DuctSelectionForm(allDucts, allLevels);

                if (form.ShowDialog() != true)
                {
                    return Result.Cancelled;
                }

                List<Duct> selectedDucts = form.SelectedDucts;
                List<Level> selectedLevels = form.SelectedLevels.OrderBy(l => l.Elevation).ToList();

                if (selectedDucts.Count == 0)
                {
                    TaskDialog.Show("No Selection", "No ducts selected.");
                    return Result.Cancelled;
                }

                if (selectedLevels.Count < 2)
                {
                    TaskDialog.Show("Insufficient Levels", "At least 2 levels must be selected.");
                    return Result.Cancelled;
                }

                // Process the selected ducts
                using (Transaction trans = new Transaction(doc, "Split Ducts by Level"))
                {
                    trans.Start();

                    int successCount = 0;
                    int failCount = 0;
                    List<string> errors = new List<string>();

                    foreach (Duct duct in selectedDucts)
                    {
                        try
                        {
                            bool result = SplitDuctByLevels(doc, duct, selectedLevels);
                            if (result)
                                successCount++;
                            else
                                failCount++;
                        }
                        catch (Exception ex)
                        {
                            failCount++;
                            errors.Add($"Duct {duct.Id}: {ex.Message}");
                        }
                    }

                    trans.Commit();

                    // Show results
                    string resultMessage = $"Split operation completed:\n" +
                                         $"Successfully split: {successCount} ducts\n" +
                                         $"Failed: {failCount} ducts";

                    if (errors.Count > 0)
                    {
                        resultMessage += "\n\nErrors:\n" + string.Join("\n", errors.Take(10));
                        if (errors.Count > 10)
                            resultMessage += $"\n... and {errors.Count - 10} more errors";
                    }

                    TaskDialog.Show("Split Results", resultMessage);
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private bool SplitDuctByLevels(Document doc, Duct duct, List<Level> levels)
        {
            try
            {
                // Get duct geometry
                LocationCurve locCurve = duct.Location as LocationCurve;
                if (locCurve == null || !(locCurve.Curve is Line))
                {
                    // Skip non-linear ducts
                    return false;
                }

                Line ductLine = locCurve.Curve as Line;
                XYZ startPoint = ductLine.GetEndPoint(0);
                XYZ endPoint = ductLine.GetEndPoint(1);

                // Check if duct is vertical or has significant vertical component
                XYZ direction = (endPoint - startPoint).Normalize();
                if (Math.Abs(direction.Z) < 0.1) // Nearly horizontal
                {
                    return false;
                }

                // Find intersection points with levels
                List<double> splitHeights = new List<double>();

                foreach (Level level in levels)
                {
                    double levelElevation = level.Elevation;

                    // Check if level intersects with duct
                    if ((levelElevation > Math.Min(startPoint.Z, endPoint.Z) + 0.001) &&
                        (levelElevation < Math.Max(startPoint.Z, endPoint.Z) - 0.001))
                    {
                        splitHeights.Add(levelElevation);
                    }
                }

                if (splitHeights.Count == 0)
                {
                    // No intersections found
                    return false;
                }

                // Sort split heights
                splitHeights.Sort();

                // Store original duct properties
                DuctType ductType = doc.GetElement(duct.GetTypeId()) as DuctType;
                MEPSystem system = duct.MEPSystem;

                // Get duct dimensions
                double width = duct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM).AsDouble();
                double height = duct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM).AsDouble();

                // Store connectors and their connections
                List<Connector> originalConnectors = GetDuctConnectors(duct);
                Dictionary<Connector, List<Connector>> originalConnections = new Dictionary<Connector, List<Connector>>();

                foreach (Connector conn in originalConnectors)
                {
                    originalConnections[conn] = GetConnectedConnectors(conn);
                }

                // Calculate split points
                List<XYZ> splitPoints = new List<XYZ>();
                splitPoints.Add(startPoint); // Add start point

                foreach (double h in splitHeights)
                {
                    // Calculate parameter on line for this height
                    double t = (h - startPoint.Z) / (endPoint.Z - startPoint.Z);
                    XYZ splitPoint = startPoint + t * (endPoint - startPoint);
                    splitPoints.Add(splitPoint);
                }

                splitPoints.Add(endPoint); // Add end point

                // Create new duct segments
                List<Duct> newDucts = new List<Duct>();

                for (int i = 0; i < splitPoints.Count - 1; i++)
                {
                    XYZ segmentStart = splitPoints[i];
                    XYZ segmentEnd = splitPoints[i + 1];

                    // Create new duct
                    Duct newDuct = Duct.Create(doc,
                        doc.GetElement(duct.LevelId) as Level,
                        ductType.Id,
                        segmentStart,
                        segmentEnd);

                    // Set dimensions
                    newDuct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM).Set(width);
                    newDuct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM).Set(height);

                    // Copy other parameters
                    CopyDuctParameters(duct, newDuct);

                    newDucts.Add(newDuct);
                }

                // Connect the new ducts together
                for (int i = 0; i < newDucts.Count - 1; i++)
                {
                    ConnectDucts(newDucts[i], newDucts[i + 1]);
                }

                // Reconnect to original connections at start and end
                if (newDucts.Count > 0)
                {
                    // Reconnect start
                    Connector firstDuctStartConnector = GetConnectorAtPoint(newDucts[0], splitPoints[0]);
                    if (firstDuctStartConnector != null)
                    {
                        Connector originalStartConnector = GetConnectorAtPoint(duct, startPoint);
                        if (originalStartConnector != null && originalConnections.ContainsKey(originalStartConnector))
                        {
                            foreach (Connector conn in originalConnections[originalStartConnector])
                            {
                                if (conn.Owner.Id != duct.Id)
                                {
                                    try { firstDuctStartConnector.ConnectTo(conn); } catch { }
                                }
                            }
                        }
                    }

                    // Reconnect end
                    Connector lastDuctEndConnector = GetConnectorAtPoint(newDucts[newDucts.Count - 1], splitPoints[splitPoints.Count - 1]);
                    if (lastDuctEndConnector != null)
                    {
                        Connector originalEndConnector = GetConnectorAtPoint(duct, endPoint);
                        if (originalEndConnector != null && originalConnections.ContainsKey(originalEndConnector))
                        {
                            foreach (Connector conn in originalConnections[originalEndConnector])
                            {
                                if (conn.Owner.Id != duct.Id)
                                {
                                    try { lastDuctEndConnector.ConnectTo(conn); } catch { }
                                }
                            }
                        }
                    }
                }

                // Delete original duct
                doc.Delete(duct.Id);

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private List<Connector> GetDuctConnectors(Duct duct)
        {
            List<Connector> connectors = new List<Connector>();
            ConnectorSet connectorSet = duct.ConnectorManager.Connectors;

            foreach (Connector connector in connectorSet)
            {
                connectors.Add(connector);
            }

            return connectors;
        }

        private List<Connector> GetConnectedConnectors(Connector connector)
        {
            List<Connector> connected = new List<Connector>();

            foreach (Connector conn in connector.AllRefs)
            {
                connected.Add(conn);
            }

            return connected;
        }

        private Connector GetConnectorAtPoint(Duct duct, XYZ point)
        {
            const double tolerance = 0.001;

            foreach (Connector connector in duct.ConnectorManager.Connectors)
            {
                if (connector.Origin.IsAlmostEqualTo(point, tolerance))
                {
                    return connector;
                }
            }

            return null;
        }

        private void ConnectDucts(Duct duct1, Duct duct2)
        {
            try
            {
                // Get the closest connectors between the two ducts
                Connector conn1 = null;
                Connector conn2 = null;
                double minDistance = double.MaxValue;

                foreach (Connector c1 in duct1.ConnectorManager.Connectors)
                {
                    foreach (Connector c2 in duct2.ConnectorManager.Connectors)
                    {
                        double distance = c1.Origin.DistanceTo(c2.Origin);
                        if (distance < minDistance)
                        {
                            minDistance = distance;
                            conn1 = c1;
                            conn2 = c2;
                        }
                    }
                }

                if (conn1 != null && conn2 != null && minDistance < 1.0) // Within 1 foot
                {
                    conn1.ConnectTo(conn2);
                }
            }
            catch { }
        }

        private void CopyDuctParameters(Duct source, Duct target)
        {
            try
            {
                // Copy system type if possible
                Parameter sourceSystemType = source.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM);
                Parameter targetSystemType = target.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM);

                if (sourceSystemType != null && targetSystemType != null && !targetSystemType.IsReadOnly)
                {
                    targetSystemType.Set(sourceSystemType.AsElementId());
                }

                // Copy other relevant parameters
                CopyParameter(source, target, BuiltInParameter.RBS_DUCT_FLOW_PARAM);
                CopyParameter(source, target, BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                CopyParameter(source, target, BuiltInParameter.ALL_MODEL_MARK);

                // Copy insulation if present
                CopyParameter(source, target, BuiltInParameter.RBS_REFERENCE_INSULATION_TYPE);
                CopyParameter(source, target, BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS);
                CopyParameter(source, target, BuiltInParameter.RBS_REFERENCE_LINING_TYPE);
                CopyParameter(source, target, BuiltInParameter.RBS_REFERENCE_LINING_THICKNESS);
            }
            catch { }
        }

        private void CopyParameter(Element source, Element target, BuiltInParameter param)
        {
            try
            {
                Parameter sourceParam = source.get_Parameter(param);
                Parameter targetParam = target.get_Parameter(param);

                if (sourceParam != null && targetParam != null && !targetParam.IsReadOnly)
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
            catch { }
        }

        internal static PushButtonData GetButtonData()
        {
            string buttonInternalName = "btnDuctSplit";
            string buttonTitle = "Split Ducts\nby Level";

            Common.ButtonDataClass myButtonData = new Common.ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Yellow_32,
                Properties.Resources.Yellow_16,
                "Splits selected ducts at level intersections");

            return myButtonData.Data;
        }
    }
}