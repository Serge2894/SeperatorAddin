using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using SeperatorAddin.Forms;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace SeperatorAddin
{
    [Transaction(TransactionMode.Manual)]
    public class cmdPipeSplit : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Get all pipes in the model
                FilteredElementCollector pipeCollector = new FilteredElementCollector(doc)
                    .OfClass(typeof(Pipe))
                    .WhereElementIsNotElementType();

                List<Pipe> allPipes = pipeCollector.Cast<Pipe>().ToList();

                if (allPipes.Count == 0)
                {
                    TaskDialog.Show("No Pipes", "No pipes found in the model.");
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
                    TaskDialog.Show("Insufficient Levels", "At least 2 levels are required to split pipes.");
                    return Result.Cancelled;
                }

                // Show pipe selection form
                PipeSelectionForm form = new PipeSelectionForm(allPipes, allLevels);

                if (form.ShowDialog() != true)
                {
                    return Result.Cancelled;
                }

                List<Pipe> selectedPipes = form.SelectedPipes;
                List<Level> selectedLevels = form.SelectedLevels.OrderBy(l => l.Elevation).ToList();

                if (selectedPipes.Count == 0)
                {
                    TaskDialog.Show("No Selection", "No pipes selected.");
                    return Result.Cancelled;
                }

                if (selectedLevels.Count < 2)
                {
                    TaskDialog.Show("Insufficient Levels", "At least 2 levels must be selected.");
                    return Result.Cancelled;
                }

                // Process the selected pipes
                using (Transaction trans = new Transaction(doc, "Split Pipes by Level"))
                {
                    trans.Start();

                    int successCount = 0;
                    int failCount = 0;
                    List<string> errors = new List<string>();

                    foreach (Pipe pipe in selectedPipes)
                    {
                        try
                        {
                            bool result = SplitPipeByLevels(doc, pipe, selectedLevels);
                            if (result)
                                successCount++;
                            else
                                failCount++;
                        }
                        catch (Exception ex)
                        {
                            failCount++;
                            errors.Add($"Pipe {pipe.Id}: {ex.Message}");
                        }
                    }

                    trans.Commit();

                    // Show results
                    string resultMessage = $"Split operation completed:\n" +
                                         $"Successfully split: {successCount} pipes\n" +
                                         $"Failed: {failCount} pipes";

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

        private bool SplitPipeByLevels(Document doc, Pipe pipe, List<Level> levels)
        {
            try
            {
                // Get pipe geometry
                LocationCurve locCurve = pipe.Location as LocationCurve;
                if (locCurve == null || !(locCurve.Curve is Line))
                {
                    // Skip non-linear pipes
                    return false;
                }

                Line pipeLine = locCurve.Curve as Line;
                XYZ startPoint = pipeLine.GetEndPoint(0);
                XYZ endPoint = pipeLine.GetEndPoint(1);

                // Check if pipe is vertical or has significant vertical component
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

                    // Check if level intersects with pipe
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

                // Store original pipe properties
                PipeType pipeType = doc.GetElement(pipe.GetTypeId()) as PipeType;
                MEPSystem system = pipe.MEPSystem;
                double diameter = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble();

                // Store connectors and their connections
                List<Connector> originalConnectors = GetPipeConnectors(pipe);
                Dictionary<Connector, List<Connector>> originalConnections = new Dictionary<Connector, List<Connector>>();

                foreach (Connector conn in originalConnectors)
                {
                    originalConnections[conn] = GetConnectedConnectors(conn);
                }

                // Calculate split points
                List<XYZ> splitPoints = new List<XYZ>();
                splitPoints.Add(startPoint); // Add start point

                foreach (double height in splitHeights)
                {
                    // Calculate parameter on line for this height
                    double t = (height - startPoint.Z) / (endPoint.Z - startPoint.Z);
                    XYZ splitPoint = startPoint + t * (endPoint - startPoint);
                    splitPoints.Add(splitPoint);
                }

                splitPoints.Add(endPoint); // Add end point

                // Create new pipe segments
                List<Pipe> newPipes = new List<Pipe>();

                for (int i = 0; i < splitPoints.Count - 1; i++)
                {
                    XYZ segmentStart = splitPoints[i];
                    XYZ segmentEnd = splitPoints[i + 1];

                    // Create new pipe
                    Pipe newPipe = Pipe.Create(doc, pipeType.Id,
                        doc.GetElement(pipe.ReferenceLevel.Id) as Level,
                        segmentStart, segmentEnd);

                    // Set diameter
                    newPipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).Set(diameter);

                    // Copy other parameters
                    CopyPipeParameters(pipe, newPipe);

                    newPipes.Add(newPipe);
                }

                // Connect the new pipes together
                for (int i = 0; i < newPipes.Count - 1; i++)
                {
                    ConnectPipes(newPipes[i], newPipes[i + 1]);
                }

                // Reconnect to original connections at start and end
                if (newPipes.Count > 0)
                {
                    // Reconnect start
                    Connector firstPipeStartConnector = GetConnectorAtPoint(newPipes[0], splitPoints[0]);
                    if (firstPipeStartConnector != null)
                    {
                        Connector originalStartConnector = GetConnectorAtPoint(pipe, startPoint);
                        if (originalStartConnector != null && originalConnections.ContainsKey(originalStartConnector))
                        {
                            foreach (Connector conn in originalConnections[originalStartConnector])
                            {
                                if (conn.Owner.Id != pipe.Id)
                                {
                                    try { firstPipeStartConnector.ConnectTo(conn); } catch { }
                                }
                            }
                        }
                    }

                    // Reconnect end
                    Connector lastPipeEndConnector = GetConnectorAtPoint(newPipes[newPipes.Count - 1], splitPoints[splitPoints.Count - 1]);
                    if (lastPipeEndConnector != null)
                    {
                        Connector originalEndConnector = GetConnectorAtPoint(pipe, endPoint);
                        if (originalEndConnector != null && originalConnections.ContainsKey(originalEndConnector))
                        {
                            foreach (Connector conn in originalConnections[originalEndConnector])
                            {
                                if (conn.Owner.Id != pipe.Id)
                                {
                                    try { lastPipeEndConnector.ConnectTo(conn); } catch { }
                                }
                            }
                        }
                    }
                }

                // Delete original pipe
                doc.Delete(pipe.Id);

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private List<Connector> GetPipeConnectors(Pipe pipe)
        {
            List<Connector> connectors = new List<Connector>();
            ConnectorSet connectorSet = pipe.ConnectorManager.Connectors;

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

        private Connector GetConnectorAtPoint(Pipe pipe, XYZ point)
        {
            const double tolerance = 0.001;

            foreach (Connector connector in pipe.ConnectorManager.Connectors)
            {
                if (connector.Origin.IsAlmostEqualTo(point, tolerance))
                {
                    return connector;
                }
            }

            return null;
        }

        private void ConnectPipes(Pipe pipe1, Pipe pipe2)
        {
            try
            {
                // Get the closest connectors between the two pipes
                Connector conn1 = null;
                Connector conn2 = null;
                double minDistance = double.MaxValue;

                foreach (Connector c1 in pipe1.ConnectorManager.Connectors)
                {
                    foreach (Connector c2 in pipe2.ConnectorManager.Connectors)
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

        private void CopyPipeParameters(Pipe source, Pipe target)
        {
            try
            {
                // Copy system type if possible
                Parameter sourceSystemType = source.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM);
                Parameter targetSystemType = target.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM);

                if (sourceSystemType != null && targetSystemType != null && !targetSystemType.IsReadOnly)
                {
                    targetSystemType.Set(sourceSystemType.AsElementId());
                }

                // Copy other relevant parameters
                CopyParameter(source, target, BuiltInParameter.RBS_PIPE_SLOPE);
                CopyParameter(source, target, BuiltInParameter.RBS_PIPE_INVERT_ELEVATION);
                CopyParameter(source, target, BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                CopyParameter(source, target, BuiltInParameter.ALL_MODEL_MARK);
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
            string buttonInternalName = "btnPipeSplit";
            string buttonTitle = "Split Pipes\nby Level";

            Common.ButtonDataClass myButtonData = new Common.ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Green_32,
                Properties.Resources.Green_16,
                "Splits selected pipes at level intersections");

            return myButtonData.Data;
        }
    }
}