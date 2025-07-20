using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI.Selection;
using SeperatorAddin.Common;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;

namespace SeperatorAddin
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class cmdPipeSplitter : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            try
            {
                // Select a pipe
                Reference pipeRef = uiDoc.Selection.PickObject(
                    ObjectType.Element,
                    new PipeSelectionFilter(),
                    "Select a pipe to split");

                Pipe selectedPipe = doc.GetElement(pipeRef) as Pipe;

                // Select ONE model line
                Reference lineRef = uiDoc.Selection.PickObject(
                    ObjectType.Element,
                    new ModelCurveSelectionFilter(),
                    "Select ONE straight model line that crosses the pipe");

                ModelCurve modelLine = doc.GetElement(lineRef) as ModelCurve;
                Line splitLine = modelLine.GeometryCurve as Line;

                if (splitLine == null)
                {
                    TaskDialog.Show("Error", "Please select a straight line.");
                    return Result.Failed;
                }

                using (Transaction trans = new Transaction(doc, "Split Pipe"))
                {
                    trans.Start();

                    try
                    {
                        // Store pipe properties before splitting
                        PipeProperties pipeProps = GetPipeProperties(selectedPipe);

                        if (pipeProps.SystemTypeId == null || pipeProps.SystemTypeId == ElementId.InvalidElementId)
                        {
                            TaskDialog.Show("Error", "The selected pipe is not assigned to a valid piping system.");
                            trans.RollBack();
                            return Result.Failed;
                        }

                        LocationCurve locCurve = selectedPipe.Location as LocationCurve;
                        if (locCurve == null)
                        {
                            trans.RollBack();
                            TaskDialog.Show("Error", "Could not get pipe location.");
                            return Result.Failed;
                        }

                        Line pipeLine = locCurve.Curve as Line;
                        if (pipeLine == null)
                        {
                            trans.RollBack();
                            TaskDialog.Show("Error", "This tool only works with straight pipes.");
                            return Result.Failed;
                        }

                        // Find intersection point
                        XYZ intersectionPoint = FindIntersectionPoint(pipeLine, splitLine);
                        if (intersectionPoint == null)
                        {
                            trans.RollBack();
                            TaskDialog.Show("Error", "The split line does not intersect with the pipe in plan view.");
                            return Result.Failed;
                        }

                        // Check if intersection point is valid
                        double minDistance = 0.1;
                        if (intersectionPoint.DistanceTo(pipeLine.GetEndPoint(0)) < minDistance ||
                            intersectionPoint.DistanceTo(pipeLine.GetEndPoint(1)) < minDistance)
                        {
                            trans.RollBack();
                            TaskDialog.Show("Error", "Split point is too close to pipe ends.");
                            return Result.Failed;
                        }

                        // Get connected elements
                        Dictionary<int, ConnectedElement> connectedElements = GetConnectedElements(selectedPipe.ConnectorManager.Connectors);

                        // Check for insulation
                        ElementId originalInsulationId = PipeInsulation.GetInsulationIds(doc, selectedPipe.Id).FirstOrDefault();
                        PipeInsulation originalInsulation = null;
                        if (originalInsulationId != null && originalInsulationId != ElementId.InvalidElementId)
                        {
                            originalInsulation = doc.GetElement(originalInsulationId) as PipeInsulation;
                        }
                        ElementId insulationTypeId = originalInsulation?.GetTypeId();
                        double insulationThickness = originalInsulation?.Thickness ?? 0;

                        // Get original connected elements before deleting pipe
                        Connector startConnector = connectedElements.ContainsKey(0) ? connectedElements[0].OtherConnector : null;
                        Connector endConnector = connectedElements.ContainsKey(1) ? connectedElements[1].OtherConnector : null;

                        // Create new pipes (we create them before deleting the original to copy parameters)
                        Pipe pipe1 = Pipe.Create(doc, pipeProps.SystemTypeId, pipeProps.PipeTypeId, pipeProps.LevelId, pipeLine.GetEndPoint(0), intersectionPoint);
                        Pipe pipe2 = Pipe.Create(doc, pipeProps.SystemTypeId, pipeProps.PipeTypeId, pipeProps.LevelId, intersectionPoint, pipeLine.GetEndPoint(1));

                        if (pipe1 == null || pipe2 == null)
                        {
                            trans.RollBack();
                            TaskDialog.Show("Error", "Failed to create new pipe segments.");
                            return Result.Failed;
                        }

                        pipe1.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)?.Set(pipeProps.Diameter);
                        pipe2.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)?.Set(pipeProps.Diameter);

                        // Copy parameters from original pipe
                        CopyPipeParameters(selectedPipe, pipe1);
                        CopyPipeParameters(selectedPipe, pipe2);

                        // Now delete the original pipe
                        doc.Delete(selectedPipe.Id);

                        // Connect the two new segments to each other first
                        ConnectPipes(doc, pipe1, pipe2, intersectionPoint);

                        // Regenerate to commit the new geometry and connection
                        doc.Regenerate();

                        // Connect the outer ends to the original connectors
                        if (startConnector != null)
                        {
                            GetConnectorAtEnd(pipe1, true)?.ConnectTo(startConnector);
                        }
                        if (endConnector != null)
                        {
                            GetConnectorAtEnd(pipe2, false)?.ConnectTo(endConnector);
                        }

                        // Add insulation if original had it
                        if (insulationTypeId != null && insulationTypeId != ElementId.InvalidElementId && insulationThickness > 0)
                        {
                            PipeInsulation.Create(doc, pipe1.Id, insulationTypeId, insulationThickness);
                            PipeInsulation.Create(doc, pipe2.Id, insulationTypeId, insulationThickness);
                        }

                        trans.Commit();
                        TaskDialog.Show("Success", "Pipe split into 2 segments successfully.");
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
        private class PipeProperties
        {
            public ElementId SystemTypeId { get; set; }
            public ElementId PipeTypeId { get; set; }
            public ElementId LevelId { get; set; }
            public double Diameter { get; set; }
        }

        private class ConnectedElement
        {
            public Element Element { get; set; }
            public Connector OtherConnector { get; set; }
        }
        #endregion

        #region Property Extraction
        private PipeProperties GetPipeProperties(Pipe pipe)
        {
            return new PipeProperties
            {
                PipeTypeId = pipe.GetTypeId(),
                LevelId = pipe.ReferenceLevel.Id,
                SystemTypeId = pipe.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM)?.AsElementId(),
                Diameter = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble()
            };
        }
        #endregion

        #region Geometry
        private XYZ FindIntersectionPoint(Line pipeLine, Line splitLine)
        {
            Line pipeLine2D = Line.CreateBound(new XYZ(pipeLine.GetEndPoint(0).X, pipeLine.GetEndPoint(0).Y, 0), new XYZ(pipeLine.GetEndPoint(1).X, pipeLine.GetEndPoint(1).Y, 0));
            Line splitLine2D = Line.CreateBound(new XYZ(splitLine.GetEndPoint(0).X, splitLine.GetEndPoint(0).Y, 0), new XYZ(splitLine.GetEndPoint(1).X, splitLine.GetEndPoint(1).Y, 0));

            pipeLine2D.MakeUnbound();
            splitLine2D.MakeUnbound();

            IntersectionResultArray results;
            if (pipeLine2D.Intersect(splitLine2D, out results) == SetComparisonResult.Overlap && results.Size > 0)
            {
                XYZ intersection2D = results.get_Item(0).XYZPoint;
                double param = pipeLine.Project(new XYZ(intersection2D.X, intersection2D.Y, pipeLine.GetEndPoint(0).Z)).Parameter;

                return pipeLine.Evaluate(param, false);
            }

            return null;
        }
        #endregion

        #region Connection Handling
        private Dictionary<int, ConnectedElement> GetConnectedElements(ConnectorSet connectors)
        {
            var connected = new Dictionary<int, ConnectedElement>();
            int index = 0;
            foreach (Connector conn in connectors)
            {
                if (conn.IsConnected)
                {
                    foreach (Connector other in conn.AllRefs)
                    {
                        if (other.Owner.Id != conn.Owner.Id)
                        {
                            connected[index] = new ConnectedElement { Element = other.Owner, OtherConnector = other };
                            break;
                        }
                    }
                }
                index++;
            }
            return connected;
        }

        private void ConnectPipes(Document doc, Pipe pipe1, Pipe pipe2, XYZ connectionPoint)
        {
            Connector conn1 = GetConnectorNearPoint(pipe1, connectionPoint);
            Connector conn2 = GetConnectorNearPoint(pipe2, connectionPoint);

            if (conn1 != null && conn2 != null)
            {
                try
                {
                    doc.Create.NewUnionFitting(conn1, conn2);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Could not create union fitting: {ex.Message}");
                }
            }
        }

        private Connector GetConnectorAtEnd(Pipe pipe, bool start)
        {
            LocationCurve locCurve = pipe.Location as LocationCurve;
            if (locCurve == null) return null;

            Line pipeLine = locCurve.Curve as Line;
            if (pipeLine == null) return null;

            XYZ endPoint = start ? pipeLine.GetEndPoint(0) : pipeLine.GetEndPoint(1);
            return GetConnectorNearPoint(pipe, endPoint);
        }

        private Connector GetConnectorNearPoint(Pipe pipe, XYZ point)
        {
            ConnectorSet connectors = pipe.ConnectorManager.Connectors;
            return connectors.Cast<Connector>().OrderBy(c => c.Origin.DistanceTo(point)).FirstOrDefault();
        }
        #endregion

        #region Parameter Copying
        private void CopyPipeParameters(Pipe source, Pipe target)
        {
            var ignoredParams = new List<BuiltInParameter>
            {
                BuiltInParameter.CURVE_ELEM_LENGTH,
                BuiltInParameter.RBS_PIPE_DIAMETER_PARAM,
                BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM,
                BuiltInParameter.RBS_START_OFFSET_PARAM,
                BuiltInParameter.RBS_END_OFFSET_PARAM
            };

            foreach (Parameter sourceParam in source.Parameters)
            {
                if (sourceParam.IsReadOnly || !sourceParam.HasValue || (sourceParam.Definition != null && ignoredParams.Contains((BuiltInParameter)sourceParam.Id.IntegerValue)))
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
                        // Fails silently if a parameter can't be set
                    }
                }
            }
        }
        #endregion

        #region Selection Filters
        public class PipeSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem) => elem is Pipe;
            public bool AllowReference(Reference reference, XYZ position) => false;
        }

        public class ModelCurveSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem) => elem is ModelCurve;
            public bool AllowReference(Reference reference, XYZ position) => false;
        }
        #endregion

        #region Button Data
        internal static PushButtonData GetButtonData()
        {
            string buttonInternalName = "btnPipeSplitter";
            string buttonTitle = "Split Pipe";

            ButtonDataClass myButtonData = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Green_32,
                Properties.Resources.Green_16,
                "Splits a pipe into two segments at a model line intersection, preserving connections and insulation.");

            return myButtonData.Data;
        }
        #endregion
    }
}
