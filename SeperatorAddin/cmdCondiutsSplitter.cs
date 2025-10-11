using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace SeperatorAddin
{
    [Transaction(TransactionMode.Manual)]
    public class cmdConduitsSplitter : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Select a conduit
                Reference conduitRef = uidoc.Selection.PickObject(
                    ObjectType.Element,
                    new ConduitSelectionFilter(),
                    "Select a conduit to split");

                Conduit originalConduit = doc.GetElement(conduitRef) as Conduit;
                if (originalConduit == null)
                {
                    TaskDialog.Show("Error", "Selected element is not a conduit.");
                    return Result.Failed;
                }

                // Get split point
                XYZ splitPoint = uidoc.Selection.PickPoint("Pick a point on the conduit to split");

                using (Transaction trans = new Transaction(doc, "Split Conduit"))
                {
                    trans.Start();
                    bool success = SplitConduit(doc, originalConduit, splitPoint);

                    if (success)
                    {
                        trans.Commit();
                        TaskDialog.Show("Success", "Conduit split successfully.");
                        return Result.Succeeded;
                    }
                    else
                    {
                        trans.RollBack();
                        message = "Failed to split conduit. The split point may be too close to an end.";
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

        public bool SplitConduit(Document doc, Conduit originalConduit, XYZ splitPoint)
        {
            if (originalConduit == null) return false;
            LocationCurve locCurve = originalConduit.Location as LocationCurve;
            if (locCurve == null || !(locCurve.Curve is Line conduitLine)) return false;

            IntersectionResult projection = conduitLine.Project(splitPoint);
            if (projection == null) return false;
            XYZ projectedPoint = projection.XYZPoint;

            // Check if split point is valid (not too close to ends)
            double minDistance = 0.5; // 6 inches minimum
            if (projectedPoint.DistanceTo(conduitLine.GetEndPoint(0)) < minDistance ||
                projectedPoint.DistanceTo(conduitLine.GetEndPoint(1)) < minDistance)
            {
                return false;
            }

            try
            {
                // Store original conduit parameters
                Dictionary<string, object> conduitParameters = StoreConduitParameters(originalConduit);
                ConduitType conduitType = doc.GetElement(originalConduit.GetTypeId()) as ConduitType;
                ElementId levelId = originalConduit.ReferenceLevel.Id;

                // Get start and end points
                XYZ startPoint = conduitLine.GetEndPoint(0);
                XYZ endPoint = conduitLine.GetEndPoint(1);

                // Create two new lines for the new segments
                Line line1 = Line.CreateBound(startPoint, projectedPoint);
                Line line2 = Line.CreateBound(projectedPoint, endPoint);

                // Create the new conduit segments
                Conduit newConduit1 = Conduit.Create(doc, conduitType.Id, line1.GetEndPoint(0), line1.GetEndPoint(1), levelId);
                Conduit newConduit2 = Conduit.Create(doc, conduitType.Id, line2.GetEndPoint(0), line2.GetEndPoint(1), levelId);


                if (newConduit1 == null || newConduit2 == null)
                {
                    return false;
                }

                // Restore parameters to both new conduits
                RestoreConduitParameters(newConduit1, conduitParameters);
                RestoreConduitParameters(newConduit2, conduitParameters);

                // Delete the original conduit
                doc.Delete(originalConduit.Id);

                return true;
            }
            catch
            {
                return false;
            }
        }


        #region Conduit Parameter Handling

        private Dictionary<string, object> StoreConduitParameters(Conduit conduit)
        {
            Dictionary<string, object> parameters = new Dictionary<string, object>();

            // Store custom parameters
            foreach (Parameter param in conduit.Parameters)
            {
                if (param == null || !param.HasValue) continue;

                string paramName = param.Definition.Name;
                if (parameters.ContainsKey(paramName)) continue;

                try
                {
                    switch (param.StorageType)
                    {
                        case StorageType.Double:
                            parameters[paramName] = param.AsDouble();
                            break;
                        case StorageType.Integer:
                            parameters[paramName] = param.AsInteger();
                            break;
                        case StorageType.String:
                            parameters[paramName] = param.AsString();
                            break;
                        case StorageType.ElementId:
                            parameters[paramName] = param.AsElementId();
                            break;
                    }
                }
                catch { }
            }

            return parameters;
        }

        private void RestoreConduitParameters(Conduit conduit, Dictionary<string, object> parameters)
        {
            foreach (var kvp in parameters)
            {
                try
                {
                    Parameter param = conduit.LookupParameter(kvp.Key);
                    if (param == null || param.IsReadOnly) continue;

                    switch (param.StorageType)
                    {
                        case StorageType.Double:
                            if (kvp.Value is double)
                                param.Set((double)kvp.Value);
                            break;
                        case StorageType.Integer:
                            if (kvp.Value is int)
                                param.Set((int)kvp.Value);
                            break;
                        case StorageType.String:
                            if (kvp.Value is string)
                                param.Set((string)kvp.Value);
                            break;
                        case StorageType.ElementId:
                            if (kvp.Value is ElementId)
                                param.Set((ElementId)kvp.Value);
                            break;
                    }
                }
                catch { }
            }
        }

        #endregion
    }

    public class ConduitSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            return elem is Conduit;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }
}