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
    public class cmdCableTraySplitter : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Select a cable tray
                Reference cableTrayRef = uidoc.Selection.PickObject(
                    ObjectType.Element,
                    new CableTraySelectionFilter(),
                    "Select a cable tray to split");

                CableTray originalCableTray = doc.GetElement(cableTrayRef) as CableTray;
                if (originalCableTray == null)
                {
                    TaskDialog.Show("Error", "Selected element is not a cable tray.");
                    return Result.Failed;
                }

                // Get split point
                XYZ splitPoint = uidoc.Selection.PickPoint("Pick a point on the cable tray to split");

                using (Transaction trans = new Transaction(doc, "Split Cable Tray"))
                {
                    trans.Start();
                    bool success = SplitCableTray(doc, originalCableTray, splitPoint);

                    if (success)
                    {
                        trans.Commit();
                        TaskDialog.Show("Success", "Cable tray split successfully.");
                        return Result.Succeeded;
                    }
                    else
                    {
                        trans.RollBack();
                        message = "Failed to split cable tray. The split point may be too close to an end.";
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

        public bool SplitCableTray(Document doc, CableTray originalCableTray, XYZ splitPoint)
        {
            if (originalCableTray == null) return false;

            LocationCurve locCurve = originalCableTray.Location as LocationCurve;
            if (locCurve == null || !(locCurve.Curve is Line cableTrayLine)) return false;

            IntersectionResult projection = cableTrayLine.Project(splitPoint);
            if (projection == null) return false;
            XYZ projectedPoint = projection.XYZPoint;

            // Check if split point is valid (not too close to ends)
            double minDistance = 0.5; // 6 inches minimum
            if (projectedPoint.DistanceTo(cableTrayLine.GetEndPoint(0)) < minDistance ||
                projectedPoint.DistanceTo(cableTrayLine.GetEndPoint(1)) < minDistance)
            {
                return false;
            }

            try
            {
                // Store original cable tray parameters
                Dictionary<string, object> cableTrayParameters = StoreCableTrayParameters(originalCableTray);
                CableTrayType cableTrayType = doc.GetElement(originalCableTray.GetTypeId()) as CableTrayType;
                ElementId levelId = originalCableTray.ReferenceLevel.Id;

                // Get start and end points
                XYZ startPoint = cableTrayLine.GetEndPoint(0);
                XYZ endPoint = cableTrayLine.GetEndPoint(1);

                // Create two new lines for the new segments
                Line line1 = Line.CreateBound(startPoint, projectedPoint);
                Line line2 = Line.CreateBound(projectedPoint, endPoint);

                // Create the new cable tray segments
                CableTray newCableTray1 = CableTray.Create(doc, cableTrayType.Id, line1.GetEndPoint(0), line1.GetEndPoint(1), levelId);
                CableTray newCableTray2 = CableTray.Create(doc, cableTrayType.Id, line2.GetEndPoint(0), line2.GetEndPoint(1), levelId);


                if (newCableTray1 == null || newCableTray2 == null)
                {
                    return false;
                }

                // Restore parameters to both new cable trays
                RestoreCableTrayParameters(newCableTray1, cableTrayParameters);
                RestoreCableTrayParameters(newCableTray2, cableTrayParameters);

                // Delete the original cable tray
                doc.Delete(originalCableTray.Id);

                return true;
            }
            catch
            {
                return false;
            }
        }


        #region CableTray Parameter Handling

        private Dictionary<string, object> StoreCableTrayParameters(CableTray cableTray)
        {
            Dictionary<string, object> parameters = new Dictionary<string, object>();

            // Store custom parameters
            foreach (Parameter param in cableTray.Parameters)
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

        private void RestoreCableTrayParameters(CableTray cableTray, Dictionary<string, object> parameters)
        {
            foreach (var kvp in parameters)
            {
                try
                {
                    Parameter param = cableTray.LookupParameter(kvp.Key);
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

    public class CableTraySelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            return elem is CableTray;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }
}