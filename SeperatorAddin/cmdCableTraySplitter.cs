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

                // Project the point onto the cable tray
                LocationCurve locCurve = originalCableTray.Location as LocationCurve;
                if (locCurve == null)
                {
                    TaskDialog.Show("Error", "Could not get cable tray location curve.");
                    return Result.Failed;
                }

                Line cableTrayLine = locCurve.Curve as Line;
                if (cableTrayLine == null)
                {
                    TaskDialog.Show("Error", "This tool only works with straight cable trays.");
                    return Result.Failed;
                }

                // Project split point onto cable tray line
                IntersectionResult projection = cableTrayLine.Project(splitPoint);
                if (projection == null)
                {
                    TaskDialog.Show("Error", "Could not project point onto cable tray.");
                    return Result.Failed;
                }

                XYZ projectedPoint = projection.XYZPoint;

                // Check if split point is valid (not too close to ends)
                double minDistance = 0.5; // 6 inches minimum
                if (projectedPoint.DistanceTo(cableTrayLine.GetEndPoint(0)) < minDistance ||
                    projectedPoint.DistanceTo(cableTrayLine.GetEndPoint(1)) < minDistance)
                {
                    TaskDialog.Show("Error", "Split point is too close to cable tray ends.");
                    return Result.Failed;
                }

                using (Transaction trans = new Transaction(doc, "Split Cable Tray"))
                {
                    trans.Start();

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
                            trans.RollBack();
                            TaskDialog.Show("Error", "Failed to create new cable tray segments.");
                            return Result.Failed;
                        }

                        // Restore parameters to both new cable trays
                        RestoreCableTrayParameters(newCableTray1, cableTrayParameters);
                        RestoreCableTrayParameters(newCableTray2, cableTrayParameters);

                        // Delete the original cable tray
                        doc.Delete(originalCableTray.Id);

                        trans.Commit();

                        TaskDialog.Show("Success", "Cable tray split successfully.");
                        return Result.Succeeded;
                    }
                    catch (Exception ex)
                    {
                        trans.RollBack();
                        TaskDialog.Show("Error", $"Failed to split cable tray: {ex.Message}");
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

        #region CableTray Parameter Handling

        private Dictionary<string, object> StoreCableTrayParameters(CableTray cableTray)
        {
            Dictionary<string, object> parameters = new Dictionary<string, object>();

            // Store important cable tray parameters
            Parameter width = cableTray.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM);
            if (width != null)
                parameters["Width"] = width.AsDouble();

            Parameter height = cableTray.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM);
            if (height != null)
                parameters["Height"] = height.AsDouble();

            // Store custom parameters
            foreach (Parameter param in cableTray.Parameters)
            {
                if (param == null || param.IsReadOnly || !param.HasValue) continue;

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

        #region Utility Methods

        internal static PushButtonData GetButtonData()
        {
            string buttonInternalName = "btnCableTraySplitter";
            string buttonTitle = "Split Cable Tray";

            Common.ButtonDataClass myButtonData = new Common.ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Yellow_32,
                Properties.Resources.Yellow_16,
                "Splits a cable tray at a selected point and copies parameters to both segments");

            return myButtonData.Data;
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