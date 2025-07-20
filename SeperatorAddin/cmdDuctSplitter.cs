using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace SeperatorAddin
{
    [Transaction(TransactionMode.Manual)]
    public class cmdDuctSplitter : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Select a duct
                Reference ductRef = uidoc.Selection.PickObject(
                    ObjectType.Element,
                    new DuctSelectionFilter(),
                    "Select a duct to split");

                Duct originalDuct = doc.GetElement(ductRef) as Duct;
                if (originalDuct == null)
                {
                    TaskDialog.Show("Error", "Selected element is not a duct.");
                    return Result.Failed;
                }

                // Get split point
                XYZ splitPoint = uidoc.Selection.PickPoint("Pick a point on the duct to split");

                // Project the point onto the duct
                LocationCurve locCurve = originalDuct.Location as LocationCurve;
                if (locCurve == null)
                {
                    TaskDialog.Show("Error", "Could not get duct location curve.");
                    return Result.Failed;
                }

                Line ductLine = locCurve.Curve as Line;
                if (ductLine == null)
                {
                    TaskDialog.Show("Error", "This tool only works with straight ducts.");
                    return Result.Failed;
                }

                // Project split point onto duct line
                IntersectionResult projection = ductLine.Project(splitPoint);
                if (projection == null)
                {
                    TaskDialog.Show("Error", "Could not project point onto duct.");
                    return Result.Failed;
                }

                XYZ projectedPoint = projection.XYZPoint;

                // Check if split point is valid (not too close to ends)
                double minDistance = 0.5; // 6 inches minimum
                if (projectedPoint.DistanceTo(ductLine.GetEndPoint(0)) < minDistance ||
                    projectedPoint.DistanceTo(ductLine.GetEndPoint(1)) < minDistance)
                {
                    TaskDialog.Show("Error", "Split point is too close to duct ends.");
                    return Result.Failed;
                }

                using (Transaction trans = new Transaction(doc, "Split Duct with Insulation and Lining"))
                {
                    trans.Start();

                    try
                    {
                        // Store original duct insulation data before splitting
                        InsulationData originalInsulationData = GetDuctInsulation(doc, originalDuct);

                        // Store original duct lining data before splitting
                        LiningData originalLiningData = GetDuctLining(doc, originalDuct);

                        // Store original duct parameters
                        Dictionary<string, object> ductParameters = StoreDuctParameters(originalDuct);

                        // Get original duct system
                        ElementId systemTypeId = originalDuct.get_Parameter(
                            BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM).AsElementId();

                        // Split the duct using Revit's built-in method
                        ElementId newDuctId = MechanicalUtils.BreakCurve(
                            doc, originalDuct.Id, projectedPoint);

                        if (newDuctId == null || newDuctId == ElementId.InvalidElementId)
                        {
                            trans.RollBack();
                            TaskDialog.Show("Error", "Failed to split duct.");
                            return Result.Failed;
                        }

                        // Get the new duct created by the split
                        Duct newDuct = doc.GetElement(newDuctId) as Duct;

                        // Restore parameters to both ducts
                        RestoreDuctParameters(originalDuct, ductParameters);
                        RestoreDuctParameters(newDuct, ductParameters);

                        // Handle insulation for split ducts
                        if (originalInsulationData != null && originalInsulationData.HasInsulation)
                        {
                            // Add insulation to the new duct
                            DuctInsulation newDuctInsulation = AddInsulationToDuct(
                                doc, newDuct, originalInsulationData);

                            // Add insulation to the original duct (it may have been removed during split)
                            DuctInsulation originalDuctNewInsulation = AddInsulationToDuct(
                                doc, originalDuct, originalInsulationData);
                        }

                        // Handle lining for split ducts
                        if (originalLiningData != null && originalLiningData.HasLining)
                        {
                            // Add lining to the new duct
                            DuctLining newDuctLining = AddLiningToDuct(
                                doc, newDuct, originalLiningData);

                            // Add lining to the original duct (it may have been removed during split)
                            DuctLining originalDuctNewLining = AddLiningToDuct(
                                doc, originalDuct, originalLiningData);
                        }

                        trans.Commit();

                        // Report results
                        string resultMessage = "Duct split successfully.";
                        if (originalInsulationData != null && originalInsulationData.HasInsulation)
                        {
                            resultMessage += "\nInsulation parameters copied to both duct segments.";
                        }
                        if (originalLiningData != null && originalLiningData.HasLining)
                        {
                            resultMessage += "\nLining parameters copied to both duct segments.";
                        }

                        TaskDialog.Show("Success", resultMessage);
                        return Result.Succeeded;
                    }
                    catch (Exception ex)
                    {
                        trans.RollBack();
                        TaskDialog.Show("Error", $"Failed to split duct: {ex.Message}");
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

        #region Insulation Data Storage and Retrieval

        private class InsulationData
        {
            public bool HasInsulation { get; set; }
            public ElementId InsulationTypeId { get; set; }
            public double Thickness { get; set; }
            public Dictionary<string, object> Parameters { get; set; }

            public InsulationData()
            {
                Parameters = new Dictionary<string, object>();
            }
        }

        private InsulationData GetDuctInsulation(Document doc, Duct duct)
        {
            InsulationData data = new InsulationData();

            // Find insulation associated with the duct
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            var insulations = collector
                .OfClass(typeof(DuctInsulation))
                .Cast<DuctInsulation>()
                .Where(ins => ins.HostElementId == duct.Id)
                .ToList();

            if (insulations.Count == 0)
            {
                data.HasInsulation = false;
                return data;
            }

            DuctInsulation insulation = insulations.First();
            data.HasInsulation = true;
            data.InsulationTypeId = insulation.GetTypeId();
            data.Thickness = insulation.Thickness;

            // Store all parameters
            foreach (Parameter param in insulation.Parameters)
            {
                if (param == null || param.IsReadOnly) continue;

                string paramName = param.Definition.Name;

                try
                {
                    switch (param.StorageType)
                    {
                        case StorageType.Double:
                            data.Parameters[paramName] = param.AsDouble();
                            break;
                        case StorageType.Integer:
                            data.Parameters[paramName] = param.AsInteger();
                            break;
                        case StorageType.String:
                            data.Parameters[paramName] = param.AsString();
                            break;
                        case StorageType.ElementId:
                            data.Parameters[paramName] = param.AsElementId();
                            break;
                    }
                }
                catch { }
            }

            return data;
        }

        #endregion

        #region Lining Data Storage and Retrieval

        private class LiningData
        {
            public bool HasLining { get; set; }
            public ElementId LiningTypeId { get; set; }
            public double Thickness { get; set; }
            public Dictionary<string, object> Parameters { get; set; }

            public LiningData()
            {
                Parameters = new Dictionary<string, object>();
            }
        }

        private LiningData GetDuctLining(Document doc, Duct duct)
        {
            LiningData data = new LiningData();

            // Find lining associated with the duct
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            var linings = collector
                .OfClass(typeof(DuctLining))
                .Cast<DuctLining>()
                .Where(lining => lining.HostElementId == duct.Id)
                .ToList();

            if (linings.Count == 0)
            {
                data.HasLining = false;
                return data;
            }

            DuctLining lining = linings.First();
            data.HasLining = true;
            data.LiningTypeId = lining.GetTypeId();
            data.Thickness = lining.Thickness;

            // Store all parameters
            foreach (Parameter param in lining.Parameters)
            {
                if (param == null || param.IsReadOnly) continue;

                string paramName = param.Definition.Name;

                try
                {
                    switch (param.StorageType)
                    {
                        case StorageType.Double:
                            data.Parameters[paramName] = param.AsDouble();
                            break;
                        case StorageType.Integer:
                            data.Parameters[paramName] = param.AsInteger();
                            break;
                        case StorageType.String:
                            data.Parameters[paramName] = param.AsString();
                            break;
                        case StorageType.ElementId:
                            data.Parameters[paramName] = param.AsElementId();
                            break;
                    }
                }
                catch { }
            }

            return data;
        }

        #endregion

        #region Duct Parameter Handling

        private Dictionary<string, object> StoreDuctParameters(Duct duct)
        {
            Dictionary<string, object> parameters = new Dictionary<string, object>();

            // Store important duct parameters
            Parameter systemType = duct.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM);
            if (systemType != null)
                parameters["SystemType"] = systemType.AsElementId();

            Parameter width = duct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
            if (width != null)
                parameters["Width"] = width.AsDouble();

            Parameter height = duct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM);
            if (height != null)
                parameters["Height"] = height.AsDouble();

            Parameter diameter = duct.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM);
            if (diameter != null)
                parameters["Diameter"] = diameter.AsDouble();

            // Store custom parameters
            foreach (Parameter param in duct.Parameters)
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

        private void RestoreDuctParameters(Duct duct, Dictionary<string, object> parameters)
        {
            foreach (var kvp in parameters)
            {
                try
                {
                    Parameter param = duct.LookupParameter(kvp.Key);
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

        #region Insulation Creation

        private DuctInsulation AddInsulationToDuct(Document doc, Duct duct, InsulationData insulationData)
        {
            try
            {
                // Check if duct already has insulation
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                var existingInsulation = collector
                    .OfClass(typeof(DuctInsulation))
                    .Cast<DuctInsulation>()
                    .FirstOrDefault(ins => ins.HostElementId == duct.Id);

                if (existingInsulation != null)
                {
                    // Update existing insulation
                    CopyInsulationParameters(existingInsulation, insulationData);
                    return existingInsulation;
                }

                // Create new insulation
                DuctInsulation newInsulation = DuctInsulation.Create(
                    doc, duct.Id, insulationData.InsulationTypeId, insulationData.Thickness);

                if (newInsulation != null)
                {
                    CopyInsulationParameters(newInsulation, insulationData);
                }

                return newInsulation;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Warning", $"Failed to add insulation to duct: {ex.Message}");
                return null;
            }
        }

        private void CopyInsulationParameters(DuctInsulation targetInsulation, InsulationData sourceData)
        {
            foreach (var kvp in sourceData.Parameters)
            {
                try
                {
                    Parameter param = targetInsulation.LookupParameter(kvp.Key);
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

        #region Lining Creation

        private DuctLining AddLiningToDuct(Document doc, Duct duct, LiningData liningData)
        {
            try
            {
                // Check if duct already has lining
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                var existingLining = collector
                    .OfClass(typeof(DuctLining))
                    .Cast<DuctLining>()
                    .FirstOrDefault(lining => lining.HostElementId == duct.Id);

                if (existingLining != null)
                {
                    // Update existing lining
                    CopyLiningParameters(existingLining, liningData);
                    return existingLining;
                }

                // Create new lining
                DuctLining newLining = DuctLining.Create(
                    doc, duct.Id, liningData.LiningTypeId, liningData.Thickness);

                if (newLining != null)
                {
                    CopyLiningParameters(newLining, liningData);
                }

                return newLining;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Warning", $"Failed to add lining to duct: {ex.Message}");
                return null;
            }
        }

        private void CopyLiningParameters(DuctLining targetLining, LiningData sourceData)
        {
            foreach (var kvp in sourceData.Parameters)
            {
                try
                {
                    Parameter param = targetLining.LookupParameter(kvp.Key);
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
            string buttonInternalName = "btnDuctSplitter";
            string buttonTitle = "Split Duct";

            Common.ButtonDataClass myButtonData = new Common.ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Green_32,
                Properties.Resources.Green_16,
                "Splits a duct at a selected point and copies insulation and lining parameters to both segments");

            return myButtonData.Data;
        }

        #endregion
    }

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
}
