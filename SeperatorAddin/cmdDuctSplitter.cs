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

                using (Transaction trans = new Transaction(doc, "Split Duct with Insulation and Lining"))
                {
                    trans.Start();
                    bool success = SplitDuct(doc, originalDuct, splitPoint);
                    if (success)
                    {
                        trans.Commit();
                        TaskDialog.Show("Success", "Duct split successfully.");
                        return Result.Succeeded;
                    }
                    else
                    {
                        trans.RollBack();
                        message = "Failed to split the duct. The split point might be too close to an end or on an invalid element.";
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

        public bool SplitDuct(Document doc, Duct originalDuct, XYZ splitPoint)
        {
            if (originalDuct == null) return false;

            LocationCurve locCurve = originalDuct.Location as LocationCurve;
            if (locCurve == null || !(locCurve.Curve is Line ductLine)) return false;

            IntersectionResult projection = ductLine.Project(splitPoint);
            if (projection == null) return false;
            XYZ projectedPoint = projection.XYZPoint;

            // Check if split point is valid (not too close to ends)
            double minDistance = 0.5; // 6 inches minimum
            if (projectedPoint.DistanceTo(ductLine.GetEndPoint(0)) < minDistance ||
                projectedPoint.DistanceTo(ductLine.GetEndPoint(1)) < minDistance)
            {
                return false;
            }

            try
            {
                // Store original duct insulation data before splitting
                InsulationData originalInsulationData = GetDuctInsulation(doc, originalDuct);
                LiningData originalLiningData = GetDuctLining(doc, originalDuct);
                Dictionary<string, object> ductParameters = StoreDuctParameters(originalDuct);

                // Split the duct using Revit's built-in method
                ElementId newDuctId = MechanicalUtils.BreakCurve(doc, originalDuct.Id, projectedPoint);
                if (newDuctId == null || newDuctId == ElementId.InvalidElementId) return false;

                Duct newDuct = doc.GetElement(newDuctId) as Duct;
                if (newDuct == null) return false;

                // Restore parameters to both ducts
                RestoreDuctParameters(originalDuct, ductParameters);
                RestoreDuctParameters(newDuct, ductParameters);

                // Handle insulation for split ducts
                if (originalInsulationData != null && originalInsulationData.HasInsulation)
                {
                    AddInsulationToDuct(doc, newDuct, originalInsulationData);
                    AddInsulationToDuct(doc, originalDuct, originalInsulationData);
                }
                // Handle lining for split ducts
                if (originalLiningData != null && originalLiningData.HasLining)
                {
                    AddLiningToDuct(doc, newDuct, originalLiningData);
                    AddLiningToDuct(doc, originalDuct, originalLiningData);
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        #region Insulation Data Storage and Retrieval
        private class InsulationData
        {
            public bool HasInsulation { get; set; }
            public ElementId InsulationTypeId { get; set; }
            public double Thickness { get; set; }
            public Dictionary<string, object> Parameters { get; set; } = new Dictionary<string, object>();
        }

        private InsulationData GetDuctInsulation(Document doc, Duct duct)
        {
            var data = new InsulationData();
            var insulations = new FilteredElementCollector(doc).OfClass(typeof(DuctInsulation)).Cast<DuctInsulation>().Where(ins => ins.HostElementId == duct.Id).ToList();

            if (insulations.Count == 0)
            {
                data.HasInsulation = false;
                return data;
            }

            DuctInsulation insulation = insulations.First();
            data.HasInsulation = true;
            data.InsulationTypeId = insulation.GetTypeId();
            data.Thickness = insulation.Thickness;
            StoreParameters(insulation, data.Parameters);
            return data;
        }

        private void AddInsulationToDuct(Document doc, Duct duct, InsulationData insulationData)
        {
            try
            {
                // Check if duct already has insulation
                var existingInsulation = new FilteredElementCollector(doc)
                    .OfClass(typeof(DuctInsulation))
                    .Cast<DuctInsulation>()
                    .FirstOrDefault(ins => ins.HostElementId == duct.Id);

                if (existingInsulation != null)
                {
                    // Update existing insulation
                    RestoreParameters(existingInsulation, insulationData.Parameters);
                    return;
                }

                // Create new insulation
                DuctInsulation newInsulation = DuctInsulation.Create(doc, duct.Id, insulationData.InsulationTypeId, insulationData.Thickness);
                if (newInsulation != null)
                {
                    RestoreParameters(newInsulation, insulationData.Parameters);
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Warning", $"Failed to add insulation to duct: {ex.Message}");
            }
        }
        #endregion

        #region Lining Data Storage and Retrieval

        private class LiningData
        {
            public bool HasLining { get; set; }
            public ElementId LiningTypeId { get; set; }
            public double Thickness { get; set; }
            public Dictionary<string, object> Parameters { get; set; } = new Dictionary<string, object>();
        }

        private LiningData GetDuctLining(Document doc, Duct duct)
        {
            LiningData data = new LiningData();
            var linings = new FilteredElementCollector(doc).OfClass(typeof(DuctLining)).Cast<DuctLining>().Where(lining => lining.HostElementId == duct.Id).ToList();

            if (linings.Count == 0)
            {
                data.HasLining = false;
                return data;
            }

            DuctLining lining = linings.First();
            data.HasLining = true;
            data.LiningTypeId = lining.GetTypeId();
            data.Thickness = lining.Thickness;
            StoreParameters(lining, data.Parameters);
            return data;
        }

        private void AddLiningToDuct(Document doc, Duct duct, LiningData liningData)
        {
            try
            {
                // Check if duct already has lining
                var existingLining = new FilteredElementCollector(doc)
                    .OfClass(typeof(DuctLining))
                    .Cast<DuctLining>()
                    .FirstOrDefault(lining => lining.HostElementId == duct.Id);

                if (existingLining != null)
                {
                    // Update existing lining
                    RestoreParameters(existingLining, liningData.Parameters);
                    return;
                }

                // Create new lining
                DuctLining newLining = DuctLining.Create(doc, duct.Id, liningData.LiningTypeId, liningData.Thickness);
                if (newLining != null)
                {
                    RestoreParameters(newLining, liningData.Parameters);
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Warning", $"Failed to add lining to duct: {ex.Message}");
            }
        }
        #endregion

        #region Duct Parameter Handling

        private Dictionary<string, object> StoreDuctParameters(Duct duct)
        {
            var parameters = new Dictionary<string, object>();
            StoreParameters(duct, parameters);
            return parameters;
        }

        private void RestoreDuctParameters(Duct duct, Dictionary<string, object> parameters) => RestoreParameters(duct, parameters);

        private void StoreParameters(Element element, Dictionary<string, object> parameters)
        {
            foreach (Parameter param in element.Parameters)
            {
                if (param == null || param.IsReadOnly || !param.HasValue) continue;
                try
                {
                    string name = param.Definition.Name;
                    if (parameters.ContainsKey(name)) continue;

                    if (param.StorageType == StorageType.Double) parameters[name] = param.AsDouble();
                    else if (param.StorageType == StorageType.Integer) parameters[name] = param.AsInteger();
                    else if (param.StorageType == StorageType.String) parameters[name] = param.AsString();
                    else if (param.StorageType == StorageType.ElementId) parameters[name] = param.AsElementId();
                }
                catch { }
            }
        }
        private void RestoreParameters(Element element, Dictionary<string, object> parameters)
        {
            foreach (var kvp in parameters)
            {
                try
                {
                    Parameter param = element.LookupParameter(kvp.Key);
                    if (param == null || param.IsReadOnly) continue;

                    if (param.StorageType == StorageType.Double && kvp.Value is double) param.Set((double)kvp.Value);
                    else if (param.StorageType == StorageType.Integer && kvp.Value is int) param.Set((int)kvp.Value);
                    else if (param.StorageType == StorageType.String && kvp.Value is string) param.Set((string)kvp.Value);
                    else if (param.StorageType == StorageType.ElementId && kvp.Value is ElementId) param.Set((ElementId)kvp.Value);
                }
                catch { }
            }
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