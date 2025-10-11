using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using SeperatorAddin.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace SeperatorAddin
{
    [Transaction(TransactionMode.Manual)]
    public class cmdFloorSeperator : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Let the user select floors
                Utils.FloorSelectionFilter filter = new Utils.FloorSelectionFilter();
                IList<Reference> references = uidoc.Selection.PickObjects(ObjectType.Element, filter, "Select floors to separate layers");

                if (references.Count == 0)
                {
                    return Result.Cancelled;
                }

                using (Transaction t = new Transaction(doc, "Separate Floor Layers"))
                {
                    t.Start();

                    foreach (Reference reference in references)
                    {
                        Floor originalFloor = doc.GetElement(reference) as Floor;
                        if (originalFloor != null)
                        {
                            ProcessFloor(doc, originalFloor);
                        }
                    }

                    t.Commit();
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        public void ProcessFloor(Document doc, Floor originalFloor)
        {
            FloorType originalFloorType = originalFloor.FloorType;
            CompoundStructure structure = originalFloorType.GetCompoundStructure();

            if (structure == null || structure.LayerCount <= 1)
            {
                return;
            }

            Level level = doc.GetElement(originalFloor.LevelId) as Level;
            double heightOffset = originalFloor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM).AsDouble();

            CurveLoop floorBoundary = null;
            Solid floorSolid = GetFloorSolid(originalFloor);
            if (floorSolid != null)
            {
                PlanarFace topFace = GetTopFace(floorSolid);
                if (topFace != null)
                {
                    floorBoundary = topFace.GetEdgesAsCurveLoops().FirstOrDefault();
                }
            }

            if (floorBoundary == null)
            {
                return;
            }

            double currentOffset = 0.0;
            int successfulLayers = 0;

            for (int i = 0; i < structure.LayerCount; i++)
            {
                CompoundStructureLayer layer = structure.GetLayers()[i];
                double layerThickness = layer.Width;

                FloorType newFloorType = GetOrCreateFloorType(doc, originalFloorType, layer, i);

                if (newFloorType == null)
                {
                    string matName = doc.GetElement(layer.MaterialId)?.Name ?? "Unnamed";
                    System.Diagnostics.Debug.Print($"Skipping layer {i + 1} ('{matName}') because it has properties (e.g., Variable Thickness) that prevent it from being a single-layer floor.");
                    currentOffset += layerThickness;
                    continue;
                }

                double newFloorOffset = heightOffset - currentOffset;

                Floor newFloor = Floor.Create(doc, new List<CurveLoop> { floorBoundary }, newFloorType.Id, level.Id);

                Parameter heightParam = newFloor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM);
                if (heightParam != null && !heightParam.IsReadOnly)
                {
                    heightParam.Set(newFloorOffset);
                }

                // --- NEW CODE: Copy parameters from the original floor to the new one ---
                CopyFloorParameters(originalFloor, newFloor);

                successfulLayers++;
                currentOffset += layerThickness;
            }

            if (successfulLayers > 0)
            {
                doc.Delete(originalFloor.Id);
            }
        }

        /// <summary>
        /// Copies instance parameters from a source element to a target element.
        /// </summary>
        /// <param name="source">The element to copy parameters from.</param>
        /// <param name="target">The element to copy parameters to.</param>
        private void CopyFloorParameters(Element source, Element target)
        {
            // List of parameters to explicitly ignore
            var ignoredParams = new List<BuiltInParameter>
    {
        // --- ADDED THIS LINE ---
        BuiltInParameter.ALL_MODEL_MARK, // This is the "Mark" parameter

        // Parameters related to type, which we are intentionally changing
        BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM,
        BuiltInParameter.ELEM_FAMILY_PARAM,
        BuiltInParameter.ELEM_TYPE_PARAM,

        // Geometric/Host parameters that are set uniquely for each new floor
        BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM,
        BuiltInParameter.HOST_AREA_COMPUTED,
        BuiltInParameter.HOST_VOLUME_COMPUTED,
        BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM,
        
        // Read-only identifiers
        BuiltInParameter.ID_PARAM
    };

            foreach (Parameter sourceParam in source.Parameters)
            {
                // Skip read-only parameters
                if (sourceParam.IsReadOnly) continue;

                // Check if the parameter is in the ignored list by its BuiltInParameter enum
                if (sourceParam.Definition is InternalDefinition internalDef)
                {
                    if (ignoredParams.Contains(internalDef.BuiltInParameter))
                    {
                        continue;
                    }
                }

                Parameter targetParam = target.get_Parameter(sourceParam.Definition);

                // If the target has the same parameter and it's not read-only, copy the value
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
                        // Fails silently if a specific parameter can't be set
                    }
                }
            }
        }


        // (The rest of the methods: GetFloorSolid, GetTopFace, GetOrCreateFloorType, GetButtonData remain the same)

        private Solid GetFloorSolid(Floor floor)
        {
            GeometryElement geoElement = floor.get_Geometry(new Options());
            foreach (GeometryObject geoObject in geoElement)
            {
                if (geoObject is Solid solid && solid.Volume > 0)
                {
                    return solid;
                }
            }
            return null;
        }

        private PlanarFace GetTopFace(Solid solid)
        {
            PlanarFace topFace = null;
            double highestZ = double.MinValue;

            foreach (Face face in solid.Faces)
            {
                if (face is PlanarFace planarFace)
                {
                    if (planarFace.FaceNormal.IsAlmostEqualTo(XYZ.BasisZ))
                    {
                        if (planarFace.Origin.Z > highestZ)
                        {
                            highestZ = planarFace.Origin.Z;
                            topFace = planarFace;
                        }
                    }
                }
            }
            return topFace;
        }

        private FloorType GetOrCreateFloorType(Document doc, FloorType originalFloorType, CompoundStructureLayer layer, int layerIndex)
        {
            if (originalFloorType.GetCompoundStructure().VariableLayerIndex == layerIndex)
            {
                return null;
            }

            string materialName = "Unnamed";
            if (layer.MaterialId != ElementId.InvalidElementId)
            {
                Material material = doc.GetElement(layer.MaterialId) as Material;
                if (material != null)
                {
                    materialName = material.Name;
                }
            }

            string newTypeName = $"{originalFloorType.Name}_Layer_{layerIndex + 1}_{materialName}";

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(FloorType));
            FloorType newFloorType = collector.FirstOrDefault(e => e.Name == newTypeName) as FloorType;

            if (newFloorType == null)
            {
                try
                {
                    newFloorType = originalFloorType.Duplicate(newTypeName) as FloorType;
                    CompoundStructure cs = newFloorType.GetCompoundStructure();

                    cs.SetLayers(new List<CompoundStructureLayer> { layer });
                    cs.SetNumberOfShellLayers(ShellLayerType.Exterior, 0);
                    cs.SetNumberOfShellLayers(ShellLayerType.Interior, 0);
                    cs.EndCap = EndCapCondition.NoEndCap;
                    cs.OpeningWrapping = OpeningWrappingCondition.None;

                    newFloorType.SetCompoundStructure(cs);
                }
                catch (Exception)
                {
                    return null;
                }
            }

            return newFloorType;
        }
    }
}