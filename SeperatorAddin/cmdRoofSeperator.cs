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
    public class cmdRoofSeperator : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                Utils.RoofSelectionFilter filter = new Utils.RoofSelectionFilter();
                IList<Reference> references = uidoc.Selection.PickObjects(ObjectType.Element, filter, "Select roofs to separate layers");

                if (references.Count == 0)
                {
                    return Result.Cancelled;
                }

                using (Transaction t = new Transaction(doc, "Separate Roof Layers"))
                {
                    t.Start();

                    foreach (Reference reference in references)
                    {
                        RoofBase originalRoof = doc.GetElement(reference) as RoofBase;
                        if (originalRoof != null)
                        {
                            ProcessRoof(doc, originalRoof);
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

        public void ProcessRoof(Document doc, RoofBase originalRoof)
        {
            RoofType originalRoofType = doc.GetElement(originalRoof.GetTypeId()) as RoofType;
            CompoundStructure structure = originalRoofType?.GetCompoundStructure();

            if (structure == null || structure.LayerCount <= 1)
            {
                return;
            }

            Level level = doc.GetElement(originalRoof.LevelId) as Level;

            // Get the base offset of the original roof (offset from level to BOTTOM of roof assembly)
            double originalBaseOffset = originalRoof.get_Parameter(BuiltInParameter.ROOF_LEVEL_OFFSET_PARAM)?.AsDouble() ?? 0.0;

            // Get roof boundary
            IList<CurveLoop> roofBoundary = GetRoofBoundary(originalRoof);

            if (roofBoundary == null || roofBoundary.Count == 0)
            {
                return;
            }

            // Calculate total thickness of all layers
            double totalThickness = structure.GetLayers().Sum(l => l.Width);

            // Calculate the offset of the TOP of the original roof assembly
            double topOffset = originalBaseOffset + totalThickness;

            // Track cumulative thickness as we create layers from top to bottom
            double cumulativeThickness = 0.0;
            int successfulLayers = 0;

            // GetLayers() returns layers from top to bottom (Exterior to Interior)
            for (int i = 0; i < structure.LayerCount; i++)
            {
                CompoundStructureLayer layer = structure.GetLayers()[i];
                double layerThickness = layer.Width;

                RoofType newRoofType = GetOrCreateRoofType(doc, originalRoofType, layer, i);

                if (newRoofType == null)
                {
                    string matName = doc.GetElement(layer.MaterialId)?.Name ?? "Unnamed";
                    System.Diagnostics.Debug.Print($"Skipping layer {i + 1} ('{matName}') because it has properties (e.g., Variable Thickness) that prevent it from being a single-layer roof.");
                    cumulativeThickness += layerThickness;
                    continue;
                }

                // Calculate the correct base offset for the NEW single-layer roof
                // It's the top of the whole assembly, minus the layers we've already created, minus the thickness of the current layer
                double newRoofBaseOffset = topOffset - cumulativeThickness - layerThickness;

                // Create the new roof
                RoofBase newRoof = CreateRoofFromBoundary(doc, roofBoundary, newRoofType.Id, level.Id);

                if (newRoof == null)
                {
                    cumulativeThickness += layerThickness;
                    continue;
                }

                // Set the calculated base offset
                Parameter offsetParam = newRoof.get_Parameter(BuiltInParameter.ROOF_LEVEL_OFFSET_PARAM);
                if (offsetParam != null && !offsetParam.IsReadOnly)
                {
                    offsetParam.Set(newRoofBaseOffset);
                }

                // Copy parameters from original roof
                CopyRoofParameters(originalRoof, newRoof);

                // Update cumulative thickness for next iteration
                cumulativeThickness += layerThickness;
                successfulLayers++;
            }

            if (successfulLayers > 0)
            {
                doc.Delete(originalRoof.Id);
            }
        }

        private void CopyRoofParameters(Element source, Element target)
        {
            var ignoredParams = new List<BuiltInParameter>
            {
                BuiltInParameter.ALL_MODEL_MARK,
                BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM,
                BuiltInParameter.ELEM_FAMILY_PARAM,
                BuiltInParameter.ELEM_TYPE_PARAM,
                BuiltInParameter.ROOF_LEVEL_OFFSET_PARAM,
                BuiltInParameter.HOST_AREA_COMPUTED,
                BuiltInParameter.HOST_VOLUME_COMPUTED,
                BuiltInParameter.ROOF_ATTR_THICKNESS_PARAM,
                BuiltInParameter.ID_PARAM
            };

            foreach (Parameter sourceParam in source.Parameters)
            {
                if (sourceParam.IsReadOnly) continue;

                if (sourceParam.Definition is InternalDefinition internalDef)
                {
                    if (ignoredParams.Contains(internalDef.BuiltInParameter))
                    {
                        continue;
                    }
                }

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
                    catch { }
                }
            }
        }

        private IList<CurveLoop> GetRoofBoundary(RoofBase roof)
        {
            GeometryElement geoElement = roof.get_Geometry(new Options());
            foreach (GeometryObject geoObject in geoElement)
            {
                if (geoObject is Solid solid && solid.Volume > 0)
                {
                    PlanarFace bottomFace = null;
                    double lowestZ = double.MaxValue;

                    foreach (Face face in solid.Faces)
                    {
                        if (face is PlanarFace planarFace)
                        {
                            // For flat roofs, find the face with normal pointing down (bottom face)
                            if (planarFace.FaceNormal.IsAlmostEqualTo(XYZ.BasisZ.Negate()))
                            {
                                if (planarFace.Origin.Z < lowestZ)
                                {
                                    lowestZ = planarFace.Origin.Z;
                                    bottomFace = planarFace;
                                }
                            }
                        }
                    }
                    if (bottomFace != null) return bottomFace.GetEdgesAsCurveLoops();
                }
            }
            return null;
        }

        private RoofBase CreateRoofFromBoundary(Document doc, IList<CurveLoop> boundary, ElementId roofTypeId, ElementId levelId)
        {
            try
            {
                // Convert CurveLoop to CurveArray
                CurveArray curveArray = new CurveArray();
                foreach (CurveLoop loop in boundary)
                {
                    foreach (Curve curve in loop)
                    {
                        curveArray.Append(curve);
                    }
                    break; // Only use the first loop (outer boundary)
                }

                // Get Level and RoofType elements
                Level level = doc.GetElement(levelId) as Level;
                RoofType roofType = doc.GetElement(roofTypeId) as RoofType;

                if (level == null || roofType == null || curveArray.Size == 0)
                {
                    return null;
                }

                // Create a FootPrintRoof using the boundary
                ModelCurveArray footprintCurves = new ModelCurveArray();
                RoofBase newRoof = doc.Create.NewFootPrintRoof(curveArray, level, roofType, out footprintCurves);

                // Set roof to be flat (no slope)
                if (newRoof != null && newRoof is FootPrintRoof footprintRoof)
                {
                    // Set all slope defining lines to have no slope
                    foreach (ModelCurve modelCurve in footprintCurves)
                    {
                        footprintRoof.set_DefinesSlope(modelCurve, false);
                    }
                }

                return newRoof;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print($"Failed to create roof: {ex.Message}");
                return null;
            }
        }

        private RoofType GetOrCreateRoofType(Document doc, RoofType originalRoofType, CompoundStructureLayer layer, int layerIndex)
        {
            if (originalRoofType.GetCompoundStructure().VariableLayerIndex == layerIndex)
            {
                return null;
            }

            string materialName = doc.GetElement(layer.MaterialId)?.Name ?? "Unnamed";
            string newTypeName = $"{originalRoofType.Name}_Layer_{layerIndex + 1}_{materialName}";

            RoofType newRoofType = new FilteredElementCollector(doc)
                .OfClass(typeof(RoofType))
                .Cast<RoofType>()
                .FirstOrDefault(e => e.Name == newTypeName);

            if (newRoofType == null)
            {
                try
                {
                    newRoofType = originalRoofType.Duplicate(newTypeName) as RoofType;
                    CompoundStructure cs = newRoofType.GetCompoundStructure();

                    cs.SetLayers(new List<CompoundStructureLayer> { layer });
                    cs.SetNumberOfShellLayers(ShellLayerType.Exterior, 0);
                    cs.SetNumberOfShellLayers(ShellLayerType.Interior, 0);

                    newRoofType.SetCompoundStructure(cs);
                }
                catch (Exception)
                {
                    return null;
                }
            }

            return newRoofType;
        }
    }
}