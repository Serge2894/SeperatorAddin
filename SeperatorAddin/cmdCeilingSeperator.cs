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
    public class cmdCeilingSeperator : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                Utils.CeilingSelectionFilter filter = new Utils.CeilingSelectionFilter();
                IList<Reference> references = uidoc.Selection.PickObjects(ObjectType.Element, filter, "Select ceilings to separate layers");

                if (references.Count == 0)
                {
                    return Result.Cancelled;
                }

                using (Transaction t = new Transaction(doc, "Separate Ceiling Layers"))
                {
                    t.Start();

                    foreach (Reference reference in references)
                    {
                        Ceiling originalCeiling = doc.GetElement(reference) as Ceiling;
                        if (originalCeiling != null)
                        {
                            ProcessCeiling(doc, originalCeiling);
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

        public void ProcessCeiling(Document doc, Ceiling originalCeiling)
        {
            CeilingType originalCeilingType = doc.GetElement(originalCeiling.GetTypeId()) as CeilingType;
            CompoundStructure structure = originalCeilingType.GetCompoundStructure();

            if (structure == null || structure.LayerCount <= 1)
            {
                return;
            }

            Level level = doc.GetElement(originalCeiling.LevelId) as Level;
            // This is the offset of the BOTTOM of the original ceiling assembly
            double originalBottomOffset = originalCeiling.get_Parameter(BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM).AsDouble();

            IList<CurveLoop> ceilingBoundary = GetCeilingBoundary(originalCeiling);

            if (ceilingBoundary == null || ceilingBoundary.Count == 0)
            {
                return;
            }

            // --- CORRECTED LOGIC ---
            // 1. Get the total thickness of all layers.
            double totalThickness = structure.GetLayers().Sum(l => l.Width);

            // 2. Calculate the offset of the TOP of the original ceiling assembly.
            double topOffset = originalBottomOffset + totalThickness;

            // 3. This will track the thickness of layers as we create them, stacking downwards.
            double cumulativeThickness = 0.0;
            int successfulLayers = 0;

            // The GetLayers() method returns layers from top to bottom (Exterior to Interior).
            for (int i = 0; i < structure.LayerCount; i++)
            {
                CompoundStructureLayer layer = structure.GetLayers()[i];
                double layerThickness = layer.Width;

                CeilingType newCeilingType = GetOrCreateCeilingType(doc, originalCeilingType, layer, i);

                if (newCeilingType == null)
                {
                    string matName = doc.GetElement(layer.MaterialId)?.Name ?? "Unnamed";
                    System.Diagnostics.Debug.Print($"Skipping layer {i + 1} ('{matName}') because it has properties (e.g., Variable Thickness) that prevent it from being a single-layer ceiling.");
                    cumulativeThickness += layerThickness; // Still need to account for thickness
                    continue;
                }

                // 4. Calculate the correct bottom offset for the NEW single-layer ceiling.
                // It's the top of the whole assembly, minus the layers we've already created, minus the thickness of the current layer.
                double newCeilingBottomOffset = topOffset - cumulativeThickness - layerThickness;

                // 5. Create the new ceiling.
                Ceiling newCeiling = Ceiling.Create(doc, ceilingBoundary, newCeilingType.Id, level.Id);

                // 6. Set its calculated height offset.
                Parameter heightParam = newCeiling.get_Parameter(BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM);
                if (heightParam != null && !heightParam.IsReadOnly)
                {
                    heightParam.Set(newCeilingBottomOffset);
                }

                CopyCeilingParameters(originalCeiling, newCeiling);

                // 7. Update the cumulative thickness for the next iteration.
                cumulativeThickness += layerThickness;
                successfulLayers++;
            }

            if (successfulLayers > 0)
            {
                doc.Delete(originalCeiling.Id);
            }
        }

        private void CopyCeilingParameters(Element source, Element target)
        {
            var ignoredParams = new List<BuiltInParameter>
            {
                BuiltInParameter.ALL_MODEL_MARK,
                BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM,
                BuiltInParameter.ELEM_FAMILY_PARAM,
                BuiltInParameter.ELEM_TYPE_PARAM,
                BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM,
                BuiltInParameter.HOST_AREA_COMPUTED,
                BuiltInParameter.HOST_VOLUME_COMPUTED,
                BuiltInParameter.CEILING_THICKNESS,
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
                    switch (sourceParam.StorageType)
                    {
                        case StorageType.Double: targetParam.Set(sourceParam.AsDouble()); break;
                        case StorageType.Integer: targetParam.Set(sourceParam.AsInteger()); break;
                        case StorageType.String: targetParam.Set(sourceParam.AsString()); break;
                        case StorageType.ElementId: targetParam.Set(sourceParam.AsElementId()); break;
                    }
                }
            }
        }

        private IList<CurveLoop> GetCeilingBoundary(Ceiling ceiling)
        {
            GeometryElement geoElement = ceiling.get_Geometry(new Options());
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

        private CeilingType GetOrCreateCeilingType(Document doc, CeilingType originalCeilingType, CompoundStructureLayer layer, int layerIndex)
        {
            if (originalCeilingType.GetCompoundStructure().VariableLayerIndex == layerIndex)
            {
                return null;
            }

            string materialName = doc.GetElement(layer.MaterialId)?.Name ?? "Unnamed";
            string newTypeName = $"{originalCeilingType.Name}_Layer_{layerIndex + 1}_{materialName}";

            CeilingType newCeilingType = new FilteredElementCollector(doc)
                .OfClass(typeof(CeilingType))
                .Cast<CeilingType>()
                .FirstOrDefault(e => e.Name == newTypeName);

            if (newCeilingType == null)
            {
                try
                {
                    newCeilingType = originalCeilingType.Duplicate(newTypeName) as CeilingType;
                    CompoundStructure cs = newCeilingType.GetCompoundStructure();

                    cs.SetLayers(new List<CompoundStructureLayer> { layer });
                    cs.SetNumberOfShellLayers(ShellLayerType.Exterior, 0);
                    cs.SetNumberOfShellLayers(ShellLayerType.Interior, 0);

                    newCeilingType.SetCompoundStructure(cs);
                }
                catch (Exception)
                {
                    return null;
                }
            }

            return newCeilingType;
        }
    }
}