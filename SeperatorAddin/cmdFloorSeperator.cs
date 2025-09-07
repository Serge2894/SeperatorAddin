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

        private void ProcessFloor(Document doc, Floor originalFloor)
        {
            FloorType originalFloorType = originalFloor.FloorType;
            CompoundStructure structure = originalFloorType.GetCompoundStructure();

            if (structure == null || structure.LayerCount <= 1)
            {
                // Not a compound floor or has only one layer, so skip it.
                return;
            }

            Level level = doc.GetElement(originalFloor.LevelId) as Level;
            double heightOffset = originalFloor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM).AsDouble();

            // Get the boundary of the floor
            CurveLoop floorBoundary = null;
            Solid floorSolid = GetFloorSolid(originalFloor);
            if (floorSolid != null)
            {
                PlanarFace topFace = GetTopFace(floorSolid);
                if (topFace != null)
                {
                    // Assuming the first loop is the exterior boundary
                    floorBoundary = topFace.GetEdgesAsCurveLoops().FirstOrDefault();
                }
            }

            if (floorBoundary == null)
            {
                // Could not determine floor boundary, skip this floor
                return;
            }


            double currentOffset = 0.0;

            // We iterate from top to bottom
            for (int i = 0; i < structure.LayerCount; i++)
            {
                CompoundStructureLayer layer = structure.GetLayers()[i];
                double layerThickness = layer.Width;

                // Create a new floor type for this layer
                FloorType newFloorType = GetOrCreateFloorType(doc, originalFloorType, layer, i);

                // The first new floor is created at the original offset.
                // Subsequent floors are offset by the thickness of the layers above them.
                double newFloorOffset = heightOffset - currentOffset;

                // Create the new floor
                Floor newFloor = Floor.Create(doc, new List<CurveLoop> { floorBoundary }, newFloorType.Id, level.Id);

                // Set the height offset for the new floor
                Parameter heightParam = newFloor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM);
                if (heightParam != null && !heightParam.IsReadOnly)
                {
                    heightParam.Set(newFloorOffset);
                }

                currentOffset += layerThickness;
            }

            // After creating all the new floors, delete the original one
            doc.Delete(originalFloor.Id);
        }

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

            // Check if this floor type already exists
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(FloorType));
            FloorType newFloorType = collector.FirstOrDefault(e => e.Name == newTypeName) as FloorType;

            if (newFloorType == null)
            {
                // It doesn't exist, so create it
                newFloorType = originalFloorType.Duplicate(newTypeName) as FloorType;

                // Create a new compound structure with just this one layer
                CompoundStructure newStructure = CompoundStructure.CreateSimpleCompoundStructure(new List<CompoundStructureLayer> { layer });
                newFloorType.SetCompoundStructure(newStructure);
            }

            return newFloorType;
        }


        internal static PushButtonData GetButtonData()
        {
            string buttonInternalName = "cmdFloorSeperator";
            string buttonTitle = "Separate\nFloors";

            ButtonDataClass myButtonData = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "Separates a multi-layer floor into individual floors for each layer.");

            return myButtonData.Data;
        }
    }
}