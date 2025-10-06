using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SeperatorAddin
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class cmdColumnSplit : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            TaskDialog.Show("Command Info", "This function is now part of the 'Vertical Split' tool.");
            return Result.Succeeded;
        }

        public bool SplitColumn(Document doc, Element column, List<Level> selectedLevels)
        {
            try
            {
                if (!(column is FamilyInstance famInst) || famInst.Location as LocationPoint == null)
                {
                    return false; // Not a valid vertical column
                }

                // Get original parameters
                Parameter baseOffsetParam = column.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM);
                Parameter topOffsetParam = column.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM);
                double baseOffset = baseOffsetParam?.AsDouble() ?? 0.0;
                double topOffset = topOffsetParam?.AsDouble() ?? 0.0;

                Level baseLevel = doc.GetElement(column.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM).AsElementId()) as Level;
                Level topLevel = doc.GetElement(column.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM).AsElementId()) as Level;

                if (baseLevel == null || topLevel == null) return false;
                
                double baseElevation = baseLevel.Elevation + baseOffset;
                double topElevation = topLevel.Elevation + topOffset;

                // Build a list of valid split points including the column's actual start and end levels
                List<Level> splitLevels = new List<Level> { baseLevel };
                foreach (Level selectedLevel in selectedLevels)
                {
                    if (selectedLevel.Elevation > baseElevation + 0.001 && selectedLevel.Elevation < topElevation - 0.001)
                    {
                        splitLevels.Add(selectedLevel);
                    }
                }
                splitLevels.Add(topLevel);

                splitLevels = splitLevels.Distinct().OrderBy(l => l.Elevation).ToList();

                if (splitLevels.Count < 2) return false; // No split occurred

                ElementId originalColumnId = column.Id;
                var newColumnIds = new List<ElementId>();

                // Create new segments
                for (int i = 0; i < splitLevels.Count - 1; i++)
                {
                    Level currentBaseLevel = splitLevels[i];
                    Level currentTopLevel = splitLevels[i + 1];

                    // Use CopyElement which is more robust
                    ICollection<ElementId> copiedIds = ElementTransformUtils.CopyElement(doc, originalColumnId, XYZ.Zero);
                    if (!copiedIds.Any()) continue;
                    
                    Element newColumn = doc.GetElement(copiedIds.First());

                    // Copy all parameters from the original before modifying levels
                    CopyColumnParameters(column, newColumn);

                    // Set base level and offset
                    newColumn.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM).Set(currentBaseLevel.Id);
                    double currentBaseOffset = (currentBaseLevel.Id == baseLevel.Id) ? baseOffset : 0.0;
                    newColumn.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM).Set(currentBaseOffset);

                    // Set top level and offset
                    newColumn.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM).Set(currentTopLevel.Id);
                    double currentTopOffset = (currentTopLevel.Id == topLevel.Id) ? topOffset : 0.0;
                    newColumn.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM).Set(currentTopOffset);

                    newColumnIds.Add(newColumn.Id);
                }

                if (newColumnIds.Any())
                {
                    doc.Delete(originalColumnId);
                    return true;
                }
                
                return false;
            }
            catch
            {
                return false;
            }
        }
        
        /// <summary>
        /// Copies all writable parameters from a source element to a target element.
        /// </summary>
        private void CopyColumnParameters(Element source, Element target)
        {
            var ignoredBips = new List<BuiltInParameter>
            {
                BuiltInParameter.FAMILY_BASE_LEVEL_PARAM,
                BuiltInParameter.FAMILY_TOP_LEVEL_PARAM,
                BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM,
                BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM,
                BuiltInParameter.SCHEDULE_LEVEL_PARAM
            };

            foreach (Parameter sourceParam in source.Parameters)
            {
                if (sourceParam.IsReadOnly || !sourceParam.HasValue) continue;

                if (sourceParam.Definition is InternalDefinition def)
                {
                    if (ignoredBips.Contains((BuiltInParameter)def.BuiltInParameter))
                    {
                        continue;
                    }
                }

                Parameter targetParam = target.get_Parameter(sourceParam.Definition);
                if (targetParam != null && !targetParam.IsReadOnly)
                {
                    try
                    {
                        if (sourceParam.StorageType == StorageType.String) targetParam.Set(sourceParam.AsString());
                        else if (sourceParam.StorageType == StorageType.Double) targetParam.Set(sourceParam.AsDouble());
                        else if (sourceParam.StorageType == StorageType.Integer) targetParam.Set(sourceParam.AsInteger());
                        else if (sourceParam.StorageType == StorageType.ElementId) targetParam.Set(sourceParam.AsElementId());
                    }
                    catch { } // Ignore parameters that can't be set
                }
            }
        }
    }
}