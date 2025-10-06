using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
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

        private bool SplitColumn(Document doc, Element column, List<Level> levels)
        {
            // Get column location
            LocationPoint locPoint = column.Location as LocationPoint;
            if (locPoint == null)
            {
                throw new InvalidOperationException("Column does not have a valid location point.");
            }

            XYZ basePoint = locPoint.Point;

            // Get current column parameters
            Parameter baseOffsetParam = column.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM);
            Parameter topOffsetParam = column.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM);

            double baseOffset = baseOffsetParam?.AsDouble() ?? 0;
            double topOffset = topOffsetParam?.AsDouble() ?? 0;

            // Get column's base and top levels
            Level baseLevel = doc.GetElement(column.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM).AsElementId()) as Level;
            Level topLevel = doc.GetElement(column.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM).AsElementId()) as Level;

            if (baseLevel == null || topLevel == null)
            {
                throw new InvalidOperationException("Cannot determine column's base or top level.");
            }

            // Calculate actual base and top elevations
            double baseElevation = baseLevel.Elevation + baseOffset;
            double topElevation = topLevel.Elevation + topOffset;

            // Filter levels that are within the column's height range
            List<Level> relevantLevels = levels
                .Where(l => l.Elevation > baseElevation && l.Elevation < topElevation)
                .ToList();

            if (relevantLevels.Count == 0)
            {
                // No levels intersect the column, no need to split
                return true;
            }

            // Add the base and top levels to create complete segments
            List<Level> allLevels = new List<Level> { baseLevel };
            allLevels.AddRange(relevantLevels);
            allLevels.Add(topLevel);
            allLevels = allLevels.OrderBy(l => l.Elevation).ToList();

            // Store original column ID for deletion
            ElementId originalColumnId = column.Id;

            // Create new column segments
            List<ElementId> newColumnIds = new List<ElementId>();
            for (int i = 0; i < allLevels.Count - 1; i++)
            {
                Level currentLevel = allLevels[i];
                Level nextLevel = allLevels[i + 1];

                // Copy the column
                ICollection<ElementId> copiedIds = ElementTransformUtils.CopyElement(
                    doc,
                    originalColumnId,
                    XYZ.Zero); // Copy at same location

                if (copiedIds.Count > 0)
                {
                    ElementId newColumnId = copiedIds.First();
                    Element newColumn = doc.GetElement(newColumnId);

                    // Set base level and offset
                    newColumn.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM).Set(currentLevel.Id);
                    if (i == 0 && currentLevel == baseLevel)
                    {
                        // First segment: maintain original base offset
                        newColumn.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM).Set(baseOffset);
                    }
                    else
                    {
                        // Other segments: no base offset
                        newColumn.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM).Set(0);
                    }

                    // Set top level and offset
                    newColumn.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM).Set(nextLevel.Id);
                    if (i == allLevels.Count - 2 && nextLevel == topLevel)
                    {
                        // Last segment: maintain original top offset
                        newColumn.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM).Set(topOffset);
                    }
                    else
                    {
                        // Other segments: no top offset
                        newColumn.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM).Set(0);
                    }

                    newColumnIds.Add(newColumnId);
                }
            }

            // Delete the original column only if we successfully created segments
            if (newColumnIds.Count > 0)
            {
                doc.Delete(originalColumnId);
                return true;
            }

            return false;
        }
    }
}