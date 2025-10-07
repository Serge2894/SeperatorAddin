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

        /// <summary>
        /// Splits a vertical column (both architectural and structural) at each of the specified levels.
        /// </summary>
        public bool SplitColumn(Document doc, Element column, List<Level> levels)
        {
            var familyInstance = column as FamilyInstance;
            if (familyInstance == null) return false;

            var locPoint = column.Location as LocationPoint;
            if (locPoint == null) return false;

            // Get the correct parameters for the column, whether it's architectural or structural
            var columnParameters = GetColumnParameters(column);
            if (columnParameters == null) return false; // Column type not supported

            var baseLevel = doc.GetElement(columnParameters.BaseLevelParam.AsElementId()) as Level;
            var topLevel = doc.GetElement(columnParameters.TopLevelParam.AsElementId()) as Level;
            if (baseLevel == null || topLevel == null) return false;

            double baseElevation = baseLevel.ProjectElevation + columnParameters.BaseOffsetParam.AsDouble();
            double topElevation = topLevel.ProjectElevation + columnParameters.TopOffsetParam.AsDouble();

            // Create a sorted list of elevations to split at
            var splitElevations = new List<double> { baseElevation };
            foreach (var level in levels)
            {
                if (level.ProjectElevation > baseElevation + 0.001 && level.ProjectElevation < topElevation - 0.001)
                {
                    splitElevations.Add(level.ProjectElevation);
                }
            }
            splitElevations.Add(topElevation);
            var distinctElevations = splitElevations.Distinct().OrderBy(e => e).ToList();

            if (distinctElevations.Count < 2) return false;

            var newColumns = new List<FamilyInstance>();
            var allLevels = levels.Concat(new[] { baseLevel, topLevel }).Distinct().ToList();

            // Create new column segments
            for (int i = 0; i < distinctElevations.Count - 1; i++)
            {
                double bottomElev = distinctElevations[i];
                double topElev = distinctElevations[i + 1];

                Level newBaseLevel = GetBestHostLevel(allLevels, bottomElev);
                Level newTopLevel = GetBestHostLevel(allLevels, topElev);
                if (newBaseLevel == null || newTopLevel == null) continue;

                double newBaseOffset = bottomElev - newBaseLevel.ProjectElevation;
                double newTopOffset = topElev - newTopLevel.ProjectElevation;

                var newColumn = doc.Create.NewFamilyInstance(locPoint.Point, familyInstance.Symbol, newBaseLevel, familyInstance.StructuralType) as FamilyInstance;

                if (newColumn != null)
                {
                    // Use the same parameter definitions we found earlier to set the new values
                    newColumn.get_Parameter(columnParameters.TopLevelParam.Definition)?.Set(newTopLevel.Id);
                    newColumn.get_Parameter(columnParameters.BaseOffsetParam.Definition)?.Set(newBaseOffset);
                    newColumn.get_Parameter(columnParameters.TopOffsetParam.Definition)?.Set(newTopOffset);
                    newColumns.Add(newColumn);
                }
            }

            if (newColumns.Any())
            {
                doc.Delete(column.Id);
                return true;
            }

            return false;
        }

        /// <summary>
        /// A helper class to hold the correct parameters for a column.
        /// </summary>
        private class ColumnParameterSet
        {
            public Parameter BaseLevelParam { get; set; }
            public Parameter TopLevelParam { get; set; }
            public Parameter BaseOffsetParam { get; set; }
            public Parameter TopOffsetParam { get; set; }
        }

        /// <summary>
        /// Finds the correct geometry-driving parameters for either an architectural or structural column.
        /// </summary>
        private ColumnParameterSet GetColumnParameters(Element column)
        {
            // First, try the parameters typically used by Structural Columns
            var structuralParams = new ColumnParameterSet
            {
                BaseLevelParam = column.get_Parameter(BuiltInParameter.SCHEDULE_BASE_LEVEL_PARAM),
                TopLevelParam = column.get_Parameter(BuiltInParameter.SCHEDULE_TOP_LEVEL_PARAM),
                BaseOffsetParam = column.get_Parameter(BuiltInParameter.SCHEDULE_BASE_LEVEL_OFFSET_PARAM),
                TopOffsetParam = column.get_Parameter(BuiltInParameter.SCHEDULE_TOP_LEVEL_OFFSET_PARAM)
            };

            if (structuralParams.BaseLevelParam != null && structuralParams.TopLevelParam != null && structuralParams.BaseOffsetParam != null && structuralParams.TopOffsetParam != null)
            {
                return structuralParams;
            }

            // If that fails, try the parameters for Architectural Columns
            var architecturalParams = new ColumnParameterSet
            {
                BaseLevelParam = column.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM),
                TopLevelParam = column.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM),
                BaseOffsetParam = column.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM),
                TopOffsetParam = column.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM)
            };

            if (architecturalParams.BaseLevelParam != null && architecturalParams.TopLevelParam != null && architecturalParams.BaseOffsetParam != null && architecturalParams.TopOffsetParam != null)
            {
                return architecturalParams;
            }

            // If neither set of parameters is found, this column type is not supported
            return null;
        }

        /// <summary>
        /// Finds the highest level that is at or below the given elevation.
        /// </summary>
        private Level GetBestHostLevel(IEnumerable<Level> levels, double elevation)
        {
            return levels
                .Where(l => l.ProjectElevation <= elevation + 0.001)
                .OrderByDescending(l => l.ProjectElevation)
                .FirstOrDefault();
        }
    }
}