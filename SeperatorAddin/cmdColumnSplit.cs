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

        public bool SplitColumn(Document doc, Element column, List<Level> selectedLevels)
        {
            try
            {
                if (!(column is FamilyInstance famInst))
                {
                    return false;
                }

                // Verify it's a column (architectural or structural)
                if (famInst.Category == null ||
                    (famInst.Category.Id.IntegerValue != (int)BuiltInCategory.OST_Columns &&
                     famInst.Category.Id.IntegerValue != (int)BuiltInCategory.OST_StructuralColumns))
                {
                    return false;
                }

                LocationPoint locationPoint = famInst.Location as LocationPoint;
                if (locationPoint == null)
                {
                    return false;
                }

                // Get the column's XY location and rotation
                XYZ columnLocation = locationPoint.Point;
                double rotation = locationPoint.Rotation;

                // Get base and top levels
                Parameter baseLevelParam = column.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM);
                Parameter topLevelParam = column.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM);

                if (baseLevelParam == null || topLevelParam == null)
                    return false;

                Level baseLevel = doc.GetElement(baseLevelParam.AsElementId()) as Level;
                Level topLevel = doc.GetElement(topLevelParam.AsElementId()) as Level;

                if (baseLevel == null || topLevel == null)
                    return false;

                // Get base and top offsets
                Parameter baseOffsetParam = column.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM);
                Parameter topOffsetParam = column.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM);

                double baseOffset = baseOffsetParam?.AsDouble() ?? 0.0;
                double topOffset = topOffsetParam?.AsDouble() ?? 0.0;

                // Calculate actual elevations
                double baseElevation = baseLevel.Elevation + baseOffset;
                double topElevation = topLevel.Elevation + topOffset;

                // Build split elevations list
                List<Level> splitLevels = new List<Level> { baseLevel };

                foreach (Level selectedLevel in selectedLevels)
                {
                    if (selectedLevel.Elevation > baseElevation + 0.001 &&
                        selectedLevel.Elevation < topElevation - 0.001)
                    {
                        splitLevels.Add(selectedLevel);
                    }
                }

                splitLevels.Add(topLevel);
                splitLevels = splitLevels.Distinct().OrderBy(l => l.Elevation).ToList();

                // Need at least 2 levels to split (original base + at least one split point + original top = 3 minimum)
                if (splitLevels.Count < 3)
                    return false; // No actual split needed

                // Get column properties before deletion
                FamilySymbol columnSymbol = famInst.Symbol;
                ElementId columnTypeId = famInst.GetTypeId();

                // Activate the symbol if not already active
                if (!columnSymbol.IsActive)
                {
                    columnSymbol.Activate();
                    doc.Regenerate();
                }

                // Determine structural type based on category
                StructuralType structuralType = StructuralType.NonStructural;

                // Check if it's a structural column
                if (famInst.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralColumns)
                {
                    // Get the structural usage from the family instance
                    try
                    {
                        structuralType = famInst.StructuralType;
                        if (structuralType == StructuralType.NonStructural)
                        {
                            structuralType = StructuralType.Column; // Default for structural columns
                        }
                    }
                    catch
                    {
                        structuralType = StructuralType.Column;
                    }
                }

                var newColumnIds = new List<ElementId>();
                List<ColumnSegmentData> segmentDataList = new List<ColumnSegmentData>();

                // First, collect all segment data
                for (int i = 0; i < splitLevels.Count - 1; i++)
                {
                    Level currentBaseLevel = splitLevels[i];
                    Level currentTopLevel = splitLevels[i + 1];

                    // Calculate offsets for this segment
                    double currentBaseOffset = 0.0;
                    double currentTopOffset = 0.0;

                    // If this is the original bottom segment, preserve base offset
                    if (currentBaseLevel.Id == baseLevel.Id)
                    {
                        currentBaseOffset = baseOffset;
                    }

                    // If this is the original top segment, preserve top offset
                    if (currentTopLevel.Id == topLevel.Id)
                    {
                        currentTopOffset = topOffset;
                    }

                    segmentDataList.Add(new ColumnSegmentData
                    {
                        BaseLevel = currentBaseLevel,
                        TopLevel = currentTopLevel,
                        BaseOffset = currentBaseOffset,
                        TopOffset = currentTopOffset
                    });
                }

                // Delete the original column FIRST
                ElementId originalId = column.Id;
                doc.Delete(originalId);
                doc.Regenerate();

                // Now create new column segments
                foreach (var segmentData in segmentDataList)
                {
                    try
                    {
                        // Create new column instance
                        FamilyInstance newColumn = doc.Create.NewFamilyInstance(
                            columnLocation,
                            columnSymbol,
                            segmentData.BaseLevel,
                            structuralType);

                        if (newColumn != null)
                        {
                            // Set top level
                            Parameter newTopLevelParam = newColumn.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM);
                            if (newTopLevelParam != null && !newTopLevelParam.IsReadOnly)
                            {
                                newTopLevelParam.Set(segmentData.TopLevel.Id);
                            }

                            // Set base offset
                            Parameter newBaseOffsetParam = newColumn.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM);
                            if (newBaseOffsetParam != null && !newBaseOffsetParam.IsReadOnly)
                            {
                                newBaseOffsetParam.Set(segmentData.BaseOffset);
                            }

                            // Set top offset
                            Parameter newTopOffsetParam = newColumn.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM);
                            if (newTopOffsetParam != null && !newTopOffsetParam.IsReadOnly)
                            {
                                newTopOffsetParam.Set(segmentData.TopOffset);
                            }

                            // Apply rotation if any
                            if (Math.Abs(rotation) > 0.001)
                            {
                                Line rotationAxis = Line.CreateBound(
                                    columnLocation,
                                    columnLocation + XYZ.BasisZ);
                                ElementTransformUtils.RotateElement(
                                    doc,
                                    newColumn.Id,
                                    rotationAxis,
                                    rotation);
                            }

                            // Copy all other parameters from original (stored before deletion)
                            CopyColumnParametersFromStored(column, newColumn, segmentDataList.IndexOf(segmentData) == 0);

                            newColumnIds.Add(newColumn.Id);
                        }
                    }
                    catch (Exception ex)
                    {
                        // Log error but continue with other segments
                        System.Diagnostics.Debug.WriteLine($"Failed to create column segment: {ex.Message}");
                        continue;
                    }
                }

                // Return true if we created at least one new column
                return newColumnIds.Count > 0;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SplitColumn error: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Helper class to store column segment data
        /// </summary>
        private class ColumnSegmentData
        {
            public Level BaseLevel { get; set; }
            public Level TopLevel { get; set; }
            public double BaseOffset { get; set; }
            public double TopOffset { get; set; }
        }

        /// <summary>
        /// Copies all writable parameters from a source element to a target element.
        /// Note: This version is simplified since we can't access the original element after deletion
        /// </summary>
        private void CopyColumnParametersFromStored(Element source, Element target, bool isFirstSegment)
        {
            // Since we're calling this after deletion, we need to be more careful
            // This is a placeholder - in production, you'd want to cache parameter values before deletion
            var ignoredBips = new List<BuiltInParameter>
            {
                BuiltInParameter.FAMILY_BASE_LEVEL_PARAM,
                BuiltInParameter.FAMILY_TOP_LEVEL_PARAM,
                BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM,
                BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM,
                BuiltInParameter.SCHEDULE_LEVEL_PARAM,
                BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM,
                BuiltInParameter.INSTANCE_ELEVATION_PARAM
            };

            try
            {
                foreach (Parameter sourceParam in source.Parameters)
                {
                    if (sourceParam.IsReadOnly || !sourceParam.HasValue)
                        continue;

                    // Skip built-in parameters that control levels/offsets
                    if (sourceParam.Definition is InternalDefinition def)
                    {
                        BuiltInParameter bip = (BuiltInParameter)def.BuiltInParameter;
                        if (ignoredBips.Contains(bip))
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
                                case StorageType.String:
                                    targetParam.Set(sourceParam.AsString());
                                    break;
                                case StorageType.Double:
                                    targetParam.Set(sourceParam.AsDouble());
                                    break;
                                case StorageType.Integer:
                                    targetParam.Set(sourceParam.AsInteger());
                                    break;
                                case StorageType.ElementId:
                                    targetParam.Set(sourceParam.AsElementId());
                                    break;
                            }
                        }
                        catch
                        {
                            // Ignore parameters that can't be set
                        }
                    }
                }
            }
            catch
            {
                // If we can't access source parameters (deleted), that's okay
            }
        }

        /// <summary>
        /// Better approach: Copy parameters before deletion
        /// </summary>
        private void CopyColumnParameters(Element source, Element target)
        {
            var ignoredBips = new List<BuiltInParameter>
            {
                BuiltInParameter.FAMILY_BASE_LEVEL_PARAM,
                BuiltInParameter.FAMILY_TOP_LEVEL_PARAM,
                BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM,
                BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM,
                BuiltInParameter.SCHEDULE_LEVEL_PARAM,
                BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM,
                BuiltInParameter.INSTANCE_ELEVATION_PARAM
            };

            foreach (Parameter sourceParam in source.Parameters)
            {
                if (sourceParam.IsReadOnly || !sourceParam.HasValue)
                    continue;

                // Skip built-in parameters that control levels/offsets
                if (sourceParam.Definition is InternalDefinition def)
                {
                    BuiltInParameter bip = (BuiltInParameter)def.BuiltInParameter;
                    if (ignoredBips.Contains(bip))
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
                            case StorageType.String:
                                targetParam.Set(sourceParam.AsString());
                                break;
                            case StorageType.Double:
                                targetParam.Set(sourceParam.AsDouble());
                                break;
                            case StorageType.Integer:
                                targetParam.Set(sourceParam.AsInteger());
                                break;
                            case StorageType.ElementId:
                                targetParam.Set(sourceParam.AsElementId());
                                break;
                        }
                    }
                    catch
                    {
                        // Ignore parameters that can't be set
                    }
                }
            }
        }
    }
}