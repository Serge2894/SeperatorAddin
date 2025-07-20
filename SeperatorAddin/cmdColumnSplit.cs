using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using SeperatorAddin.Forms;

namespace SeperatorAddin
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class cmdColumnSplit : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Check if we're in a 3D view
                View3D view3D = doc.ActiveView as View3D;
                if (view3D == null)
                {
                    TaskDialog.Show("Error", "Please switch to a 3D view before running this command.");
                    return Result.Failed;
                }

                // Get all levels in the project
                List<Level> allLevels = new FilteredElementCollector(doc)
                    .OfClass(typeof(Level))
                    .Cast<Level>()
                    .OrderBy(l => l.Elevation)
                    .ToList();

                if (allLevels.Count < 2)
                {
                    TaskDialog.Show("Error", "The project must have at least 2 levels to split columns.");
                    return Result.Failed;
                }

                // Show level selection dialog
                LevelSelectionForm levelForm = new LevelSelectionForm(allLevels);
                if (levelForm.ShowDialog() != true)
                {
                    return Result.Cancelled;
                }

                List<Level> selectedLevels = levelForm.SelectedLevels
                    .OrderBy(l => l.Elevation)
                    .ToList();

                if (selectedLevels.Count < 2)
                {
                    TaskDialog.Show("Error", "Please select at least 2 levels to split columns between.");
                    return Result.Failed;
                }

                // Select columns
                IList<Reference> columnRefs;
                try
                {
                    ColumnSelectionFilter columnFilter = new ColumnSelectionFilter();
                    columnRefs = uidoc.Selection.PickObjects(
                        ObjectType.Element,
                        columnFilter,
                        "Select columns to split (ESC to cancel)");
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    return Result.Cancelled;
                }

                if (columnRefs.Count == 0)
                {
                    TaskDialog.Show("Error", "No columns selected.");
                    return Result.Failed;
                }

                // Process columns
                using (Transaction trans = new Transaction(doc, "Split Columns by Levels"))
                {
                    trans.Start();

                    int successCount = 0;
                    int failCount = 0;
                    List<string> errors = new List<string>();

                    foreach (Reference columnRef in columnRefs)
                    {
                        Element columnElement = doc.GetElement(columnRef);

                        try
                        {
                            bool success = SplitColumn(doc, columnElement, selectedLevels);
                            if (success)
                                successCount++;
                            else
                            {
                                failCount++;
                                errors.Add($"Failed to split column ID: {columnElement.Id}");
                            }
                        }
                        catch (Exception ex)
                        {
                            failCount++;
                            errors.Add($"Column ID {columnElement.Id}: {ex.Message}");
                        }
                    }

                    trans.Commit();

                    // Show results
                    string resultMessage = $"Split operation completed:\n" +
                                         $"- Successfully split: {successCount} columns\n" +
                                         $"- Failed: {failCount} columns";

                    if (errors.Count > 0)
                    {
                        resultMessage += "\n\nErrors:\n" + string.Join("\n", errors.Take(10));
                        if (errors.Count > 10)
                            resultMessage += $"\n... and {errors.Count - 10} more errors";
                    }

                    TaskDialog.Show("Split Columns Result", resultMessage);
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
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

    // Selection filter for columns
    public class ColumnSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            // Check for both architectural and structural columns
            if (elem.Category != null)
            {
                var catId = elem.Category.Id.IntegerValue;
                return catId == (int)BuiltInCategory.OST_Columns ||
                       catId == (int)BuiltInCategory.OST_StructuralColumns;
            }
            return false;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }
}