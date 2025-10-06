using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SeperatorAddin
{
    [Transaction(TransactionMode.Manual)]
    public class cmdWallSplit : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            TaskDialog.Show("Command Info", "This function is now part of the 'Vertical Split' tool.");
            return Result.Succeeded;
        }

        public List<Wall> SplitWallByLevels(Document doc, Wall originalWall, List<Level> levels)
        {
            var newWalls = new List<Wall>();
            try
            {
                LocationCurve locationCurve = originalWall.Location as LocationCurve;
                if (locationCurve == null) return newWalls;

                Curve wallCurve = locationCurve.Curve;
                WallType wallType = originalWall.WallType;

                Level originalBaseLevel = doc.GetElement(originalWall.LevelId) as Level;
                double originalBaseOffset = originalWall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET).AsDouble();
                double originalBaseElevation = originalBaseLevel.ProjectElevation + originalBaseOffset;

                Parameter topConstraintParam = originalWall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE);
                ElementId topLevelId = topConstraintParam.AsElementId();

                double originalTopElevation;
                if (topLevelId != ElementId.InvalidElementId)
                {
                    Level topLevel = doc.GetElement(topLevelId) as Level;
                    double topOffset = originalWall.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET).AsDouble();
                    originalTopElevation = topLevel.ProjectElevation + topOffset;
                }
                else
                {
                    double height = originalWall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).AsDouble();
                    originalTopElevation = originalBaseElevation + height;
                }

                // CORRECTED LOGIC: Build a list of valid split points
                List<double> splitElevations = new List<double> { originalBaseElevation };
                foreach (Level selectedLevel in levels)
                {
                    if (selectedLevel.ProjectElevation > originalBaseElevation + 0.001 && selectedLevel.ProjectElevation < originalTopElevation - 0.001)
                    {
                        splitElevations.Add(selectedLevel.ProjectElevation);
                    }
                }
                splitElevations.Add(originalTopElevation);

                splitElevations = splitElevations.Distinct().OrderBy(e => e).ToList();

                if (splitElevations.Count < 2) return newWalls;

                // Create wall segments
                for (int i = 0; i < splitElevations.Count - 1; i++)
                {
                    double bottomElev = splitElevations[i];
                    double topElev = splitElevations[i + 1];
                    double height = topElev - bottomElev;

                    if (height <= 0.001) continue;

                    // Find the best level to host this segment
                    Level hostLevel = FindBestHostLevel(doc.GetElement(originalWall.LevelId) as Level, levels, bottomElev);

                    double baseOffset = bottomElev - hostLevel.ProjectElevation;

                    Wall newWall = Wall.Create(doc, wallCurve, wallType.Id, hostLevel.Id, height, baseOffset, originalWall.Flipped, false);

                    if (newWall != null)
                    {
                        CopyAllWallParameters(originalWall, newWall);
                        newWalls.Add(newWall);
                    }
                }
                return newWalls;
            }
            catch
            {
                return new List<Wall>();
            }
        }

        private Level FindBestHostLevel(Level originalLevel, List<Level> allLevels, double elevation)
        {
            return allLevels
                .Where(l => l.ProjectElevation <= elevation + 0.001)
                .OrderByDescending(l => l.ProjectElevation)
                .FirstOrDefault() ?? originalLevel;
        }

        private void CopyAllWallParameters(Wall sourceWall, Wall targetWall)
        {
            var ignoredBips = new List<BuiltInParameter>
            {
                BuiltInParameter.WALL_BASE_CONSTRAINT,
                BuiltInParameter.WALL_HEIGHT_TYPE,
                BuiltInParameter.WALL_BASE_OFFSET,
                BuiltInParameter.WALL_TOP_OFFSET,
                BuiltInParameter.WALL_USER_HEIGHT_PARAM,
                BuiltInParameter.HOST_AREA_COMPUTED,
                BuiltInParameter.HOST_VOLUME_COMPUTED,
                BuiltInParameter.CURVE_ELEM_LENGTH
            };

            foreach (Parameter sourceParam in sourceWall.Parameters)
            {
                if (sourceParam == null || !sourceParam.HasValue || sourceParam.IsReadOnly) continue;

                if (sourceParam.Definition is InternalDefinition def)
                {
                    if (ignoredBips.Contains((BuiltInParameter)def.BuiltInParameter))
                    {
                        continue;
                    }
                }

                Parameter targetParam = targetWall.get_Parameter(sourceParam.Definition);
                if (targetParam != null && !targetParam.IsReadOnly)
                {
                    try
                    {
                        if (sourceParam.StorageType == StorageType.String) targetParam.Set(sourceParam.AsString());
                        else if (sourceParam.StorageType == StorageType.Double) targetParam.Set(sourceParam.AsDouble());
                        else if (sourceParam.StorageType == StorageType.Integer) targetParam.Set(sourceParam.AsInteger());
                        else if (sourceParam.StorageType == StorageType.ElementId) targetParam.Set(sourceParam.AsElementId());
                    }
                    catch { }
                }
            }
        }
    }
}