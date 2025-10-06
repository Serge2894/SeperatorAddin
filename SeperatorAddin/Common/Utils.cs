using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;

namespace SeperatorAddin.Common
{
    internal static class Utils
    {
        // Filter for selecting multi-layer walls (excluding curtain walls)
        internal class WallSelectionilter : ISelectionFilter
        {
            public bool AllowElement(Element elem)
            {
                if (elem is Wall wall)
                {
                    if (wall.WallType.Kind == WallKind.Curtain)
                    {
                        return false; // Exclude curtain walls
                    }

                    if (wall.WallType.GetCompoundStructure() == null || wall.WallType.GetCompoundStructure().LayerCount <= 1)
                    {
                        return false; // Exclude single-layer walls
                    }

                    return true; // Allow other walls
                }
                return false; // Exclude non-wall elements
            }

            public bool AllowReference(Reference reference, XYZ position) => false;
        }

        // Filter for selecting floors with a compound structure
        internal class FloorSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem)
            {
                if (elem is Floor floor)
                {
                    if (floor.FloorType.GetCompoundStructure() == null || floor.FloorType.GetCompoundStructure().LayerCount <= 1)
                    {
                        return false;
                    }
                    return true;
                }
                return false;
            }

            public bool AllowReference(Reference reference, XYZ position) => false;
        }

        // Filter for selecting walls (any kind)
        internal class WallSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem) => elem is Wall;
            public bool AllowReference(Reference reference, XYZ position) => false;
        }
        // Filter for selecting multi -layer ceilings
        internal class CeilingSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem)
            {
                if (elem is Ceiling ceiling)
                {
                    CeilingType ceilingType = elem.Document.GetElement(ceiling.GetTypeId()) as CeilingType;
                    if (ceilingType?.GetCompoundStructure() == null || ceilingType.GetCompoundStructure().LayerCount <= 1)
                    {
                        return false; // Exclude single-layer or invalid ceilings
                    }
                    return true;
                }
                return false;
            }
            public bool AllowReference(Reference reference, XYZ position) => false;
        }
        // Filter for selecting multi -layer roofs
        internal class RoofSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem)
            {
                if (elem is RoofBase roof)
                {
                    RoofType roofType = elem.Document.GetElement(roof.GetTypeId()) as RoofType;
                    // Ensure it's a flat roof and has multiple layers
                    bool isSloped = roof.get_Parameter(BuiltInParameter.ROOF_SLOPE)?.AsDouble() > 0;
                    if (isSloped) return false;

                    if (roofType?.GetCompoundStructure() == null || roofType.GetCompoundStructure().LayerCount <= 1)
                    {
                        return false;
                    }
                    return true;
                }
                return false;
            }

            public bool AllowReference(Reference reference, XYZ position) => false;
        }

        // Filter for selecting model curves
        internal class ModelCurveSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem) => elem is ModelCurve;
            public bool AllowReference(Reference reference, XYZ position) => false;
        }

        // ADDED FILTERS START HERE

        // Filter for selecting columns (Architectural and Structural)
        public class ColumnSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem)
            {
                if (elem.Category != null)
                {
                    var catId = elem.Category.Id.IntegerValue;
                    return catId == (int)BuiltInCategory.OST_Columns ||
                           catId == (int)BuiltInCategory.OST_StructuralColumns;
                }
                return false;
            }

            public bool AllowReference(Reference reference, XYZ position) => false;
        }

        // Filter for selecting Ducts
        public class DuctSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem)
            {
                return elem is Duct;
            }

            public bool AllowReference(Reference reference, XYZ position) => false;
        }

        // Filter for selecting Pipes
        public class PipeSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem)
            {
                return elem is Pipe;
            }

            public bool AllowReference(Reference reference, XYZ position) => false;
        }

        // ADDED FILTERS END HERE

        // Method to get all parameters from an element
        public static List<Parameter> GetAllParametersFromElement(Element element)
        {
            List<Parameter> parameters = new List<Parameter>();
            foreach (Parameter param in element.Parameters)
            {
                if (param != null)
                {
                    parameters.Add(param);
                }
            }
            return parameters;
        }

        // Helper methods for creating Ribbon items
        internal static RibbonPanel CreateRibbonPanel(UIControlledApplication app, string tabName, string panelName)
        {
            RibbonPanel curPanel = GetRibbonPanelByName(app, tabName, panelName);
            if (curPanel == null)
                curPanel = app.CreateRibbonPanel(tabName, panelName);
            return curPanel;
        }

        internal static RibbonPanel GetRibbonPanelByName(UIControlledApplication app, string tabName, string panelName)
        {
            foreach (RibbonPanel tmpPanel in app.GetRibbonPanels(tabName))
            {
                if (tmpPanel.Name == panelName)
                    return tmpPanel;
            }
            return null;
        }
    }
}