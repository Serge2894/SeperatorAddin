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

        // Filter for selecting model curves
        internal class ModelCurveSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem) => elem is ModelCurve;
            public bool AllowReference(Reference reference, XYZ position) => false;
        }


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