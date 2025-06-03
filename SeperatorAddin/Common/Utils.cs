using Autodesk.Revit.UI.Selection;

namespace SeperatorAddin.Common
{
    internal static class Utils
    {
        internal class WallSelectionilter : ISelectionFilter
        {
            public bool AllowElement(Element elem)
            {
                if (elem is Wall wall)
                {
                    // Check if the wall is a curtain wall
                    if (wall.WallType.Kind == WallKind.Curtain)
                    {
                        return false; // Exclude curtain walls
                    }

                    // Check if the wall is single-layered
                    if (wall.WallType.GetCompoundStructure() == null || wall.WallType.GetCompoundStructure().LayerCount <= 1)
                    {
                        return false; // Exclude single-layer walls
                    }

                    return true; // Allow other walls
                }
                return false; // Exclude non-wall elements
            }

            public bool AllowReference(Reference reference, XYZ position)
            {
                return false;
            }
        }

        internal class FloorSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem)
            {
                if (elem is Floor floor)
                {
                    // Check if the floor is single-layered
                    if (floor.FloorType.GetCompoundStructure() == null || floor.FloorType.GetCompoundStructure().LayerCount <= 1)
                    {
                        return false; // Exclude single-layer floors
                    }

                    return true; // Allow other floors
                }
                return false; // Exclude non-floor elements
            }

            public bool AllowReference(Reference reference, XYZ position)
            {
                return false;
            }
        }

        internal static RibbonPanel CreateRibbonPanel(UIControlledApplication app, string tabName, string panelName)
        {
            RibbonPanel curPanel;

            if (GetRibbonPanelByName(app, tabName, panelName) == null)
                curPanel = app.CreateRibbonPanel(tabName, panelName);

            else
                curPanel = GetRibbonPanelByName(app, tabName, panelName);

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
