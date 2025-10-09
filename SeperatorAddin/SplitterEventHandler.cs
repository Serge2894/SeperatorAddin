using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using SeperatorAddin.Common;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SeperatorAddin
{
    public enum SplitterCommandType
    {
        Floor,
        Wall,
        Roof,
        Pipe,
        Framing,
        Duct,
        Conduit,
        Ceiling,
        CableTray
    }

    public class SplitterEventHandler : IExternalEventHandler
    {
        private List<SplitterCommandType> _commandTypes = new List<SplitterCommandType>();
        private List<Reference> _modelLineRefs = new List<Reference>();

        public void SetData(List<SplitterCommandType> commandTypes, List<Reference> modelLineRefs)
        {
            _commandTypes = commandTypes;
            _modelLineRefs = modelLineRefs;
        }

        public void Execute(UIApplication app)
        {
            UIDocument uidoc = app.ActiveUIDocument;
            Document doc = uidoc.Document;

            if (_modelLineRefs == null || !_modelLineRefs.Any()) return;

            using (Transaction t = new Transaction(doc, "Split Elements"))
            {
                t.Start();
                foreach (var commandType in _commandTypes)
                {
                    try
                    {
                        // Get the appropriate filter and prompt for the current category
                        (ISelectionFilter filter, string prompt) = GetSelectionDetails(commandType);

                        // Prompt the user to select elements of that category
                        var elementsToSplitRefs = uidoc.Selection.PickObjects(ObjectType.Element, filter, prompt);
                        if (elementsToSplitRefs == null || !elementsToSplitRefs.Any()) continue;

                        foreach (var elemRef in elementsToSplitRefs)
                        {
                            var element = doc.GetElement(elemRef);
                            if (element == null) continue;

                            // We'll just use the first model line for splitting for simplicity
                            var modelLineRef = _modelLineRefs.First();
                            var modelLine = doc.GetElement(modelLineRef) as ModelCurve;
                            if (modelLine == null) continue;

                            // Execute the split using the appropriate command
                            ExecuteCommand(commandType, app, element, modelLine);
                        }
                    }
                    catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                    {
                        // User cancelled selection for a category, so we continue to the next
                        continue;
                    }
                    catch (Exception ex)
                    {
                        TaskDialog.Show("Error", $"An error occurred while splitting {commandType}: {ex.Message}");
                    }
                }
                t.Commit();
            }
        }

        private (ISelectionFilter, string) GetSelectionDetails(SplitterCommandType commandType)
        {
            switch (commandType)
            {
                case SplitterCommandType.Floor:
                    return (new FloorSelectionFilter(), "Select Floors to Split");
                case SplitterCommandType.Wall:
                    return (new Utils.WallSelectionFilter(), "Select Walls to Split");
                case SplitterCommandType.Roof:
                    return (new RoofSelectionFilter(), "Select Roofs to Split");
                case SplitterCommandType.Pipe:
                    return (new Utils.PipeSelectionFilter(), "Select Pipes to Split");
                case SplitterCommandType.Framing:
                    return (new cmdFramingSplitter.FramingSelectionFilter(), "Select Framing to Split");
                case SplitterCommandType.Duct:
                    return (new Utils.DuctSelectionFilter(), "Select Ducts to Split");
                case SplitterCommandType.Conduit:
                    return (new ConduitSelectionFilter(), "Select Conduits to Split");
                case SplitterCommandType.Ceiling:
                    return (new CeilingSelectionFilter(), "Select Ceilings to Split");
                case SplitterCommandType.CableTray:
                    return (new CableTraySelectionFilter(), "Select Cable Trays to Split");
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        private void ExecuteCommand(SplitterCommandType commandType, UIApplication app, Element element, ModelCurve modelLine)
        {
            // In a real-world scenario, you would refactor the core logic of each command
            // to be callable from here. For this example, we will just call the command's Execute method.
            string message = "";
            IExternalCommand command = null;
            switch (commandType)
            {
                case SplitterCommandType.Floor:
                    command = new cmdSplitTool();
                    break;
                case SplitterCommandType.Wall:
                    command = new cmdWallSplitter();
                    break;
                case SplitterCommandType.Roof:
                    command = new cmdRoofSplitter();
                    break;
                case SplitterCommandType.Pipe:
                    command = new cmdPipeSplitter();
                    break;
                case SplitterCommandType.Framing:
                    command = new cmdFramingSplitter();
                    break;
                case SplitterCommandType.Duct:
                    command = new cmdDuctSplitter();
                    break;
                case SplitterCommandType.Conduit:
                    command = new cmdConduitsSplitter();
                    break;
                case SplitterCommandType.Ceiling:
                    command = new cmdCeilingSplitter();
                    break;
                case SplitterCommandType.CableTray:
                    command = new cmdCableTraySplitter();
                    break;
            }
            command?.Execute(new ExternalCommandData(app.Application, new View(app.ActiveUIDocument.Document), app.ActiveUIDocument.ActiveView, null), ref message, new ElementSet());

        }


        public string GetName() => "Splitter Event Handler";
    }
}