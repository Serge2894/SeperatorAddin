using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using SeperatorAddin.Forms;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace SeperatorAddin
{
    public class SplitterEventHandler : IExternalEventHandler
    {
        private List<Reference> _modelLineRefs = new List<Reference>();
        private List<Reference> _floorRefs = new List<Reference>();
        private List<Reference> _wallRefs = new List<Reference>();
        private List<Reference> _roofRefs = new List<Reference>();
        private List<Reference> _pipeRefs = new List<Reference>();
        private List<Reference> _framingRefs = new List<Reference>();
        private List<Reference> _ductRefs = new List<Reference>();
        private List<Reference> _conduitRefs = new List<Reference>();
        private List<Reference> _ceilingRefs = new List<Reference>();
        private List<Reference> _cableTrayRefs = new List<Reference>();

        public void SetData(List<Reference> modelLineRefs, List<Reference> floorRefs, List<Reference> wallRefs, List<Reference> roofRefs, List<Reference> pipeRefs, List<Reference> framingRefs, List<Reference> ductRefs, List<Reference> conduitRefs, List<Reference> ceilingRefs, List<Reference> cableTrayRefs)
        {
            _modelLineRefs = modelLineRefs;
            _floorRefs = floorRefs;
            _wallRefs = wallRefs;
            _roofRefs = roofRefs;
            _pipeRefs = pipeRefs;
            _framingRefs = framingRefs;
            _ductRefs = ductRefs;
            _conduitRefs = conduitRefs;
            _ceilingRefs = ceilingRefs;
            _cableTrayRefs = cableTrayRefs;
        }

        public void Execute(UIApplication app)
        {
            UIDocument uidoc = app.ActiveUIDocument;
            Document doc = uidoc.Document;

            if (_modelLineRefs == null || !_modelLineRefs.Any()) return;

            int totalElementsProcessed = 0;

            using (Transaction t = new Transaction(doc, "Split Elements by Line"))
            {
                t.Start();

                var floorSplitter = new cmdFloorSplitter();
                var wallSplitter = new cmdWallSplitter();
                var roofSplitter = new cmdRoofSplitter();
                var ductSplitter = new cmdDuctSplitter();
                var pipeSplitter = new cmdPipeSplitter();
                var framingSplitter = new cmdFramingSplitter();
                var conduitSplitter = new cmdConduitsSplitter();
                var cableTraySplitter = new cmdCableTraySplitter();
                var ceilingSplitter = new cmdCeilingSplitter();

                foreach (var lineRef in _modelLineRefs)
                {
                    var modelCurve = doc.GetElement(lineRef) as ModelCurve;
                    if (modelCurve == null) continue;

                    XYZ splitPoint = modelCurve.GeometryCurve.Evaluate(0.5, true);

                    totalElementsProcessed += ProcessElements(doc, _floorRefs, e => floorSplitter.SplitFloor(doc, e as Floor, modelCurve));
                    totalElementsProcessed += ProcessElements(doc, _wallRefs, e => wallSplitter.SplitWall(doc, e as Wall, modelCurve));
                    totalElementsProcessed += ProcessElements(doc, _roofRefs, e => roofSplitter.SplitRoof(doc, e as RoofBase, modelCurve));
                    totalElementsProcessed += ProcessElements(doc, _ductRefs, e => ductSplitter.SplitDuct(doc, e as Duct, splitPoint));
                    totalElementsProcessed += ProcessElements(doc, _pipeRefs, e => pipeSplitter.SplitPipe(doc, e as Pipe, splitPoint));
                    totalElementsProcessed += ProcessElements(doc, _framingRefs, e => framingSplitter.SplitFraming(doc, e as FamilyInstance, splitPoint));
                    totalElementsProcessed += ProcessElements(doc, _conduitRefs, e => conduitSplitter.SplitConduit(doc, e as Conduit, splitPoint));
                    totalElementsProcessed += ProcessElements(doc, _cableTrayRefs, e => cableTraySplitter.SplitCableTray(doc, e as CableTray, splitPoint));
                    totalElementsProcessed += ProcessElements(doc, _ceilingRefs, e => ceilingSplitter.SplitCeiling(doc, e as Ceiling, modelCurve));
                }

                t.Commit();
            }

            if (totalElementsProcessed > 0)
            {
                new frmInfoDialog("Element(s) were split successfully.", "Split Results").ShowDialog();
            }
        }

        private int ProcessElements(Document doc, List<Reference> references, Func<Element, bool> splitAction)
        {
            if (references == null || !references.Any()) return 0;

            int successCount = 0;
            var elementIds = references.Select(r => r.ElementId).ToList();

            foreach (var id in elementIds)
            {
                Element element = doc.GetElement(id);
                if (element != null)
                {
                    if (splitAction(element))
                    {
                        successCount++;
                    }
                }
            }
            return successCount;
        }

        public string GetName() => "Splitter Event Handler";
    }
}