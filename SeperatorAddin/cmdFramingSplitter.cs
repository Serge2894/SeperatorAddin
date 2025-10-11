using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using SeperatorAddin.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace SeperatorAddin
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class cmdFramingSplitter : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            try
            {
                // 1. Select a structural framing element
                Reference framingRef = uiDoc.Selection.PickObject(ObjectType.Element, new FramingSelectionFilter(), "Select a structural framing element to split");
                FamilyInstance framingElement = doc.GetElement(framingRef) as FamilyInstance;

                if (framingElement == null)
                {
                    message = "The selected element is not a structural framing element.";
                    return Result.Failed;
                }

                // 2. Pick a point to define the split location
                XYZ splitPoint = uiDoc.Selection.PickPoint("Pick a point on the framing element to split");

                using (Transaction trans = new Transaction(doc, "Split Structural Framing"))
                {
                    trans.Start();
                    bool success = SplitFraming(doc, framingElement, splitPoint);
                    if (success)
                    {
                        trans.Commit();
                        return Result.Succeeded;
                    }
                    else
                    {
                        trans.RollBack();
                        message = "Failed to split the framing element. The split point may be too close to an end.";
                        return Result.Failed;
                    }
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        public bool SplitFraming(Document doc, FamilyInstance framingElement, XYZ splitPoint)
        {
            if (framingElement == null) return false;
            LocationCurve framingLocationCurve = framingElement.Location as LocationCurve;
            if (framingLocationCurve == null || !(framingLocationCurve.Curve is Line framingLine)) return false;

            IntersectionResult projection = framingLine.Project(splitPoint);
            if (projection == null) return false;
            XYZ intersectionPoint = projection.XYZPoint;

            // Check if split point is too close to the ends
            double minDistance = 0.01; // A small tolerance
            if (intersectionPoint.IsAlmostEqualTo(framingLine.GetEndPoint(0), minDistance) ||
                intersectionPoint.IsAlmostEqualTo(framingLine.GetEndPoint(1), minDistance))
            {
                return false;
            }

            try
            {
                // Store original parameters
                Dictionary<string, Parameter> originalParams = GetElementParameters(framingElement);
                double startExtension = framingElement.get_Parameter(BuiltInParameter.START_EXTENSION).AsDouble();
                double endExtension = framingElement.get_Parameter(BuiltInParameter.END_EXTENSION).AsDouble();

                // Get structural type and level
                StructuralType structuralType = framingElement.StructuralType;
                Level level = doc.GetElement(framingElement.LevelId) as Level;
                if (level == null)
                {
                    level = doc.ActiveView.GenLevel ?? new FilteredElementCollector(doc).OfClass(typeof(Level)).FirstElement() as Level;
                }
                if (level == null) return false;


                // Create two new framing segments
                Line segment1 = Line.CreateBound(framingLine.GetEndPoint(0), intersectionPoint);
                Line segment2 = Line.CreateBound(intersectionPoint, framingLine.GetEndPoint(1));

                FamilyInstance newFraming1 = doc.Create.NewFamilyInstance(segment1, framingElement.Symbol, level, structuralType);
                FamilyInstance newFraming2 = doc.Create.NewFamilyInstance(segment2, framingElement.Symbol, level, structuralType);

                if (newFraming1 == null || newFraming2 == null) return false;

                // Copy all parameters first
                CopyParameters(originalParams, newFraming1);
                CopyParameters(originalParams, newFraming2);

                // Preserve original start/end properties
                newFraming1.get_Parameter(BuiltInParameter.START_EXTENSION).Set(startExtension);
                newFraming1.get_Parameter(BuiltInParameter.END_EXTENSION).Set(0);
                newFraming2.get_Parameter(BuiltInParameter.START_EXTENSION).Set(0);
                newFraming2.get_Parameter(BuiltInParameter.END_EXTENSION).Set(endExtension);

                // Safely set JOIN cutback parameters to zero to remove the gap
                SetJoinCutbackParameter(newFraming1, "END_JOIN_CUTBACK", 0);
                SetJoinCutbackParameter(newFraming2, "START_JOIN_CUTBACK", 0);

                // Delete the original element
                doc.Delete(framingElement.Id);
                return true;
            }
            catch
            {
                return false;
            }
        }

        #region Helper Methods
        /// <summary>
        /// Safely sets a BuiltInParameter value by its string name, avoiding compile-time errors for different API versions.
        /// </summary>
        private void SetJoinCutbackParameter(Element element, string paramName, double value)
        {
            try
            {
                // Use reflection to get the enum value from the string name
                var bip = (BuiltInParameter)Enum.Parse(typeof(BuiltInParameter), paramName);

                // If the parameter exists on the element, set its value
                Parameter param = element.get_Parameter(bip);
                if (param != null && !param.IsReadOnly)
                {
                    param.Set(value);
                }
            }
            catch
            {
                // The parameter does not exist in this version of the Revit API, so we do nothing.
            }
        }

        private Dictionary<string, Parameter> GetElementParameters(Element element)
        {
            var paramDict = new Dictionary<string, Parameter>();
            foreach (Parameter p in element.Parameters)
            {
                if (!paramDict.ContainsKey(p.Definition.Name))
                {
                    paramDict.Add(p.Definition.Name, p);
                }
            }
            return paramDict;
        }

        private void CopyParameters(Dictionary<string, Parameter> sourceParams, Element target)
        {
            foreach (Parameter targetParam in target.Parameters)
            {
                if (!targetParam.IsReadOnly && sourceParams.TryGetValue(targetParam.Definition.Name, out Parameter sourceParam))
                {
                    if (sourceParam.StorageType == targetParam.StorageType)
                    {
                        try
                        {
                            switch (sourceParam.StorageType)
                            {
                                case StorageType.Double:
                                    targetParam.Set(sourceParam.AsDouble());
                                    break;
                                case StorageType.Integer:
                                    targetParam.Set(sourceParam.AsInteger());
                                    break;
                                case StorageType.String:
                                    targetParam.Set(sourceParam.AsString());
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

        public class FramingSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem)
            {
                return elem.Category != null && elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming;
            }

            public bool AllowReference(Reference reference, XYZ position)
            {
                return false;
            }
        }
        #endregion
    }
}