using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using SeperatorAddin.Common;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace SeperatorAddin.Forms
{
    public partial class frmSplitter : Window
    {
        private readonly ExternalCommandData _commandData;
        private readonly SplitterEventHandler _handler;
        private readonly ExternalEvent _externalEvent;
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

        public frmSplitter(ExternalCommandData commandData, SplitterEventHandler handler, ExternalEvent externalEvent)
        {
            InitializeComponent();
            _commandData = commandData;
            _handler = handler;
            _externalEvent = externalEvent;
        }

        private void Border_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                this.DragMove();
            }
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }

        private void SelectAndCount(ISelectionFilter filter, string prompt, ref List<Reference> refList, TextBlock countLabel)
        {
            try
            {
                this.Hide();
                refList = _commandData.Application.ActiveUIDocument.Selection.PickObjects(ObjectType.Element, filter, prompt).ToList();
                countLabel.Text = $"{refList.Count} selected";
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException) { }
            finally
            {
                this.ShowDialog();
            }
        }

        private void SelectModelLines_Click(object sender, RoutedEventArgs e)
        {
            SelectAndCount(new Utils.ModelCurveSelectionFilter(), "Select Model Lines", ref _modelLineRefs, ModelLineCount);
        }

        private void SelectFloors_Click(object sender, RoutedEventArgs e)
        {
            SelectAndCount(new Utils.FloorSelectionFilter(), "Select Floors", ref _floorRefs, FloorCount);
        }

        private void SelectWalls_Click(object sender, RoutedEventArgs e)
        {
            SelectAndCount(new Utils.WallSelectionFilter(), "Select Walls", ref _wallRefs, WallCount);
        }

        private void SelectRoofs_Click(object sender, RoutedEventArgs e)
        {
            SelectAndCount(new RoofSelectionFilter(), "Select Roofs", ref _roofRefs, RoofCount);
        }

        private void SelectPipes_Click(object sender, RoutedEventArgs e)
        {
            SelectAndCount(new Utils.PipeSelectionFilter(), "Select Pipes", ref _pipeRefs, PipeCount);
        }

        private void SelectFraming_Click(object sender, RoutedEventArgs e)
        {
            SelectAndCount(new cmdFramingSplitter.FramingSelectionFilter(), "Select Framing", ref _framingRefs, FramingCount);
        }

        private void SelectDucts_Click(object sender, RoutedEventArgs e)
        {
            SelectAndCount(new Utils.DuctSelectionFilter(), "Select Ducts", ref _ductRefs, DuctCount);
        }

        private void SelectConduits_Click(object sender, RoutedEventArgs e)
        {
            SelectAndCount(new ConduitSelectionFilter(), "Select Conduits", ref _conduitRefs, ConduitCount);
        }

        private void SelectCeilings_Click(object sender, RoutedEventArgs e)
        {
            SelectAndCount(new CeilingSelectionFilter(), "Select Ceilings", ref _ceilingRefs, CeilingCount);
        }

        private void SelectCableTrays_Click(object sender, RoutedEventArgs e)
        {
            SelectAndCount(new CableTraySelectionFilter(), "Select Cable Trays", ref _cableTrayRefs, CableTrayCount);
        }

        private void Split_Click(object sender, RoutedEventArgs e)
        {
            if (_modelLineRefs.Count == 0)
            {
                var warningDialog = new frmInfoDialog("Please select at least one model line to split with.", "Selection Required");
                warningDialog.ShowDialog();
                return;
            }

            if (!_floorRefs.Any() && !_wallRefs.Any() && !_roofRefs.Any() && !_pipeRefs.Any() && !_framingRefs.Any() && !_ductRefs.Any() && !_conduitRefs.Any() && !_ceilingRefs.Any() && !_cableTrayRefs.Any())
            {
                var warningDialog = new frmInfoDialog("Please select elements from at least one category to split.", "Selection Required");
                warningDialog.ShowDialog();
                return;
            }

            _handler.SetData(_modelLineRefs, _floorRefs, _wallRefs, _roofRefs, _pipeRefs, _framingRefs, _ductRefs, _conduitRefs, _ceilingRefs, _cableTrayRefs);
            _externalEvent.Raise();
            this.Close();
        }
    }
}