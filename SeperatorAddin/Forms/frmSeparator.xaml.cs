using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace SeperatorAddin.Forms
{
    public partial class frmSeparator : Window
    {
        private readonly ExternalCommandData _commandData;
        private readonly SeparatorEventHandler _handler;
        private readonly ExternalEvent _externalEvent;

        private List<Reference> _floorRefs = new List<Reference>();
        private List<Reference> _roofRefs = new List<Reference>();
        private List<Reference> _wallRefs = new List<Reference>();
        private List<Reference> _ceilingRefs = new List<Reference>();

        public frmSeparator(ExternalCommandData commandData, SeparatorEventHandler handler, ExternalEvent externalEvent)
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

        private void SelectFloors_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                this.Hide();
                _floorRefs = _commandData.Application.ActiveUIDocument.Selection.PickObjects(ObjectType.Element, new Common.Utils.FloorSelectionFilter(), "Select Floors").ToList();
                FloorCount.Text = $"{_floorRefs.Count} selected";
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException) { }
            finally
            {
                this.ShowDialog();
            }
        }

        private void SelectRoofs_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                this.Hide();
                _roofRefs = _commandData.Application.ActiveUIDocument.Selection.PickObjects(ObjectType.Element, new Common.Utils.RoofSelectionFilter(), "Select Roofs").ToList();
                RoofCount.Text = $"{_roofRefs.Count} selected";
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException) { }
            finally
            {
                this.ShowDialog();
            }
        }

        private void SelectWalls_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                this.Hide();
                _wallRefs = _commandData.Application.ActiveUIDocument.Selection.PickObjects(ObjectType.Element, new Common.Utils.WallSelectionilter(), "Select Walls").ToList();
                WallCount.Text = $"{_wallRefs.Count} selected";
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException) { }
            finally
            {
                this.ShowDialog();
            }
        }

        private void SelectCeilings_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                this.Hide();
                _ceilingRefs = _commandData.Application.ActiveUIDocument.Selection.PickObjects(ObjectType.Element, new Common.Utils.CeilingSelectionFilter(), "Select Ceilings").ToList();
                CeilingCount.Text = $"{_ceilingRefs.Count} selected";
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException) { }
            finally
            {
                this.ShowDialog();
            }
        }

        private void Separate_Click(object sender, RoutedEventArgs e)
        {
            // ADDED: Check if any elements have been selected
            if (!_floorRefs.Any() && !_roofRefs.Any() && !_wallRefs.Any() && !_ceilingRefs.Any())
            {
                var warningDialog = new frmInfoDialog("Please select elements from at least one category to separate.", "Selection Required");
                warningDialog.ShowDialog();
                return; // Stop execution
            }

            // Pass the selected element references to the handler
            _handler.SetData(_floorRefs, _roofRefs, _wallRefs, _ceilingRefs);

            // Raise the external event to run the separation logic
            _externalEvent.Raise();

            this.Close();
        }
    }
}