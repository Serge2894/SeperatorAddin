using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using SeperatorAddin.Common;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace SeperatorAddin.Forms
{
    public partial class frmVerticalSplit : Window
    {
        private readonly ExternalCommandData _commandData;
        private readonly VerticalSplitEventHandler _handler;
        private readonly ExternalEvent _externalEvent;

        private List<Reference> _columnRefs = new List<Reference>();
        private List<Reference> _wallRefs = new List<Reference>();
        private List<Reference> _ductRefs = new List<Reference>();
        private List<Reference> _pipeRefs = new List<Reference>();

        private ObservableCollection<LevelItem> _levelItems;

        public List<Level> SelectedLevels
        {
            get
            {
                if (_levelItems == null) return new List<Level>();
                return _levelItems
                    .Where(item => item.IsSelected)
                    .Select(item => item.Level)
                    .ToList();
            }
        }

        public frmVerticalSplit(ExternalCommandData commandData, VerticalSplitEventHandler handler, ExternalEvent externalEvent)
        {
            InitializeComponent();
            _commandData = commandData;
            _handler = handler;
            _externalEvent = externalEvent;

            LoadLevels();
        }

        private void LoadLevels()
        {
            var doc = _commandData.Application.ActiveUIDocument.Document;
            var allLevels = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .OrderBy(l => l.Elevation)
                .ToList();

            _levelItems = new ObservableCollection<LevelItem>();
            foreach (var level in allLevels)
            {
                _levelItems.Add(new LevelItem(level));
            }
            lvLevels.ItemsSource = _levelItems;
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

        private void SelectColumns_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                this.Hide();
                ISelectionFilter filter = new Utils.ColumnSelectionFilter();
                _columnRefs = _commandData.Application.ActiveUIDocument.Selection.PickObjects(ObjectType.Element, filter, "Select Columns").ToList();
                ColumnCount.Text = $"{_columnRefs.Count} selected";
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
                ISelectionFilter filter = new Utils.WallSelectionFilter();
                _wallRefs = _commandData.Application.ActiveUIDocument.Selection.PickObjects(ObjectType.Element, filter, "Select Walls").ToList();
                WallCount.Text = $"{_wallRefs.Count} selected";
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException) { }
            finally
            {
                this.ShowDialog();
            }
        }

        private void SelectDucts_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                this.Hide();
                ISelectionFilter filter = new Utils.DuctSelectionFilter();
                _ductRefs = _commandData.Application.ActiveUIDocument.Selection.PickObjects(ObjectType.Element, filter, "Select Ducts").ToList();
                DuctCount.Text = $"{_ductRefs.Count} selected";
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException) { }
            finally
            {
                this.ShowDialog();
            }
        }

        private void SelectPipes_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                this.Hide();
                ISelectionFilter filter = new Utils.PipeSelectionFilter();
                _pipeRefs = _commandData.Application.ActiveUIDocument.Selection.PickObjects(ObjectType.Element, filter, "Select Pipes").ToList();
                PipeCount.Text = $"{_pipeRefs.Count} selected";
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException) { }
            finally
            {
                this.ShowDialog();
            }
        }

        private void Split_Click(object sender, RoutedEventArgs e)
        {
            if (!_columnRefs.Any() && !_wallRefs.Any() && !_ductRefs.Any() && !_pipeRefs.Any())
            {
                var warningDialog = new frmInfoDialog("Please select elements from at least one category to split.", "Selection Required");
                warningDialog.ShowDialog();
                return;
            }

            if (!SelectedLevels.Any())
            {
                var warningDialog = new frmInfoDialog("Please select at least one level to split the elements.", "Selection Required");
                warningDialog.ShowDialog();
                return;
            }

            _handler.SetData(_columnRefs, _wallRefs, _ductRefs, _pipeRefs, SelectedLevels);
            _externalEvent.Raise();
            this.Close();
        }

        private void cbSelectAll_Checked(object sender, RoutedEventArgs e)
        {
            if (_levelItems == null) return;
            foreach (var item in _levelItems)
            {
                item.IsSelected = true;
            }
        }

        private void cbSelectAll_Unchecked(object sender, RoutedEventArgs e)
        {
            if (_levelItems == null) return;
            foreach (var item in _levelItems)
            {
                item.IsSelected = false;
            }
        }

        private void LevelCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            if (_levelItems == null) return;
            int selectedCount = _levelItems.Count(item => item.IsSelected);
            if (selectedCount == 0)
            {
                cbSelectAll.IsChecked = false;
            }
            else if (selectedCount == _levelItems.Count)
            {
                cbSelectAll.IsChecked = true;
            }
            else
            {
                cbSelectAll.IsChecked = null; // Indeterminate state
            }
        }
    }

    public class LevelItem : INotifyPropertyChanged
    {
        private bool _isSelected;
        public Level Level { get; }
        public string DisplayName { get; }
        public bool IsSelected
        {
            get { return _isSelected; }
            set
            {
                if (_isSelected != value)
                {
                    _isSelected = value;
                    OnPropertyChanged(nameof(IsSelected));
                }
            }
        }

        public LevelItem(Level level)
        {
            Level = level;
            DisplayName = $"{level.Name}";
            IsSelected = false;
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}