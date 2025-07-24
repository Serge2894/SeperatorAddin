using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace SeperatorAddin.Forms
{
    public partial class PipeSelectionForm : Window, INotifyPropertyChanged
    {
        private ObservableCollection<PipeItem> _pipes;
        private ObservableCollection<LevelItem> _levels;
        private List<PipeItem> _allPipes;

        public List<Pipe> SelectedPipes { get; private set; }
        public List<Level> SelectedLevels { get; private set; }

        public PipeSelectionForm(List<Pipe> pipes, List<Level> levels)
        {
            InitializeComponent();
            DataContext = this;

            // Initialize collections
            _allPipes = new List<PipeItem>();
            foreach (var pipe in pipes)
            {
                _allPipes.Add(new PipeItem(pipe));
            }
            _pipes = new ObservableCollection<PipeItem>(_allPipes.OrderBy(p => p.Name));

            _levels = new ObservableCollection<LevelItem>();
            foreach (var level in levels.OrderBy(l => l.Elevation))
            {
                _levels.Add(new LevelItem(level));
            }

            // Set ItemsSource
            lbPipes.ItemsSource = _pipes;
            lbLevels.ItemsSource = _levels;

            // Initialize
            UpdateCounts();
            UpdateOKButton();
        }

        private void txtFilter_TextChanged(object sender, TextChangedEventArgs e)
        {
            string filterText = txtFilter.Text.ToLower();

            if (string.IsNullOrEmpty(filterText))
            {
                _pipes.Clear();
                foreach (var pipe in _allPipes.OrderBy(p => p.Name))
                {
                    _pipes.Add(pipe);
                }
            }
            else
            {
                _pipes.Clear();
                foreach (var pipe in _allPipes.Where(p => p.DisplayName.ToLower().Contains(filterText)).OrderBy(p => p.Name))
                {
                    _pipes.Add(pipe);
                }
            }

            UpdatePipeCount();
        }

        private void btnSelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var pipe in _pipes)
            {
                pipe.IsSelected = true;
            }
            UpdatePipeCount();
        }

        private void btnSelectNone_Click(object sender, RoutedEventArgs e)
        {
            foreach (var pipe in _pipes)
            {
                pipe.IsSelected = false;
            }
            UpdatePipeCount();
        }

        private void PipeCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            UpdatePipeCount();
        }

        private void PipeCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            UpdatePipeCount();
        }

        private void LevelCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            UpdateLevelCount();
            UpdateOKButton();
        }

        private void LevelCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            UpdateLevelCount();
            UpdateOKButton();
        }

        private void UpdateCounts()
        {
            UpdatePipeCount();
            UpdateLevelCount();
        }

        private void UpdatePipeCount()
        {
            int count = _pipes.Count(p => p.IsSelected);
            txtPipeCount.Text = $"{count} selected";
        }

        private void UpdateLevelCount()
        {
            int count = _levels.Count(l => l.IsSelected);
            txtLevelCount.Text = $"{count} selected";
        }

        private void UpdateOKButton()
        {
            int levelCount = _levels.Count(l => l.IsSelected);
            btnOK.IsEnabled = levelCount >= 2;
        }

        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            SelectedPipes = new List<Pipe>();
            SelectedLevels = new List<Level>();

            foreach (var item in _pipes.Where(p => p.IsSelected))
            {
                SelectedPipes.Add(item.Pipe);
            }

            foreach (var item in _levels.Where(l => l.IsSelected))
            {
                SelectedLevels.Add(item.Level);
            }

            DialogResult = true;
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class PipeItem : INotifyPropertyChanged
    {
        private bool _isSelected;

        public Pipe Pipe { get; }
        public string Name { get; }
        public string DisplayName { get; }

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected != value)
                {
                    _isSelected = value;
                    OnPropertyChanged(nameof(IsSelected));
                }
            }
        }

        public PipeItem(Pipe pipe)
        {
            Pipe = pipe;
            Name = GetPipeName(pipe);
            DisplayName = GetPipeDisplayName(pipe);
        }

        private string GetPipeName(Pipe pipe)
        {
            try
            {
                string mark = pipe.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)?.AsString();
                if (!string.IsNullOrEmpty(mark))
                    return mark;

                return $"Pipe {pipe.Id.IntegerValue}";
            }
            catch
            {
                return $"Pipe {pipe.Id.IntegerValue}";
            }
        }

        private string GetPipeDisplayName(Pipe pipe)
        {
            try
            {
                string name = GetPipeName(pipe);
                string systemName = "Unknown System";

                Parameter systemParam = pipe.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM);
                if (systemParam != null)
                {
                    Element systemType = pipe.Document.GetElement(systemParam.AsElementId());
                    if (systemType != null)
                        systemName = systemType.Name;
                }

                double diameter = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble();
                double diameterInches = diameter * 12; // Convert from feet to inches

                LocationCurve locCurve = pipe.Location as LocationCurve;
                if (locCurve != null && locCurve.Curve is Line line)
                {
                    double length = line.Length;
                    return $"{name} - {systemName} - Ø{diameterInches:F1}\" - L: {length:F1}'";
                }
                else
                {
                    return $"{name} - {systemName} - Ø{diameterInches:F1}\"";
                }
            }
            catch
            {
                return $"Pipe {pipe.Id.IntegerValue}";
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class LevelItem : INotifyPropertyChanged
    {
        private bool _isSelected;

        public Level Level { get; }
        public double Elevation { get; }
        public string DisplayName { get; }

        public bool IsSelected
        {
            get => _isSelected;
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
            Elevation = level.Elevation;
            DisplayName = $"{level.Name} (Elev: {level.Elevation:F2}')";
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}