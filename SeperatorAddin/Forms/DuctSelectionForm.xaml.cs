using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace SeperatorAddin.Forms
{
    public partial class DuctSelectionForm : Window, INotifyPropertyChanged
    {
        private ObservableCollection<DuctItem> _ducts;
        private ObservableCollection<LevelItem> _levels;
        private List<DuctItem> _allDucts;

        public List<Duct> SelectedDucts { get; private set; }
        public List<Level> SelectedLevels { get; private set; }

        public DuctSelectionForm(List<Duct> ducts, List<Level> levels)
        {
            InitializeComponent();
            DataContext = this;

            // Initialize collections
            _allDucts = new List<DuctItem>();
            foreach (var duct in ducts)
            {
                _allDucts.Add(new DuctItem(duct));
            }
            _ducts = new ObservableCollection<DuctItem>(_allDucts.OrderBy(d => d.Name));

            _levels = new ObservableCollection<LevelItem>();
            foreach (var level in levels.OrderBy(l => l.Elevation))
            {
                _levels.Add(new LevelItem(level));
            }

            // Set ItemsSource
            lbDucts.ItemsSource = _ducts;
            lbLevels.ItemsSource = _levels;

            // Initialize system filter
            InitializeSystemFilter();

            // Initialize
            UpdateCounts();
            UpdateOKButton();
        }

        private void InitializeSystemFilter()
        {
            var systems = new HashSet<string> { "All Systems" };
            foreach (var duct in _allDucts)
            {
                if (!string.IsNullOrEmpty(duct.SystemName))
                    systems.Add(duct.SystemName);
            }

            cmbSystem.ItemsSource = systems.OrderBy(s => s).ToList();
            cmbSystem.SelectedIndex = 0;
        }

        private void ApplyFilters()
        {
            string textFilter = txtFilter.Text.ToLower();
            string systemFilter = cmbSystem.SelectedItem?.ToString() ?? "All Systems";

            var filtered = _allDucts.Where(d =>
            {
                bool matchesText = string.IsNullOrEmpty(textFilter) ||
                                   d.DisplayName.ToLower().Contains(textFilter);
                bool matchesSystem = systemFilter == "All Systems" ||
                                     d.SystemName == systemFilter;
                return matchesText && matchesSystem;
            }).OrderBy(d => d.Name);

            _ducts.Clear();
            foreach (var duct in filtered)
            {
                _ducts.Add(duct);
            }

            UpdateDuctCount();
        }

        private void txtFilter_TextChanged(object sender, TextChangedEventArgs e)
        {
            ApplyFilters();
        }

        private void cmbSystem_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_allDucts != null)
                ApplyFilters();
        }

        private void btnSelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var duct in _ducts)
            {
                duct.IsSelected = true;
            }
            UpdateDuctCount();
        }

        private void btnSelectNone_Click(object sender, RoutedEventArgs e)
        {
            foreach (var duct in _ducts)
            {
                duct.IsSelected = false;
            }
            UpdateDuctCount();
        }

        private void DuctCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            UpdateDuctCount();
        }

        private void DuctCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            UpdateDuctCount();
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
            UpdateDuctCount();
            UpdateLevelCount();
        }

        private void UpdateDuctCount()
        {
            int count = _ducts.Count(d => d.IsSelected);
            txtDuctCount.Text = $"{count} selected";
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
            SelectedDucts = new List<Duct>();
            SelectedLevels = new List<Level>();

            foreach (var item in _ducts.Where(d => d.IsSelected))
            {
                SelectedDucts.Add(item.Duct);
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

    public class DuctItem : INotifyPropertyChanged
    {
        private bool _isSelected;

        public Duct Duct { get; }
        public string Name { get; }
        public string SystemName { get; }
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

        public DuctItem(Duct duct)
        {
            Duct = duct;
            Name = GetDuctName(duct);
            SystemName = GetDuctSystemName(duct);
            DisplayName = GetDuctDisplayName(duct);
        }

        private string GetDuctName(Duct duct)
        {
            try
            {
                string mark = duct.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)?.AsString();
                if (!string.IsNullOrEmpty(mark))
                    return mark;

                return $"Duct {duct.Id.IntegerValue}";
            }
            catch
            {
                return $"Duct {duct.Id.IntegerValue}";
            }
        }

        private string GetDuctSystemName(Duct duct)
        {
            try
            {
                Parameter systemParam = duct.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM);
                if (systemParam != null)
                {
                    Element systemType = duct.Document.GetElement(systemParam.AsElementId());
                    if (systemType != null)
                        return systemType.Name;
                }
                return "Unknown System";
            }
            catch
            {
                return "Unknown System";
            }
        }

        private string GetDuctDisplayName(Duct duct)
        {
            try
            {
                string name = GetDuctName(duct);
                string systemName = SystemName;

                // Get duct dimensions and shape
                string sizeString = "";
                ConnectorProfileType shape = ConnectorProfileType.Invalid;

                foreach (Connector conn in duct.ConnectorManager.Connectors)
                {
                    shape = conn.Shape;
                    break;
                }

                if (shape == ConnectorProfileType.Round)
                {
                    double diameter = duct.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM).AsDouble() * 12;
                    sizeString = $"Ø{diameter:F1}\"";
                }
                else
                {
                    double width = duct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM).AsDouble();
                    double height = duct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM).AsDouble();
                    double widthInches = width * 12;
                    double heightInches = height * 12;
                    sizeString = $"{widthInches:F1}\"x{heightInches:F1}\"";
                }

                LocationCurve locCurve = duct.Location as LocationCurve;
                if (locCurve != null && locCurve.Curve is Line line)
                {
                    double length = line.Length;
                    return $"{name} - {systemName} - {sizeString} - L: {length:F1}'";
                }
                else
                {
                    return $"{name} - {systemName} - {sizeString}";
                }
            }
            catch
            {
                return $"Duct {duct.Id.IntegerValue}";
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