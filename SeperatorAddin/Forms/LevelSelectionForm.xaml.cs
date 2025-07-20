using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace SeperatorAddin.Forms
{
    public partial class LevelSelectionForm : Window
    {
        private ObservableCollection<LevelItem> _levelItems;

        public List<Level> SelectedLevels
        {
            get
            {
                return _levelItems
                    .Where(item => item.IsSelected)
                    .Select(item => item.Level)
                    .ToList();
            }
        }

        public LevelSelectionForm(List<Level> levels)
        {
            InitializeComponent();

            // Initialize level items
            _levelItems = new ObservableCollection<LevelItem>();

            foreach (var level in levels)
            {
                _levelItems.Add(new LevelItem(level));
            }

            lvLevels.ItemsSource = _levelItems;
        }

        private void Border_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                DragMove();
            }
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedLevels.Count < 2)
            {
                System.Windows.MessageBox.Show("Please select at least 2 levels to split columns between.",
                    "Selection Required",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            DialogResult = true;
            Close();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void cbSelectAll_Checked(object sender, RoutedEventArgs e)
        {
            foreach (var item in _levelItems)
            {
                item.IsSelected = true;
            }
        }

        private void cbSelectAll_Unchecked(object sender, RoutedEventArgs e)
        {
            foreach (var item in _levelItems)
            {
                item.IsSelected = false;
            }
        }

        private void LevelCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            // Update Select All checkbox state
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

    // Helper class for level items in the list
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
            DisplayName = $"{level.Name} (Elev: {Math.Round(level.Elevation, 3)} ft)";
            IsSelected = false;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}