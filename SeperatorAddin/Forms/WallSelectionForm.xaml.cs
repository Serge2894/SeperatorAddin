using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace SeperatorAddin.Forms
{
    /// <summary>
    /// Interaction logic for WallSelectionForm.xaml
    /// </summary>
    public partial class WallSelectionForm : Window
    {
        private List<Level> allLevels;
        private Dictionary<System.Windows.Controls.CheckBox, Level> checkBoxToLevel;

        public WallSelectionForm(List<Level> levels)
        {
            InitializeComponent();

            allLevels = levels;
            checkBoxToLevel = new Dictionary<System.Windows.Controls.CheckBox, Level>();

            PopulateLevels();
            UpdateSelectedCount();
        }

        private void PopulateLevels()
        {
            levelStackPanel.Children.Clear();
            checkBoxToLevel.Clear();

            // Add levels in order (they should already be sorted by elevation)
            foreach (Level level in allLevels)
            {
                System.Windows.Controls.CheckBox cb = new System.Windows.Controls.CheckBox();
                cb.Content = $"{level.Name} (Elev: {Math.Round(level.ProjectElevation, 3)} ft)";
                cb.Tag = level;
                cb.Checked += CheckBox_CheckedChanged;
                cb.Unchecked += CheckBox_CheckedChanged;

                levelStackPanel.Children.Add(cb);
                checkBoxToLevel[cb] = level;
            }

            // Pre-select all levels by default
            foreach (System.Windows.Controls.CheckBox cb in checkBoxToLevel.Keys)
            {
                cb.IsChecked = true;
            }
        }

        private void CheckBox_CheckedChanged(object sender, RoutedEventArgs e)
        {
            UpdateSelectedCount();
            ValidateSelection();
        }

        private void UpdateSelectedCount()
        {
            int count = checkBoxToLevel.Keys.Count(cb => cb.IsChecked == true);
            txtSelectedCount.Text = $"{count} level{(count != 1 ? "s" : "")} selected";
        }

        private void ValidateSelection()
        {
            int selectedCount = checkBoxToLevel.Keys.Count(cb => cb.IsChecked == true);
            btnOK.IsEnabled = selectedCount >= 2;

            if (selectedCount < 2)
            {
                txtSelectedCount.Foreground = System.Windows.Media.Brushes.Red;
                txtSelectedCount.Text += " (minimum 2 required)";
            }
            else
            {
                txtSelectedCount.Foreground = System.Windows.Media.Brushes.Black;
            }
        }

        private void SelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (System.Windows.Controls.CheckBox cb in checkBoxToLevel.Keys)
            {
                cb.IsChecked = true;
            }
        }

        private void DeselectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (System.Windows.Controls.CheckBox cb in checkBoxToLevel.Keys)
            {
                cb.IsChecked = false;
            }
        }

        private void OK_Click(object sender, RoutedEventArgs e)
        {
            if (GetSelectedLevels().Count < 2)
            {
                System.Windows.MessageBox.Show("Please select at least 2 levels.", "Invalid Selection",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        public List<Level> GetSelectedLevels()
        {
            List<Level> selectedLevels = new List<Level>();

            foreach (var kvp in checkBoxToLevel)
            {
                if (kvp.Key.IsChecked == true)
                {
                    selectedLevels.Add(kvp.Value);
                }
            }

            return selectedLevels.OrderBy(l => l.ProjectElevation).ToList();
        }
    }
}