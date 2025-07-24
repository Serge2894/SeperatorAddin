using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace SeperatorAddin.Forms
{
    /// <summary>
    /// Interaction logic for PipeLevelSelectionForm.xaml
    /// </summary>
    public partial class PipeSelectionForm : Window
    {
        /// <summary>
        /// A public property to access the list of selected level names after the dialog is closed.
        /// </summary>
        public List<string> SelectedLevelNames { get; private set; }

        /// <summary>
        /// Constructor for the level selection form.
        /// </summary>
        /// <param name="availableLevelNames">A list of strings representing the level names to display.</param>
        public PipeSelectionForm(List<string> availableLevelNames)
        {
            InitializeComponent();

            // Populate the ListBox with the provided level names
            LevelsListBox.ItemsSource = availableLevelNames;

            // Initialize the selection count
            UpdateSelectionCount();
        }

        /// <summary>
        /// Updates the text block that shows how many items are currently selected.
        /// </summary>
        private void UpdateSelectionCount()
        {
            int count = LevelsListBox.SelectedItems.Count;
            SelectionCountLabel.Text = $"{count} level{(count == 1 ? "" : "s")} selected";
        }

        private void LevelsListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateSelectionCount();
        }

        private void SelectAllButton_Click(object sender, RoutedEventArgs e)
        {
            LevelsListBox.SelectAll();
        }

        private void DeselectButton_Click(object sender, RoutedEventArgs e)
        {
            LevelsListBox.UnselectAll();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            // Save the selected items to the public property
            SelectedLevelNames = LevelsListBox.SelectedItems.Cast<string>().ToList();
            this.DialogResult = true;
            this.Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }
    }
}