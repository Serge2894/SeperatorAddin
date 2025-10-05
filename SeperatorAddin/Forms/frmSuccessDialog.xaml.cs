using System.Windows;
using System.Windows.Input;

namespace SeperatorAddin.Forms
{
    /// <summary>
    /// Interaction logic for frmSuccessDialog.xaml
    /// </summary>
    public partial class frmSuccessDialog : Window
    {
        /// <summary>
        /// Default constructor.
        /// </summary>
        public frmSuccessDialog()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Constructor that accepts a custom message.
        /// </summary>
        /// <param name="message">The message to display in the dialog.</param>
        public frmSuccessDialog(string message) : this()
        {
            txtSuccessMessage.Text = message;
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        // Allows the window to be dragged by clicking and holding the title bar.
        private void Border_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                DragMove();
            }
        }
    }
}