using System.Windows;
using System.Windows.Input;

namespace SeperatorAddin.Forms
{
    public partial class frmInfoDialog : Window
    {
        public frmInfoDialog()
        {
            InitializeComponent();
        }

        public frmInfoDialog(string message) : this(message, "Operation Complete")
        {
        }

        public frmInfoDialog(string message, string title) : this()
        {
            txtSuccessMessage.Text = message;
            TitleText.Text = $"STools - {title}";
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Border_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                DragMove();
            }
        }
    }
}