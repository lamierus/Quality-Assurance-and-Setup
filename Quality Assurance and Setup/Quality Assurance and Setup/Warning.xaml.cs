using System.Windows;

namespace Quality_Assurance_and_Setup {
    /// <summary>
    /// Interaction logic for Warning.xaml
    /// </summary>
    public partial class Warning : Window {
        public Warning(string message) {
            InitializeComponent();
            WarningText.Text = message;
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e) {
            DialogResult = false;
            Close();
        }

        private void btnContinue_Click(object sender, RoutedEventArgs e) {
            DialogResult = true;
            Close();
        }
    }
}
