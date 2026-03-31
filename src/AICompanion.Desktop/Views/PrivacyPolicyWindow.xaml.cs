using System.Windows;
using AICompanion.Desktop.Services.Database;

namespace AICompanion.Desktop.Views
{
    /// <summary>
    /// First-run privacy policy acceptance window.
    /// Show before login; close only when the user ticks the checkbox and clicks Continue.
    /// Acceptance is persisted via <see cref="DatabaseService.MarkFirstRunConsent"/>.
    /// </summary>
    public partial class PrivacyPolicyWindow : Window
    {
        /// <summary>True when the user clicked Continue after ticking the checkbox.</summary>
        public bool Accepted { get; private set; }

        public PrivacyPolicyWindow()
        {
            InitializeComponent();
        }

        private void AgreeCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            ContinueButton.IsEnabled = AgreeCheckBox.IsChecked == true;
        }

        private void Continue_Click(object sender, RoutedEventArgs e)
        {
            if (AgreeCheckBox.IsChecked != true) return;

            // Persist first-run consent flag
            DatabaseService.MarkFirstRunConsent();

            Accepted = true;
            DialogResult = true;
            Close();
        }

        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            Accepted = false;
            DialogResult = false;
            Close();
        }
    }
}
