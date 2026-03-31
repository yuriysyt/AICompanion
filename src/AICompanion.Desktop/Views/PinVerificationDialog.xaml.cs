using System.Windows;
using System.Windows.Input;

namespace AICompanion.Desktop.Views
{
    /// <summary>
    /// Dialog that asks the user to enter their security PIN before a high-risk operation proceeds.
    /// Set <see cref="OperationDescription"/> before calling ShowDialog().
    /// After ShowDialog() returns true, read <see cref="EnteredPin"/>.
    /// </summary>
    public partial class PinVerificationDialog : Window
    {
        /// <summary>Human-readable description of the operation being authorized.</summary>
        public string OperationDescription
        {
            get => DescriptionText.Text;
            set => DescriptionText.Text = $"Enter your security PIN to confirm:\n\"{value}\"";
        }

        /// <summary>The PIN the user typed (only valid when DialogResult == true).</summary>
        public string EnteredPin { get; private set; } = "";

        public PinVerificationDialog()
        {
            InitializeComponent();
            Loaded += (_, _) => PinInputBox.Focus();

            // Allow Enter key to confirm
            PinInputBox.KeyDown += (_, e) =>
            {
                if (e.Key == Key.Enter) Confirm_Click(this, new RoutedEventArgs());
            };
        }

        private void Confirm_Click(object sender, RoutedEventArgs e)
        {
            var pin = PinInputBox.Password;

            if (string.IsNullOrEmpty(pin) || pin.Length < 4)
            {
                ErrorText.Text = "PIN must be at least 4 digits.";
                ErrorText.Visibility = Visibility.Visible;
                return;
            }

            foreach (var c in pin)
            {
                if (!char.IsDigit(c))
                {
                    ErrorText.Text = "PIN must contain only digits.";
                    ErrorText.Visibility = Visibility.Visible;
                    return;
                }
            }

            EnteredPin = pin;
            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            EnteredPin = "";
            DialogResult = false;
            Close();
        }
    }
}
