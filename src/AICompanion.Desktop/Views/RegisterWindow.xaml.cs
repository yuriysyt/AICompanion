using System;
using System.Windows;
using System.Windows.Input;
using Microsoft.Extensions.DependencyInjection;
using AICompanion.Desktop.Services.Security;

namespace AICompanion.Desktop.Views
{
    public partial class RegisterWindow : Window
    {
        private readonly SecurityService? _securityService;

        public RegisterWindow()
        {
            InitializeComponent();
            _securityService = App.ServiceProvider?.GetService<SecurityService>();
            Loaded += (s, e) => UsernameBox.Focus();
        }

        private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount == 1) DragMove();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void BackToLogin_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private async void Register_Click(object sender, RoutedEventArgs e)
        {
            ErrorMessage.Visibility = Visibility.Collapsed;

            var username = UsernameBox.Text.Trim();
            var email    = EmailBox.Text.Trim();
            var password = PasswordBox.Password;
            var confirm  = ConfirmPasswordBox.Password;
            var pin      = PinBox.Password;

            // ── Validation ───────────────────────────────────────────────────────
            if (string.IsNullOrEmpty(username) || username.Length < 3)
            {
                ShowError("Username must be at least 3 characters");
                return;
            }

            if (password.Length < 8)
            {
                ShowError("Password must be at least 8 characters");
                return;
            }

            bool hasUpper = false, hasDigit = false;
            foreach (var c in password)
            {
                if (char.IsUpper(c)) hasUpper = true;
                if (char.IsDigit(c)) hasDigit = true;
            }
            if (!hasUpper || !hasDigit)
            {
                ShowError("Password must contain at least 1 uppercase letter and 1 digit");
                return;
            }

            if (password != confirm)
            {
                ShowError("Passwords do not match");
                return;
            }

            if (string.IsNullOrEmpty(pin) || pin.Length < 4 || pin.Length > 6)
            {
                ShowError("PIN must be 4 to 6 digits");
                return;
            }

            foreach (var c in pin)
            {
                if (!char.IsDigit(c))
                {
                    ShowError("PIN must contain digits only");
                    return;
                }
            }

            // ── Register ─────────────────────────────────────────────────────────
            try
            {
                if (_securityService == null)
                {
                    ShowError("Security service unavailable. Please restart the application.");
                    return;
                }

                var result = await _securityService.RegisterWithPinAsync(username, email, password, pin);
                if (result.Success)
                {
                    System.Windows.MessageBox.Show(
                        "Account created successfully!\nPlease sign in with your new credentials.",
                        "Registration Complete",
                        System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Information);
                    DialogResult = true;
                    Close();
                }
                else
                {
                    ShowError(result.ErrorMessage ?? "Registration failed. Please try again.");
                }
            }
            catch (Exception ex)
            {
                ShowError($"Registration failed: {ex.Message}");
            }
        }

        private void ShowError(string message)
        {
            ErrorMessage.Text = message;
            ErrorMessage.Visibility = Visibility.Visible;
        }
    }
}
