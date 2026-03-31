using System;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using Microsoft.Extensions.DependencyInjection;
using AICompanion.Desktop.Services.Security;

namespace AICompanion.Desktop.Views
{
    public partial class LoginWindow : Window
    {
        private readonly SecurityService? _securityService;
        private DispatcherTimer? _codeTimer;
        private int _codeSecondsRemaining = 60;
        private string _currentSecurityCode = "";
        private bool _isSecurityCodeMode;

        public bool IsAuthenticated { get; private set; }
        public string? AuthenticatedUser { get; private set; }

        public LoginWindow()
        {
            InitializeComponent();

            _securityService = App.ServiceProvider?.GetService<SecurityService>();

            // Focus username on load
            Loaded += (s, e) => UsernameBox.Focus();
        }

        /// <summary>
        /// Show login window for security code verification (not full login)
        /// </summary>
        public void ShowSecurityCodeVerification(string code)
        {
            _isSecurityCodeMode = true;
            _currentSecurityCode = code;
            SecurityCodeDisplay.Text = code;

            LoginPanel.Visibility = Visibility.Collapsed;
            RegisterPanel.Visibility = Visibility.Collapsed;
            SecurityCodePanel.Visibility = Visibility.Visible;

            StartCodeTimer();
        }

        private void StartCodeTimer()
        {
            _codeSecondsRemaining = 60;
            UpdateTimerDisplay();

            _codeTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(1)
            };

            _codeTimer.Tick += (s, e) =>
            {
                _codeSecondsRemaining--;
                UpdateTimerDisplay();

                if (_codeSecondsRemaining <= 0)
                {
                    _codeTimer.Stop();
                    ErrorMessage.Text = "Code expired. Please request a new one.";
                    ErrorMessage.Visibility = Visibility.Visible;
                    DialogResult = false;
                    Close();
                }
            };

            _codeTimer.Start();
        }

        private void UpdateTimerDisplay()
        {
            CodeTimerText.Text = $"Code expires in: {_codeSecondsRemaining} seconds";

            if (_codeSecondsRemaining <= 10)
            {
                CodeTimerText.Foreground = (System.Windows.Media.Brush)FindResource("ErrorBrush");
            }
        }

        private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount == 1)
            {
                DragMove();
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            _codeTimer?.Stop();
            DialogResult = false;
            Close();
        }

        private async void Login_Click(object sender, RoutedEventArgs e)
        {
            ErrorMessage.Visibility = Visibility.Collapsed;

            var username = UsernameBox.Text.Trim();
            var password = PasswordBox.Password;

            if (string.IsNullOrEmpty(username))
            {
                ShowError("Please enter your username or email");
                return;
            }

            if (string.IsNullOrEmpty(password))
            {
                ShowError("Please enter your password");
                return;
            }

            try
            {
                // Validate with security service
                if (_securityService != null)
                {
                    var result = await _securityService.AuthenticateAsync(username, password);
                    if (result.Success)
                    {
                        IsAuthenticated = true;
                        AuthenticatedUser = username;

                        if (RememberMeCheck.IsChecked == true && _securityService.CurrentUserId.HasValue)
                        {
                            await _securityService.CreatePersistentSessionAsync(_securityService.CurrentUserId.Value);
                        }
                        else
                        {
                            // Clear any old Remember Me token when logging in without the checkbox
                            await _securityService.ClearPersistentSessionAsync();
                        }

                        DialogResult = true;
                        Close();
                        return;
                    }
                    else
                    {
                        ShowError(result.ErrorMessage ?? "Invalid username or password");
                        return;
                    }
                }

                ShowError("Security service unavailable. Please restart the application.");
            }
            catch (Exception ex)
            {
                ShowError($"Login failed: {ex.Message}");
            }
        }

        private async void Register_Click(object sender, RoutedEventArgs e)
        {
            RegisterErrorMessage.Visibility = Visibility.Collapsed;

            var username = RegisterUsernameBox.Text.Trim();
            var email = RegisterEmailBox.Text.Trim();
            var password = RegisterPasswordBox.Password;
            var confirmPassword = RegisterConfirmPasswordBox.Password;
            var pin = RegisterPinBox.Password;

            // Validation
            if (string.IsNullOrEmpty(username) || username.Length < 3)
            {
                ShowRegisterError("Username must be at least 3 characters");
                return;
            }

            if (string.IsNullOrEmpty(email) || !email.Contains("@"))
            {
                ShowRegisterError("Please enter a valid email address");
                return;
            }

            if (string.IsNullOrEmpty(password) || password.Length < 8)
            {
                ShowRegisterError("Password must be at least 8 characters");
                return;
            }

            var hasUpper = false;
            var hasDigit = false;
            foreach (var c in password)
            {
                if (char.IsUpper(c)) hasUpper = true;
                if (char.IsDigit(c)) hasDigit = true;
            }
            if (!hasUpper || !hasDigit)
            {
                ShowRegisterError("Password must contain at least 1 uppercase letter and 1 number");
                return;
            }

            if (password != confirmPassword)
            {
                ShowRegisterError("Passwords do not match");
                return;
            }

            // PIN validation
            if (string.IsNullOrEmpty(pin) || pin.Length < 4 || pin.Length > 6)
            {
                ShowRegisterError("PIN must be 4-6 digits");
                return;
            }
            foreach (var c in pin)
            {
                if (!char.IsDigit(c))
                {
                    ShowRegisterError("PIN must contain only digits");
                    return;
                }
            }

            try
            {
                if (_securityService != null)
                {
                    var result = await _securityService.RegisterAsync(username, email, password, pin);
                    if (result.Success)
                    {
                        System.Windows.MessageBox.Show(
                            "Account created successfully!\nYour security PIN has been saved.\nPlease sign in.",
                            "Registration Complete",
                            System.Windows.MessageBoxButton.OK,
                            System.Windows.MessageBoxImage.Information);
                        ShowLogin_Click(sender, e);
                        UsernameBox.Text = username;
                        return;
                    }
                    else
                    {
                        ShowRegisterError(result.ErrorMessage ?? "Registration failed");
                        return;
                    }
                }

                ShowRegisterError("Security service unavailable. Please restart the application.");
            }
            catch (Exception ex)
            {
                ShowRegisterError($"Registration failed: {ex.Message}");
            }
        }

        private void VerifyCode_Click(object sender, RoutedEventArgs e)
        {
            var enteredCode = SecurityCodeInput.Text.Trim();

            if (enteredCode == _currentSecurityCode)
            {
                _codeTimer?.Stop();
                IsAuthenticated = true;
                DialogResult = true;
                Close();
            }
            else
            {
                SecurityCodeInput.Clear();
                System.Windows.MessageBox.Show("Invalid code. Please try again.", "Verification Failed",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
            }
        }

        private void ForgotPassword_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.MessageBox.Show(
                "Password reset instructions will be sent to your registered email address.\n\n" +
                "If you don't have access to your email, please contact your administrator.",
                "Forgot Password",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Information);
        }

        private void ShowRegister_Click(object sender, RoutedEventArgs e)
        {
            LoginPanel.Visibility = Visibility.Collapsed;
            RegisterPanel.Visibility = Visibility.Visible;
            RegisterUsernameBox.Focus();
        }

        private void ShowLogin_Click(object sender, RoutedEventArgs e)
        {
            RegisterPanel.Visibility = Visibility.Collapsed;
            LoginPanel.Visibility = Visibility.Visible;
            UsernameBox.Focus();
        }

        // SkipLogin removed — all authentication flows through SecurityService + Database

        private void ShowError(string message)
        {
            ErrorMessage.Text = message;
            ErrorMessage.Visibility = Visibility.Visible;
        }

        private void ShowRegisterError(string message)
        {
            RegisterErrorMessage.Text = message;
            RegisterErrorMessage.Visibility = Visibility.Visible;
        }

        /// <summary>
        /// Process voice input for security code
        /// </summary>
        public bool TryVoiceCode(string spokenText)
        {
            if (!_isSecurityCodeMode) return false;

            // Extract digits from spoken text
            var digits = new System.Text.StringBuilder();
            foreach (var c in spokenText)
            {
                if (char.IsDigit(c))
                {
                    digits.Append(c);
                }
            }

            // Handle spoken numbers like "one two three"
            var words = spokenText.ToLowerInvariant().Split(' ');
            foreach (var word in words)
            {
                var digit = word switch
                {
                    "zero" or "oh" => "0",
                    "one" => "1",
                    "two" or "to" or "too" => "2",
                    "three" => "3",
                    "four" or "for" => "4",
                    "five" => "5",
                    "six" => "6",
                    "seven" => "7",
                    "eight" => "8",
                    "nine" => "9",
                    _ => null
                };

                if (digit != null)
                {
                    digits.Append(digit);
                }
            }

            var code = digits.ToString();
            if (code.Length >= 6)
            {
                code = code.Substring(0, 6);
                SecurityCodeInput.Text = code;

                if (code == _currentSecurityCode)
                {
                    _codeTimer?.Stop();
                    IsAuthenticated = true;
                    DialogResult = true;
                    Close();
                    return true;
                }
            }

            return false;
        }
    }
}
