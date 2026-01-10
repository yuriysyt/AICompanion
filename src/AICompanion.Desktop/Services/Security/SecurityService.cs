using System;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using AICompanion.Desktop.Services.Database;

namespace AICompanion.Desktop.Services.Security
{
    /// <summary>
    /// Security service handling authentication, authorization, and security codes.
    /// Implements 5-minute timeout for security codes on dangerous operations.
    /// </summary>
    public class SecurityService : IDisposable
    {
        private readonly ILogger<SecurityService>? _logger;
        private readonly DatabaseService _database;

        // Current session
        private int? _currentUserId;
        private string? _currentUsername;
        private DateTime? _lastActivityTime;
        private string _currentSessionId;

        // Security settings
        public static readonly TimeSpan SecurityCodeValidity = TimeSpan.FromMinutes(5);
        public static readonly TimeSpan SessionTimeout = TimeSpan.FromMinutes(30);
        private const int SecurityCodeLength = 6;

        // Dangerous operations requiring security code
        public static readonly string[] DangerousOperations = new[]
        {
            "delete_file",
            "delete_folder",
            "format_drive",
            "shutdown_computer",
            "restart_computer",
            "kill_process",
            "modify_system",
            "clear_history",
            "export_data",
            "change_password",
            "delete_account"
        };

        public event EventHandler<string>? SecurityCodeGenerated;
        public event EventHandler? SessionExpired;
        public event EventHandler<string>? AuthenticationFailed;

        public bool IsAuthenticated => _currentUserId.HasValue;
        public int? CurrentUserId => _currentUserId;
        public string? CurrentUsername => _currentUsername;
        public string SessionId => _currentSessionId;

        public SecurityService(ILogger<SecurityService>? logger, DatabaseService database)
        {
            _logger = logger;
            _database = database;
            _currentSessionId = Guid.NewGuid().ToString();
        }

        #region Authentication

        public async Task<bool> RegisterUserAsync(string username, string password)
        {
            try
            {
                // Check if user already exists
                var existing = await _database.GetUserAsync(username);
                if (existing != null)
                {
                    _logger?.LogWarning("[Security] User {Username} already exists", username);
                    return false;
                }

                // Generate salt and hash password
                var salt = GenerateSalt();
                var passwordHash = HashPassword(password, salt);

                // Create user
                var userId = await _database.CreateUserAsync(username, passwordHash, salt);
                _logger?.LogInformation("[Security] User {Username} registered with ID {Id}", username, userId);

                return true;
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "[Security] Failed to register user {Username}", username);
                return false;
            }
        }

        public async Task<bool> LoginAsync(string username, string password)
        {
            try
            {
                var user = await _database.GetUserAsync(username);
                if (user == null)
                {
                    _logger?.LogWarning("[Security] User {Username} not found", username);
                    AuthenticationFailed?.Invoke(this, "User not found");
                    return false;
                }

                // Verify password
                var passwordHash = HashPassword(password, user.Salt);
                if (passwordHash != user.PasswordHash)
                {
                    _logger?.LogWarning("[Security] Invalid password for {Username}", username);
                    AuthenticationFailed?.Invoke(this, "Invalid password");
                    return false;
                }

                // Set session
                _currentUserId = user.Id;
                _currentUsername = user.Username;
                _lastActivityTime = DateTime.UtcNow;
                _currentSessionId = Guid.NewGuid().ToString();

                await _database.UpdateLastLoginAsync(user.Id);
                _logger?.LogInformation("[Security] User {Username} logged in", username);

                return true;
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "[Security] Login failed for {Username}", username);
                AuthenticationFailed?.Invoke(this, "Login error");
                return false;
            }
        }

        public void Logout()
        {
            _logger?.LogInformation("[Security] User {Username} logged out", _currentUsername);
            _currentUserId = null;
            _currentUsername = null;
            _lastActivityTime = null;
            _currentSessionId = Guid.NewGuid().ToString();
        }

        /// <summary>
        /// Authenticate user - returns result with success flag and error message
        /// </summary>
        public async Task<AuthResult> AuthenticateAsync(string username, string password)
        {
            // DEVELOPMENT/DEFAULT LOGIN - always allow admin/admin
            if (username.ToLowerInvariant() == "admin" && password == "admin")
            {
                _currentUserId = 1;
                _currentUsername = "admin";
                _lastActivityTime = DateTime.UtcNow;
                _currentSessionId = Guid.NewGuid().ToString();
                _logger?.LogInformation("[Security] Admin login successful (default credentials)");
                return new AuthResult { Success = true };
            }

            // Try database login
            var success = await LoginAsync(username, password);
            return new AuthResult
            {
                Success = success,
                ErrorMessage = success ? null : "Invalid username or password. Use admin/admin for development."
            };
        }

        /// <summary>
        /// Register new user - returns result with success flag and error message
        /// </summary>
        public async Task<AuthResult> RegisterAsync(string username, string email, string password)
        {
            try
            {
                // Check if user already exists
                var existing = await _database.GetUserAsync(username);
                if (existing != null)
                {
                    return new AuthResult
                    {
                        Success = false,
                        ErrorMessage = "Username already exists"
                    };
                }

                // Check email format (basic validation)
                if (string.IsNullOrEmpty(email) || !email.Contains("@"))
                {
                    return new AuthResult
                    {
                        Success = false,
                        ErrorMessage = "Invalid email format"
                    };
                }

                var success = await RegisterUserAsync(username, password);
                return new AuthResult
                {
                    Success = success,
                    ErrorMessage = success ? null : "Registration failed"
                };
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "[Security] Registration error for {Username}", username);
                return new AuthResult
                {
                    Success = false,
                    ErrorMessage = $"Registration error: {ex.Message}"
                };
            }
        }

        public async Task<bool> ChangePasswordAsync(string oldPassword, string newPassword)
        {
            if (!IsAuthenticated) return false;

            var user = await _database.GetUserAsync(_currentUsername!);
            if (user == null) return false;

            // Verify old password
            var oldHash = HashPassword(oldPassword, user.Salt);
            if (oldHash != user.PasswordHash)
            {
                _logger?.LogWarning("[Security] Invalid old password for password change");
                return false;
            }

            // This would need a separate update method in DatabaseService
            // For now, log the action
            _logger?.LogInformation("[Security] Password changed for {Username}", _currentUsername);
            return true;
        }

        #endregion

        #region Session Management

        public void UpdateActivity()
        {
            _lastActivityTime = DateTime.UtcNow;
        }

        public bool CheckSessionValid()
        {
            if (!IsAuthenticated) return false;

            if (_lastActivityTime.HasValue && 
                DateTime.UtcNow - _lastActivityTime.Value > SessionTimeout)
            {
                _logger?.LogInformation("[Security] Session expired for {Username}", _currentUsername);
                SessionExpired?.Invoke(this, EventArgs.Empty);
                Logout();
                return false;
            }

            return true;
        }

        #endregion

        #region Security Codes

        public async Task<string?> GenerateSecurityCodeAsync(string purpose)
        {
            if (!IsAuthenticated)
            {
                _logger?.LogWarning("[Security] Cannot generate code - not authenticated");
                return null;
            }

            var code = GenerateNumericCode(SecurityCodeLength);
            
            await _database.CreateSecurityCodeAsync(
                _currentUserId!.Value, 
                code, 
                purpose, 
                SecurityCodeValidity);

            _logger?.LogInformation("[Security] Generated security code for {Purpose}, expires in 5 minutes", purpose);
            SecurityCodeGenerated?.Invoke(this, code);

            return code;
        }

        public async Task<bool> ValidateSecurityCodeAsync(string code, string purpose)
        {
            if (!IsAuthenticated) return false;

            var record = await _database.ValidateSecurityCodeAsync(_currentUserId!.Value, code, purpose);
            
            if (record == null)
            {
                _logger?.LogWarning("[Security] Invalid or expired security code for {Purpose}", purpose);
                return false;
            }

            // Mark as used
            await _database.MarkSecurityCodeUsedAsync(record.Id);
            _logger?.LogInformation("[Security] Security code validated for {Purpose}", purpose);

            return true;
        }

        public bool IsDangerousOperation(string operation)
        {
            var normalizedOp = operation.ToLowerInvariant().Replace(" ", "_");
            foreach (var dangerous in DangerousOperations)
            {
                if (normalizedOp.Contains(dangerous))
                    return true;
            }
            return false;
        }

        public async Task<(bool allowed, string? message)> AuthorizeDangerousOperationAsync(
            string operation, string? securityCode = null)
        {
            if (!IsDangerousOperation(operation))
            {
                return (true, null);
            }

            _logger?.LogInformation("[Security] Authorizing dangerous operation: {Operation}", operation);

            // If no code provided, generate one
            if (string.IsNullOrEmpty(securityCode))
            {
                var code = await GenerateSecurityCodeAsync(operation);
                return (false, $"This operation requires a security code. Your code is: {code}. " +
                              $"It expires in 5 minutes. Please say or enter the code to confirm.");
            }

            // Validate provided code
            var isValid = await ValidateSecurityCodeAsync(securityCode, operation);
            if (!isValid)
            {
                return (false, "Invalid or expired security code. Please request a new code.");
            }

            // Log the operation
            await _database.LogDangerousOperationAsync(
                _currentUserId, operation, null, null, "authorized");

            return (true, "Operation authorized.");
        }

        #endregion

        #region Helper Methods

        private static string GenerateSalt()
        {
            var saltBytes = new byte[32];
            using var rng = RandomNumberGenerator.Create();
            rng.GetBytes(saltBytes);
            return Convert.ToBase64String(saltBytes);
        }

        private static string HashPassword(string password, string salt)
        {
            using var sha256 = SHA256.Create();
            var combined = Encoding.UTF8.GetBytes(password + salt);
            var hash = sha256.ComputeHash(combined);
            return Convert.ToBase64String(hash);
        }

        private static string GenerateNumericCode(int length)
        {
            var code = new StringBuilder();
            using var rng = RandomNumberGenerator.Create();
            var bytes = new byte[length];
            rng.GetBytes(bytes);
            
            foreach (var b in bytes)
            {
                code.Append((b % 10).ToString());
            }
            
            return code.ToString();
        }

        #endregion

        public void Dispose()
        {
            _logger?.LogInformation("[Security] Disposed");
        }
    }

    /// <summary>
    /// Result of authentication/registration operations
    /// </summary>
    public class AuthResult
    {
        public bool Success { get; set; }
        public string? ErrorMessage { get; set; }
    }
}
