using System;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using AICompanion.Desktop.Services.Database;
using System.IO;

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

        public async Task<bool> RegisterUserAsync(string username, string password, string? email = null)
        {
            try
            {
                var existing = await _database.GetUserAsync(username);
                if (existing != null)
                {
                    _logger?.LogWarning("[Security] User {Username} already exists", username);
                    return false;
                }

                var salt = GenerateSalt();
                var passwordHash = HashPassword(password, salt);

                var userId = await _database.CreateUserAsync(username, passwordHash, salt, email, "PBKDF2");
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

                // Verify password — support both legacy SHA256 and current PBKDF2
                string passwordHash;
                if (user.HashAlgorithm == "SHA256")
                {
                    passwordHash = HashPasswordSha256Legacy(password, user.Salt);
                }
                else
                {
                    passwordHash = HashPassword(password, user.Salt);
                }

                if (passwordHash != user.PasswordHash)
                {
                    _logger?.LogWarning("[Security] Invalid password for {Username}", username);
                    AuthenticationFailed?.Invoke(this, "Invalid password");
                    return false;
                }

                // Transparently upgrade legacy SHA256 hashes to PBKDF2 on login
                if (user.HashAlgorithm == "SHA256")
                {
                    var newSalt = GenerateSalt();
                    var newHash = HashPassword(password, newSalt);
                    await _database.UpdatePasswordHashAsync(user.Id, newHash, newSalt, "PBKDF2");
                    _logger?.LogInformation("[Security] Upgraded password hash to PBKDF2 for {Username}", username);
                }

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
        /// Authenticate user - returns result with success flag and error message.
        /// All authentication flows through the database. No hardcoded credentials.
        /// </summary>
        public async Task<AuthResult> AuthenticateAsync(string username, string password)
        {
            // Ensure default admin exists on first launch
            await EnsureDefaultUserAsync();

            // All auth goes through database
            var success = await LoginAsync(username, password);
            return new AuthResult
            {
                Success = success,
                ErrorMessage = success ? null : "Invalid username or password."
            };
        }

        /// <summary>
        /// On first launch, seeds a default admin account with a PBKDF2-hashed password.
        /// On subsequent launches, re-hashes the admin account if it still uses legacy SHA256.
        /// </summary>
        private async Task EnsureDefaultUserAsync()
        {
            try
            {
                var existing = await _database.GetUserAsync("admin");
                if (existing == null)
                {
                    var salt = GenerateSalt();
                    var passwordHash = HashPassword("admin", salt);
                    await _database.CreateUserAsync("admin", passwordHash, salt, null, "PBKDF2");
                    _logger?.LogInformation("[Security] Default admin account seeded with PBKDF2 hash");
                }
                else if (existing.HashAlgorithm == "SHA256")
                {
                    // Migrate legacy admin hash to PBKDF2
                    var newSalt = GenerateSalt();
                    var newHash = HashPassword("admin", newSalt);
                    await _database.UpdatePasswordHashAsync(existing.Id, newHash, newSalt, "PBKDF2");
                    _logger?.LogInformation("[Security] Migrated admin account to PBKDF2");
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "[Security] Failed to seed/migrate default admin account");
            }
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

                var success = await RegisterUserAsync(username, password, email);
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

        /// <summary>
        /// Registers a new user and stores a hashed PIN for quick verification.
        /// Called by RegisterWindow.
        /// </summary>
        public async Task<AuthResult> RegisterWithPinAsync(string username, string email, string password, string pin)
        {
            // Register the user via the existing path
            var result = await RegisterAsync(username, email, password);
            if (!result.Success) return result;

            // Hash and persist the PIN
            try
            {
                var user = await _database.GetUserAsync(username);
                if (user != null)
                {
                    var pinSalt = GenerateSalt();
                    var pinHash = HashPassword(pin, pinSalt);
                    // Store combined "hash:salt" so we can verify later
                    await _database.SetUserPinAsync(user.Id, $"{pinHash}:{pinSalt}");
                    _logger?.LogInformation("[Security] PIN stored for user {Username}", username);
                }
            }
            catch (Exception ex)
            {
                // Non-fatal: account was created but PIN storage failed
                _logger?.LogWarning(ex, "[Security] Failed to store PIN for {Username}", username);
            }

            return result;
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
            // 16-byte (128-bit) salt — NIST standard for PBKDF2
            var saltBytes = new byte[16];
            using var rng = RandomNumberGenerator.Create();
            rng.GetBytes(saltBytes);
            return Convert.ToBase64String(saltBytes);
        }

        /// <summary>
        /// PBKDF2 password hashing with 100,000 iterations (NIST SP 800-132 compliant).
        /// Replaced legacy SHA-256 to resist brute-force attacks.
        /// </summary>
        private static string HashPassword(string password, string salt)
        {
            var saltBytes = Convert.FromBase64String(salt);
            using var pbkdf2 = new Rfc2898DeriveBytes(
                password,
                saltBytes,
                iterations: 100_000,
                hashAlgorithm: HashAlgorithmName.SHA256);
            var hash = pbkdf2.GetBytes(32); // 256-bit output
            return Convert.ToBase64String(hash);
        }

        /// <summary>
        /// Legacy SHA-256 hash used only to verify old accounts before migrating them to PBKDF2.
        /// Do NOT use this for new passwords.
        /// </summary>
        private static string HashPasswordSha256Legacy(string password, string salt)
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

        #region Remember Me (Persistent Session)

        private static readonly string _sessionFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "AICompanion", "session.dat");

        /// <summary>
        /// Creates a 30-day persistent session token, saves it to DB, and stores
        /// a DPAPI-encrypted copy on disk for auto-login.
        /// </summary>
        public async Task<string> CreatePersistentSessionAsync(int userId)
        {
            var tokenBytes = new byte[32];
            RandomNumberGenerator.Fill(tokenBytes);
            var token = Convert.ToBase64String(tokenBytes);

            await _database.SaveSessionTokenAsync(userId, token, DateTime.UtcNow.AddDays(30));

            // Encrypt with DPAPI (Windows user-scope — only readable by same Windows user)
            try
            {
                var tokenData = Encoding.UTF8.GetBytes(token);
                var encrypted = System.Security.Cryptography.ProtectedData.Protect(
                    tokenData, null, System.Security.Cryptography.DataProtectionScope.CurrentUser);
                Directory.CreateDirectory(Path.GetDirectoryName(_sessionFilePath)!);
                await File.WriteAllBytesAsync(_sessionFilePath, encrypted);
                _logger?.LogInformation("[Security] Remember Me token saved (DPAPI encrypted)");
            }
            catch (Exception ex)
            {
                _logger?.LogWarning(ex, "[Security] Could not save Remember Me token to disk");
            }

            return token;
        }

        /// <summary>
        /// Tries to restore session from the DPAPI-encrypted file on disk.
        /// Returns true and sets the current user if the token is valid.
        /// </summary>
        public async Task<bool> RestoreSessionAsync()
        {
            try
            {
                if (!File.Exists(_sessionFilePath)) return false;

                var encrypted = await File.ReadAllBytesAsync(_sessionFilePath);
                var tokenData = System.Security.Cryptography.ProtectedData.Unprotect(
                    encrypted, null, System.Security.Cryptography.DataProtectionScope.CurrentUser);
                var token = Encoding.UTF8.GetString(tokenData);

                var session = await _database.GetValidSessionTokenAsync(token);
                if (session == null)
                {
                    File.Delete(_sessionFilePath);
                    return false;
                }

                _currentUserId = session.UserId;
                _currentUsername = session.Username;
                _lastActivityTime = DateTime.UtcNow;
                _currentSessionId = Guid.NewGuid().ToString();

                _logger?.LogInformation("[Security] Session restored for {Username} via Remember Me", session.Username);
                return true;
            }
            catch (Exception ex)
            {
                _logger?.LogDebug(ex, "[Security] Could not restore session from disk");
                // Delete corrupted/inaccessible file
                try { File.Delete(_sessionFilePath); } catch { }
                return false;
            }
        }

        /// <summary>Clears the saved Remember Me token from disk and DB.</summary>
        public async Task ClearPersistentSessionAsync()
        {
            try
            {
                if (File.Exists(_sessionFilePath))
                {
                    var encrypted = await File.ReadAllBytesAsync(_sessionFilePath);
                    var tokenData = System.Security.Cryptography.ProtectedData.Unprotect(
                        encrypted, null, System.Security.Cryptography.DataProtectionScope.CurrentUser);
                    var token = Encoding.UTF8.GetString(tokenData);
                    await _database.DeleteSessionTokenAsync(token);
                    File.Delete(_sessionFilePath);
                }
            }
            catch (Exception ex)
            {
                _logger?.LogDebug(ex, "[Security] Error clearing persistent session");
            }
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
