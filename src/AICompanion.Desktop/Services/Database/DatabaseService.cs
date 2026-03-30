using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.Data.Sqlite;
using Microsoft.Extensions.Logging;

namespace AICompanion.Desktop.Services.Database
{
    /// <summary>
    /// SQLite database service for storing context, history, settings, and security data.
    /// Implements ChromaDB-like vector storage for conversation context.
    /// </summary>
    public class DatabaseService : IDisposable
    {
        private readonly ILogger<DatabaseService>? _logger;
        private readonly string _dbPath;
        private SqliteConnection? _connection;
        private bool _isInitialized;

        // AES-256 field encryption key (32 bytes), loaded via DPAPI on first InitializeAsync.
        // Null when running in-memory test mode (encryption disabled for test isolation).
        private byte[]? _fieldEncryptionKey;

        // Path to the DPAPI-encrypted field key file (sibling to the database file)
        private readonly string? _fieldKeyPath;

        public bool IsInitialized => _isInitialized;

        public DatabaseService(ILogger<DatabaseService>? logger = null)
        {
            _logger = logger;
            var appDataPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "AICompanion", "Data");
            Directory.CreateDirectory(appDataPath);
            _dbPath       = Path.Combine(appDataPath, "aicompanion.db");
            _fieldKeyPath = Path.Combine(appDataPath, "fieldkey.dat");
        }

        /// <summary>
        /// Test-only constructor: pass ":memory:" for an isolated SQLite in-memory database.
        /// Encryption is disabled in test mode to keep tests deterministic.
        /// </summary>
        internal DatabaseService(ILogger<DatabaseService>? logger, string dbPath)
        {
            _logger       = logger;
            _dbPath       = dbPath;
            _fieldKeyPath = null;  // no encryption in test mode
        }

        public async Task InitializeAsync()
        {
            try
            {
                _connection = new SqliteConnection($"Data Source={_dbPath}");
                await _connection.OpenAsync();

                await CreateTablesAsync();

                // Load or generate the AES-256 field encryption key (production only)
                if (_fieldKeyPath != null)
                    _fieldEncryptionKey = LoadOrCreateFieldKey(_fieldKeyPath);

                _isInitialized = true;
                _logger?.LogInformation("[Database] Initialized at {Path}", _dbPath);
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "[Database] Failed to initialize");
                throw;
            }
        }

        /// <summary>
        /// Generates a fresh 32-byte AES-256 key on first run and persists it in
        /// <paramref name="keyPath"/> protected by DPAPI (DataProtectionScope.CurrentUser).
        /// On subsequent runs, decrypts and returns the stored key.
        /// </summary>
        private byte[] LoadOrCreateFieldKey(string keyPath)
        {
            if (File.Exists(keyPath))
            {
                try
                {
                    var encrypted = File.ReadAllBytes(keyPath);
                    return ProtectedData.Unprotect(encrypted, null, DataProtectionScope.CurrentUser);
                }
                catch (Exception ex)
                {
                    _logger?.LogWarning(ex, "[Database] Could not decrypt field key — generating a new one (existing data will be unreadable)");
                }
            }

            // Generate a cryptographically random 256-bit key
            var key = new byte[32];
            RandomNumberGenerator.Fill(key);
            var blob = ProtectedData.Protect(key, null, DataProtectionScope.CurrentUser);
            File.WriteAllBytes(keyPath, blob);
            _logger?.LogInformation("[Database] Generated new AES-256 field encryption key");
            return key;
        }

        private async Task CreateTablesAsync()
        {
            var commands = new[]
            {
                // Users table for authentication
                @"CREATE TABLE IF NOT EXISTS Users (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Username TEXT UNIQUE NOT NULL,
                    Email TEXT,
                    PasswordHash TEXT NOT NULL,
                    Salt TEXT NOT NULL,
                    HashAlgorithm TEXT NOT NULL DEFAULT 'SHA256',
                    CreatedAt TEXT NOT NULL,
                    LastLoginAt TEXT,
                    IsActive INTEGER DEFAULT 1
                )",

                // Migration: add Email column to existing DB (safe IF NOT EXISTS workaround)
                @"ALTER TABLE Users ADD COLUMN Email TEXT",
                // Migration: add HashAlgorithm column (for PBKDF2 upgrade tracking)
                @"ALTER TABLE Users ADD COLUMN HashAlgorithm TEXT NOT NULL DEFAULT 'SHA256'",

                // Persistent session tokens for Remember Me
                @"CREATE TABLE IF NOT EXISTS SessionTokens (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    UserId INTEGER NOT NULL,
                    Token TEXT UNIQUE NOT NULL,
                    CreatedAt TEXT NOT NULL,
                    ExpiresAt TEXT NOT NULL,
                    FOREIGN KEY (UserId) REFERENCES Users(Id)
                )",

                // Security codes for dangerous operations
                @"CREATE TABLE IF NOT EXISTS SecurityCodes (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    UserId INTEGER NOT NULL,
                    Code TEXT NOT NULL,
                    Purpose TEXT NOT NULL,
                    CreatedAt TEXT NOT NULL,
                    ExpiresAt TEXT NOT NULL,
                    UsedAt TEXT,
                    IsUsed INTEGER DEFAULT 0,
                    FOREIGN KEY (UserId) REFERENCES Users(Id)
                )",

                // Conversation history
                @"CREATE TABLE IF NOT EXISTS ConversationHistory (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    SessionId TEXT NOT NULL,
                    UserId INTEGER,
                    Command TEXT NOT NULL,
                    Response TEXT,
                    ActionType TEXT,
                    ActionResult TEXT,
                    Confidence REAL,
                    CreatedAt TEXT NOT NULL,
                    ContextVector TEXT,
                    FOREIGN KEY (UserId) REFERENCES Users(Id)
                )",

                // Context memory (ChromaDB-like)
                @"CREATE TABLE IF NOT EXISTS ContextMemory (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    UserId INTEGER,
                    Key TEXT NOT NULL,
                    Value TEXT NOT NULL,
                    Category TEXT,
                    Importance REAL DEFAULT 0.5,
                    CreatedAt TEXT NOT NULL,
                    LastAccessedAt TEXT,
                    AccessCount INTEGER DEFAULT 0,
                    FOREIGN KEY (UserId) REFERENCES Users(Id)
                )",

                // User settings
                @"CREATE TABLE IF NOT EXISTS UserSettings (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    UserId INTEGER NOT NULL,
                    SettingKey TEXT NOT NULL,
                    SettingValue TEXT NOT NULL,
                    UpdatedAt TEXT NOT NULL,
                    UNIQUE(UserId, SettingKey),
                    FOREIGN KEY (UserId) REFERENCES Users(Id)
                )",

                // Command aliases (custom voice commands)
                @"CREATE TABLE IF NOT EXISTS CommandAliases (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    UserId INTEGER,
                    Alias TEXT NOT NULL,
                    ActualCommand TEXT NOT NULL,
                    Description TEXT,
                    CreatedAt TEXT NOT NULL,
                    FOREIGN KEY (UserId) REFERENCES Users(Id)
                )",

                // Dangerous operations log
                @"CREATE TABLE IF NOT EXISTS DangerousOperationsLog (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    UserId INTEGER,
                    Operation TEXT NOT NULL,
                    Parameters TEXT,
                    SecurityCodeId INTEGER,
                    Result TEXT,
                    ExecutedAt TEXT NOT NULL,
                    FOREIGN KEY (UserId) REFERENCES Users(Id),
                    FOREIGN KEY (SecurityCodeId) REFERENCES SecurityCodes(Id)
                )",

                // Create indexes for performance
                @"CREATE INDEX IF NOT EXISTS idx_conversation_session ON ConversationHistory(SessionId)",
                @"CREATE INDEX IF NOT EXISTS idx_conversation_user ON ConversationHistory(UserId)",
                @"CREATE INDEX IF NOT EXISTS idx_context_user ON ContextMemory(UserId)",
                @"CREATE INDEX IF NOT EXISTS idx_context_key ON ContextMemory(Key)",
                @"CREATE INDEX IF NOT EXISTS idx_security_codes_user ON SecurityCodes(UserId)"
            };

            foreach (var sql in commands)
            {
                try
                {
                    using var cmd = new SqliteCommand(sql, _connection);
                    await cmd.ExecuteNonQueryAsync();
                }
                catch (SqliteException ex) when (sql.TrimStart().StartsWith("ALTER TABLE"))
                {
                    // ALTER TABLE throws if the column already exists in an existing DB — safe to ignore
                    _logger?.LogDebug("[Database] Migration already applied (skipping): {Msg}", ex.Message);
                }
            }

            _logger?.LogInformation("[Database] Tables created/verified");
        }

        #region User Management

        public async Task<int> CreateUserAsync(string username, string passwordHash, string salt,
            string? email = null, string hashAlgorithm = "PBKDF2")
        {
            var sql = @"INSERT INTO Users (Username, Email, PasswordHash, Salt, HashAlgorithm, CreatedAt)
                        VALUES (@username, @email, @passwordHash, @salt, @hashAlgorithm, @createdAt);
                        SELECT last_insert_rowid();";

            using var cmd = new SqliteCommand(sql, _connection);
            cmd.Parameters.AddWithValue("@username", username);
            cmd.Parameters.AddWithValue("@email", email ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@passwordHash", passwordHash);
            cmd.Parameters.AddWithValue("@salt", salt);
            cmd.Parameters.AddWithValue("@hashAlgorithm", hashAlgorithm);
            cmd.Parameters.AddWithValue("@createdAt", DateTime.UtcNow.ToString("O"));

            var result = await cmd.ExecuteScalarAsync();
            return Convert.ToInt32(result);
        }

        public async Task<UserRecord?> GetUserAsync(string username)
        {
            var sql = @"SELECT Id, Username, Email, PasswordHash, Salt, HashAlgorithm, CreatedAt, LastLoginAt
                        FROM Users WHERE Username = @username AND IsActive = 1";
            using var cmd = new SqliteCommand(sql, _connection);
            cmd.Parameters.AddWithValue("@username", username);

            using var reader = await cmd.ExecuteReaderAsync();
            if (await reader.ReadAsync())
            {
                return new UserRecord
                {
                    Id = reader.GetInt32(0),
                    Username = reader.GetString(1),
                    Email = reader.IsDBNull(2) ? null : reader.GetString(2),
                    PasswordHash = reader.GetString(3),
                    Salt = reader.GetString(4),
                    HashAlgorithm = reader.IsDBNull(5) ? "SHA256" : reader.GetString(5),
                    CreatedAt = DateTime.Parse(reader.GetString(6)),
                    LastLoginAt = reader.IsDBNull(7) ? null : DateTime.Parse(reader.GetString(7))
                };
            }
            return null;
        }

        public async Task UpdatePasswordHashAsync(int userId, string newHash, string newSalt, string algorithm)
        {
            var sql = "UPDATE Users SET PasswordHash = @hash, Salt = @salt, HashAlgorithm = @algo WHERE Id = @userId";
            using var cmd = new SqliteCommand(sql, _connection);
            cmd.Parameters.AddWithValue("@hash", newHash);
            cmd.Parameters.AddWithValue("@salt", newSalt);
            cmd.Parameters.AddWithValue("@algo", algorithm);
            cmd.Parameters.AddWithValue("@userId", userId);
            await cmd.ExecuteNonQueryAsync();
        }

        public async Task UpdateLastLoginAsync(int userId)
        {
            var sql = "UPDATE Users SET LastLoginAt = @now WHERE Id = @userId";
            using var cmd = new SqliteCommand(sql, _connection);
            cmd.Parameters.AddWithValue("@now", DateTime.UtcNow.ToString("O"));
            cmd.Parameters.AddWithValue("@userId", userId);
            await cmd.ExecuteNonQueryAsync();
        }

        #endregion

        #region Security Codes

        public async Task<int> CreateSecurityCodeAsync(int userId, string code, string purpose, TimeSpan validity)
        {
            var now = DateTime.UtcNow;
            var sql = @"INSERT INTO SecurityCodes (UserId, Code, Purpose, CreatedAt, ExpiresAt) 
                        VALUES (@userId, @code, @purpose, @createdAt, @expiresAt);
                        SELECT last_insert_rowid();";

            using var cmd = new SqliteCommand(sql, _connection);
            cmd.Parameters.AddWithValue("@userId", userId);
            cmd.Parameters.AddWithValue("@code", code);
            cmd.Parameters.AddWithValue("@purpose", purpose);
            cmd.Parameters.AddWithValue("@createdAt", now.ToString("O"));
            cmd.Parameters.AddWithValue("@expiresAt", now.Add(validity).ToString("O"));

            var result = await cmd.ExecuteScalarAsync();
            return Convert.ToInt32(result);
        }

        public async Task<SecurityCodeRecord?> ValidateSecurityCodeAsync(int userId, string code, string purpose)
        {
            var sql = @"SELECT * FROM SecurityCodes 
                        WHERE UserId = @userId AND Code = @code AND Purpose = @purpose 
                        AND IsUsed = 0 AND ExpiresAt > @now";

            using var cmd = new SqliteCommand(sql, _connection);
            cmd.Parameters.AddWithValue("@userId", userId);
            cmd.Parameters.AddWithValue("@code", code);
            cmd.Parameters.AddWithValue("@purpose", purpose);
            cmd.Parameters.AddWithValue("@now", DateTime.UtcNow.ToString("O"));

            using var reader = await cmd.ExecuteReaderAsync();
            if (await reader.ReadAsync())
            {
                return new SecurityCodeRecord
                {
                    Id = reader.GetInt32(0),
                    UserId = reader.GetInt32(1),
                    Code = reader.GetString(2),
                    Purpose = reader.GetString(3),
                    CreatedAt = DateTime.Parse(reader.GetString(4)),
                    ExpiresAt = DateTime.Parse(reader.GetString(5))
                };
            }
            return null;
        }

        public async Task MarkSecurityCodeUsedAsync(int codeId)
        {
            var sql = "UPDATE SecurityCodes SET IsUsed = 1, UsedAt = @now WHERE Id = @id";
            using var cmd = new SqliteCommand(sql, _connection);
            cmd.Parameters.AddWithValue("@now", DateTime.UtcNow.ToString("O"));
            cmd.Parameters.AddWithValue("@id", codeId);
            await cmd.ExecuteNonQueryAsync();
        }

        #endregion

        #region Conversation History

        public async Task SaveConversationAsync(string sessionId, int? userId, string command,
            string? response, string? actionType, string? actionResult, float confidence)
        {
            // Encrypt sensitive fields (command text + response) at rest when the key is available.
            // appsettings.json EncryptSensitiveData flag is now honoured by this implementation.
            var storedCommand  = _fieldEncryptionKey != null
                ? FieldEncryptionHelper.Encrypt(command, _fieldEncryptionKey)
                : command;
            var storedResponse = (_fieldEncryptionKey != null && response != null)
                ? FieldEncryptionHelper.Encrypt(response, _fieldEncryptionKey)
                : response;

            var sql = @"INSERT INTO ConversationHistory
                        (SessionId, UserId, Command, Response, ActionType, ActionResult, Confidence, CreatedAt)
                        VALUES (@sessionId, @userId, @command, @response, @actionType, @actionResult, @confidence, @createdAt)";

            using var cmd = new SqliteCommand(sql, _connection);
            cmd.Parameters.AddWithValue("@sessionId", sessionId);
            cmd.Parameters.AddWithValue("@userId", userId.HasValue ? userId.Value : DBNull.Value);
            cmd.Parameters.AddWithValue("@command", storedCommand);
            cmd.Parameters.AddWithValue("@response", storedResponse ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@actionType", actionType ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@actionResult", actionResult ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@confidence", confidence);
            cmd.Parameters.AddWithValue("@createdAt", DateTime.UtcNow.ToString("O"));

            await cmd.ExecuteNonQueryAsync();
        }

        public async Task<List<ConversationRecord>> GetRecentConversationsAsync(string sessionId, int limit = 10)
        {
            var sql = @"SELECT * FROM ConversationHistory
                        WHERE SessionId = @sessionId
                        ORDER BY CreatedAt DESC LIMIT @limit";

            var records = new List<ConversationRecord>();
            using var cmd = new SqliteCommand(sql, _connection);
            cmd.Parameters.AddWithValue("@sessionId", sessionId);
            cmd.Parameters.AddWithValue("@limit", limit);

            using var reader = await cmd.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                var rawCommand  = reader.GetString(3);
                var rawResponse = reader.IsDBNull(4) ? null : reader.GetString(4);

                // Decrypt if the key is available and the stored value is an encrypted blob.
                // The IsEncrypted guard handles rows that were written before encryption was
                // enabled (plaintext rows are returned as-is without error).
                string command;
                string? response;
                try
                {
                    command = (_fieldEncryptionKey != null && FieldEncryptionHelper.IsEncrypted(rawCommand))
                        ? FieldEncryptionHelper.Decrypt(rawCommand, _fieldEncryptionKey)
                        : rawCommand;

                    response = (_fieldEncryptionKey != null && FieldEncryptionHelper.IsEncrypted(rawResponse))
                        ? FieldEncryptionHelper.Decrypt(rawResponse!, _fieldEncryptionKey)
                        : rawResponse;
                }
                catch (Exception ex)
                {
                    _logger?.LogWarning(ex, "[Database] Failed to decrypt conversation record {Id} — returning as-is", reader.GetInt32(0));
                    command  = rawCommand;
                    response = rawResponse;
                }

                records.Add(new ConversationRecord
                {
                    Id         = reader.GetInt32(0),
                    SessionId  = reader.GetString(1),
                    Command    = command,
                    Response   = response,
                    ActionType = reader.IsDBNull(5) ? null : reader.GetString(5),
                    CreatedAt  = DateTime.Parse(reader.GetString(8))
                });
            }
            return records;
        }

        #endregion

        #region Context Memory

        public async Task SaveContextAsync(int? userId, string key, string value, string? category = null, float importance = 0.5f)
        {
            var sql = @"INSERT OR REPLACE INTO ContextMemory 
                        (UserId, Key, Value, Category, Importance, CreatedAt, LastAccessedAt, AccessCount) 
                        VALUES (@userId, @key, @value, @category, @importance, @now, @now, 
                            COALESCE((SELECT AccessCount FROM ContextMemory WHERE Key = @key), 0) + 1)";

            using var cmd = new SqliteCommand(sql, _connection);
            cmd.Parameters.AddWithValue("@userId", userId.HasValue ? userId.Value : DBNull.Value);
            cmd.Parameters.AddWithValue("@key", key);
            cmd.Parameters.AddWithValue("@value", value);
            cmd.Parameters.AddWithValue("@category", category ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@importance", importance);
            cmd.Parameters.AddWithValue("@now", DateTime.UtcNow.ToString("O"));

            await cmd.ExecuteNonQueryAsync();
        }

        public async Task<string?> GetContextAsync(string key)
        {
            var sql = @"UPDATE ContextMemory SET LastAccessedAt = @now, AccessCount = AccessCount + 1 
                        WHERE Key = @key;
                        SELECT Value FROM ContextMemory WHERE Key = @key";

            using var cmd = new SqliteCommand(sql, _connection);
            cmd.Parameters.AddWithValue("@key", key);
            cmd.Parameters.AddWithValue("@now", DateTime.UtcNow.ToString("O"));

            var result = await cmd.ExecuteScalarAsync();
            return result?.ToString();
        }

        public async Task<Dictionary<string, string>> GetAllContextAsync(int? userId = null, string? category = null)
        {
            var sql = "SELECT Key, Value FROM ContextMemory WHERE 1=1";
            if (userId.HasValue) sql += " AND UserId = @userId";
            if (!string.IsNullOrEmpty(category)) sql += " AND Category = @category";
            sql += " ORDER BY Importance DESC, AccessCount DESC";

            var context = new Dictionary<string, string>();
            using var cmd = new SqliteCommand(sql, _connection);
            if (userId.HasValue) cmd.Parameters.AddWithValue("@userId", userId.Value);
            if (!string.IsNullOrEmpty(category)) cmd.Parameters.AddWithValue("@category", category);

            using var reader = await cmd.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                context[reader.GetString(0)] = reader.GetString(1);
            }
            return context;
        }

        #endregion

        #region User Settings

        public async Task SaveSettingAsync(int userId, string key, string value)
        {
            var sql = @"INSERT OR REPLACE INTO UserSettings (UserId, SettingKey, SettingValue, UpdatedAt) 
                        VALUES (@userId, @key, @value, @now)";

            using var cmd = new SqliteCommand(sql, _connection);
            cmd.Parameters.AddWithValue("@userId", userId);
            cmd.Parameters.AddWithValue("@key", key);
            cmd.Parameters.AddWithValue("@value", value);
            cmd.Parameters.AddWithValue("@now", DateTime.UtcNow.ToString("O"));

            await cmd.ExecuteNonQueryAsync();
        }

        public async Task<string?> GetSettingAsync(int userId, string key)
        {
            var sql = "SELECT SettingValue FROM UserSettings WHERE UserId = @userId AND SettingKey = @key";
            using var cmd = new SqliteCommand(sql, _connection);
            cmd.Parameters.AddWithValue("@userId", userId);
            cmd.Parameters.AddWithValue("@key", key);

            var result = await cmd.ExecuteScalarAsync();
            return result?.ToString();
        }

        public async Task<Dictionary<string, string>> GetAllSettingsAsync(int userId)
        {
            var sql = "SELECT SettingKey, SettingValue FROM UserSettings WHERE UserId = @userId";
            var settings = new Dictionary<string, string>();
            
            using var cmd = new SqliteCommand(sql, _connection);
            cmd.Parameters.AddWithValue("@userId", userId);

            using var reader = await cmd.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                settings[reader.GetString(0)] = reader.GetString(1);
            }
            return settings;
        }

        #endregion

        #region Dangerous Operations Log

        public async Task LogDangerousOperationAsync(int? userId, string operation, string? parameters, int? securityCodeId, string result)
        {
            var sql = @"INSERT INTO DangerousOperationsLog 
                        (UserId, Operation, Parameters, SecurityCodeId, Result, ExecutedAt) 
                        VALUES (@userId, @operation, @parameters, @securityCodeId, @result, @now)";

            using var cmd = new SqliteCommand(sql, _connection);
            cmd.Parameters.AddWithValue("@userId", userId.HasValue ? userId.Value : DBNull.Value);
            cmd.Parameters.AddWithValue("@operation", operation);
            cmd.Parameters.AddWithValue("@parameters", parameters ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@securityCodeId", securityCodeId.HasValue ? securityCodeId.Value : DBNull.Value);
            cmd.Parameters.AddWithValue("@result", result);
            cmd.Parameters.AddWithValue("@now", DateTime.UtcNow.ToString("O"));

            await cmd.ExecuteNonQueryAsync();
        }

        #endregion

        #region Session Tokens (Remember Me)

        public async Task SaveSessionTokenAsync(int userId, string token, DateTime expiresAt)
        {
            // Remove any existing tokens for this user first
            var deleteSql = "DELETE FROM SessionTokens WHERE UserId = @userId";
            using var deleteCmd = new SqliteCommand(deleteSql, _connection);
            deleteCmd.Parameters.AddWithValue("@userId", userId);
            await deleteCmd.ExecuteNonQueryAsync();

            var sql = @"INSERT INTO SessionTokens (UserId, Token, CreatedAt, ExpiresAt)
                        VALUES (@userId, @token, @createdAt, @expiresAt)";
            using var cmd = new SqliteCommand(sql, _connection);
            cmd.Parameters.AddWithValue("@userId", userId);
            cmd.Parameters.AddWithValue("@token", token);
            cmd.Parameters.AddWithValue("@createdAt", DateTime.UtcNow.ToString("O"));
            cmd.Parameters.AddWithValue("@expiresAt", expiresAt.ToString("O"));
            await cmd.ExecuteNonQueryAsync();
        }

        public async Task<SessionTokenRecord?> GetValidSessionTokenAsync(string token)
        {
            var sql = @"SELECT st.Id, st.UserId, u.Username, st.ExpiresAt
                        FROM SessionTokens st
                        JOIN Users u ON st.UserId = u.Id
                        WHERE st.Token = @token AND st.ExpiresAt > @now AND u.IsActive = 1";
            using var cmd = new SqliteCommand(sql, _connection);
            cmd.Parameters.AddWithValue("@token", token);
            cmd.Parameters.AddWithValue("@now", DateTime.UtcNow.ToString("O"));

            using var reader = await cmd.ExecuteReaderAsync();
            if (await reader.ReadAsync())
            {
                return new SessionTokenRecord
                {
                    Id = reader.GetInt32(0),
                    UserId = reader.GetInt32(1),
                    Username = reader.GetString(2),
                    ExpiresAt = DateTime.Parse(reader.GetString(3))
                };
            }
            return null;
        }

        public async Task DeleteSessionTokenAsync(string token)
        {
            var sql = "DELETE FROM SessionTokens WHERE Token = @token";
            using var cmd = new SqliteCommand(sql, _connection);
            cmd.Parameters.AddWithValue("@token", token);
            await cmd.ExecuteNonQueryAsync();
        }

        #endregion

        public void Dispose()
        {
            _connection?.Close();
            _connection?.Dispose();
            _logger?.LogInformation("[Database] Disposed");
        }
    }

    #region Records

    public class UserRecord
    {
        public int Id { get; set; }
        public string Username { get; set; } = "";
        public string? Email { get; set; }
        public string PasswordHash { get; set; } = "";
        public string Salt { get; set; } = "";
        public string HashAlgorithm { get; set; } = "SHA256";
        public DateTime CreatedAt { get; set; }
        public DateTime? LastLoginAt { get; set; }
    }

    public class SessionTokenRecord
    {
        public int Id { get; set; }
        public int UserId { get; set; }
        public string Username { get; set; } = "";
        public DateTime ExpiresAt { get; set; }
    }

    public class SecurityCodeRecord
    {
        public int Id { get; set; }
        public int UserId { get; set; }
        public string Code { get; set; } = "";
        public string Purpose { get; set; } = "";
        public DateTime CreatedAt { get; set; }
        public DateTime ExpiresAt { get; set; }
    }

    public class ConversationRecord
    {
        public int Id { get; set; }
        public string SessionId { get; set; } = "";
        public string Command { get; set; } = "";
        public string? Response { get; set; }
        public string? ActionType { get; set; }
        public DateTime CreatedAt { get; set; }
    }

    #endregion
}
