using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Microsoft.Extensions.Logging;

namespace AICompanion.Desktop.Services.Security
{
    /// <summary>
    /// Manages sensitive API keys using Windows DPAPI encryption.
    ///
    /// Security design:
    ///   - Keys are encrypted with ProtectedData (DPAPI) scoped to DataProtectionScope.CurrentUser.
    ///   - Encrypted blob is stored in %LOCALAPPDATA%\AICompanion\keys.dat.
    ///   - keys.dat is excluded from source control via .gitignore.
    ///   - No key material ever appears in appsettings.json or the process environment.
    ///
    /// Why DPAPI over AES with a hard-coded password?
    ///   - DPAPI uses the logged-in Windows user's credential as the implicit key — no
    ///     password management burden, no risk of a leaked master password in source.
    ///   - The encrypted blob is useless on any other machine or user account.
    ///   - .NET ProtectedData.Protect is FIPS-compliant on all supported Windows versions.
    /// </summary>
    public class SecureApiKeyManager
    {
        private readonly ILogger<SecureApiKeyManager>? _logger;
        private readonly string _keysFilePath;

        public const string ElevenLabsKeyName = "ElevenLabs_ApiKey";

        public SecureApiKeyManager(ILogger<SecureApiKeyManager>? logger = null)
        {
            _logger = logger;
            var appDataPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "AICompanion");
            Directory.CreateDirectory(appDataPath);
            _keysFilePath = Path.Combine(appDataPath, "keys.dat");
        }

        /// <summary>
        /// Persists <paramref name="apiKey"/> under <paramref name="keyName"/> using DPAPI.
        /// Merges with any existing keys so other entries are not overwritten.
        /// </summary>
        public void SaveApiKey(string keyName, string apiKey)
        {
            if (string.IsNullOrWhiteSpace(keyName)) throw new ArgumentNullException(nameof(keyName));
            if (string.IsNullOrWhiteSpace(apiKey))  throw new ArgumentNullException(nameof(apiKey));

            try
            {
                var all = LoadAllKeys();
                all[keyName] = apiKey;
                PersistKeys(all);
                _logger?.LogInformation("[SecureKeys] Saved key '{Name}'", keyName);
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "[SecureKeys] Failed to save key '{Name}'", keyName);
                throw;
            }
        }

        /// <summary>
        /// Decrypts and returns the stored value for <paramref name="keyName"/>,
        /// or <c>null</c> if the key has never been saved.
        /// </summary>
        public string? LoadApiKey(string keyName)
        {
            try
            {
                var all = LoadAllKeys();
                return all.TryGetValue(keyName, out var val) ? val : null;
            }
            catch (Exception ex)
            {
                _logger?.LogWarning(ex, "[SecureKeys] Failed to load key '{Name}'", keyName);
                return null;
            }
        }

        /// <summary>Returns true if keys.dat exists and contains <paramref name="keyName"/>.</summary>
        public bool HasKey(string keyName)
        {
            try { return LoadAllKeys().ContainsKey(keyName); }
            catch { return false; }
        }

        /// <summary>Removes a single key from the encrypted store.</summary>
        public void DeleteApiKey(string keyName)
        {
            try
            {
                var all = LoadAllKeys();
                if (all.Remove(keyName))
                {
                    PersistKeys(all);
                    _logger?.LogInformation("[SecureKeys] Deleted key '{Name}'", keyName);
                }
            }
            catch (Exception ex)
            {
                _logger?.LogWarning(ex, "[SecureKeys] Failed to delete key '{Name}'", keyName);
            }
        }

        // ── private helpers ──────────────────────────────────────────────────────

        private Dictionary<string, string> LoadAllKeys()
        {
            if (!File.Exists(_keysFilePath))
                return new Dictionary<string, string>(StringComparer.Ordinal);

            var encrypted = File.ReadAllBytes(_keysFilePath);
            var plainBytes = ProtectedData.Unprotect(encrypted, null, DataProtectionScope.CurrentUser);
            var json = Encoding.UTF8.GetString(plainBytes);
            return JsonSerializer.Deserialize<Dictionary<string, string>>(json)
                   ?? new Dictionary<string, string>(StringComparer.Ordinal);
        }

        private void PersistKeys(Dictionary<string, string> keys)
        {
            var json = JsonSerializer.Serialize(keys);
            var plainBytes = Encoding.UTF8.GetBytes(json);
            var encrypted = ProtectedData.Protect(plainBytes, null, DataProtectionScope.CurrentUser);
            File.WriteAllBytes(_keysFilePath, encrypted);
        }
    }
}
