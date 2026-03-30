using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace AICompanion.Desktop.Services.Database
{
    /// <summary>
    /// AES-256-GCM field-level encryption for sensitive database columns.
    ///
    /// Why AES-GCM over AES-CBC?
    ///   - GCM is an authenticated encryption mode: it produces a 16-byte authentication
    ///     tag that detects both accidental corruption AND deliberate tampering.
    ///   - CBC without MAC (HMAC-then-Encrypt or Encrypt-then-MAC) is vulnerable to
    ///     padding-oracle attacks; GCM eliminates that attack surface entirely.
    ///   - .NET AesGcm (System.Security.Cryptography) is available in .NET 5+ and
    ///     uses hardware AES-NI acceleration on modern CPUs.
    ///
    /// Wire format (Base64 of): [ 12-byte nonce | 16-byte tag | ciphertext ]
    ///   - Nonce is freshly generated per encryption (CSPRNG) — safe to reuse the same key.
    ///   - Key is 32 bytes (256-bit), derived and stored via DPAPI (see DatabaseService).
    /// </summary>
    public static class FieldEncryptionHelper
    {
        private const int NonceSizeBytes = 12;  // GCM standard nonce
        private const int TagSizeBytes   = 16;  // GCM authentication tag (128-bit)

        /// <summary>
        /// Encrypts <paramref name="plaintext"/> and returns a Base64 string that encodes
        /// the nonce, authentication tag, and ciphertext concatenated together.
        /// </summary>
        public static string Encrypt(string plaintext, byte[] key)
        {
            if (key == null || key.Length != 32)
                throw new ArgumentException("Key must be exactly 32 bytes (AES-256).", nameof(key));

            var plaintextBytes = Encoding.UTF8.GetBytes(plaintext);
            var nonce          = new byte[NonceSizeBytes];
            var tag            = new byte[TagSizeBytes];
            var ciphertext     = new byte[plaintextBytes.Length];

            RandomNumberGenerator.Fill(nonce);   // fresh CSPRNG nonce each call

            using var aes = new AesGcm(key, TagSizeBytes);
            aes.Encrypt(nonce, plaintextBytes, ciphertext, tag);

            // Layout: nonce || tag || ciphertext → Base64
            var combined = new byte[NonceSizeBytes + TagSizeBytes + ciphertext.Length];
            Buffer.BlockCopy(nonce,      0, combined, 0,                         NonceSizeBytes);
            Buffer.BlockCopy(tag,        0, combined, NonceSizeBytes,            TagSizeBytes);
            Buffer.BlockCopy(ciphertext, 0, combined, NonceSizeBytes + TagSizeBytes, ciphertext.Length);

            return Convert.ToBase64String(combined);
        }

        /// <summary>
        /// Decrypts a value produced by <see cref="Encrypt"/>.
        /// Throws <see cref="CryptographicException"/> if the tag is invalid (tampered data).
        /// </summary>
        public static string Decrypt(string encryptedBase64, byte[] key)
        {
            if (key == null || key.Length != 32)
                throw new ArgumentException("Key must be exactly 32 bytes (AES-256).", nameof(key));

            var combined       = Convert.FromBase64String(encryptedBase64);
            var nonce          = combined[..NonceSizeBytes];
            var tag            = combined[NonceSizeBytes..(NonceSizeBytes + TagSizeBytes)];
            var ciphertext     = combined[(NonceSizeBytes + TagSizeBytes)..];
            var plaintext      = new byte[ciphertext.Length];

            using var aes = new AesGcm(key, TagSizeBytes);
            aes.Decrypt(nonce, ciphertext, tag, plaintext);

            return Encoding.UTF8.GetString(plaintext);
        }

        /// <summary>
        /// Returns true if <paramref name="value"/> looks like a value produced by
        /// <see cref="Encrypt"/> (valid Base64, minimum length for nonce+tag).
        /// Used to handle databases that contain a mix of old plaintext and new encrypted rows.
        /// </summary>
        public static bool IsEncrypted(string? value)
        {
            if (string.IsNullOrEmpty(value)) return false;
            try
            {
                var bytes = Convert.FromBase64String(value);
                return bytes.Length > NonceSizeBytes + TagSizeBytes;
            }
            catch
            {
                return false;
            }
        }
    }
}
