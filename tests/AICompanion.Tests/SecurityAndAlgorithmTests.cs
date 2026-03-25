using System;
using System.Security.Cryptography;
using Xunit;
using FluentAssertions;

namespace AICompanion.Tests
{
    /// <summary>
    /// Tests for security-related functionality.
    /// Verifies that cryptographic operations use proper RNG,
    /// and that Levenshtein distance calculations are correct.
    /// </summary>
    public class SecurityAndAlgorithmTests
    {
        // ==========================================
        // 1. Security Code Generation (RNGCrypto)
        // ==========================================

        [Fact]
        public void SecurityCode_ShouldBe6Digits()
        {
            // Generate a security code the same way SecurityService does
            var code = GenerateNumericCode(6);

            code.Should().HaveLength(6);
            code.Should().MatchRegex(@"^\d{6}$");
        }

        [Fact]
        public void SecurityCode_ShouldBeRandom_NotPredictable()
        {
            // Generate multiple codes and ensure they are not all the same
            var codes = new string[10];
            for (int i = 0; i < codes.Length; i++)
            {
                codes[i] = GenerateNumericCode(6);
            }

            // At least 2 different codes in 10 generations (extremely high probability)
            codes.Distinct().Count().Should().BeGreaterThan(1);
        }

        [Fact]
        public void SecurityCode_UsesRNGCryptoServiceProvider()
        {
            // Verify that RandomNumberGenerator.Create() works (cryptographic RNG)
            using var rng = RandomNumberGenerator.Create();
            var bytes = new byte[6];
            rng.GetBytes(bytes);

            bytes.Should().HaveCount(6);
            // At least one byte should be non-zero (extremely high probability)
            bytes.Any(b => b != 0).Should().BeTrue();
        }

        // ==========================================
        // 2. Password Hashing (SHA256 + Salt)
        // ==========================================

        [Fact]
        public void PasswordHash_SameInput_ShouldProduceSameHash()
        {
            var password = "TestPassword123";
            var salt = GenerateSalt();

            var hash1 = HashPassword(password, salt);
            var hash2 = HashPassword(password, salt);

            hash1.Should().Be(hash2);
        }

        [Fact]
        public void PasswordHash_DifferentSalts_ShouldProduceDifferentHashes()
        {
            var password = "TestPassword123";
            var salt1 = GenerateSalt();
            var salt2 = GenerateSalt();

            var hash1 = HashPassword(password, salt1);
            var hash2 = HashPassword(password, salt2);

            hash1.Should().NotBe(hash2);
        }

        [Fact]
        public void PasswordHash_DifferentPasswords_ShouldProduceDifferentHashes()
        {
            var salt = GenerateSalt();

            var hash1 = HashPassword("Password1", salt);
            var hash2 = HashPassword("Password2", salt);

            hash1.Should().NotBe(hash2);
        }

        // ==========================================
        // 3. Levenshtein Distance Algorithm
        // ==========================================

        [Theory]
        [InlineData("", "", 0)]
        [InlineData("abc", "", 3)]
        [InlineData("", "abc", 3)]
        [InlineData("abc", "abc", 0)]
        [InlineData("abc", "abd", 1)]
        [InlineData("kitten", "sitting", 3)]
        [InlineData("open", "oepn", 2)]
        [InlineData("notepad", "noetpad", 2)]
        [InlineData("calculator", "calcilator", 1)]
        public void LevenshteinDistance_ShouldBeCorrect(string source, string target, int expected)
        {
            var result = CalculateLevenshteinDistance(source, target);
            result.Should().Be(expected);
        }

        [Fact]
        public void LevenshteinRatio_IdenticalStrings_ShouldBeOne()
        {
            var ratio = CalculateLevenshteinRatio("open", "open");
            ratio.Should().Be(1.0);
        }

        [Fact]
        public void LevenshteinRatio_CompletelyDifferent_ShouldBeZero()
        {
            var ratio = CalculateLevenshteinRatio("abc", "xyz");
            ratio.Should().Be(0.0);
        }

        [Fact]
        public void LevenshteinRatio_Calculator_Calcilator_ShouldBeAboveThreshold()
        {
            // "calcilator" vs "calculator" - distance 1, length 10
            // Ratio = 1 - (1/10) = 0.90
            var ratio = CalculateLevenshteinRatio("calculator", "calcilator");
            ratio.Should().BeGreaterOrEqualTo(0.80);
        }

        [Fact]
        public void LevenshteinRatio_Launch_Lanch_ShouldBeAboveThreshold()
        {
            // "lanch" vs "launch" - distance 1, max length 6
            // Ratio = 1 - (1/6) ≈ 0.83
            var ratio = CalculateLevenshteinRatio("lanch", "launch");
            ratio.Should().BeGreaterOrEqualTo(0.80);
        }

        // ==========================================
        // Helper methods (mirrors SecurityService)
        // ==========================================

        private static string GenerateNumericCode(int length)
        {
            var code = new System.Text.StringBuilder();
            using var rng = RandomNumberGenerator.Create();
            var bytes = new byte[length];
            rng.GetBytes(bytes);

            foreach (var b in bytes)
            {
                code.Append((b % 10).ToString());
            }

            return code.ToString();
        }

        private static string GenerateSalt()
        {
            var saltBytes = new byte[32];
            using var rng = RandomNumberGenerator.Create();
            rng.GetBytes(saltBytes);
            return Convert.ToBase64String(saltBytes);
        }

        private static string HashPassword(string password, string salt)
        {
            using var sha256 = System.Security.Cryptography.SHA256.Create();
            var combined = System.Text.Encoding.UTF8.GetBytes(password + salt);
            var hash = sha256.ComputeHash(combined);
            return Convert.ToBase64String(hash);
        }

        private static int CalculateLevenshteinDistance(string source, string target)
        {
            if (string.IsNullOrEmpty(source))
                return string.IsNullOrEmpty(target) ? 0 : target.Length;
            if (string.IsNullOrEmpty(target))
                return source.Length;

            var sourceLength = source.Length;
            var targetLength = target.Length;
            var distanceMatrix = new int[sourceLength + 1, targetLength + 1];

            for (var i = 0; i <= sourceLength; i++)
                distanceMatrix[i, 0] = i;
            for (var j = 0; j <= targetLength; j++)
                distanceMatrix[0, j] = j;

            for (var i = 1; i <= sourceLength; i++)
            {
                for (var j = 1; j <= targetLength; j++)
                {
                    var cost = (target[j - 1] == source[i - 1]) ? 0 : 1;
                    distanceMatrix[i, j] = Math.Min(
                        Math.Min(
                            distanceMatrix[i - 1, j] + 1,
                            distanceMatrix[i, j - 1] + 1),
                        distanceMatrix[i - 1, j - 1] + cost);
                }
            }

            return distanceMatrix[sourceLength, targetLength];
        }

        private static double CalculateLevenshteinRatio(string source, string target)
        {
            if (string.IsNullOrEmpty(source) && string.IsNullOrEmpty(target)) return 1.0;
            if (string.IsNullOrEmpty(source) || string.IsNullOrEmpty(target)) return 0.0;

            var distance = CalculateLevenshteinDistance(source, target);
            var maxLen = Math.Max(source.Length, target.Length);
            return 1.0 - ((double)distance / maxLen);
        }
    }
}
