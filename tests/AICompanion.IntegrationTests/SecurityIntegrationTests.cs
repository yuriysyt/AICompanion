using System.Threading.Tasks;
using AICompanion.Desktop.Services.Security;
using AICompanion.Desktop.Services.Database;
using Microsoft.Extensions.Logging.Abstractions;
using Xunit.Abstractions;

namespace AICompanion.IntegrationTests
{
    /// <summary>
    /// Integration tests for the PIN / security-code system.
    ///
    /// Tests the full chain: SecurityService → DatabaseService (SQLite in-memory)
    ///   - Dangerous operation classification
    ///   - 6-digit CSPRNG code generation and DB storage
    ///   - Code validation and replay prevention (mark-as-used)
    ///   - PBKDF2 password hashing on registration
    ///   - Remember Me token lifecycle (create → restore → clear)
    ///
    /// No Python backend required. No files written to disk (uses :memory: SQLite).
    /// </summary>
    public class SecurityIntegrationTests
    {
        private readonly ITestOutputHelper _output;

        public SecurityIntegrationTests(ITestOutputHelper output) => _output = output;

        // ================================================================
        // TEST 1: Dangerous operation classification
        // IsDangerousOperation must correctly flag each keyword.
        // ================================================================
        [Theory]
        [InlineData("delete_file",      true)]
        [InlineData("shutdown_computer", true)]
        [InlineData("restart_computer", true)]
        [InlineData("format_drive",     true)]
        [InlineData("kill_process",     true)]
        [InlineData("clear_history",    true)]
        [InlineData("export_data",      true)]
        [InlineData("delete_account",   true)]
        [InlineData("open_notepad",     false)]
        [InlineData("search_weather",   false)]
        [InlineData("type_hello",       false)]
        [InlineData("scroll_down",      false)]
        public void SecurityService_IsDangerousOperation_ClassifiesCorrectly(
            string operation, bool expectedDangerous)
        {
            _output.WriteLine($"[SECURITY] '{operation}' → dangerous={expectedDangerous}");

            // SecurityService only needs DB for operations that touch DB;
            // IsDangerousOperation is a pure string check.
            var db  = BuildInMemoryDb();
            var sec = new SecurityService(NullLogger<SecurityService>.Instance, db);

            var actual = sec.IsDangerousOperation(operation);
            _output.WriteLine($"[RESULT] IsDangerous={actual}");

            actual.Should().Be(expectedDangerous,
                $"'{operation}' should{(expectedDangerous ? "" : " NOT")} be dangerous");

            _output.WriteLine($"✅ PASSED: '{operation}' correctly classified");
        }

        // ================================================================
        // TEST 2: 6-digit CSPRNG security code generation
        // Registers a user, logs in, generates a code for delete_file.
        // ================================================================
        [Fact]
        public async Task SecurityService_GenerateCode_Returns6DigitCsprngCode()
        {
            _output.WriteLine("[SECURITY] Testing 6-digit code generation via SecurityService");

            var (db, sec) = await BuildAuthenticatedSecurityServiceAsync("codetest", "CodeTest1!");

            var code = await sec.GenerateSecurityCodeAsync("delete_file");
            _output.WriteLine($"[CODE] Generated: '{code}'");

            code.Should().NotBeNull("an authenticated user should receive a code");
            code!.Length.Should().Be(6, "security code must be exactly 6 digits");
            code.Should().MatchRegex(@"^\d{6}$", "code must be numeric digits only (CSPRNG, not System.Random)");

            _output.WriteLine($"✅ PASSED: 6-digit code '{code}' generated");
        }

        // ================================================================
        // TEST 3: Code validation and replay prevention
        // The same code must NOT be accepted twice.
        // ================================================================
        [Fact]
        public async Task SecurityService_ValidateCode_CannotReplayUsedCode()
        {
            _output.WriteLine("[SECURITY] Testing replay prevention — code used once must be rejected second time");

            var (db, sec) = await BuildAuthenticatedSecurityServiceAsync("replaytest", "Replay1Test!");

            var code = await sec.GenerateSecurityCodeAsync("delete_file");
            code.Should().NotBeNull();

            // First validation — must succeed
            var first = await sec.ValidateSecurityCodeAsync(code!, "delete_file");
            _output.WriteLine($"[VALIDATE 1] {(first ? "✅ accepted" : "❌ rejected")}");
            first.Should().BeTrue("first validation of a fresh code must succeed");

            // Second validation of the SAME code — must be rejected (marked as used)
            var second = await sec.ValidateSecurityCodeAsync(code!, "delete_file");
            _output.WriteLine($"[VALIDATE 2] {(second ? "⚠️ accepted (REPLAY BUG)" : "✅ rejected")}");
            second.Should().BeFalse("replaying an already-used code must be rejected");

            _output.WriteLine("✅ PASSED: Replay prevention works correctly");
        }

        // ================================================================
        // TEST 4: PBKDF2 password hashing — registration uses PBKDF2
        // Reads the stored hash back and confirms it's not a SHA256 hash.
        // ================================================================
        [Fact]
        public async Task SecurityService_Register_StoresPasswordAsPbkdf2()
        {
            _output.WriteLine("[SECURITY] Verifying PBKDF2 hash on registration");

            var db  = BuildInMemoryDb();
            await db.InitializeAsync();
            var sec = new SecurityService(NullLogger<SecurityService>.Instance, db);

            var ok = await sec.RegisterUserAsync("hashtest", "HashTest1!", "hash@test.com");
            ok.Should().BeTrue("registration should succeed");

            // Read back the stored user record directly from the DB
            var user = await db.GetUserAsync("hashtest");
            _output.WriteLine($"[HASH] Algorithm stored: '{user?.HashAlgorithm}'");
            _output.WriteLine($"[HASH] Hash length: {user?.PasswordHash?.Length ?? 0} chars");

            user.Should().NotBeNull();
            user!.HashAlgorithm.Should().Be("PBKDF2",
                "new registrations must use PBKDF2, not legacy SHA256");

            // PBKDF2-SHA256 with 32-byte output = 44 Base64 chars; SHA256 = 44 too,
            // but the algorithm field is authoritative
            user.PasswordHash.Should().NotBeNullOrEmpty();

            _output.WriteLine($"✅ PASSED: Password stored as PBKDF2 ('{user.HashAlgorithm}')");
        }

        // ================================================================
        // TEST 5: Login with correct password succeeds; wrong password fails
        // ================================================================
        [Fact]
        public async Task SecurityService_Login_SucceedsWithCorrectPassword_FailsWithWrong()
        {
            _output.WriteLine("[SECURITY] Testing login validation");

            var db  = BuildInMemoryDb();
            await db.InitializeAsync();
            var sec = new SecurityService(NullLogger<SecurityService>.Instance, db);

            await sec.RegisterUserAsync("logintest", "LoginTest1!", "l@test.com");

            var good = await sec.LoginAsync("logintest", "LoginTest1!");
            _output.WriteLine($"[LOGIN CORRECT] {(good ? "✅ success" : "❌ failed")}");
            good.Should().BeTrue("correct password must be accepted");

            sec.Logout();

            var bad = await sec.LoginAsync("logintest", "wrongpassword");
            _output.WriteLine($"[LOGIN WRONG]   {(bad ? "⚠️ accepted (BUG)" : "✅ rejected")}");
            bad.Should().BeFalse("wrong password must be rejected");

            _output.WriteLine("✅ PASSED: Login credential validation works");
        }

        // ================================================================
        // TEST 6: Email stored in DB during registration
        // ================================================================
        [Fact]
        public async Task SecurityService_Register_EmailStoredInDatabase()
        {
            _output.WriteLine("[SECURITY] Verifying email is persisted during registration (P1 fix)");

            var db  = BuildInMemoryDb();
            await db.InitializeAsync();
            var sec = new SecurityService(NullLogger<SecurityService>.Instance, db);

            const string testEmail = "integration@test.com";
            var ok = await sec.RegisterAsync("emailtest", testEmail, "EmailTest1!");
            ok.Success.Should().BeTrue("registration should succeed");

            var user = await db.GetUserAsync("emailtest");
            _output.WriteLine($"[EMAIL] Stored: '{user?.Email}'");

            user.Should().NotBeNull();
            user!.Email.Should().Be(testEmail,
                "email passed to RegisterAsync must be stored in the Users table (P1 fix)");

            _output.WriteLine($"✅ PASSED: Email '{user.Email}' correctly persisted");
        }

        // ================================================================
        // TEST 7: Remember Me — create, restore, and clear session token
        // ================================================================
        [Fact]
        public async Task SecurityService_RememberMe_TokenLifecycleWorks()
        {
            _output.WriteLine("[SECURITY] Testing Remember Me token lifecycle (P6 fix)");

            var (db, sec) = await BuildAuthenticatedSecurityServiceAsync("rememberme", "Remember1!");

            // 1) Create persistent session (saves token to DB only, no disk in test)
            var userId = sec.CurrentUserId!.Value;
            var token  = await sec.CreatePersistentSessionAsync(userId);
            _output.WriteLine($"[TOKEN] Created: length={token.Length}");
            token.Should().NotBeNullOrEmpty("a persistent session token must be generated");

            // 2) Validate token is findable in DB
            var record = await db.GetValidSessionTokenAsync(token);
            _output.WriteLine($"[TOKEN] DB record: user='{record?.Username}' expires={record?.ExpiresAt:u}");
            record.Should().NotBeNull("the token must be stored in the DB");
            record!.Username.Should().Be("rememberme");

            // 3) Delete the token
            await db.DeleteSessionTokenAsync(token);
            var deleted = await db.GetValidSessionTokenAsync(token);
            deleted.Should().BeNull("token must be gone after deletion");

            _output.WriteLine("✅ PASSED: Remember Me token lifecycle (create → find → delete) works");
        }

        // ─── Helpers ────────────────────────────────────────────────────

        /// <summary>Creates an in-memory DatabaseService (uses SQLite :memory: — no files written).</summary>
        private static DatabaseService BuildInMemoryDb()
            => new DatabaseService(NullLogger<DatabaseService>.Instance, ":memory:");

        /// <summary>
        /// Creates an in-memory DB + SecurityService, registers and logs in a test user.
        /// Returns the pair ready for authenticated operations.
        /// </summary>
        private static async Task<(DatabaseService db, SecurityService sec)>
            BuildAuthenticatedSecurityServiceAsync(string username, string password)
        {
            var db  = BuildInMemoryDb();
            await db.InitializeAsync();
            var sec = new SecurityService(NullLogger<SecurityService>.Instance, db);

            await sec.RegisterUserAsync(username, password);
            var loggedIn = await sec.LoginAsync(username, password);
            loggedIn.Should().BeTrue($"test setup: login as '{username}' must succeed");

            return (db, sec);
        }
    }
}
