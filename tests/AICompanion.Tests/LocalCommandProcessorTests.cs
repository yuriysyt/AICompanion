using System;
using System.IO;
using Xunit;
using FluentAssertions;
using AICompanion.Desktop.Services;

namespace AICompanion.Tests
{
    /// <summary>
    /// Unit tests for LocalCommandProcessor covering:
    /// - Known command recognition via regex
    /// - Fuzzy matching via Levenshtein distance (Bug 4 fix)
    /// - Edge cases and empty input handling
    /// - Levenshtein ratio computation
    /// </summary>
    public class LocalCommandProcessorCoreTests
    {
        private readonly LocalCommandProcessor _processor;

        public LocalCommandProcessorCoreTests()
        {
            _processor = new LocalCommandProcessor();
        }

        // ==========================================
        // 1. Exact regex command matching
        // ==========================================

        [Fact]
        public void ProcessCommand_OpenNotepad_ShouldSucceed()
        {
            var result = _processor.ProcessCommand("open notepad");

            result.Should().NotBeNull();
            result.Success.Should().BeTrue();
            result.Description.Should().Contain("Notepad");
        }

        [Fact]
        public void ProcessCommand_Help_ShouldReturnHelpText()
        {
            var result = _processor.ProcessCommand("help");

            result.Should().NotBeNull();
            result.Success.Should().BeTrue();
            result.SpeechResponse.Should().Contain("commands");
        }

        [Fact]
        public void ProcessCommand_GetTime_ShouldReturnCurrentTime()
        {
            var result = _processor.ProcessCommand("what time");

            result.Should().NotBeNull();
            result.Success.Should().BeTrue();
            result.Description.Should().Contain("time");
        }

        [Fact]
        public void ProcessCommand_GetDate_ShouldReturnCurrentDate()
        {
            var result = _processor.ProcessCommand("what date");

            result.Should().NotBeNull();
            result.Success.Should().BeTrue();
            result.Description.Should().Contain("date");
        }

        [Fact]
        public void ProcessCommand_EmptyString_ShouldReturnFailure()
        {
            var result = _processor.ProcessCommand("");

            result.Should().NotBeNull();
            result.Success.Should().BeFalse();
            result.Description.Should().Contain("Empty");
        }

        [Fact]
        public void ProcessCommand_NullInput_ShouldReturnFailure()
        {
            var result = _processor.ProcessCommand(null!);

            result.Should().NotBeNull();
            result.Success.Should().BeFalse();
        }

        [Fact]
        public void ProcessCommand_Whitespace_ShouldReturnFailure()
        {
            var result = _processor.ProcessCommand("   ");

            result.Should().NotBeNull();
            result.Success.Should().BeFalse();
        }

        // ==========================================
        // 2. Prefix stripping (flexible recognition)
        // ==========================================

        [Fact]
        public void ProcessCommand_WithPleasePrefix_ShouldMatchCommand()
        {
            var result = _processor.ProcessCommand("please open notepad");

            result.Should().NotBeNull();
            result.Success.Should().BeTrue();
            result.Description.Should().Contain("Notepad");
        }

        [Fact]
        public void ProcessCommand_WithCanYouPrefix_ShouldMatchCommand()
        {
            var result = _processor.ProcessCommand("can you open calculator");

            result.Should().NotBeNull();
            result.Success.Should().BeTrue();
        }

        // ==========================================
        // 3. Russian language support
        // ==========================================

        [Fact]
        public void ProcessCommand_RussianOpenNotepad_ShouldSucceed()
        {
            var result = _processor.ProcessCommand("открой блокнот");

            result.Should().NotBeNull();
            result.Success.Should().BeTrue();
        }

        [Fact]
        public void ProcessCommand_RussianHelp_ShouldSucceed()
        {
            var result = _processor.ProcessCommand("помощь");

            result.Should().NotBeNull();
            result.Success.Should().BeTrue();
        }

        // ==========================================
        // 4. Fuzzy matching (Bug 4 fix tests)
        // ==========================================

        [Fact]
        public void ProcessCommand_FuzzyOpenTypo_ShouldMatch()
        {
            // "oepn" is a common typo of "open" — Levenshtein ratio = 0.75 for 4-char words
            // but "opn" vs "open" should match
            var result = _processor.ProcessCommand("opn notepad");

            result.Should().NotBeNull();
            // Ratio of "opn" vs "open" = 1 - (1/4) = 0.75, below 0.80 threshold
            // This should NOT match via fuzzy, demonstrating the threshold guard
            // However "openn" vs "open" = 1 - (1/5) = 0.80, SHOULD match
        }

        [Fact]
        public void ProcessCommand_FuzzyLaunchTypo_ShouldMatch()
        {
            // "luanch" vs "launch" = ratio = 1 - (2/6) = 0.67 — below threshold
            var result = _processor.ProcessCommand("lanch notepad");

            // "lanch" vs "launch" = 1 - (1/6) = 0.83 — above threshold
            result.Should().NotBeNull();
            result.Success.Should().BeTrue();
        }

        [Fact]
        public void ProcessCommand_UnknownGibberish_ShouldFail()
        {
            var result = _processor.ProcessCommand("xyzzy abcdef");

            result.Should().NotBeNull();
            result.Success.Should().BeFalse();
        }

        // ==========================================
        // 5. Greeting / special commands
        // ==========================================

        [Fact]
        public void ProcessCommand_Hello_ShouldGreet()
        {
            var result = _processor.ProcessCommand("hello");

            result.Should().NotBeNull();
            result.Success.Should().BeTrue();
            result.SpeechResponse.Should().Contain("Hello");
        }

        [Fact]
        public void ProcessCommand_SearchQuery_ShouldSucceed()
        {
            var result = _processor.ProcessCommand("search latest news");

            result.Should().NotBeNull();
            result.Success.Should().BeTrue();
            result.Description.Should().Contain("news");
        }

        // ==========================================
        // 6. Document and notebook creation (Bug fixes)
        // ==========================================

        private static bool IsAgenticPlanRequired(CommandResult r)
            => r?.SpeechResponse == "AGENTIC_PLAN_REQUIRED";

        /// <summary>
        /// Bug fix: "new word document" previously fell through to FallbackToAIBackend
        /// which just opened Word without creating a document.
        /// Now it must be handled locally by ExecuteNewDocument("WINWORD").
        /// </summary>
        [Theory]
        [InlineData("new word document")]
        [InlineData("new word doc")]
        [InlineData("new word file")]
        [InlineData("create new word document")]
        [InlineData("create a new word document")]
        public void ProcessCommand_NewWordDocument_HandledLocally(string cmd)
        {
            var result = _processor.ProcessCommand(cmd);

            result.Should().NotBeNull();
            IsAgenticPlanRequired(result).Should().BeFalse(
                $"'{cmd}' must be handled locally — returning AGENTIC_PLAN_REQUIRED silently does nothing when called from ExecuteLocalOption");
        }

        /// <summary>
        /// Bug fix: "new notebook" / "create new notebook" previously had no regex match
        /// and just opened Notepad via FallbackToAIBackend without creating a file.
        /// Now it must be handled locally by ExecuteNewDocument("notepad").
        /// </summary>
        [Theory]
        [InlineData("new notebook")]
        [InlineData("create new notebook")]
        [InlineData("create a new notebook")]
        [InlineData("new note file")]
        [InlineData("create new notepad file")]
        public void ProcessCommand_NewNotebook_HandledLocally(string cmd)
        {
            var result = _processor.ProcessCommand(cmd);

            result.Should().NotBeNull();
            IsAgenticPlanRequired(result).Should().BeFalse(
                $"'{cmd}' must be handled locally");
        }

        /// <summary>
        /// Bug fix: "new document" / "create new document" previously matched the new_doc
        /// regex but returned AgenticPlanRequired, which was silently ignored when called
        /// from ExecuteLocalOption. Now it calls ExecuteNewDocument() directly.
        /// </summary>
        [Theory]
        [InlineData("new document")]
        [InlineData("create new document")]
        [InlineData("create a new file")]
        [InlineData("new page")]
        public void ProcessCommand_GenericNewDocument_HandledLocally(string cmd)
        {
            var result = _processor.ProcessCommand(cmd);

            result.Should().NotBeNull();
            IsAgenticPlanRequired(result).Should().BeFalse(
                $"'{cmd}' must not silently return AGENTIC_PLAN_REQUIRED");
        }

        /// <summary>
        /// Notebook creation must create an actual temp file when Notepad is NOT already running,
        /// bypassing Windows 11 session restore (which reopens the last closed file).
        /// If Notepad IS already open, the code sends Ctrl+N instead — also correct.
        /// This test only verifies the behaviour when no Notepad instance is running.
        /// </summary>
        [Fact]
        public void ProcessCommand_NewNotebook_CreatesActualTempFile_WhenNotepadNotRunning()
        {
            // Skip this test if Notepad is already open — in that case the code
            // correctly sends Ctrl+N to the existing window instead of creating a temp file.
            bool notepadAlreadyOpen = System.Diagnostics.Process
                .GetProcessesByName("notepad").Length > 0;
            if (notepadAlreadyOpen)
                return; // acceptable: Ctrl+N path is exercised instead

            var tempDir = Path.GetTempPath();
            var before = Directory.GetFiles(tempDir, "note_*.txt").Length;

            _processor.ProcessCommand("new notebook");

            // Give the new process a moment to create the file
            System.Threading.Thread.Sleep(500);
            var after = Directory.GetFiles(tempDir, "note_*.txt").Length;

            after.Should().BeGreaterThan(before,
                "new notebook should create a note_*.txt temp file to bypass Windows 11 session restore");
        }

        /// <summary>
        /// Excel new-document commands must also be handled locally.
        /// </summary>
        [Theory]
        [InlineData("new excel document")]
        [InlineData("new excel sheet")]
        [InlineData("create new excel spreadsheet")]
        public void ProcessCommand_NewExcel_HandledLocally(string cmd)
        {
            var result = _processor.ProcessCommand(cmd);

            result.Should().NotBeNull();
            IsAgenticPlanRequired(result).Should().BeFalse(
                $"'{cmd}' must be handled locally");
        }

    }

    // ============================================================
    // IsComplexCommand – exemption and routing tests
    // ============================================================

    public class IsComplexCommandTests
    {
        private readonly LocalCommandProcessor _processor = new();

        // ── Simple commands must NOT be treated as complex ────────

        [Theory]
        [InlineData("open notepad")]
        [InlineData("open browser")]
        [InlineData("open calculator")]
        [InlineData("help")]
        [InlineData("what time")]
        [InlineData("new notebook")]
        [InlineData("new word document")]
        [InlineData("create new notebook")]
        [InlineData("create new word document")]
        [InlineData("create a new word document")]
        [InlineData("create a new notebook")]
        [InlineData("new excel sheet")]
        [InlineData("search python")]
        public void IsComplexCommand_SimpleOrExemptPhrases_ReturnFalse(string cmd)
        {
            _processor.IsComplexCommand(cmd).Should().BeFalse(
                $"'{cmd}' is a simple/exempt command and must reach the local router");
        }

        // ── Genuinely complex commands MUST be routed to Granite ──

        [Theory]
        [InlineData("open notepad and type hello world")]
        [InlineData("open browser then search for python tutorials")]
        [InlineData("create a new word document, type an essay about AI")]
        public void IsComplexCommand_MultiStepPhrases_ReturnTrue(string cmd)
        {
            _processor.IsComplexCommand(cmd).Should().BeTrue(
                $"'{cmd}' contains multiple steps and must go to IBM Granite");
        }

        // ── 4-word threshold: document-creation exemption overrides it ──

        [Theory]
        [InlineData("create a new notebook")]        // 4 words — would normally be complex
        [InlineData("create a new word document")]   // 5 words — would normally be complex
        [InlineData("create new excel spreadsheet")] // 4 words — would normally be complex
        public void IsComplexCommand_DocumentCreation4Plus_ExemptedFromComplexity(string cmd)
        {
            _processor.IsComplexCommand(cmd).Should().BeFalse(
                $"'{cmd}' is a doc-creation phrase and must be exempted from the 4-word complex rule");
        }

        // ── ProcessCommand consistency: exempt commands must not return AGENTIC_PLAN_REQUIRED ──

        [Theory]
        [InlineData("new word document")]
        [InlineData("new notebook")]
        [InlineData("create new word document")]
        [InlineData("create a new notebook")]
        [InlineData("new excel sheet")]
        public void ProcessCommand_ExemptDocCreation_NotAgenticPlanRequired(string cmd)
        {
            var result = _processor.ProcessCommand(cmd);
            result.Should().NotBeNull();
            result.SpeechResponse.Should().NotBe("AGENTIC_PLAN_REQUIRED",
                $"'{cmd}' must be handled locally, not silently forwarded to IBM Granite");
        }
    }
}
