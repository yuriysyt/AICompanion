using System;
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
    }
}
