using Xunit;
using FluentAssertions;
using AICompanion.Desktop.Models;

namespace AICompanion.Tests
{
    /*
        Unit tests for the VoiceCommand model class.
        
        These tests verify that voice commands are properly created,
        initialized with correct defaults, and handle edge cases
        appropriately.
    */
    public class VoiceCommandTests
    {
        [Fact]
        public void FromTranscription_ShouldCreateCommandWithCorrectText()
        {
            var text = "open notepad";
            var confidence = 0.85f;

            var command = VoiceCommand.FromTranscription(text, confidence);

            command.TranscribedText.Should().Be(text);
            command.RecognitionConfidence.Should().Be(confidence);
        }

        [Fact]
        public void FromTranscription_ShouldTrimWhitespace()
        {
            var text = "  open notepad  ";

            var command = VoiceCommand.FromTranscription(text, 0.9f);

            command.TranscribedText.Should().Be("open notepad");
        }

        [Fact]
        public void FromTranscription_ShouldHandleNullText()
        {
            var command = VoiceCommand.FromTranscription(null!, 0.5f);

            command.TranscribedText.Should().BeEmpty();
        }

        [Fact]
        public void NewCommand_ShouldHaveUniqueId()
        {
            var command1 = new VoiceCommand();
            var command2 = new VoiceCommand();

            command1.CommandId.Should().NotBe(command2.CommandId);
        }

        [Fact]
        public void NewCommand_ShouldHaveRecentTimestamp()
        {
            var before = DateTime.UtcNow;
            var command = new VoiceCommand();
            var after = DateTime.UtcNow;

            command.CapturedAt.Should().BeOnOrAfter(before);
            command.CapturedAt.Should().BeOnOrBefore(after);
        }

        [Fact]
        public void ToString_ShouldIncludeIdAndText()
        {
            var command = new VoiceCommand
            {
                TranscribedText = "test command",
                RecognitionConfidence = 0.95f
            };

            var result = command.ToString();

            result.Should().Contain(command.CommandId);
            result.Should().Contain("test command");
            result.Should().Contain("95");
        }
    }
}
