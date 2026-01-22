using Xunit;
using FluentAssertions;
using AICompanion.Desktop.Models;

namespace AICompanion.Tests
{
    /*
        Unit tests for the ConversationHistory model class.
        
        These tests verify conversation turn management, capacity limits,
        and prompt context generation.
    */
    public class ConversationHistoryTests
    {
        [Fact]
        public void NewHistory_ShouldBeEmpty()
        {
            var history = new ConversationHistory();

            history.TurnCount.Should().Be(0);
            history.GetAllTurns().Should().BeEmpty();
        }

        [Fact]
        public void AddTurn_ShouldIncrementCount()
        {
            var history = new ConversationHistory();
            var command = new VoiceCommand { TranscribedText = "open notepad" };
            var result = ActionResult.Success("Open", "Opened", "Done");

            history.AddTurn(command, result);

            history.TurnCount.Should().Be(1);
        }

        [Fact]
        public void AddTurn_ShouldStoreCorrectData()
        {
            var history = new ConversationHistory();
            var command = new VoiceCommand { TranscribedText = "test command" };
            var result = ActionResult.Success("TestAction", "Test result", "Test feedback");

            history.AddTurn(command, result);

            var turns = history.GetAllTurns();
            turns.Should().HaveCount(1);
            turns[0].UserInput.Should().Be("test command");
            turns[0].AssistantResponse.Should().Be("Test feedback");
            turns[0].WasSuccessful.Should().BeTrue();
        }

        [Fact]
        public void History_ShouldRespectMaxTurns()
        {
            var history = new ConversationHistory(maxTurns: 3);

            for (int i = 0; i < 5; i++)
            {
                var command = new VoiceCommand { TranscribedText = $"command {i}" };
                var result = ActionResult.Success("Test", "Test", "Test");
                history.AddTurn(command, result);
            }

            history.TurnCount.Should().Be(3);
            history.GetAllTurns()[0].UserInput.Should().Be("command 2");
        }

        [Fact]
        public void Clear_ShouldRemoveAllTurns()
        {
            var history = new ConversationHistory();
            var command = new VoiceCommand { TranscribedText = "test" };
            var result = ActionResult.Success("Test", "Test", "Test");
            
            history.AddTurn(command, result);
            history.AddTurn(command, result);
            history.Clear();

            history.TurnCount.Should().Be(0);
        }

        [Fact]
        public void GetRecentTurns_ShouldReturnLastN()
        {
            var history = new ConversationHistory();

            for (int i = 0; i < 5; i++)
            {
                var command = new VoiceCommand { TranscribedText = $"command {i}" };
                var result = ActionResult.Success("Test", "Test", "Test");
                history.AddTurn(command, result);
            }

            var recent = history.GetRecentTurns(2);
            recent.Should().HaveCount(2);
        }

        [Fact]
        public void ToPromptContext_ShouldReturnFormattedString()
        {
            var history = new ConversationHistory();
            var command = new VoiceCommand { TranscribedText = "open notepad" };
            var result = ActionResult.Success("Open", "Opened", "I opened Notepad");

            history.AddTurn(command, result);

            var context = history.ToPromptContext();
            context.Should().Contain("open notepad");
            context.Should().Contain("I opened Notepad");
        }

        [Fact]
        public void ToPromptContext_WhenEmpty_ShouldReturnNoConversation()
        {
            var history = new ConversationHistory();

            var context = history.ToPromptContext();

            context.Should().Be("No previous conversation.");
        }
    }
}
