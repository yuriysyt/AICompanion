using Xunit;
using FluentAssertions;
using AICompanion.Desktop.Models;

namespace AICompanion.Tests
{
    /*
        Unit tests for the ActionResult model class.
        
        These tests verify the factory methods and state transitions
        for action results used throughout the application.
    */
    public class ActionResultTests
    {
        [Fact]
        public void Success_ShouldCreateSuccessfulResult()
        {
            var result = ActionResult.Success(
                "OpenApplication",
                "Opened Notepad",
                "I have opened Notepad for you.");

            result.IsSuccess.Should().BeTrue();
            result.ActionType.Should().Be("OpenApplication");
            result.ResultDescription.Should().Be("Opened Notepad");
            result.SpeechFeedback.Should().Be("I have opened Notepad for you.");
        }

        [Fact]
        public void Success_ShouldSetHappyEmotion()
        {
            var result = ActionResult.Success("Test", "Test", "Test");

            result.AvatarState.Should().Be(AvatarEmotion.Happy);
        }

        [Fact]
        public void Failure_ShouldCreateFailedResult()
        {
            var result = ActionResult.Failure(
                "ClickElement",
                "Could not find button",
                "I could not find that button.");

            result.IsSuccess.Should().BeFalse();
            result.ActionType.Should().Be("ClickElement");
            result.ResultDescription.Should().Be("Could not find button");
        }

        [Fact]
        public void Failure_ShouldSetConfusedEmotion()
        {
            var result = ActionResult.Failure("Test", "Test", "Test");

            result.AvatarState.Should().Be(AvatarEmotion.Confused);
        }

        [Theory]
        [InlineData(AvatarEmotion.Neutral)]
        [InlineData(AvatarEmotion.Listening)]
        [InlineData(AvatarEmotion.Thinking)]
        [InlineData(AvatarEmotion.Speaking)]
        [InlineData(AvatarEmotion.Happy)]
        [InlineData(AvatarEmotion.Confused)]
        public void AvatarState_ShouldAcceptAllEmotions(AvatarEmotion emotion)
        {
            var result = new ActionResult { AvatarState = emotion };

            result.AvatarState.Should().Be(emotion);
        }

        [Fact]
        public void DefaultResult_ShouldBeFailure()
        {
            var result = new ActionResult();

            result.IsSuccess.Should().BeFalse();
        }
    }
}
