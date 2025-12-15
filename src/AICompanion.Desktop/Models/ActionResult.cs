using System;

namespace AICompanion.Desktop.Models
{
    /*
        ActionResult encapsulates the outcome of executing a computer control action.
        
        After the AI engine determines what action to take (open an application, click
        a button, type text, etc.), the automation service executes that action and
        returns an ActionResult describing what happened. This result is used to:
        
        1. Provide verbal feedback to the user through text-to-speech
        2. Update the companion avatar's emotional state (success = smile, error = frown)
        3. Log the interaction for debugging and conversation history
        4. Determine if follow-up actions or clarifications are needed
        
        The system aims for 95% action execution success rate, and this class helps
        track and analyze failures to improve reliability over time.
    */
    public class ActionResult
    {
        /*
            Indicates whether the action completed successfully.
            A successful result means the intended operation was performed
            without errors, though the outcome may still need user verification.
        */
        public bool IsSuccess { get; set; }

        /*
            The type of action that was attempted, matching the ActionType enum
            from the gRPC protocol definition. This helps correlate results
            with the original AI decision.
        */
        public string ActionType { get; set; } = string.Empty;

        /*
            Human-readable description of what happened during execution.
            For successful actions: "Opened Microsoft Word"
            For failures: "Could not find the specified file"
        */
        public string ResultDescription { get; set; } = string.Empty;

        /*
            Text to be spoken back to the user via text-to-speech.
            This is crafted to be clear and helpful, avoiding technical jargon.
            For errors, it includes suggestions for what the user can try next.
        */
        public string SpeechFeedback { get; set; } = string.Empty;

        /*
            Reference to the original command that triggered this action.
            Enables tracking the complete flow from voice input to execution result.
        */
        public string OriginalCommandId { get; set; } = string.Empty;

        /*
            Time taken to execute the action in milliseconds.
            The target is under 1 second for most actions to maintain
            responsive interaction.
        */
        public int ExecutionTimeMs { get; set; }

        /*
            Timestamp when the action completed, used for logging
            and calculating total end-to-end latency.
        */
        public DateTime CompletedAt { get; set; } = DateTime.UtcNow;

        /*
            If the action failed, this contains the error details.
            Used for logging and debugging but not shown directly to users.
        */
        public string? ErrorDetails { get; set; }

        /*
            If the action requires user confirmation before proceeding,
            this flag is set to true. For example, file deletion actions
            always require confirmation for safety.
        */
        public bool RequiresConfirmation { get; set; }

        /*
            The emotional state to display on the companion avatar.
            Maps to avatar animation states: happy, thinking, confused, error.
        */
        public AvatarEmotion AvatarState { get; set; } = AvatarEmotion.Neutral;

        /*
            Factory method to create a successful action result.
            Simplifies the common case where an action completes without issues.
        */
        public static ActionResult Success(string actionType, string description, string feedback)
        {
            return new ActionResult
            {
                IsSuccess = true,
                ActionType = actionType,
                ResultDescription = description,
                SpeechFeedback = feedback,
                AvatarState = AvatarEmotion.Happy
            };
        }

        /*
            Factory method to create a failed action result.
            Ensures error cases are properly structured with helpful feedback.
        */
        public static ActionResult Failure(string actionType, string description, string feedback, string? errorDetails = null)
        {
            return new ActionResult
            {
                IsSuccess = false,
                ActionType = actionType,
                ResultDescription = description,
                SpeechFeedback = feedback,
                ErrorDetails = errorDetails,
                AvatarState = AvatarEmotion.Confused
            };
        }
    }

    /*
        Enumeration of emotional states for the companion avatar.
        
        The avatar provides visual feedback to the user about the system's
        current state and the outcome of their commands. This helps users
        with hearing impairments and makes the interaction more engaging.
    */
    public enum AvatarEmotion
    {
        /*
            Default resting state with subtle idle animations (breathing, blinking).
        */
        Neutral = 0,

        /*
            Displayed when the system is actively listening for voice input.
            The avatar's ears perk up and eyes focus on the user.
        */
        Listening = 1,

        /*
            Shown while the AI is processing a command.
            Includes a thinking animation with a loading indicator.
        */
        Thinking = 2,

        /*
            Positive feedback when an action completes successfully.
            The avatar smiles and may perform a brief celebration.
        */
        Happy = 3,

        /*
            Displayed when something goes wrong or the system is confused.
            Encourages the user to rephrase or provide clarification.
        */
        Confused = 4,

        /*
            Shown during text-to-speech output.
            The avatar's mouth movements sync with the spoken words.
        */
        Speaking = 5
    }
}
