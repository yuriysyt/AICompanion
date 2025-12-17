using System;
using System.Collections.Generic;
using System.Linq;

namespace AICompanion.Desktop.Models
{
    /*
        ConversationHistory maintains a record of recent interactions between
        the user and the AI assistant.
        
        This history provides context for the AI model to understand references
        like "do that again" or "no, the other one". The IBM Granite model uses
        this context to generate more accurate and contextually appropriate
        responses.
        
        The history is limited to a configurable number of turns to balance
        context quality with memory usage and processing time.
    */
    public class ConversationHistory
    {
        /*
            Maximum number of conversation turns to retain.
            Each turn consists of a user command and assistant response pair.
        */
        private readonly int _maxTurns;
        
        /*
            Internal storage for conversation turns in chronological order.
        */
        private readonly List<ConversationTurn> _turns;

        /*
            Creates a new conversation history with the specified capacity.
        */
        public ConversationHistory(int maxTurns = 10)
        {
            _maxTurns = maxTurns;
            _turns = new List<ConversationTurn>();
        }

        /*
            Adds a new turn to the conversation history.
            If the history exceeds the maximum capacity, the oldest turn
            is automatically removed.
        */
        public void AddTurn(VoiceCommand userCommand, ActionResult assistantResponse)
        {
            var turn = new ConversationTurn
            {
                UserInput = userCommand.TranscribedText,
                AssistantResponse = assistantResponse.SpeechFeedback ?? assistantResponse.ResultDescription,
                ActionPerformed = assistantResponse.ActionType,
                WasSuccessful = assistantResponse.IsSuccess,
                Timestamp = DateTime.UtcNow
            };

            _turns.Add(turn);

            /*
                Remove oldest turns if we exceed capacity.
            */
            while (_turns.Count > _maxTurns)
            {
                _turns.RemoveAt(0);
            }
        }

        /*
            Returns all turns in the conversation history.
        */
        public IReadOnlyList<ConversationTurn> GetAllTurns()
        {
            return _turns.AsReadOnly();
        }

        /*
            Returns the most recent N turns for context building.
        */
        public IEnumerable<ConversationTurn> GetRecentTurns(int count)
        {
            return _turns.TakeLast(count);
        }

        /*
            Formats the conversation history as a string suitable for
            inclusion in an AI prompt.
        */
        public string ToPromptContext()
        {
            if (_turns.Count == 0)
            {
                return "No previous conversation.";
            }

            var lines = new List<string>();
            lines.Add("Recent conversation:");

            foreach (var turn in _turns.TakeLast(5))
            {
                lines.Add($"User: {turn.UserInput}");
                lines.Add($"Assistant: {turn.AssistantResponse}");
            }

            return string.Join("\n", lines);
        }

        /*
            Clears all conversation history.
        */
        public void Clear()
        {
            _turns.Clear();
        }

        /*
            Returns the number of turns currently stored.
        */
        public int TurnCount => _turns.Count;
    }

    /*
        Represents a single turn in the conversation, containing both
        the user's input and the assistant's response.
    */
    public class ConversationTurn
    {
        /*
            The transcribed text of what the user said.
        */
        public string UserInput { get; set; } = string.Empty;

        /*
            The verbal response provided by the assistant.
        */
        public string AssistantResponse { get; set; } = string.Empty;

        /*
            The type of action that was performed (e.g., "OpenApplication").
        */
        public string ActionPerformed { get; set; } = string.Empty;

        /*
            Whether the action completed successfully.
        */
        public bool WasSuccessful { get; set; }

        /*
            When this turn occurred.
        */
        public DateTime Timestamp { get; set; }
    }
}
