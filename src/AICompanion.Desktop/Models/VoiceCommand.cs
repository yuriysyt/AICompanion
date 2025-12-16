using System;

namespace AICompanion.Desktop.Models
{
    /*
        VoiceCommand represents a single voice input captured from the user's microphone.
        
        This class serves as the primary data transfer object between the voice recognition
        service and the command processing pipeline. Each instance contains the transcribed
        text along with metadata about the capture process such as timing and confidence.
        
        ElevenLabs Speech-to-Text API provides the audio-to-text conversion,
        and this class packages that output with additional context needed for AI processing.
    */
    public class VoiceCommand
    {
        /*
            The transcribed text from the user's spoken command.
            This is the primary input that gets sent to the IBM Granite AI model
            for natural language understanding and intent classification.
        */
        public string TranscribedText { get; set; } = string.Empty;

        /*
            Unique identifier generated when the command is captured.
            This ID follows the command through the entire processing pipeline,
            enabling correlation of logs and debugging of specific interactions.
        */
        public string CommandId { get; set; } = Guid.NewGuid().ToString();

        /*
            Timestamp when the voice input was captured by the microphone.
            Used to calculate end-to-end latency from speech to action execution.
        */
        public DateTime CapturedAt { get; set; } = DateTime.UtcNow;

        /*
            Confidence score from the speech recognition engine, ranging from 0.0 to 1.0.
            Higher values indicate more certainty about the transcription accuracy.
            If this value falls below the configured threshold, the system may ask
            the user to repeat their command.
        */
        public float RecognitionConfidence { get; set; }

        /*
            Duration of the audio input in milliseconds.
            Helps distinguish between short commands and longer dictation requests.
        */
        public int AudioDurationMs { get; set; }

        /*
            Flag indicating whether this command was triggered by the wake word
            ("Hey Assistant") or by the push-to-talk button. Wake word activation
            requires additional processing to strip the wake phrase from the command.
        */
        public bool WasWakeWordTriggered { get; set; }

        /*
            Optional reference to any previous command in the conversation.
            This enables context tracking for follow-up commands like "do that again"
            or references to previously mentioned files and applications.
        */
        public string? PreviousCommandId { get; set; }

        /*
            Indicates if the command processing has been cancelled by the user.
            The user can interrupt processing by saying "cancel" or pressing Escape.
        */
        public bool IsCancelled { get; set; }

        /*
            Factory method to create a VoiceCommand from raw transcription output.
            This encapsulates the initialization logic and ensures all required
            fields are properly set.
        */
        public static VoiceCommand FromTranscription(string text, float confidence)
        {
            return new VoiceCommand
            {
                TranscribedText = text?.Trim() ?? string.Empty,
                RecognitionConfidence = confidence,
                CapturedAt = DateTime.UtcNow
            };
        }

        /*
            Returns a string representation useful for logging and debugging.
        */
        public override string ToString()
        {
            return $"[{CommandId}] \"{TranscribedText}\" (confidence: {RecognitionConfidence:P0})";
        }
    }
}
