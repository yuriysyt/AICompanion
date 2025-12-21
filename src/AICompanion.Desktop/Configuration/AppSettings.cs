using System;
using System.IO;
using System.Text.Json;

namespace AICompanion.Desktop.Configuration
{
    /*
        AppSettings manages application configuration and user preferences.
        
        Settings are stored in a JSON file in the user's local application data
        folder, ensuring they persist between sessions and are isolated per user.
        The class provides strongly-typed access to all configurable options
        with sensible defaults for new installations.
        
        Configuration categories:
        - Voice: Speech recognition and synthesis settings
        - AI: Model selection and processing options
        - UI: Avatar appearance and window behavior
        - Privacy: Data storage and logging preferences
        
        Settings are loaded once at startup and saved whenever changes are made.
    */
    public class AppSettings
    {
        /*
            File path where settings are persisted.
        */
        private static readonly string SettingsFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "AICompanion", "settings.json");

        /*
            Voice recognition settings.
        */
        public VoiceSettings Voice { get; set; } = new();

        /*
            AI processing settings.
        */
        public AISettings AI { get; set; } = new();

        /*
            User interface settings.
        */
        public UISettings UI { get; set; } = new();

        /*
            Privacy and data settings.
        */
        public PrivacySettings Privacy { get; set; } = new();

        /*
            Loads settings from the JSON file, or creates defaults if not found.
        */
        public static AppSettings Load()
        {
            try
            {
                if (File.Exists(SettingsFilePath))
                {
                    var json = File.ReadAllText(SettingsFilePath);
                    var settings = JsonSerializer.Deserialize<AppSettings>(json);
                    return settings ?? new AppSettings();
                }
            }
            catch
            {
                /*
                    If loading fails, return defaults.
                    This handles corrupted settings files gracefully.
                */
            }

            return new AppSettings();
        }

        /*
            Saves current settings to the JSON file.
        */
        public void Save()
        {
            try
            {
                var directory = Path.GetDirectoryName(SettingsFilePath);
                if (!string.IsNullOrEmpty(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                var options = new JsonSerializerOptions { WriteIndented = true };
                var json = JsonSerializer.Serialize(this, options);
                File.WriteAllText(SettingsFilePath, json);
            }
            catch
            {
                /*
                    Log the error but do not crash.
                    Settings will use in-memory values until next save.
                */
            }
        }
    }

    /*
        Voice recognition and text-to-speech configuration.
    */
    public class VoiceSettings
    {
        /*
            The wake word phrase that activates the assistant.
            Users can customize this to avoid conflicts with other assistants.
        */
        public string WakeWord { get; set; } = "Hey Assistant";

        /*
            Whether wake word detection is enabled.
            When disabled, only push-to-talk works.
        */
        public bool WakeWordEnabled { get; set; } = true;

        /*
            Minimum confidence threshold for accepting speech recognition results.
            Higher values reduce false positives but may miss valid commands.
        */
        public float RecognitionConfidenceThreshold { get; set; } = 0.65f;

        /*
            Name of the text-to-speech voice to use.
            Available voices depend on installed Windows language packs.
        */
        public string SpeechVoice { get; set; } = "Microsoft David Desktop";

        /*
            Speech rate from -10 (slowest) to 10 (fastest).
            Default of -1 is slightly slower than normal for accessibility.
        */
        public int SpeechRate { get; set; } = -1;

        /*
            Speech volume from 0 (silent) to 100 (maximum).
        */
        public int SpeechVolume { get; set; } = 90;
    }

    /*
        AI model and processing configuration.
    */
    public class AISettings
    {
        /*
            Address of the Python AI engine gRPC server.
            Default is localhost for local deployment.
        */
        public string EngineAddress { get; set; } = "http://localhost:50051";

        /*
            Timeout for AI requests in seconds.
        */
        public int RequestTimeoutSeconds { get; set; } = 10;

        /*
            Whether to use the cloud API as fallback when local engine is unavailable.
        */
        public bool UseCloudFallback { get; set; } = false;

        /*
            Number of previous interactions to include in context.
        */
        public int ConversationHistoryLength { get; set; } = 10;
    }

    /*
        User interface and visual settings.
    */
    public class UISettings
    {
        /*
            Avatar character type: Cat, Dog, or Robot.
        */
        public string AvatarType { get; set; } = "Robot";

        /*
            Whether the main window stays on top of other windows.
        */
        public bool AlwaysOnTop { get; set; } = true;

        /*
            Whether to minimize to system tray instead of taskbar.
        */
        public bool MinimizeToTray { get; set; } = false;

        /*
            Whether to start the application when Windows starts.
        */
        public bool StartWithWindows { get; set; } = false;

        /*
            Initial window position X coordinate.
            Negative value means use default (bottom right corner).
        */
        public int WindowPositionX { get; set; } = -1;

        /*
            Initial window position Y coordinate.
        */
        public int WindowPositionY { get; set; } = -1;
    }

    /*
        Privacy and data handling settings.
    */
    public class PrivacySettings
    {
        /*
            Whether to save conversation history between sessions.
        */
        public bool SaveConversationHistory { get; set; } = true;

        /*
            Whether to send anonymous usage statistics.
            No personal data is included in these reports.
        */
        public bool SendAnonymousStatistics { get; set; } = false;

        /*
            Whether to send crash reports for debugging.
        */
        public bool SendCrashReports { get; set; } = false;

        /*
            Number of days to retain conversation history.
            Set to 0 for unlimited retention.
        */
        public int HistoryRetentionDays { get; set; } = 30;

        /*
            Whether to require password/PIN to launch the application.
        */
        public bool RequireAuthentication { get; set; } = false;
    }
}
