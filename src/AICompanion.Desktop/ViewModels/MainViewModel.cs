using System;
using System.Threading.Tasks;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Extensions.Logging;
using AICompanion.Desktop.Models;
using AICompanion.Desktop.Services.Voice;
using AICompanion.Desktop.Services.Screen;
using AICompanion.Desktop.Services.Automation;
using AICompanion.Desktop.Services.Communication;
using AICompanion.Desktop.Services.Database;
using AICompanion.Desktop.Services.Security;
using AICompanion.Desktop.Services.Dictation;

namespace AICompanion.Desktop.ViewModels
{
    /*
        MainViewModel coordinates all application logic following the MVVM pattern.
        
        This view model acts as the central orchestrator, connecting the voice
        recognition service to the AI engine and routing action responses to
        the appropriate automation handlers. It maintains the application state
        and exposes bindable properties for the UI.
        
        The command processing flow:
        1. VoiceRecognitionService captures user speech
        2. ScreenCaptureService captures current screen state
        3. AIEngineClient sends command + context to Python backend
        4. Response is parsed and routed to UIAutomationService
        5. TextToSpeechService provides verbal feedback
        6. UI is updated to reflect the result
        
        Reference: https://docs.microsoft.com/en-us/dotnet/communitytoolkit/mvvm/
    */
    public partial class MainViewModel : ObservableObject
    {
        private readonly ILogger<MainViewModel> _logger;
        private readonly UnifiedVoiceManager _voiceManager;
        private readonly ScreenCaptureService _screenCapture;
        private readonly UIAutomationService _uiAutomation;
        private readonly AIEngineClient _aiClient;
        private readonly DatabaseService _database;
        private readonly SecurityService _security;
        private readonly DictationService _dictation;

        /*
            Observable properties that the UI binds to for display.
            Changes to these properties automatically update the view.
        */
        [ObservableProperty]
        private string _statusMessage = "Initializing...";

        [ObservableProperty]
        private bool _isListening;

        [ObservableProperty]
        private bool _isProcessing;

        [ObservableProperty]
        private bool _isConnected;

        [ObservableProperty]
        private AvatarEmotion _currentEmotion = AvatarEmotion.Neutral;

        [ObservableProperty]
        private string _lastCommand = string.Empty;

        [ObservableProperty]
        private bool _isDictating;

        [ObservableProperty]
        private string _activeVoiceEngine = "Windows";

        [ObservableProperty]
        private bool _isAuthenticated;

        [ObservableProperty]
        private string _currentUser = "";

        /*
            Stores the last processed command for "do that again" functionality.
        */
        private VoiceCommand? _previousCommand;
        private string _sessionId = Guid.NewGuid().ToString();

        public MainViewModel(
            ILogger<MainViewModel> logger,
            UnifiedVoiceManager voiceManager,
            ScreenCaptureService screenCapture,
            UIAutomationService uiAutomation,
            AIEngineClient aiClient,
            DatabaseService database,
            SecurityService security,
            DictationService dictation)
        {
            _logger = logger;
            _voiceManager = voiceManager;
            _screenCapture = screenCapture;
            _uiAutomation = uiAutomation;
            _aiClient = aiClient;
            _database = database;
            _security = security;
            _dictation = dictation;

            /*
                Subscribe to unified voice manager events for UI state updates only.
                Command processing is handled by MainWindow.ProcessLocalCommand to avoid
                duplicate speech responses. MainWindow subscribes to CommandRecognized
                through UnifiedVoiceManager and processes commands via LocalCommandProcessor.
            */
            _voiceManager.ListeningStateChanged += OnListeningStateChanged;
            _voiceManager.RecognitionError += OnRecognitionError;
            _voiceManager.HypothesisGenerated += OnHypothesisGenerated;
            _voiceManager.SpeakingStarted += OnSpeakingStarted;
            _voiceManager.SpeakingCompleted += OnSpeakingCompleted;

            /*
                Subscribe to dictation events.
            */
            _dictation.ModeChanged += OnDictationModeChanged;
            _dictation.TextDictated += OnTextDictated;

            /*
                Subscribe to security events.
            */
            _security.SessionExpired += OnSessionExpired;
            _security.SecurityCodeGenerated += OnSecurityCodeGenerated;

            /*
                Subscribe to AI connection state changes.
            */
            _aiClient.ConnectionStateChanged += OnConnectionStateChanged;
        }

        /*
            Initializes all services and establishes connection to AI engine.
            Called when the main window loads.
        */
        public async Task InitializeAsync()
        {
            try
            {
                _logger.LogInformation("Initializing application services");
                StatusMessage = "Starting up...";

                /*
                    Initialize database service.
                */
                await _database.InitializeAsync();

                /*
                    Initialize unified voice manager.
                    Only ONE voice engine (ElevenLabs OR Watson) will be active.
                */
                StatusMessage = "Initializing voice system...";
                await _voiceManager.InitializeAsync();
                ActiveVoiceEngine = "ElevenLabs";

                /*
                    Connect to the Python AI engine.
                */
                StatusMessage = "Connecting to AI engine...";
                var connected = await _aiClient.ConnectAsync();

                if (connected)
                {
                    StatusMessage = $"Ready. Voice: {ActiveVoiceEngine}. Say 'Hey Assistant' or click the microphone.";
                    CurrentEmotion = AvatarEmotion.Neutral;

                    /*
                        Start listening for wake word in background.
                    */
                    _voiceManager.StartListening();
                }
                else
                {
                    StatusMessage = $"AI engine not available. Voice: {ActiveVoiceEngine}. Please ensure AI engine is running.";
                    CurrentEmotion = AvatarEmotion.Confused;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error during initialization");
                StatusMessage = "Error starting up. Check logs for details.";
                CurrentEmotion = AvatarEmotion.Confused;
            }
        }

        /*
            Starts active listening mode for push-to-talk interaction.
        */
        public void StartListening()
        {
            if (!_voiceManager.IsListening)
            {
                _voiceManager.StartListening();
            }
            IsListening = true;
            CurrentEmotion = AvatarEmotion.Listening;
            StatusMessage = $"Listening ({ActiveVoiceEngine})...";
        }

        /*
            Stops active listening mode.
        */
        public void StopListening()
        {
            _voiceManager.StopListening();
            IsListening = false;
            if (CurrentEmotion == AvatarEmotion.Listening)
            {
                CurrentEmotion = AvatarEmotion.Neutral;
            }
        }


        /*
            Toggle dictation mode for typing into Word/Notepad.
        */
        public async Task ToggleDictationAsync()
        {
            var success = await _dictation.ToggleDictationAsync();
            IsDictating = _dictation.IsDictating;
            if (IsDictating)
            {
                StatusMessage = $"Dictation started to {_dictation.TargetApplication}";
                await _voiceManager.SpeakAsync("Dictation mode started. Speak and I'll type for you.");
            }
            else
            {
                StatusMessage = "Dictation stopped";
                await _voiceManager.SpeakAsync("Dictation mode stopped.");
            }
        }

        /*
            Login user.
        */
        public async Task<bool> LoginAsync(string username, string password)
        {
            var success = await _security.LoginAsync(username, password);
            IsAuthenticated = success;
            CurrentUser = success ? username : "";
            return success;
        }

        /*
            Register new user.
        */
        public async Task<bool> RegisterAsync(string username, string password)
        {
            return await _security.RegisterUserAsync(username, password);
        }

        /*
            Logout current user.
        */
        public void Logout()
        {
            _security.Logout();
            IsAuthenticated = false;
            CurrentUser = "";
        }

        /*
            Handles recognized voice commands from the speech recognition service.
        */
        private async void OnCommandRecognized(object? sender, VoiceCommand command)
        {
            if (IsProcessing)
            {
                _logger.LogDebug("Ignoring command while processing: {Text}", command.TranscribedText);
                return;
            }

            await ProcessCommandAsync(command);
        }

        /*
            Main command processing pipeline.
            
            This method orchestrates the flow from voice input to action execution:
            1. Check for dictation mode
            2. Check for dangerous operations requiring security code
            3. Capture screen context
            4. Send to AI engine
            5. Execute returned action
            6. Provide feedback and save to database
        */
        private async Task ProcessCommandAsync(VoiceCommand command)
        {
            try
            {
                IsProcessing = true;
                CurrentEmotion = AvatarEmotion.Thinking;
                StatusMessage = $"Understanding: \"{command.TranscribedText}\"";
                LastCommand = command.TranscribedText;

                _logger.LogInformation("Processing command: {Command}", command);

                /*
                    Check if this is a dictation command.
                */
                var (dictationHandled, dictationMessage) = await _dictation.ProcessVoiceCommandAsync(command.TranscribedText);
                if (dictationHandled)
                {
                    if (!string.IsNullOrEmpty(dictationMessage))
                    {
                        StatusMessage = dictationMessage;
                        await _voiceManager.SpeakAsync(dictationMessage);
                    }
                    IsProcessing = false;
                    return;
                }


                /*
                    Check for dangerous operations requiring security code.
                */
                if (_security.IsDangerousOperation(command.TranscribedText))
                {
                    var (allowed, message) = await _security.AuthorizeDangerousOperationAsync(command.TranscribedText);
                    if (!allowed)
                    {
                        StatusMessage = message ?? "Security code required";
                        await _voiceManager.SpeakAsync(message ?? "This operation requires a security code.");
                        IsProcessing = false;
                        return;
                    }
                }

                /*
                    Link to previous command for context.
                */
                if (_previousCommand != null)
                {
                    command.PreviousCommandId = _previousCommand.CommandId;
                }

                /*
                    Capture current screen state for AI context.
                */
                var screenContext = await _screenCapture.CaptureScreenAsync();

                /*
                    Add UI element information from automation.
                */
                screenContext.DiscoveredElements = _uiAutomation.GetInteractiveElements();
                screenContext.ActiveWindow = _uiAutomation.GetActiveWindowInfo();

                /*
                    Send to AI engine for processing.
                */
                var response = await _aiClient.ProcessCommandAsync(command, screenContext);

                if (response == null)
                {
                    await HandleErrorAsync("I could not process that command. Please try again.");
                    return;
                }

                /*
                    Route the response to appropriate handler.
                */
                var result = await ExecuteActionAsync(response);

                /*
                    Provide feedback to user.
                */
                await ProvideFeedbackAsync(result);

                /*
                    Save conversation to database.
                */
                await _database.SaveConversationAsync(
                    _sessionId,
                    _security.CurrentUserId,
                    command.TranscribedText,
                    result.SpeechFeedback,
                    result.ActionType,
                    result.IsSuccess ? "success" : "failed",
                    command.RecognitionConfidence);

                /*
                    Store for context in next command.
                */
                _previousCommand = command;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error processing command");
                await HandleErrorAsync("Something went wrong. Please try again.");
            }
            finally
            {
                IsProcessing = false;
            }
        }

        /*
            Routes AI response to the appropriate action handler.
        */
        private async Task<ActionResult> ExecuteActionAsync(Grpc.CommandResponse response)
        {
            _logger.LogDebug("Executing action: {Type}", response.ActionType);

            return response.ActionType switch
            {
                Grpc.ActionType.ActionOpenApplication => 
                    await ExecuteOpenApplicationAsync(response),
                
                Grpc.ActionType.ActionCloseWindow => 
                    await ExecuteCloseWindowAsync(response),
                
                Grpc.ActionType.ActionClickElement => 
                    await ExecuteClickElementAsync(response),
                
                Grpc.ActionType.Text =>
                    await ExecuteTypeTextAsync(response),
                
                Grpc.ActionType.ActionReadScreen => 
                    ExecuteReadScreen(response),
                
                Grpc.ActionType.ActionExplainHelp => 
                    ExecuteExplainHelp(response),
                
                Grpc.ActionType.ActionClarificationNeeded => 
                    ExecuteClarificationNeeded(response),
                
                _ => ActionResult.Failure(
                    response.ActionType.ToString(),
                    "Unknown action type",
                    response.ResponseText)
            };
        }

        private async Task<ActionResult> ExecuteOpenApplicationAsync(Grpc.CommandResponse response)
        {
            var appName = ExtractParameter(response.ActionParameters, "app_name");
            return await _uiAutomation.OpenApplicationAsync(appName);
        }

        private async Task<ActionResult> ExecuteCloseWindowAsync(Grpc.CommandResponse response)
        {
            var windowTitle = ExtractParameter(response.ActionParameters, "window_title");
            return await _uiAutomation.CloseWindowAsync(windowTitle);
        }

        private async Task<ActionResult> ExecuteClickElementAsync(Grpc.CommandResponse response)
        {
            var elementName = ExtractParameter(response.ActionParameters, "element_name");
            return await _uiAutomation.ClickElementAsync(elementName);
        }

        private async Task<ActionResult> ExecuteTypeTextAsync(Grpc.CommandResponse response)
        {
            var text = ExtractParameter(response.ActionParameters, "text");
            return await _uiAutomation.TypeTextAsync(text);
        }

        private ActionResult ExecuteReadScreen(Grpc.CommandResponse response)
        {
            return ActionResult.Success(
                "ReadScreen",
                "Reading screen content",
                response.ResponseText);
        }

        private ActionResult ExecuteExplainHelp(Grpc.CommandResponse response)
        {
            return ActionResult.Success(
                "ExplainHelp",
                "Providing help",
                response.ResponseText);
        }

        private ActionResult ExecuteClarificationNeeded(Grpc.CommandResponse response)
        {
            return new ActionResult
            {
                IsSuccess = true,
                ActionType = "Clarification",
                SpeechFeedback = response.ClarificationQuestion,
                AvatarState = AvatarEmotion.Thinking
            };
        }

        /*
            Extracts a parameter value from JSON action parameters.
        */
        private string ExtractParameter(string json, string key)
        {
            try
            {
                var doc = System.Text.Json.JsonDocument.Parse(json);
                if (doc.RootElement.TryGetProperty(key, out var value))
                {
                    return value.GetString() ?? string.Empty;
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Failed to parse action parameters");
            }
            return string.Empty;
        }

        /*
            Provides verbal and visual feedback based on action result.
        */
        private async Task ProvideFeedbackAsync(ActionResult result)
        {
            CurrentEmotion = result.AvatarState;
            StatusMessage = result.ResultDescription;

            if (!string.IsNullOrEmpty(result.SpeechFeedback))
            {
                await _voiceManager.SpeakAsync(result.SpeechFeedback);
            }
        }

        /*
            Handles error conditions with appropriate feedback.
        */
        private async Task HandleErrorAsync(string message)
        {
            CurrentEmotion = AvatarEmotion.Confused;
            StatusMessage = message;
            await _voiceManager.SpeakAsync(message);
        }

        /*
            Event handlers for service state changes.
        */
        private void OnListeningStateChanged(object? sender, bool isListening)
        {
            IsListening = isListening;
        }

        private void OnRecognitionError(object? sender, string error)
        {
            StatusMessage = error;
            CurrentEmotion = AvatarEmotion.Confused;
        }

        private void OnHypothesisGenerated(object? sender, string hypothesis)
        {
            // Show real-time transcription feedback
            if (!string.IsNullOrWhiteSpace(hypothesis))
            {
                StatusMessage = $"Hearing: {hypothesis}...";
            }
        }

        private void OnConnectionStateChanged(object? sender, bool isConnected)
        {
            IsConnected = isConnected;
        }

        private void OnSpeakingStarted(object? sender, EventArgs e)
        {
            CurrentEmotion = AvatarEmotion.Speaking;
        }

        private void OnSpeakingCompleted(object? sender, EventArgs e)
        {
            if (CurrentEmotion == AvatarEmotion.Speaking)
            {
                CurrentEmotion = AvatarEmotion.Neutral;
            }
        }

        private void OnDictationModeChanged(object? sender, DictationMode mode)
        {
            IsDictating = mode != DictationMode.Disabled;
            if (IsDictating)
            {
                StatusMessage = $"Dictation mode: {mode}";
            }
        }

        private void OnTextDictated(object? sender, string text)
        {
            _logger.LogDebug("Dictated text: {Text}", text);
        }

        private void OnSessionExpired(object? sender, EventArgs e)
        {
            IsAuthenticated = false;
            CurrentUser = "";
            StatusMessage = "Session expired. Please login again.";
            _ = _voiceManager.SpeakAsync("Your session has expired. Please login again.");
        }

        private void OnSecurityCodeGenerated(object? sender, string code)
        {
            StatusMessage = $"Security code: {code} (valid for 5 minutes)";
            _ = _voiceManager.SpeakAsync($"Your security code is {string.Join(" ", code.ToCharArray())}. It expires in 5 minutes.");
        }

        /*
            Cleanup when application is closing.
        */
        public void Cleanup()
        {
            _logger.LogInformation("Cleaning up view model");
            
            // Unsubscribe from voice manager events
            _voiceManager.ListeningStateChanged -= OnListeningStateChanged;
            _voiceManager.RecognitionError -= OnRecognitionError;
            _voiceManager.HypothesisGenerated -= OnHypothesisGenerated;
            _voiceManager.SpeakingStarted -= OnSpeakingStarted;
            _voiceManager.SpeakingCompleted -= OnSpeakingCompleted;

            // Unsubscribe from dictation events
            _dictation.ModeChanged -= OnDictationModeChanged;
            _dictation.TextDictated -= OnTextDictated;

            // Unsubscribe from security events
            _security.SessionExpired -= OnSessionExpired;
            _security.SecurityCodeGenerated -= OnSecurityCodeGenerated;

            // Unsubscribe from AI client events
            _aiClient.ConnectionStateChanged -= OnConnectionStateChanged;

            // Dispose services
            _voiceManager.Dispose();
            _dictation.Dispose();
            _security.Dispose();
            _database.Dispose();
            _aiClient.Dispose();
        }
    }
}
