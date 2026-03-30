using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media.Animation;
using Microsoft.Extensions.DependencyInjection;
using AICompanion.Desktop.ViewModels;
using AICompanion.Desktop.Services;
using AICompanion.Desktop.Services.Voice;
using AICompanion.Desktop.Services.Tutorial;
using AICompanion.Desktop.Services.Dictation;
using AICompanion.Desktop.Helpers;
using AICompanion.Desktop.Services.Automation;
using AICompanion.Desktop.Services.Database;
using WpfBrush = System.Windows.Media.Brush;
using WpfBrushes = System.Windows.Media.Brushes;
using WpfColor = System.Windows.Media.Color;
using WpfSolidColorBrush = System.Windows.Media.SolidColorBrush;

namespace AICompanion.Desktop.Views
{
    public partial class MainWindow : System.Windows.Window
    {
        private readonly MainViewModel? _viewModel;
        private readonly LocalCommandProcessor _commandProcessor;
        private readonly ElevenLabsService? _elevenLabs;
        private readonly ElevenLabsSpeechService? _elevenLabsSTT;
        private readonly TutorialService _tutorialService;
        private readonly DictationService? _dictationService;
        private readonly UnifiedVoiceManager? _unifiedVoiceManager;
        private readonly AICompanion.Desktop.Services.Security.SecureApiKeyManager? _secureKeyManager;
        private readonly AICompanion.Desktop.Services.Database.DatabaseService? _databaseService;
        private AgenticExecutionService? _agenticService;
        private readonly ObservableCollection<ActivityLogItem> _activityItems;
        private bool _isDictationMode;
        private bool _isListening;
        private string _elevenLabsApiKey = "";
        private string _elevenLabsVoiceId = "21m00Tcm4TlvDq8ikWAM";
        private string? _lastHypothesis;
        private DateTime _lastHypothesisAt;
        private string? _lastProcessedCommand;
        private DateTime _lastProcessedAt;
        private string _sessionId = System.Guid.NewGuid().ToString();

        // IBM Granite smart suggestions dialog
        private AgenticPlanDto? _pendingPlan;
        private TaskCompletionSource<bool>? _planConfirmationTcs;
        private List<IntentOption> _currentOptions = new();
        private Action? _pendingLocalAction;
        /// <summary>
        /// When user picks an option that implies a specific target app (e.g. "Type in Word"),
        /// this overrides the first open_app/focus_window/new_document step in the Granite plan.
        /// </summary>
        private string? _planOverrideTarget;
        /// <summary>Alternative plans fetched from server, keyed by option label.</summary>
        private readonly System.Net.Http.HttpClient _uiHttpClient = new();
        private string _dialogCommandText = "";

        public MainWindow()
        {
            InitializeComponent();

            try
            {
                Icon = IconGenerator.CreateAppIcon(256);
            }
            catch { }

            _activityItems = new ObservableCollection<ActivityLogItem>();
            ActivityLog.ItemsSource = _activityItems;

            _commandProcessor = new LocalCommandProcessor();
            _tutorialService  = new TutorialService();
            _databaseService  = App.ServiceProvider?.GetService<AICompanion.Desktop.Services.Database.DatabaseService>();

            _tutorialService.SpeechRequested += (s, speech) =>
            {
                Dispatcher.Invoke(() =>
                {
                    _ = SpeakAsync(speech);
                });
            };

            _tutorialService.TutorialEvent += (s, args) =>
            {
                Dispatcher.Invoke(() =>
                {
                    switch (args.EventType)
                    {
                        case TutorialEventType.Started:
                            AddActivity("🎓 Tutorial started! Follow the instructions.", true);
                            TutorialProgressPanel.Visibility = Visibility.Visible;
                            UpdateTutorialProgress();
                            break;
                        case TutorialEventType.StepStarted:
                            AddActivity($"📝 Step {_tutorialService.CurrentStepIndex + 1}: {args.StepTitle}", true);
                            UpdateTutorialProgress();
                            break;
                        case TutorialEventType.StepCompleted:
                            AddActivity($"✅ Step completed!", true);
                            UpdateAvatarExpression(true);
                            UpdateTutorialProgress();
                            break;
                        case TutorialEventType.Completed:
                            AddActivity("🎉 Tutorial completed! You're ready to use AI Companion!", true);
                            TutorialProgressPanel.Visibility = Visibility.Collapsed;
                            break;
                        case TutorialEventType.Stopped:
                            AddActivity("Tutorial stopped.", true);
                            TutorialProgressPanel.Visibility = Visibility.Collapsed;
                            break;
                    }
                });
            };

            try
            {
                _viewModel = App.ServiceProvider?.GetService<MainViewModel>();
                _elevenLabs = App.ServiceProvider?.GetService<ElevenLabsService>();
                _elevenLabsSTT = App.ServiceProvider?.GetService<ElevenLabsSpeechService>();

                if (_viewModel != null)
                {
                    DataContext = _viewModel;
                }

                _unifiedVoiceManager = App.ServiceProvider?.GetService<UnifiedVoiceManager>();
                _secureKeyManager   = App.ServiceProvider?.GetService<AICompanion.Desktop.Services.Security.SecureApiKeyManager>();
                _dictationService = App.ServiceProvider?.GetService<DictationService>();
                // _agenticService is initialized lazily via EnsureAgenticService()
                // to allow ConfirmationRequired callback to be wired up cleanly.

                if (_unifiedVoiceManager != null)
                {
                    _unifiedVoiceManager.CommandRecognized += OnUnifiedCommandRecognized;
                    _unifiedVoiceManager.ListeningStateChanged += OnListeningStateChanged;
                    _unifiedVoiceManager.HypothesisGenerated += OnHypothesisGenerated;
                    _unifiedVoiceManager.RecognitionError += OnRecognitionError;
                }
            }
            catch (Exception ex)
            {
                Serilog.Log.Error(ex, "[MainWindow] Service init failed — voice may be unavailable");
                AddActivity($"Service initialization warning: {ex.Message}", false);
            }

            Loaded += MainWindow_Loaded;
            Closing += MainWindow_Closing;

            AddActivity("AI Companion ready! Hold the microphone button to speak, or use quick commands.", true);
        }

        private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_viewModel != null)
                {
                    await _viewModel.InitializeAsync();
                    AddActivity("Database initialized", true);
                }

                // ── Load last 5 conversations and restore app context from DB ────────────
                if (_databaseService != null && _databaseService.IsInitialized)
                {
                    try
                    {
                        // Restore saved app context (e.g. user was in Word last session)
                        var savedCtx = await _databaseService.GetContextAsync("app_context");
                        if (!string.IsNullOrEmpty(savedCtx))
                        {
                            _commandProcessor.RestoreContext(savedCtx);
                            AddActivity($"Restored context: {savedCtx}", true);
                        }

                        // Load and display last 5 conversations
                        var recent = await _databaseService.GetRecentConversationsAsync(_sessionId, limit: 5);
                        if (recent.Count > 0)
                        {
                            AddActivity("── Recent history ──────────────────", true);
                            foreach (var rec in System.Linq.Enumerable.Reverse(recent))
                            {
                                var ts = rec.CreatedAt.ToLocalTime().ToString("HH:mm");
                                AddActivity($"[{ts}] {rec.Command}", true);
                                if (!string.IsNullOrEmpty(rec.Response))
                                    AddActivity($"  → {rec.Response}", true);
                            }
                            AddActivity("── End of history ──────────────────", true);
                        }
                    }
                    catch (Exception dbEx)
                    {
                        Serilog.Log.Warning(dbEx, "[MainWindow] Could not load conversation history");
                    }
                }

                // API key stored in DPAPI; non-sensitive settings from appsettings.json
                var config = await LoadVoiceConfigAsync();
                _elevenLabsApiKey = _secureKeyManager?.LoadApiKey(
                    AICompanion.Desktop.Services.Security.SecureApiKeyManager.ElevenLabsKeyName) ?? "";
                if (string.IsNullOrEmpty(_elevenLabsApiKey))
                    _elevenLabsApiKey = config.ElevenLabsApiKey; // fallback: migrate from JSON
                _elevenLabsVoiceId = config.ElevenLabsVoiceId;

                if (_unifiedVoiceManager != null)
                {
                    var settings = new Services.Voice.VoiceSettings
                    {
                        ElevenLabsApiKey = _elevenLabsApiKey,
                        ElevenLabsVoiceId = _elevenLabsVoiceId
                    };
                    _unifiedVoiceManager.Configure(settings);

                    var initialized = await _unifiedVoiceManager.InitializeAsync();

                    if (initialized)
                    {
                        AddActivity("ElevenLabs voice engine ready", true);
                    }
                    else
                    {
                        AddActivity("ElevenLabs initialization failed. Click Settings to configure API key.", false);
                        RecognizedText.Text = "Click Settings to configure ElevenLabs API key";
                    }
                }

                UpdateStatus("Ready", true);
            }
            catch (Exception ex)
            {
                Serilog.Log.Error(ex, "[MainWindow] Loaded init failed");
                AddActivity($"Initialization error: {ex.Message}", false);
                UpdateStatus("Error", false);
            }
        }

        private static string? _lastSpokenText;
        private static DateTime _lastSpokenAt;

        private async System.Threading.Tasks.Task SpeakAsync(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return;

            if (_lastSpokenText == text && DateTime.UtcNow - _lastSpokenAt < TimeSpan.FromSeconds(3))
            {
                AddActivity($"[TTS] Skipped duplicate: {text.Substring(0, Math.Min(30, text.Length))}...", false);
                return;
            }

            _lastSpokenText = text;
            _lastSpokenAt = DateTime.UtcNow;

            if (_unifiedVoiceManager != null)
            {
                await _unifiedVoiceManager.SpeakAsync(text);
            }
        }

        private void StopAllSpeech()
        {
            _unifiedVoiceManager?.StopSpeaking();
        }

        private void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
        {
            if (_unifiedVoiceManager != null)
            {
                _unifiedVoiceManager.CommandRecognized -= OnUnifiedCommandRecognized;
                _unifiedVoiceManager.ListeningStateChanged -= OnListeningStateChanged;
                _unifiedVoiceManager.HypothesisGenerated -= OnHypothesisGenerated;
                _unifiedVoiceManager.RecognitionError -= OnRecognitionError;
            }

            _viewModel?.Cleanup();
        }

        private void OnHypothesisGenerated(object? sender, string hypothesis)
        {
            Dispatcher.Invoke(() =>
            {
                _lastHypothesis = hypothesis;
                _lastHypothesisAt = DateTime.UtcNow;
                RecognizedText.Text = $"{hypothesis}...";
                RecognizedText.FontStyle = FontStyles.Italic;
                RecognizedText.Foreground = (WpfBrush)FindResource("PrimaryBrush");
            });
        }

        private void OnRecognitionError(object? sender, string error)
        {
            Dispatcher.Invoke(() =>
            {
                AddActivity($"Voice error: {error}", false);
                _lastHypothesis = null;
                StopListening();
                RecognizedText.Text = $"Voice error: {error}";
                RecognizedText.FontStyle = FontStyles.Italic;
                RecognizedText.Foreground = (WpfBrush)FindResource("TextSecondaryBrush");
                UpdateStatus("Error", false);
            });
        }

        // ═══════════════════════════════════════════════════════════════════════
        // BLOCK 1 — INPUT DISPATCH PIPELINE
        // Single entry point for both voice and text input.
        // Routes: simple commands → LocalCommandProcessor
        //         complex commands → IBM Granite via AgenticExecutionService
        // ═══════════════════════════════════════════════════════════════════════

        private void OnUnifiedCommandRecognized(object? sender, Models.VoiceCommand command)
        {
            Dispatcher.Invoke(() => DispatchInput(command.TranscribedText, isVoice: true));
        }

        private void OnListeningStateChanged(object? sender, bool isListening)
        {
            Dispatcher.Invoke(() => { _isListening = isListening; UpdateListeningUI(isListening); });
        }

        /// <summary>Single entry point for all input — voice or text.</summary>
        private void DispatchInput(string rawText, bool isVoice = false)
        {
            var text = rawText?.Trim() ?? "";
            if (string.IsNullOrEmpty(text)) return;

            // Confirmation mode: user is answering YES/NO for a pending IBM Granite plan
            if (_planConfirmationTcs != null && !_planConfirmationTcs.Task.IsCompleted)
            {
                HandleConfirmationInput(text);
                return;
            }

            // Dedup: ignore the same command within 2 seconds
            if (string.Equals(_lastProcessedCommand, text, StringComparison.OrdinalIgnoreCase)
                && DateTime.UtcNow - _lastProcessedAt < TimeSpan.FromSeconds(2))
                return;
            _lastProcessedCommand = text;
            _lastProcessedAt      = DateTime.UtcNow;
            _lastHypothesis       = null;

            // Strip voice filler words
            text = StripFillerWords(text);
            if (string.IsNullOrEmpty(text)) return;

            // Update display
            RecognizedText.Text       = $"\"{text}\"";
            RecognizedText.FontStyle  = FontStyles.Normal;
            RecognizedText.Foreground = (WpfBrush)FindResource("TextPrimaryBrush");
            AddActivity($"{(isVoice ? "🎤" : "⌨️")} \"{text}\"", true);

            // Tutorial intercept
            if (_tutorialService.IsActive && _tutorialService.ProcessCommand(text)) return;

            _commandProcessor.CaptureTargetWindow();
            UpdateStatus("Processing...", true);

            // Voice auto-type: when a text editor is open and the input is free-form speech
            if (isVoice && IsTextEditorActive() && !LooksLikeCommand(text))
            {
                AddActivity($"✍️ Auto-typing to {_commandProcessor.GetTargetWindowTitle()}", true);
                _commandProcessor.ProcessCommand($"type: {text}");
                UpdateStatus("Typed", true);
                StartResetTimer(2);
                return;
            }

            // Route through command processor
            var result = _commandProcessor.ProcessCommand(text);

            // Tutorial signals
            switch (result.SpeechResponse)
            {
                case "TUTORIAL_START":
                    _tutorialService.StartTutorial();
                    AddActivity("🎓 Starting interactive tutorial...", true);
                    UpdateAvatarExpression(true);
                    return;
                case "TUTORIAL_STOP":  _tutorialService.StopTutorial();  return;
                case "TUTORIAL_SKIP":  _tutorialService.SkipStep();      return;
                case "TUTORIAL_HINT":  _tutorialService.RequestHint();   return;
            }

            // Route complex commands to IBM Granite
            if (result.SpeechResponse == "AGENTIC_PLAN_REQUIRED")
            {
                var svc = EnsureAgenticService();
                if (svc != null)
                {
                    // Inject session context so Granite knows which app is open
                    svc.SessionContext = _commandProcessor.GetSessionContextSnapshot();
                    AddActivity($"🤖 IBM Granite AI → \"{text}\"", true);
                    UpdateStatus("IBM Granite planning...", true);
                    _ = ExecuteAgenticPlanAsync(text);
                }
                else
                {
                    AddActivity("⚠️ IBM Granite unavailable. Start the Python backend server.", false);
                    UpdateStatus("Error", false);
                    StartResetTimer(4);
                }
                return;
            }

            // Local result
            AddActivity(result.Description, result.Success);
            UpdateAvatarExpression(result.Success);
            if (!string.IsNullOrEmpty(result.SpeechResponse))
                _ = SpeakAsync(result.SpeechResponse);
            UpdateStatus(result.Success ? "Done" : "Error", result.Success);
            StartResetTimer(3);

            // Persist command + response to SQLite history (fire-and-forget)
            _ = SaveCommandHistoryAsync(text, result.Description, result.Success);
        }

        private async System.Threading.Tasks.Task SaveCommandHistoryAsync(string command, string response, bool success)
        {
            if (_databaseService == null || !_databaseService.IsInitialized) return;
            try
            {
                await _databaseService.SaveConversationAsync(
                    sessionId:    _sessionId,
                    userId:       null,
                    command:      command,
                    response:     response,
                    actionType:   "local",
                    actionResult: success ? "success" : "failure",
                    confidence:   success ? 0.95f : 0.5f);

                // Persist the current app context so it can be restored next session
                var ctx = _commandProcessor.CurrentContextName;
                if (ctx != "None")
                    await _databaseService.SaveContextAsync(null, "app_context", ctx, "system", 0.9f);
            }
            catch (Exception ex)
            {
                Serilog.Log.Warning(ex, "[MainWindow] DB save failed");
            }
        }

        private static string StripFillerWords(string text)
        {
            text = System.Text.RegularExpressions.Regex.Replace(
                text, @"^(?:uh|um|so|well|like|okay|ok|right|hey|oh)\s*[,.]?\s*", "",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase).Trim();
            text = System.Text.RegularExpressions.Regex.Replace(
                text, @"^(?:can\s+you|could\s+you|would\s+you|please)\s+", "",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase).Trim();
            return text;
        }

        private bool IsTextEditorActive()
        {
            var t = _commandProcessor.GetTargetWindowTitle().ToLowerInvariant();
            return t.Contains("word") || t.Contains(".docx") || t.Contains("notepad")
                || t.Contains("блокнот") || t.Contains(".txt");
        }

        private static bool LooksLikeCommand(string text)
        {
            var lower = text.ToLowerInvariant();
            string[] prefixes =
            [
                "open", "close", "save", "copy", "paste", "search", "find", "start",
                "launch", "create", "new", "what", "how", "help", "scroll",
                "minimize", "maximize", "click", "press", "type", "write",
                "select", "delete", "undo", "redo", "run", "go to",
                "открой", "закрой", "сохрани", "найди", "создай", "новый", "новая",
                "напиши", "введи", "выдели", "запусти", "помощь", "привет", "hello", "hi"
            ];
            return prefixes.Any(p => lower.StartsWith(p));
        }

        private void StartResetTimer(int seconds)
        {
            var timer = new System.Windows.Threading.DispatcherTimer
                { Interval = TimeSpan.FromSeconds(seconds) };
            timer.Tick += (s, _) =>
            {
                ((System.Windows.Threading.DispatcherTimer)s!).Stop();
                UpdateStatus("Ready", true);
                ResetAvatarExpression();
            };
            timer.Start();
        }

        /// <summary>
        /// Returns the agentic service, creating it on first call (lazy init).
        /// Also wires up ConfirmationRequired so the plan dialog shows before execution.
        /// </summary>
        private AgenticExecutionService? EnsureAgenticService()
        {
            if (_agenticService != null) return _agenticService;

            // Try DI
            AgenticExecutionService? svc = App.ServiceProvider?.GetService<AgenticExecutionService>();

            if (svc == null)
            {
                // Direct creation fallback (DI unavailable)
                try
                {
                    var winHelper = App.ServiceProvider?.GetService<WindowAutomationHelper>()
                                    ?? new WindowAutomationHelper(null);
                    svc = new AgenticExecutionService(winHelper, null);
                    AddActivity("⚙️ AgenticService initialized (fallback)", true);
                }
                catch (Exception ex)
                {
                    AddActivity($"⚠️ Cannot start AgenticService: {ex.Message}", false);
                    return null;
                }
            }

            // Wire up status messages
            svc.StatusMessage += (_, msg) =>
                Dispatcher.Invoke(() => { AddActivity(msg, true); UpdateStatus(msg, true); });

            // Show options dialog — auto-execute top option after 5 seconds if user doesn't respond.
            svc.ConfirmationRequired = async (plan) =>
            {
                var tcs = new System.Threading.Tasks.TaskCompletionSource<bool>();
                Dispatcher.Invoke(() =>
                {
                    _planConfirmationTcs = tcs;
                    ShowOptionsDialog(plan);
                });

                // Auto-execute after 5 seconds — don't block the user
                var autoTimeout = System.Threading.Tasks.Task.Delay(5000);
                var completed   = await System.Threading.Tasks.Task.WhenAny(tcs.Task, autoTimeout);

                if (!tcs.Task.IsCompleted)
                {
                    // Timeout: auto-confirm with top option
                    tcs.TrySetResult(true);
                    Dispatcher.Invoke(() =>
                    {
                        GranitePlanPanel.Visibility = Visibility.Collapsed;
                        AddActivity("⚡ Auto-executing top suggestion...", true);
                    });
                    return true;
                }

                return await tcs.Task;
            };

            _agenticService = svc;
            return _agenticService;
        }

        private async System.Threading.Tasks.Task ExecuteAgenticPlanAsync(string commandText)
        {
            if (_agenticService == null)
            {
                AddActivity("Agentic service not available", false);
                return;
            }

            try
            {
                void OnStep(object? s, StepExecutedArgs args)
                {
                    Dispatcher.Invoke(() =>
                    {
                        var icon = args.Success ? "✅" : "❌";
                        AddActivity($"{icon} Step {args.StepNumber}/{args.TotalSteps}: {args.Action} — {args.Description}", args.Success);
                        UpdateStatus($"Step {args.StepNumber}/{args.TotalSteps}: {args.Action}", true);
                    });
                }

                void OnPlan(object? s, AgenticPlanDto plan)
                {
                    Dispatcher.Invoke(() =>
                    {
                        AddActivity($"📋 Plan ready: {plan.TotalSteps} steps — {plan.Reasoning}", true);
                        for (int i = 0; i < plan.Steps.Count; i++)
                        {
                            var st = plan.Steps[i];
                            AddActivity($"   Step {st.StepNumber}: {st.Action} → {st.Target ?? st.Params?.Substring(0, Math.Min(30, st.Params?.Length ?? 0)) ?? "—"}", true);
                        }
                        // Do NOT say "executing now" here — user hasn't confirmed yet
                    });
                }

                void OnMessage(object? s, string message)
                {
                    Dispatcher.Invoke(() =>
                    {
                        AddActivity(message, true);
                        UpdateStatus(message, true);
                    });
                }

                _agenticService.StepExecuted += OnStep;
                _agenticService.PlanReceived += OnPlan;
                _agenticService.StatusMessage += OnMessage;

                Serilog.Log.Information("[MW] ExecuteAgenticPlanAsync: sending '{Cmd}'", commandText);
                var result = await _agenticService.ExecuteCommandAsync(commandText);
                Serilog.Log.Information("[MW] ExecuteAgenticPlanAsync: result={Success} — {Summary}", result.Success, result.Summary);

                _agenticService.StepExecuted -= OnStep;
                _agenticService.PlanReceived -= OnPlan;
                _agenticService.StatusMessage -= OnMessage;

                Dispatcher.Invoke(() =>
                {
                    if (result.Success)
                    {
                        AddActivity($"🎉 {result.Summary}", true);
                        UpdateAvatarExpression(true);

                        // Build a context-aware completion message based on plan steps
                        string completionMsg = "All done! Plan executed successfully.";
                        if (_agenticService != null)
                        {
                            // Check if the plan involved typing text into an app
                            var lastPlan = _pendingPlan; // may be null after HidePlanDialog but capture before
                            bool hadTypeText = result.Summary?.Contains("type_text", StringComparison.OrdinalIgnoreCase) == true
                                              || result.Summary?.Contains("typed", StringComparison.OrdinalIgnoreCase) == true
                                              || result.Summary?.Contains("written", StringComparison.OrdinalIgnoreCase) == true
                                              || result.Summary?.Contains("wrote", StringComparison.OrdinalIgnoreCase) == true;
                            string targetAppName = "";
                            bool wordRunning    = System.Diagnostics.Process.GetProcessesByName("WINWORD").Length > 0;
                            bool notepadRunning = System.Diagnostics.Process.GetProcessesByName("notepad").Length > 0;
                            bool excelRunning   = System.Diagnostics.Process.GetProcessesByName("EXCEL").Length > 0;
                            if      (_planOverrideTarget?.Contains("winword")  == true || wordRunning)    targetAppName = "Word";
                            else if (_planOverrideTarget?.Contains("notepad")  == true || notepadRunning) targetAppName = "Notepad";
                            else if (_planOverrideTarget?.Contains("excel")    == true || excelRunning)   targetAppName = "Excel";

                            if (hadTypeText && !string.IsNullOrEmpty(targetAppName))
                            {
                                completionMsg = $"Text written to {targetAppName}. Check the {targetAppName} window.";
                                AddActivity($"✅ Essay written to {targetAppName}. Check the {targetAppName} window.", true);
                            }
                        }
                        _ = SpeakAsync(completionMsg);
                    }
                    else
                    {
                        // If user picked a local option, _pendingLocalAction was set before TCS=false
                        var localAction = _pendingLocalAction;
                        _pendingLocalAction = null;
                        if (localAction != null)
                        {
                            localAction();
                        }
                        else
                        {
                            AddActivity($"⚠️ {result.Summary}", false);
                            UpdateAvatarExpression(false);
                            _ = SpeakAsync($"There was an issue: {result.Summary}");
                        }
                    }

                    UpdateStatus("Ready", true);

                    var timer = new System.Windows.Threading.DispatcherTimer
                    {
                        Interval = TimeSpan.FromSeconds(3)
                    };
                    timer.Tick += (ts, te) =>
                    {
                        ((System.Windows.Threading.DispatcherTimer)ts!).Stop();
                        ResetAvatarExpression();
                    };
                    timer.Start();
                });
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() =>
                {
                    AddActivity($"Agentic execution error: {ex.Message}", false);
                    UpdateStatus("Error", false);
                });
            }
        }

        private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2)
            {
                WindowState = WindowState == WindowState.Maximized
                    ? WindowState.Normal
                    : WindowState.Maximized;
            }
            else
            {
                DragMove();
            }
        }

        private void MinimizeButton_Click(object sender, RoutedEventArgs e)
        {
            WindowState = WindowState.Minimized;
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void SettingsButton_Click(object sender, RoutedEventArgs e)
        {
            var settingsWindow = new SettingsWindow();
            settingsWindow.Owner = this;

            if (settingsWindow.ShowDialog() == true)
            {
                AddActivity("Settings updated", true);
                _ = ReloadSettingsAsync();
            }
        }

        private async System.Threading.Tasks.Task ReloadSettingsAsync()
        {
            try
            {
                var config = await LoadVoiceConfigAsync();
                _elevenLabsApiKey = _secureKeyManager?.LoadApiKey(
                    AICompanion.Desktop.Services.Security.SecureApiKeyManager.ElevenLabsKeyName) ?? "";
                if (string.IsNullOrEmpty(_elevenLabsApiKey))
                    _elevenLabsApiKey = config.ElevenLabsApiKey;
                _elevenLabsVoiceId = config.ElevenLabsVoiceId;

                if (_unifiedVoiceManager != null)
                {
                    var settings = new Services.Voice.VoiceSettings
                    {
                        ElevenLabsApiKey = _elevenLabsApiKey,
                        ElevenLabsVoiceId = _elevenLabsVoiceId
                    };
                    _unifiedVoiceManager.Configure(settings);
                    await _unifiedVoiceManager.InitializeAsync();
                }

                AddActivity("ElevenLabs settings reloaded", true);
            }
            catch (Exception ex)
            {
                AddActivity($"Failed to reload settings: {ex.Message}", false);
            }
        }

        private void HistoryButton_Click(object sender, RoutedEventArgs e)
        {
            var historyWindow = new HistoryWindow();
            historyWindow.Owner = this;
            historyWindow.ShowDialog();
        }

        private async void DictationButton_Click(object sender, RoutedEventArgs e)
        {
            if (_dictationService == null)
            {
                AddActivity("Dictation service not available", false);
                return;
            }

            _isDictationMode = !_isDictationMode;

            if (_isDictationMode)
            {
                var mode = _dictationService.DetectTargetApplication();
                var success = await _dictationService.StartDictationAsync(mode);

                if (success)
                {
                    var targetApp = _dictationService.TargetApplication;
                    AddActivity($"Dictation mode ON - Target: {targetApp}", true);
                    UpdateDictationButtonUI(true);
                    ShowToast($"Dictation to: {targetApp}", true);
                }
                else
                {
                    _isDictationMode = false;
                    AddActivity("Could not start dictation - open Word or Notepad first", false);
                    ShowToast("Open Word or Notepad first", false);
                }
            }
            else
            {
                _dictationService.StopDictation();
                AddActivity("Dictation mode OFF", true);
                UpdateDictationButtonUI(false);
            }
        }

        private void UpdateDictationButtonUI(bool isActive)
        {
            if (DictationButton != null)
            {
                DictationButton.Background = isActive
                    ? (WpfBrush)FindResource("PrimaryBrush")
                    : new WpfSolidColorBrush(WpfColor.FromRgb(0xEC, 0xF8, 0xF6));
            }
        }

        private async void LegacySettingsButton_Click(object sender, RoutedEventArgs e)
        {
            await ShowElevenLabsSettingsAsync();
        }

        private async System.Threading.Tasks.Task ShowElevenLabsSettingsAsync()
        {
            var inputWindow = new System.Windows.Window
            {
                Title = "ElevenLabs Configuration",
                Width = 520,
                Height = 330,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = this,
                ResizeMode = ResizeMode.NoResize
            };

            var panel = new System.Windows.Controls.StackPanel { Margin = new Thickness(15) };

            panel.Children.Add(new System.Windows.Controls.TextBlock
            {
                Text = "API Key (from ElevenLabs):",
                Margin = new Thickness(0, 0, 0, 5),
                FontWeight = FontWeights.SemiBold
            });

            var apiKeyBox = new System.Windows.Controls.TextBox
            {
                Margin = new Thickness(0, 0, 0, 15),
                Text = _elevenLabsApiKey
            };
            panel.Children.Add(apiKeyBox);

            panel.Children.Add(new System.Windows.Controls.TextBlock
            {
                Text = "Voice ID (for TTS):",
                Margin = new Thickness(0, 0, 0, 5),
                FontWeight = FontWeights.SemiBold
            });

            var voiceIdBox = new System.Windows.Controls.TextBox
            {
                Margin = new Thickness(0, 0, 0, 10),
                Text = _elevenLabsVoiceId
            };
            panel.Children.Add(voiceIdBox);

            var ttsCheck = new System.Windows.Controls.CheckBox
            {
                Content = "Enable ElevenLabs Text-to-Speech",
                IsChecked = true,
                Margin = new Thickness(0, 0, 0, 6)
            };
            panel.Children.Add(ttsCheck);

            var sttCheck = new System.Windows.Controls.CheckBox
            {
                Content = "Enable ElevenLabs Speech-to-Text",
                IsChecked = true,
                Margin = new Thickness(0, 0, 0, 12)
            };
            panel.Children.Add(sttCheck);

            panel.Children.Add(new System.Windows.Controls.TextBlock
            {
                Text = "ElevenLabs STT uses cloud recognition (Scribe).",
                Margin = new Thickness(0, 0, 0, 10),
                FontSize = 11,
                Foreground = new WpfSolidColorBrush(WpfColor.FromRgb(0x66, 0x66, 0x66))
            });

            var buttonPanel = new System.Windows.Controls.StackPanel
            {
                Orientation = System.Windows.Controls.Orientation.Horizontal,
                HorizontalAlignment = System.Windows.HorizontalAlignment.Right
            };

            var saveButton = new System.Windows.Controls.Button
            {
                Content = "Save & Test",
                Width = 100,
                Margin = new Thickness(0, 0, 10, 0),
                Padding = new Thickness(5)
            };

            var cancelButton = new System.Windows.Controls.Button
            {
                Content = "Cancel",
                Width = 80,
                Padding = new Thickness(5)
            };

            saveButton.Click += async (s, args) =>
            {
                var apiKey = apiKeyBox.Text.Trim();
                if (string.IsNullOrEmpty(apiKey))
                {
                    System.Windows.MessageBox.Show("Please enter ElevenLabs API key", "Error",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                _elevenLabsApiKey = apiKey;
                _elevenLabsVoiceId = string.IsNullOrWhiteSpace(voiceIdBox.Text)
                    ? _elevenLabsVoiceId
                    : voiceIdBox.Text.Trim();

                // Store API key in DPAPI (encrypted) — never in plain-text JSON
                _secureKeyManager?.SaveApiKey(
                    AICompanion.Desktop.Services.Security.SecureApiKeyManager.ElevenLabsKeyName,
                    _elevenLabsApiKey);

                // Save only non-sensitive settings to appsettings.json
                var config = await LoadVoiceConfigAsync();
                config.ElevenLabsApiKey = ""; // key lives in DPAPI, not here
                config.ElevenLabsVoiceId = _elevenLabsVoiceId;
                config.ElevenLabsTtsEnabled = ttsCheck.IsChecked == true;
                config.ElevenLabsSttEnabled = sttCheck.IsChecked == true;
                await SaveVoiceConfigAsync(config);

                if (_unifiedVoiceManager != null)
                {
                    _unifiedVoiceManager.Configure(new Services.Voice.VoiceSettings
                    {
                        ElevenLabsApiKey = _elevenLabsApiKey,
                        ElevenLabsVoiceId = _elevenLabsVoiceId
                    });
                    await _unifiedVoiceManager.InitializeAsync();
                }

                AddActivity("ElevenLabs settings updated.", true);
                inputWindow.Close();
            };

            cancelButton.Click += (s, args) => inputWindow.Close();

            buttonPanel.Children.Add(saveButton);
            buttonPanel.Children.Add(cancelButton);
            panel.Children.Add(buttonPanel);

            inputWindow.Content = panel;
            inputWindow.ShowDialog();
        }

        private async System.Threading.Tasks.Task<VoiceConfig> LoadVoiceConfigAsync()
        {
            var configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "appsettings.json");
            var config = new VoiceConfig();

            if (!File.Exists(configPath))
                return config;

            try
            {
                var json = await File.ReadAllTextAsync(configPath);
                using var doc = JsonDocument.Parse(json);

                if (doc.RootElement.TryGetProperty("ElevenLabs", out var elevenLabs))
                {
                    config.ElevenLabsApiKey = elevenLabs.TryGetProperty("ApiKey", out var key) ? key.GetString() ?? "" : "";
                    config.ElevenLabsVoiceId = elevenLabs.TryGetProperty("VoiceId", out var voice) ? voice.GetString() ?? config.ElevenLabsVoiceId : config.ElevenLabsVoiceId;
                    config.ElevenLabsTtsEnabled = !elevenLabs.TryGetProperty("TTSEnabled", out var tts) || tts.GetBoolean();
                    config.ElevenLabsSttEnabled = !elevenLabs.TryGetProperty("STTEnabled", out var stt) || stt.GetBoolean();
                }
            }
            catch { }

            return config;
        }

        private async System.Threading.Tasks.Task SaveVoiceConfigAsync(VoiceConfig config)
        {
            var configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "appsettings.json");

            Dictionary<string, object>? existing = null;
            if (File.Exists(configPath))
            {
                try
                {
                    var existingJson = await File.ReadAllTextAsync(configPath);
                    existing = JsonSerializer.Deserialize<Dictionary<string, object>>(existingJson);
                }
                catch { }
            }
            existing ??= new Dictionary<string, object>();

            existing["ElevenLabs"] = new
            {
                ApiKey = config.ElevenLabsApiKey,
                VoiceId = config.ElevenLabsVoiceId,
                TTSEnabled = config.ElevenLabsTtsEnabled,
                STTEnabled = config.ElevenLabsSttEnabled
            };

            var json = JsonSerializer.Serialize(existing, new JsonSerializerOptions { WriteIndented = true });
            await File.WriteAllTextAsync(configPath, json);
        }

        private class VoiceConfig
        {
            public string ElevenLabsApiKey { get; set; } = "";
            public string ElevenLabsVoiceId { get; set; } = "21m00Tcm4TlvDq8ikWAM";
            public bool ElevenLabsTtsEnabled { get; set; } = true;
            public bool ElevenLabsSttEnabled { get; set; } = true;
        }

        private void HelpButton_Click(object sender, RoutedEventArgs e)
        {
            var helpText = $@"AI Companion Help
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Current Voice Engine: ElevenLabs

🎤 ACTIVATION
• Hold the microphone button and speak
• Keyboard shortcut: Ctrl+Shift+A (when focused)

📂 LAUNCH APPLICATIONS
• ""Open Word"" / ""Open Notepad""
• ""Open Calculator"" / ""Open Browser""

📝 DOCUMENT EDITING (in Word/Notepad)
• ""Type: Hello world"" - Dictate text
• ""Select all"" - Select all text
• ""Copy"" / ""Paste"" / ""Cut"" - Clipboard
• ""Undo"" / ""Redo"" - Undo/redo actions
• ""Bold"" / ""Italic"" / ""Underline"" - Formatting

🎯 WINDOW-AWARE COMMANDS (NEW!)
• ""In Word write Hello"" - Type directly in Word
• ""In Notepad write memo"" - Type in Notepad
• ""In browser open google.com"" - Open URL
• ""В ворде напиши привет"" - Russian support

📁 FILE OPERATIONS
• ""Open file report.docx"" - Open a file
• ""Save file"" - Save (Ctrl+S)
• ""Save as report-final.docx"" - Save As
• ""New file"" / ""Create new document""

🔧 UTILITIES
• ""Search weather today"" - Web search
• ""What time is it?"" - Current time
• ""What is today's date?"" - Current date
• ""Minimize"" - Minimize windows

🎓 LEARNING
• ""Start tutorial"" - Interactive guide
• ""How do I copy text?"" - Get help
• ""What commands"" - List all commands

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Supported: English, Russian
Tip: Speak clearly and wait for recognition";

            System.Windows.MessageBox.Show(helpText, "AI Companion - Voice Commands",
                MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void MicrophoneButton_MouseDown(object sender, MouseButtonEventArgs e)
        {
            StartListening();
        }

        private void MicrophoneButton_MouseUp(object sender, MouseButtonEventArgs e)
        {
            StopListening();
        }

        private void TextCommandInput_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                SendTextCommand();
                e.Handled = true;
            }
        }

        private void TextCommandSend_Click(object sender, RoutedEventArgs e)
        {
            SendTextCommand();
        }

        private void SendTextCommand()
        {
            var text = TextCommandInput.Text?.Trim();
            if (string.IsNullOrEmpty(text)) return;

            TextCommandInput.Clear();
            DispatchInput(text, isVoice: false);
        }

        // ═══════════════════════════════════════════════════════════════════════
        // BLOCK 2 — VOICE INPUT  (ElevenLabs only)
        // ═══════════════════════════════════════════════════════════════════════

        private void StartListening()
        {
            _commandProcessor.CaptureTargetWindow();
            _dictationService?.CaptureTargetWindow();
            UpdateActiveWindowStatus(_commandProcessor.GetTargetWindowTitle(), IntPtr.Zero);

            if (_unifiedVoiceManager == null)
            {
                AddActivity("⚠️ Voice system unavailable. Check Settings → ElevenLabs API key.", false);
                RecognizedText.Text = "Voice unavailable — type command below";
                RecognizedText.FontStyle = FontStyles.Italic;
                RecognizedText.Foreground = (WpfBrush)FindResource("TextSecondaryBrush");
                UpdateStatus("No voice", false);
                return;
            }

            if (!_unifiedVoiceManager.IsInitialized)
            {
                AddActivity("⚠️ ElevenLabs not initialized. Check Settings → enter API key.", false);
                RecognizedText.Text = "ElevenLabs not ready — check Settings";
                RecognizedText.FontStyle = FontStyles.Italic;
                RecognizedText.Foreground = (WpfBrush)FindResource("TextSecondaryBrush");
                UpdateStatus("No voice", false);
                return;
            }

            _unifiedVoiceManager.StopSpeaking();
            _isListening = true;
            UpdateListeningUI(true);
            RecognizedText.Text = $"Listening ({_unifiedVoiceManager.ActiveEngine})... speak now";
            RecognizedText.FontStyle = FontStyles.Italic;
            RecognizedText.Foreground = (WpfBrush)FindResource("PrimaryBrush");
            _unifiedVoiceManager.StartListening();
            UpdateStatus("Listening", true);
        }

        private void StopListening()
        {
            if (!_isListening) return;
            _isListening = false;
            UpdateListeningUI(false);
            _unifiedVoiceManager?.StopListening();

            // ElevenLabs transcribes asynchronously: the real text arrives via CommandRecognized.
            // _lastHypothesis is always a status string ("Listening...", "Processing speech...") —
            // dispatching it would send garbage to the command pipeline. Just show "Processing…".
            _lastHypothesis = null;
            RecognizedText.Text = "Processing speech...";
            RecognizedText.FontStyle = FontStyles.Italic;
            RecognizedText.Foreground = (WpfBrush)FindResource("PrimaryBrush");
            UpdateStatus("Processing...", true);
        }

        // ═══════════════════════════════════════════════════════════════════════
        // BLOCK 3 — IBM GRANITE SMART SUGGESTIONS DIALOG
        // Generates context-aware options with confidence percentages.
        // User picks by saying "1"/"2"/"3" or clicking.
        // ═══════════════════════════════════════════════════════════════════════

        /// <summary>Option shown in the smart suggestions dialog.</summary>
        private sealed class IntentOption
        {
            public int    Number            { get; set; }
            public string Icon              { get; set; } = "•";
            public string Label             { get; set; } = "";
            public string ContextHint       { get; set; } = "";
            public int    ConfidencePercent { get; set; }
            /// <summary>true = execute locally (no Granite). false = let Granite proceed.</summary>
            public bool   IsLocal           { get; set; }
            /// <summary>When IsLocal=true: action key + optional payload separated by ":"</summary>
            public string? LocalAction      { get; set; }
        }

        /// <summary>
        /// Generates smart suggestions based on plan + current context.
        /// Returns 2-3 options ordered by confidence (descending).
        /// </summary>
        private List<IntentOption> GenerateIntentOptions(AgenticPlanDto plan, string commandText)
        {
            var lower       = commandText.ToLowerInvariant();
            var lastApp     = _commandProcessor.LastOpenedApp;
            var ctxJson     = _commandProcessor.GetSessionContextSnapshot();

            // Parse session context to get last editor process
            string lastEditorProc = "";
            try
            {
                var ctx = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, object>>(ctxJson);
                if (ctx?.TryGetValue("last_text_editor_app", out var v) == true)
                    lastEditorProc = v?.ToString() ?? "";
            }
            catch { }

            // Check session context AND running processes so options are correct even after agentic actions
            bool wordIsOpen     = lastApp.Contains("Word",    StringComparison.OrdinalIgnoreCase) || lastEditorProc.Equals("WINWORD", StringComparison.OrdinalIgnoreCase)
                                  || System.Diagnostics.Process.GetProcessesByName("WINWORD").Length > 0;
            bool notepadIsOpen  = lastApp.Contains("Notepad", StringComparison.OrdinalIgnoreCase) || lastEditorProc.Equals("notepad", StringComparison.OrdinalIgnoreCase)
                                  || System.Diagnostics.Process.GetProcessesByName("notepad").Length > 0;
            bool excelIsOpen    = lastApp.Contains("Excel",   StringComparison.OrdinalIgnoreCase) || lastEditorProc.Equals("EXCEL",   StringComparison.OrdinalIgnoreCase)
                                  || System.Diagnostics.Process.GetProcessesByName("EXCEL").Length > 0;
            bool browserIsOpen  = lastApp.Contains("Chrome",  StringComparison.OrdinalIgnoreCase) || lastApp.Contains("Edge", StringComparison.OrdinalIgnoreCase) || lastApp.Contains("Opera", StringComparison.OrdinalIgnoreCase);

            // Determine what app Granite's plan actually targets
            string planTargetApp = "";
            foreach (var step in plan.Steps ?? new())
            {
                if (step.Action is "open_app" or "focus_window" or "new_document")
                {
                    planTargetApp = (step.Target ?? "").ToLowerInvariant();
                    break;
                }
            }
            bool planTargetsWord    = planTargetApp.Contains("winword") || planTargetApp.Contains("word");
            bool planTargetsNotepad = planTargetApp.Contains("notepad") || planTargetApp.Contains("блокнот");
            bool planTargetsExcel   = planTargetApp.Contains("excel");
            bool planTargetsBrowser = planTargetApp.Contains("chrome") || planTargetApp.Contains("opera") || planTargetApp.Contains("firefox") || planTargetApp.Contains("edge");

            bool mentionsWord    = lower.Contains("word") || lower.Contains("ворд");
            bool mentionsNotepad = lower.Contains("notepad") || lower.Contains("блокнот") || lower.Contains(".txt") || lower.Contains("notebook");
            bool mentionsExcel   = lower.Contains("excel") || lower.Contains("эксель") || lower.Contains("spreadsheet");
            bool isNewDoc        = System.Text.RegularExpressions.Regex.IsMatch(lower, @"\b(new|create|make|blank|новый|создай)\b")
                                   && System.Text.RegularExpressions.Regex.IsMatch(lower, @"\b(document|doc|file|page|файл|документ)\b");
            bool isSearch        = System.Text.RegularExpressions.Regex.IsMatch(lower, @"\b(search|find|look\s+up|поиск|найди)\b");
            bool isWrite         = System.Text.RegularExpressions.Regex.IsMatch(lower, @"\b(write|type|напиши|введи)\b");
            bool isOpen          = System.Text.RegularExpressions.Regex.IsMatch(lower, @"\b(open|launch|start|открой|запусти)\b");
            bool isEssay         = lower.Contains("essay") || lower.Contains("эссе");

            // ── Document creation ────────────────────────────────────────────
            // "notebook" in English/Russian context can mean Notepad — but check the PLAN first
            bool mentionsNotebook = lower.Contains("notebook");
            if (mentionsNotebook && !planTargetsNotepad && planTargetsWord)
                mentionsNotepad = false; // Granite interpreted "notebook" as Word — trust the plan

            if (isNewDoc || mentionsWord || mentionsNotepad || mentionsExcel ||
                planTargetsWord || planTargetsNotepad || planTargetsExcel)
            {
                int wordConf    = mentionsWord    ? 92 : wordIsOpen    ? 83 : 55;
                int notepadConf = mentionsNotepad ? 92 : notepadIsOpen ? 78 : 22;
                int excelConf   = mentionsExcel   ? 92 : excelIsOpen   ? 72 : 12;

                // PLAN TARGET OVERRIDES TEXT MENTIONS — Granite analyzed full context
                if (planTargetsWord)    { wordConf    = Math.Max(wordConf,    95); notepadConf = Math.Min(notepadConf, 25); excelConf = Math.Min(excelConf, 15); }
                if (planTargetsNotepad) { notepadConf = Math.Max(notepadConf, 95); wordConf    = Math.Min(wordConf,    25); excelConf = Math.Min(excelConf, 15); }
                if (planTargetsExcel)   { excelConf   = Math.Max(excelConf,   95); wordConf    = Math.Min(wordConf,    30); }

                var opts = new List<IntentOption>
                {
                    new() { Icon="📝", Label="New Word Document",   ConfidencePercent=wordConf,    ContextHint=wordIsOpen    ? "Word is open — Ctrl+N"       : "Will open Microsoft Word",  IsLocal=false },
                    new() { Icon="📄", Label="New Notepad File",    ConfidencePercent=notepadConf, ContextHint=notepadIsOpen ? "Notepad is open — Ctrl+N"    : "Quick plain-text file",     IsLocal=true, LocalAction="new_doc_notepad" },
                    new() { Icon="📊", Label="New Excel Sheet",     ConfidencePercent=excelConf,   ContextHint=excelIsOpen   ? "Excel is open — Ctrl+N"      : "Spreadsheet document",      IsLocal=true, LocalAction="new_doc_excel"   },
                };
                opts = opts.OrderByDescending(o => o.ConfidencePercent).ToList();
                for (int i = 0; i < opts.Count; i++) opts[i].Number = i + 1;
                return opts;
            }

            // ── Web search ──────────────────────────────────────────────────
            // Skip pure-search options when the command also contains "write/type" —
            // that is a research+write request that Granite handles as a full plan.
            if (isSearch && !isWrite)
            {
                var rawQuery = System.Text.RegularExpressions.Regex.Replace(
                    lower, @"^(search\s+(for\s+)?|find\s+(out\s+)?|look\s+up\s+|поиск\s+|найди\s+)", "").Trim();
                bool isYT = lower.Contains("youtube") || lower.Contains("video") || lower.Contains("watch");

                var opts = new List<IntentOption>
                {
                    new() { Icon="🔍", Label="Google Search",  ConfidencePercent=isYT?28:70, ContextHint=$"\"{rawQuery}\" in browser",         IsLocal=false },
                    new() { Icon="▶",  Label="YouTube Search", ConfidencePercent=isYT?88:18, ContextHint=$"Find videos: \"{rawQuery}\"",        IsLocal=true, LocalAction=$"search_youtube:{rawQuery}" },
                    new() { Icon="🌐", Label="Bing Search",    ConfidencePercent=10,          ContextHint=$"Microsoft Bing: \"{rawQuery}\"",     IsLocal=true, LocalAction=$"search_bing:{rawQuery}"    },
                };
                opts = opts.OrderByDescending(o => o.ConfidencePercent).ToList();
                for (int i = 0; i < opts.Count; i++) opts[i].Number = i + 1;
                return opts;
            }

            // ── Write / type text ────────────────────────────────────────────
            if (isWrite && (wordIsOpen || notepadIsOpen || excelIsOpen || planTargetsWord || planTargetsNotepad))
            {
                // Primary option label: prefer Granite's plan target, then open-app context
                string primaryLabel;
                int    primaryConf;
                string primaryHint;
                string altWordLabel;
                string altNotepadLabel;

                if (planTargetsWord)
                {
                    primaryLabel    = wordIsOpen    ? "Type in current Word"    : "Open Word + type";
                    primaryConf     = 90;
                    primaryHint     = "Granite plan targets Word";
                    altWordLabel    = "";
                    altNotepadLabel = notepadIsOpen ? "Type in current Notepad" : "Type in Notepad";
                }
                else if (planTargetsNotepad)
                {
                    primaryLabel    = notepadIsOpen ? "Type in current Notepad" : "Open Notepad + type";
                    primaryConf     = 90;
                    primaryHint     = "Granite plan targets Notepad";
                    altWordLabel    = wordIsOpen    ? "Type in current Word"    : "Type in Word";
                    altNotepadLabel = "";
                }
                else if (wordIsOpen)
                {
                    primaryLabel    = "Type in current Word";
                    primaryConf     = 82;
                    primaryHint     = "Word is open — type in existing doc";
                    altWordLabel    = "";
                    altNotepadLabel = notepadIsOpen ? "Type in current Notepad" : "Type in Notepad";
                }
                else
                {
                    primaryLabel    = "Type in Notepad";
                    primaryConf     = 75;
                    primaryHint     = notepadIsOpen ? "Notepad is open — existing file" : "Open Notepad";
                    altWordLabel    = "Type in Word";
                    altNotepadLabel = "";
                }

                // Override hint with essay-specific description when applicable
                if (isEssay)
                {
                    if (planTargetsWord || wordIsOpen)
                        primaryHint = "Will type full essay in Word";
                    else if (planTargetsNotepad || notepadIsOpen)
                        primaryHint = "Will type essay in Notepad";
                }

                var opts = new List<IntentOption>
                {
                    new() { Icon="✍️", Label=primaryLabel, ConfidencePercent=primaryConf, ContextHint=primaryHint, IsLocal=false },
                };
                if (!string.IsNullOrEmpty(altWordLabel))
                    opts.Add(new() { Icon="📝", Label=altWordLabel,
                                     ConfidencePercent = wordIsOpen ? 70 : 45,
                                     ContextHint       = wordIsOpen ? "Word is open — type in existing doc" : "Open Microsoft Word",
                                     IsLocal=false });
                if (!string.IsNullOrEmpty(altNotepadLabel))
                    opts.Add(new() { Icon="📄", Label=altNotepadLabel,
                                     ConfidencePercent = notepadIsOpen ? 70 : 30,
                                     ContextHint       = notepadIsOpen ? "Notepad is open — existing file" : "Open Notepad",
                                     IsLocal=false });
                opts.Add(new() { Icon="📝", Label="New blank document",
                                 ConfidencePercent = 20,
                                 ContextHint       = (planTargetsWord || wordIsOpen) ? "New Word doc (Ctrl+N)" : "New Notepad file",
                                 IsLocal=true, LocalAction= (planTargetsWord || wordIsOpen) ? "new_doc_word" : "new_doc_notepad" });

                opts = opts.OrderByDescending(o => o.ConfidencePercent).ToList();
                for (int i = 0; i < opts.Count; i++) opts[i].Number = i + 1;
                return opts;
            }

            // ── Default: Granite's plan + context-driven alternatives ────────
            int graniteConf = 65;
            var planText = (plan.Reasoning ?? "").ToLowerInvariant();
            if ((planText.Contains("word") && wordIsOpen) ||
                (planText.Contains("notepad") && notepadIsOpen) ||
                (planText.Contains("excel") && excelIsOpen))
                graniteConf = 85;
            else if (wordIsOpen || notepadIsOpen || excelIsOpen || browserIsOpen)
                graniteConf = 72;

            var stepsDesc = string.Join(" → ", (plan.Steps ?? new()).Take(3).Select(s =>
                s.Action + (s.Target != null ? $"({s.Target})" : "")));

            // If plan has steps, generate a readable label
            var planLabel = plan.Steps?.Count > 0
                ? string.Join(" → ", (plan.Steps ?? new()).Take(2).Select(s =>
                    (s.Action == "open_app" || s.Action == "new_document" ? (s.Target ?? s.Action) : s.Action)))
                : (plan.Reasoning is { Length: > 50 } r ? r[..50] + "…" : (plan.Reasoning ?? "Execute plan"));

            var defaultOpts = new List<IntentOption>
            {
                new() { Number=1, Icon="🤖", Label=planLabel, ConfidencePercent=graniteConf,
                        ContextHint=string.IsNullOrEmpty(stepsDesc) ? $"{plan.TotalSteps} steps" : stepsDesc, IsLocal=false },
            };

            // Add a local alternative if we have context
            if (wordIsOpen)
                defaultOpts.Add(new() { Number=2, Icon="📝", Label="New Word document (quick)",    ConfidencePercent=30, ContextHint="Word is open — Ctrl+N",   IsLocal=true, LocalAction="new_doc_word"    });
            else if (notepadIsOpen)
                defaultOpts.Add(new() { Number=2, Icon="📄", Label="New Notepad file (quick)",     ConfidencePercent=25, ContextHint="Notepad is open — Ctrl+N", IsLocal=true, LocalAction="new_doc_notepad" });

            // Always offer "Describe screen" and "Freeform AI" as last options
            defaultOpts.Add(new() { Icon="📸", Label="Describe screen first",
                ConfidencePercent=18, ContextHint="Take screenshot → AI analyzes → then act",
                IsLocal=true, LocalAction="describe_screen_first" });
            defaultOpts.Add(new() { Icon="💡", Label="Let AI decide freely",
                ConfidencePercent=15, ContextHint="Freeform: AI picks the best approach",
                IsLocal=true, LocalAction="freeform_granite" });

            for (int i = 0; i < defaultOpts.Count; i++) defaultOpts[i].Number = i + 1;
            return defaultOpts;
        }

        private void ShowOptionsDialog(AgenticPlanDto plan)
        {
            _pendingPlan        = plan;
            _dialogCommandText  = _lastProcessedCommand ?? "";
            _currentOptions     = GenerateIntentOptions(plan, _dialogCommandText);

            // Context line — show active window + open apps
            var lastApp = _commandProcessor.LastOpenedApp;
            var openApps = new List<string>();
            foreach (var pn in new[] { "WINWORD", "notepad", "EXCEL", "msedge", "opera", "chrome" })
                if (System.Diagnostics.Process.GetProcessesByName(pn).Length > 0)
                    openApps.Add(pn == "WINWORD" ? "Word" : pn == "msedge" ? "Edge" : pn);
            GraniteContextText.Text = openApps.Count > 0
                ? $"Open: {string.Join(", ", openApps)}"
                : string.IsNullOrEmpty(lastApp) ? "No active context" : $"Context: {lastApp} is open";

            GraniteOptionsContainer.Children.Clear();
            foreach (var opt in _currentOptions)
                GraniteOptionsContainer.Children.Add(BuildOptionRow(opt));

            int topConf = _currentOptions.FirstOrDefault()?.ConfidencePercent ?? 0;
            GraniteHintText.Text = $"Say 1–{_currentOptions.Count} or YES · auto-executes in 5s ({topConf}%)";
            GranitePlanPanel.Visibility = Visibility.Visible;
            UpdateStatus("Choose option (auto in 5s)", true);

            // Countdown update in hint text
            _ = System.Threading.Tasks.Task.Run(async () =>
            {
                for (int remaining = 4; remaining >= 1; remaining--)
                {
                    await System.Threading.Tasks.Task.Delay(1000);
                    if (_planConfirmationTcs?.Task.IsCompleted == true) break;
                    int r = remaining; int cnt = _currentOptions.Count;
                    Dispatcher.Invoke(() =>
                        GraniteHintText.Text = $"Say 1–{cnt} or YES · auto-executes in {r}s ({topConf}%)");
                }
            });

            // Async: load IBM Granite alternatives and add them as extra options
            _ = LoadAlternativesAsync(_dialogCommandText);

            var topLabel = _currentOptions.FirstOrDefault()?.Label ?? "option 1";
            _ = SpeakAsync($"I suggest: {topLabel}. Say 1, 2 or 3 to choose.");
        }

        private async Task LoadAlternativesAsync(string commandText)
        {
            try
            {
                var payload = System.Text.Json.JsonSerializer.Serialize(new
                {
                    text          = commandText,
                    window_title  = _agenticService != null ? "" : "",
                    open_windows  = Array.Empty<string>()
                });
                var content  = new System.Net.Http.StringContent(payload, System.Text.Encoding.UTF8, "application/json");
                Dispatcher.Invoke(() => {
                    if (_planConfirmationTcs?.Task.IsCompleted != true)
                        GraniteHintText.Text = $"Loading AI alternatives... (say 1–{_currentOptions.Count} or YES)";
                });
                var response = await _uiHttpClient.PostAsync("http://localhost:8000/api/alternatives", content)
                                                  .ConfigureAwait(false);
                if (!response.IsSuccessStatusCode) return;

                var json = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                using var doc = System.Text.Json.JsonDocument.Parse(json);
                var alts = doc.RootElement.GetProperty("alternatives");

                int startNum = _currentOptions.Count + 1;
                var newOpts  = new List<IntentOption>();

                foreach (var alt in alts.EnumerateArray())
                {
                    if (startNum > 5) break;  // Max 5 options total
                    var label  = alt.TryGetProperty("label",      out var lv) ? lv.GetString() ?? "" : $"Option {startNum}";
                    var icon   = alt.TryGetProperty("icon",       out var iv) ? iv.GetString() ?? "🤖" : "🤖";
                    var conf   = alt.TryGetProperty("confidence", out var cv) ? cv.GetInt32()        : 60;
                    var reason = alt.TryGetProperty("reasoning",  out var rv) ? rv.GetString() ?? "" : "";

                    // Build a short description of the steps for context hint
                    var stepsHint = "";
                    if (alt.TryGetProperty("steps", out var stepsEl))
                    {
                        var actions = new List<string>();
                        foreach (var s in stepsEl.EnumerateArray())
                            if (s.TryGetProperty("action", out var av)) actions.Add(av.GetString() ?? "");
                        stepsHint = string.Join(" → ", actions.Take(3));
                    }

                    // Skip if this option is a duplicate of existing
                    if (_currentOptions.Any(o => string.Equals(o.Label, label, StringComparison.OrdinalIgnoreCase)))
                        continue;

                    newOpts.Add(new IntentOption
                    {
                        Number            = startNum++,
                        Icon              = icon,
                        Label             = label,
                        ContextHint       = string.IsNullOrEmpty(stepsHint) ? reason : stepsHint,
                        ConfidencePercent = conf,
                        IsLocal           = true,
                        LocalAction       = $"granite_plan:{commandText} [{label}]"
                    });
                }

                if (newOpts.Count == 0) return;

                // Re-sort all options by confidence and renumber
                Dispatcher.Invoke(() =>
                {
                    if (_planConfirmationTcs?.Task.IsCompleted == true) return;

                    // Add new opts to list
                    foreach (var opt in newOpts)
                        _currentOptions.Add(opt);

                    // Sort all by confidence
                    _currentOptions = _currentOptions.OrderByDescending(o => o.ConfidencePercent).ToList();
                    for (int i = 0; i < _currentOptions.Count; i++) _currentOptions[i].Number = i + 1;

                    // Rebuild the options container
                    GraniteOptionsContainer.Children.Clear();
                    foreach (var opt in _currentOptions)
                        GraniteOptionsContainer.Children.Add(BuildOptionRow(opt));

                    GraniteHintText.Text = $"Say 1–{_currentOptions.Count} or YES · auto-executes in 5s ({_currentOptions.FirstOrDefault()?.ConfidencePercent ?? 0}%)";
                });
            }
            catch { /* Alternatives are best-effort */ }
        }

        private System.Windows.Controls.Border BuildOptionRow(IntentOption opt)
        {
            // Colour: top option green, others neutral
            var bgColor = opt.Number == 1
                ? WpfColor.FromRgb(0xE8, 0xF5, 0xE9)
                : WpfColor.FromRgb(0xF5, 0xF5, 0xF5);
            var fgColor = opt.Number == 1
                ? WpfColor.FromRgb(0x1B, 0x5E, 0x20)
                : WpfColor.FromRgb(0x33, 0x33, 0x33);

            var outer = new System.Windows.Controls.Border
            {
                Margin = new Thickness(0, 0, 0, 6),
                Padding = new Thickness(10, 8, 10, 8),
                CornerRadius = new CornerRadius(8),
                Background = new WpfSolidColorBrush(bgColor),
                Cursor = System.Windows.Input.Cursors.Hand
            };

            // Click anywhere on the row to select it
            int capturedNumber = opt.Number;
            outer.MouseLeftButtonUp += (s, e) => SelectOptionByNumber(capturedNumber);

            var grid = new System.Windows.Controls.Grid();
            grid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(28) });
            grid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = GridLength.Auto });
            grid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = new GridLength(6) });

            // Number badge
            var badge = new System.Windows.Controls.Border
            {
                Width = 22, Height = 22,
                CornerRadius = new CornerRadius(11),
                Background = opt.Number == 1
                    ? new WpfSolidColorBrush(WpfColor.FromRgb(0x0F, 0x76, 0x6E))
                    : new WpfSolidColorBrush(WpfColor.FromRgb(0x99, 0x99, 0x99)),
                VerticalAlignment = VerticalAlignment.Center,
                Child = new System.Windows.Controls.TextBlock
                {
                    Text = opt.Number.ToString(),
                    HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center,
                    Foreground = WpfBrushes.White,
                    FontSize = 11,
                    FontWeight = FontWeights.Bold
                }
            };
            System.Windows.Controls.Grid.SetRow(badge, 0);
            System.Windows.Controls.Grid.SetColumn(badge, 0);
            System.Windows.Controls.Grid.SetRowSpan(badge, 2);

            // Label row
            var labelPanel = new System.Windows.Controls.StackPanel { Orientation = System.Windows.Controls.Orientation.Horizontal };
            labelPanel.Children.Add(new System.Windows.Controls.TextBlock
            {
                Text = opt.Icon + "  ",
                FontSize = 13,
                VerticalAlignment = VerticalAlignment.Center
            });
            labelPanel.Children.Add(new System.Windows.Controls.TextBlock
            {
                Text = opt.Label,
                FontSize = 13,
                FontWeight = opt.Number == 1 ? FontWeights.SemiBold : FontWeights.Normal,
                Foreground = new WpfSolidColorBrush(fgColor),
                VerticalAlignment = VerticalAlignment.Center,
                TextWrapping = TextWrapping.Wrap
            });
            System.Windows.Controls.Grid.SetRow(labelPanel, 0);
            System.Windows.Controls.Grid.SetColumn(labelPanel, 1);

            // Confidence percent (top-right)
            var confText = new System.Windows.Controls.TextBlock
            {
                Text = $"{opt.ConfidencePercent}%",
                FontSize = 12,
                FontWeight = FontWeights.Bold,
                Foreground = opt.Number == 1
                    ? new WpfSolidColorBrush(WpfColor.FromRgb(0x0F, 0x76, 0x6E))
                    : new WpfSolidColorBrush(WpfColor.FromRgb(0x88, 0x88, 0x88)),
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = System.Windows.HorizontalAlignment.Right,
                Margin = new Thickness(6, 0, 0, 0)
            };
            System.Windows.Controls.Grid.SetRow(confText, 0);
            System.Windows.Controls.Grid.SetColumn(confText, 2);

            // Context hint + confidence bar
            var hintGrid = new System.Windows.Controls.Grid();
            hintGrid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            hintGrid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(80) });

            var hint = new System.Windows.Controls.TextBlock
            {
                Text = opt.ContextHint,
                FontSize = 10,
                Foreground = new WpfSolidColorBrush(WpfColor.FromRgb(0x77, 0x77, 0x77)),
                VerticalAlignment = VerticalAlignment.Center,
                TextTrimming = TextTrimming.CharacterEllipsis
            };
            System.Windows.Controls.Grid.SetColumn(hint, 0);

            // Confidence bar
            var barOuter = new System.Windows.Controls.Border
            {
                Height = 4,
                CornerRadius = new CornerRadius(2),
                Background = new WpfSolidColorBrush(WpfColor.FromRgb(0xE0, 0xE0, 0xE0)),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(6, 0, 0, 0)
            };
            var barInner = new System.Windows.Controls.Border
            {
                Height = 4,
                CornerRadius = new CornerRadius(2),
                HorizontalAlignment = System.Windows.HorizontalAlignment.Left,
                Width = opt.ConfidencePercent * 0.80,  // 80px = 100%
                Background = opt.Number == 1
                    ? new WpfSolidColorBrush(WpfColor.FromRgb(0x0F, 0x76, 0x6E))
                    : new WpfSolidColorBrush(WpfColor.FromRgb(0xBB, 0xBB, 0xBB))
            };
            barOuter.Child = barInner;
            System.Windows.Controls.Grid.SetColumn(barOuter, 1);

            hintGrid.Children.Add(hint);
            hintGrid.Children.Add(barOuter);

            System.Windows.Controls.Grid.SetRow(hintGrid, 1);
            System.Windows.Controls.Grid.SetColumn(hintGrid, 1);
            System.Windows.Controls.Grid.SetColumnSpan(hintGrid, 2);

            grid.Children.Add(badge);
            grid.Children.Add(labelPanel);
            grid.Children.Add(confText);
            grid.Children.Add(hintGrid);
            outer.Child = grid;
            return outer;
        }

        private void HidePlanDialog()
        {
            GranitePlanPanel.Visibility = Visibility.Collapsed;
            GraniteOptionsContainer.Children.Clear();
            _pendingPlan    = null;
            _currentOptions = new();
        }

        private void GraniteExecute_Click(object sender, RoutedEventArgs e)
            => SelectOptionByNumber(1); // Execute = pick top option

        private void GraniteCancel_Click(object sender, RoutedEventArgs e)
        {
            var tcs = _planConfirmationTcs;
            _planConfirmationTcs = null;
            _pendingLocalAction  = null;
            HidePlanDialog();
            AddActivity("Cancelled.", true);
            UpdateStatus("Ready", true);
            tcs?.TrySetResult(false);
        }

        private void HandleConfirmationInput(string text)
        {
            var lower = text.ToLowerInvariant().Trim();

            // Parse number ("1"/"2"/"3"/"one"/"two"/"three"/"первый"/"второй"/"третий")
            int? pick = lower switch
            {
                "1" or "one"   or "первый" or "первое" or "option 1" or "вариант 1" => 1,
                "2" or "two"   or "второй" or "второе" or "option 2" or "вариант 2" => 2,
                "3" or "three" or "третий" or "третье" or "option 3" or "вариант 3" => 3,
                _ => null
            };
            if (pick == null && (lower.StartsWith("yes") || lower is "да" or "confirm" or "ok" or "okay" or "execute" or "go" or "do it"))
                pick = 1;

            if (pick.HasValue)
            {
                SelectOptionByNumber(pick.Value);
                return;
            }

            bool isNo = lower is "no" or "нет" or "cancel" or "stop" or "abort" or "отмена"
                        || lower.StartsWith("no ") || lower.StartsWith("нет ");
            if (isNo)
            {
                GraniteCancel_Click(this, new RoutedEventArgs());
                return;
            }

            var optCount = _currentOptions.Count;
            AddActivity($"Say 1–{optCount} to pick an option, or NO to cancel.", true);
            _ = SpeakAsync($"Say 1 through {optCount} to choose, or NO to cancel.");
        }

        private void SelectOptionByNumber(int number)
        {
            var opt = _currentOptions.FirstOrDefault(o => o.Number == number);
            if (opt == null)
            {
                AddActivity($"Option {number} not available. Say 1–{_currentOptions.Count}.", true);
                return;
            }

            var tcs = _planConfirmationTcs;
            _planConfirmationTcs = null;
            HidePlanDialog();
            AddActivity($"Option {opt.Number} selected: {opt.Label}", true);

            if (opt.IsLocal)
            {
                // Cancel Granite, run local action
                _pendingLocalAction = () => ExecuteLocalOption(opt.LocalAction);
                _planOverrideTarget = null;
                tcs?.TrySetResult(false);
            }
            else
            {
                // Infer target app from the option label so we can patch the Granite plan
                var label = opt.Label;
                _planOverrideTarget =
                    label.Contains("Word",    StringComparison.OrdinalIgnoreCase) ? "winword"  :
                    label.Contains("Notepad", StringComparison.OrdinalIgnoreCase) ? "notepad"  :
                    label.Contains("Excel",   StringComparison.OrdinalIgnoreCase) ? "excel"    :
                    null;
                // Confirm Granite proceeds
                tcs?.TrySetResult(true);
            }
        }

        private void ExecuteLocalOption(string? localAction)
        {
            if (string.IsNullOrEmpty(localAction)) { UpdateStatus("Ready", true); return; }

            // IBM Granite alternative plan: re-run with a specific approach hint
            if (localAction.StartsWith("granite_plan:"))
            {
                var altCommand = localAction["granite_plan:".Length..];
                AddActivity($"🤖 Running: {altCommand}", true);
                _ = ExecuteAgenticPlanAsync(altCommand);
                return;
            }

            if (localAction.StartsWith("search_youtube:"))
            {
                var q = localAction["search_youtube:".Length..];
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    { FileName = $"https://www.youtube.com/results?search_query={Uri.EscapeDataString(q)}", UseShellExecute = true });
                AddActivity($"▶ YouTube: {q}", true);
            }
            else if (localAction.StartsWith("search_bing:"))
            {
                var q = localAction["search_bing:".Length..];
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    { FileName = $"https://www.bing.com/search?q={Uri.EscapeDataString(q)}", UseShellExecute = true });
                AddActivity($"🌐 Bing: {q}", true);
            }
            else if (localAction == "describe_screen_first")
            {
                AddActivity("📸 Taking screenshot to analyze screen...", true);
                _ = ExecuteAgenticPlanAsync($"screenshot then {_dialogCommandText}");
            }
            else if (localAction == "freeform_granite")
            {
                AddActivity("💡 Asking AI to freely decide the best approach...", true);
                _ = ExecuteAgenticPlanAsync(_dialogCommandText);
            }
            else
            {
                var cmd = localAction switch
                {
                    "new_doc_word"    => "new word document",
                    "new_doc_notepad" => "new notebook",
                    "new_doc_excel"   => "new excel document",
                    _ => null
                };
                if (cmd != null)
                {
                    var result = _commandProcessor.ProcessCommand(cmd);
                    // Safety net: if the command still requires the agentic pipeline, honour it
                    if (result.SpeechResponse == "AGENTIC_PLAN_REQUIRED")
                        _ = ExecuteAgenticPlanAsync(cmd);
                    else
                        AddActivity(result.Description, result.Success);
                }
            }

            UpdateAvatarExpression(true);
            UpdateStatus("Done", true);
            StartResetTimer(3);
        }

        private void QuickCommand_Click(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.Button button && button.Tag is string tag)
            {
                string command = tag switch
                {
                    "word" => "WINWORD",
                    "notepad" => "notepad",
                    "calc" => "calc",
                    "browser" => "msedge",
                    "explorer" => "explorer",
                    _ => tag
                };

                try
                {
                    var friendlyName = tag switch
                    {
                        "word" => "Microsoft Word",
                        "notepad" => "Notepad",
                        "calc" => "Calculator",
                        "browser" => "Browser",
                        "explorer" => "File Explorer",
                        _ => tag
                    };

                    var existingHwnd = FindRunningWindow(command);
                    if (existingHwnd != IntPtr.Zero)
                    {
                        QuickLaunchShowWindow(existingHwnd, 9 /*SW_RESTORE*/);
                        QuickLaunchSetForeground(existingHwnd);
                        AddActivity($"Switched to {friendlyName}", true);
                    }
                    else
                    {
                        Process.Start(new ProcessStartInfo { FileName = command, UseShellExecute = true });
                        AddActivity($"Opened {friendlyName}", true);
                    }
                    UpdateAvatarExpression(true);

                    var timer = new System.Windows.Threading.DispatcherTimer
                    {
                        Interval = TimeSpan.FromSeconds(2)
                    };
                    timer.Tick += (s, args) =>
                    {
                        timer.Stop();
                        ResetAvatarExpression();
                    };
                    timer.Start();
                }
                catch (Exception ex)
                {
                    AddActivity($"Failed to open {tag}: {ex.Message}", false);
                }
            }
        }

        private void TutorialButton_Click(object sender, RoutedEventArgs e)
        {
            if (_tutorialService.IsActive)
            {
                var result = System.Windows.MessageBox.Show(
                    "Tutorial is already running. Would you like to restart it?",
                    "Tutorial",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);

                if (result == MessageBoxResult.Yes)
                {
                    _tutorialService.StopTutorial();
                    _tutorialService.StartTutorial();
                }
            }
            else
            {
                _tutorialService.StartTutorial();
            }
        }

        private void UpdateListeningUI(bool isListening)
        {
            var targetOpacity = isListening ? 1.0 : 0.0;
            var animation = new DoubleAnimation(targetOpacity, TimeSpan.FromMilliseconds(200));
            ListeningIndicator.BeginAnimation(OpacityProperty, animation);

            VoiceIndicator.Fill = isListening
                ? new WpfSolidColorBrush(WpfColor.FromRgb(0x34, 0xC7, 0x59))
                : new WpfSolidColorBrush(WpfColor.FromRgb(0xCC, 0xCC, 0xCC));
        }

        private void UpdateStatus(string status, bool isSuccess)
        {
            StatusBadgeText.Text = status;
            StatusBadge.Background = isSuccess
                ? new WpfSolidColorBrush(WpfColor.FromRgb(0x34, 0xC7, 0x59))
                : new WpfSolidColorBrush(WpfColor.FromRgb(0xFF, 0x3B, 0x30));
        }

        private void UpdateAvatarExpression(bool happy)
        {
            AvatarMouth.Data = happy
                ? System.Windows.Media.Geometry.Parse("M 35,82 Q 60,100 85,82")
                : System.Windows.Media.Geometry.Parse("M 45,88 Q 60,80 75,88");
        }

        private void ResetAvatarExpression()
        {
            AvatarMouth.Data = System.Windows.Media.Geometry.Parse("M 40,82 Q 60,95 80,82");
        }

        private void AddActivity(string message, bool isSuccess)
        {
            var item = new ActivityLogItem
            {
                Time = DateTime.Now.ToString("HH:mm:ss"),
                Message = message,
                Background = isSuccess
                    ? new WpfSolidColorBrush(WpfColor.FromRgb(0xE8, 0xF5, 0xE9))
                    : new WpfSolidColorBrush(WpfColor.FromRgb(0xFF, 0xEB, 0xEE)),
                TextColor = isSuccess
                    ? new WpfSolidColorBrush(WpfColor.FromRgb(0x2E, 0x7D, 0x32))
                    : new WpfSolidColorBrush(WpfColor.FromRgb(0xC6, 0x28, 0x28))
            };

            _activityItems.Insert(0, item);

            while (_activityItems.Count > 20)
            {
                _activityItems.RemoveAt(_activityItems.Count - 1);
            }
        }

        private System.Windows.Threading.DispatcherTimer? _toastTimer;

        private void UpdateTutorialProgress()
        {
            var current = _tutorialService.CurrentStepIndex + 1;
            var total = _tutorialService.TotalSteps;
            TutorialStepText.Text = $"Step {current}/{total}";
        }

        private void UpdateVoiceEngineStatus()
        {
            VoiceEngineStatus.Text = "ElevenLabs STT (Cloud) ✓";
        }

        private void UpdateTTSEngineStatus()
        {
            try
            {
                if (_elevenLabs != null && _elevenLabs.IsInitialized)
                {
                    TTSEngineStatus.Text = "ElevenLabs (Premium) ✓";
                    TTSStatusIndicator.Background = (WpfBrush)FindResource("SuccessBrush");
                }
                else
                {
                    TTSEngineStatus.Text = "ElevenLabs (Not Initialized)";
                    TTSStatusIndicator.Background = (WpfBrush)FindResource("ErrorBrush");
                }
            }
            catch { }
        }

        private async void TestTTS_Click(object sender, RoutedEventArgs e)
        {
            var testMessage = "Hello! This is ElevenLabs premium voice. I am ready to assist you.";
            AddActivity("Testing TTS: ElevenLabs", true);
            await SpeakAsync(testMessage);
        }

        private void UpdateActiveWindowStatus(string windowTitle, IntPtr handle)
        {
            var editorTitle = _commandProcessor.GetActiveTextEditorTitle();
            var hasEditor = editorTitle != "(no text editor)";

            if (string.IsNullOrWhiteSpace(windowTitle) || windowTitle == "(none)")
            {
                if (hasEditor)
                {
                    var truncatedEditor = editorTitle.Length > 25 ? editorTitle.Substring(0, 22) + "..." : editorTitle;
                    ActiveWindowStatus.Text = $"Editor: {truncatedEditor}";
                }
                else
                {
                    ActiveWindowStatus.Text = "Ready (click on Word/Notepad)";
                }
            }
            else
            {
                var truncatedTitle = windowTitle.Length > 30
                    ? windowTitle.Substring(0, 27) + "..."
                    : windowTitle;
                ActiveWindowStatus.Text = $"Target: {truncatedTitle}";
            }
        }

        private void ShowToast(string message, bool isSuccess)
        {
            ToastMessage.Text = message;
            ToastIcon.Text = isSuccess ? "✓" : "⚠";
            ToastBackground.Color = isSuccess
                ? WpfColor.FromRgb(0x34, 0xC7, 0x59)
                : WpfColor.FromRgb(0xFF, 0x3B, 0x30);

            _toastTimer?.Stop();

            var fadeIn = new DoubleAnimation(0, 1, TimeSpan.FromMilliseconds(200));
            ToastOverlay.BeginAnimation(OpacityProperty, fadeIn);

            _toastTimer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(2500)
            };
            _toastTimer.Tick += (s, e) =>
            {
                _toastTimer.Stop();
                var fadeOut = new DoubleAnimation(1, 0, TimeSpan.FromMilliseconds(300));
                ToastOverlay.BeginAnimation(OpacityProperty, fadeOut);
            };
            _toastTimer.Start();
        }

        [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "EnumWindows")]
        private static extern bool QuickLaunchEnumWindows(QuickLaunchEnumProc cb, IntPtr lp);
        [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "IsWindowVisible")]
        private static extern bool QuickLaunchIsVisible(IntPtr h);
        [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "GetWindowTextW", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
        private static extern int QuickLaunchGetTitle(IntPtr h, System.Text.StringBuilder sb, int n);
        [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "GetWindowThreadProcessId")]
        private static extern uint QuickLaunchGetPid(IntPtr h, out uint pid);
        [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "ShowWindow")]
        private static extern bool QuickLaunchShowWindow(IntPtr h, int cmd);
        [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "SetForegroundWindow")]
        private static extern bool QuickLaunchSetForeground(IntPtr h);
        private delegate bool QuickLaunchEnumProc(IntPtr h, IntPtr lp);

        private static IntPtr FindRunningWindow(string processName)
        {
            var stem = System.IO.Path.GetFileNameWithoutExtension(processName).ToLowerInvariant();
            IntPtr found = IntPtr.Zero;
            QuickLaunchEnumWindows((h, _) =>
            {
                if (!QuickLaunchIsVisible(h)) return true;
                var sb = new System.Text.StringBuilder(256);
                QuickLaunchGetTitle(h, sb, 256);
                var title = sb.ToString();
                if (string.IsNullOrWhiteSpace(title)) return true;
                QuickLaunchGetPid(h, out uint pid);
                try
                {
                    var proc = System.Diagnostics.Process.GetProcessById((int)pid);
                    var procName = proc.ProcessName.ToLowerInvariant();
                    // Match by process name first (handles WPS=WINWORD, Russian Notepad=Блокнот, etc.)
                    if (procName.Equals(stem, StringComparison.OrdinalIgnoreCase))
                    { found = h; return false; }
                    // Also match by title for apps whose process name differs from the search term
                    if (title.Contains(stem, StringComparison.OrdinalIgnoreCase) ||
                        (stem == "notepad" && title.Contains("Блокнот", StringComparison.OrdinalIgnoreCase)))
                    { found = h; return false; }
                }
                catch { }
                return true;
            }, IntPtr.Zero);
            return found;
        }
    }

    public class ActivityLogItem
    {
        public string Time { get; set; } = "";
        public string Message { get; set; } = "";
        public WpfBrush Background { get; set; } = WpfBrushes.White;
        public WpfBrush TextColor { get; set; } = WpfBrushes.Black;
    }
}
