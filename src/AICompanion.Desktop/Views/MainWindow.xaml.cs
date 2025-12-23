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
        private readonly ObservableCollection<ActivityLogItem> _activityItems;
        private bool _isDictationMode;
        private bool _isListening;
        private string _elevenLabsApiKey = "";
        private string _elevenLabsVoiceId = "21m00Tcm4TlvDq8ikWAM";
        private string? _lastHypothesis;
        private DateTime _lastHypothesisAt;
        private string? _lastProcessedCommand;
        private DateTime _lastProcessedAt;

        public MainWindow()
        {
            InitializeComponent();

            // Set application icon
            try
            {
                Icon = IconGenerator.CreateAppIcon(256);
            }
            catch { /* Icon generation failed, continue without */ }

            _activityItems = new ObservableCollection<ActivityLogItem>();
            ActivityLog.ItemsSource = _activityItems;

            _commandProcessor = new LocalCommandProcessor();
            _tutorialService = new TutorialService();

            // Subscribe to command processor events for toast notifications
            _commandProcessor.ActionExecuted += (s, message) =>
            {
                Dispatcher.Invoke(() =>
                {
                    ShowToast(message, true);
                    var windowInfo = _commandProcessor.GetCurrentActiveWindowInfo();
                    UpdateActiveWindowStatus(windowInfo, IntPtr.Zero);
                });
            };
            _commandProcessor.FocusError += (s, error) =>
            {
                Dispatcher.Invoke(() =>
                {
                    ShowToast(error, false);
                    UpdateActiveWindowStatus("Focus Error", IntPtr.Zero);
                });
            };
            
            // Subscribe to tutorial events for speech feedback
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

            // Try to get services from DI - ElevenLabs only
            try
            {
                _viewModel = App.ServiceProvider?.GetService<MainViewModel>();
                _elevenLabs = App.ServiceProvider?.GetService<ElevenLabsService>();
                _elevenLabsSTT = App.ServiceProvider?.GetService<ElevenLabsSpeechService>();

                if (_viewModel != null)
                {
                    DataContext = _viewModel;
                }

                // Get unified voice manager and dictation service
                _unifiedVoiceManager = App.ServiceProvider?.GetService<UnifiedVoiceManager>();
                _dictationService = App.ServiceProvider?.GetService<DictationService>();

                // Subscribe to UnifiedVoiceManager events
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
                AddActivity($"Service initialization warning: {ex.Message}", false);
            }

            Loaded += MainWindow_Loaded;
            Closing += MainWindow_Closing;

            // Welcome message
            AddActivity("AI Companion ready! Hold the microphone button to speak, or use quick commands.", true);
        }

        private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                // Initialize ViewModel (includes database initialization)
                if (_viewModel != null)
                {
                    await _viewModel.InitializeAsync();
                    AddActivity("Database initialized", true);
                }

                var config = await LoadVoiceConfigAsync();
                _elevenLabsApiKey = config.ElevenLabsApiKey;
                _elevenLabsVoiceId = config.ElevenLabsVoiceId;

                // Initialize voice through UnifiedVoiceManager - ElevenLabs only
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

            // Prevent duplicate TTS calls within short time
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

        private void OnUnifiedCommandRecognized(object? sender, Models.VoiceCommand command)
        {
            Dispatcher.Invoke(() =>
            {
                HandleRecognizedText(command.TranscribedText);
            });
        }

        private void HandleRecognizedText(string text, bool fromHypothesis = false)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return;
            }

            if (!string.IsNullOrWhiteSpace(_lastProcessedCommand)
                && string.Equals(_lastProcessedCommand, text, StringComparison.OrdinalIgnoreCase)
                && DateTime.UtcNow - _lastProcessedAt < TimeSpan.FromSeconds(2))
            {
                return;
            }

            _lastProcessedCommand = text;
            _lastProcessedAt = DateTime.UtcNow;
            _lastHypothesis = null;

            RecognizedText.Text = $"\"{text}\"";
            RecognizedText.FontStyle = FontStyles.Normal;
            RecognizedText.Foreground = new WpfSolidColorBrush(WpfColor.FromRgb(0x1D, 0x1D, 0x1F));

            var source = "[ElevenLabs]";
            var interimLabel = fromHypothesis ? " (interim)" : "";
            AddActivity($"{source}{interimLabel} You said: \"{text}\"", true);

            // SMART AUTO-DICTATION: If target window is Word/Notepad and no command prefix, auto-type
            var targetWindowTitle = _commandProcessor.GetTargetWindowTitle();
            var targetLower = targetWindowTitle.ToLowerInvariant();
            
            // Detect text editors more accurately
            var isWordDocument = targetLower.Contains("word") || targetLower.Contains(".docx") || targetLower.Contains(".doc");
            var isNotepad = targetLower.Contains("notepad") || targetLower.Contains("блокнот") || targetLower.Contains(".txt");
            var isTextEditor = isWordDocument || isNotepad;

            var textLower = text.ToLowerInvariant().Trim();
            
            // Check if speech recognition already added "type" prefix (some engines do this)
            var speechHasTypePrefix = textLower.StartsWith("type ") || 
                                     textLower.StartsWith("write ") ||
                                     textLower.StartsWith("напиши ") ||
                                     textLower.StartsWith("введи ");

            // Check if text looks like a command (starts with action verbs or command patterns)
            var looksLikeCommand = textLower.StartsWith("open") ||
                                  textLower.StartsWith("close") ||
                                  textLower.StartsWith("save") ||
                                  textLower.StartsWith("copy") ||
                                  textLower.StartsWith("paste") ||
                                  textLower.StartsWith("search") ||
                                  textLower.StartsWith("select") ||
                                  textLower.StartsWith("delete") ||
                                  textLower.StartsWith("undo") ||
                                  textLower.StartsWith("redo") ||
                                  textLower.StartsWith("start") ||
                                  textLower.StartsWith("launch") ||
                                  textLower.StartsWith("run") ||
                                  textLower.StartsWith("go to") ||
                                  textLower.StartsWith("navigate") ||
                                  textLower.StartsWith("switch") ||
                                  textLower.StartsWith("minimize") ||
                                  textLower.StartsWith("maximize") ||
                                  textLower.StartsWith("new") ||
                                  textLower.StartsWith("create") ||
                                  textLower.StartsWith("stop") ||
                                  textLower.StartsWith("play") ||
                                  textLower.StartsWith("pause") ||
                                  textLower.StartsWith("mute") ||
                                  textLower.StartsWith("volume") ||
                                  textLower.StartsWith("scroll") ||
                                  textLower.StartsWith("click") ||
                                  textLower.StartsWith("press") ||
                                  textLower.StartsWith("shut") ||
                                  textLower.StartsWith("restart") ||
                                  textLower.StartsWith("in ") ||           // "in word type...", "in notepad..."
                                  textLower.StartsWith("into ") ||         // "into word..."
                                  textLower.StartsWith("hello") ||         // greetings
                                  textLower.StartsWith("hi") ||
                                  textLower.StartsWith("hey") ||
                                  textLower.StartsWith("what") ||          // "what time", "what date"
                                  textLower.StartsWith("открой") ||
                                  textLower.StartsWith("закрой") ||
                                  textLower.StartsWith("сохрани") ||
                                  textLower.StartsWith("скопируй") ||
                                  textLower.StartsWith("вставь") ||
                                  textLower.StartsWith("найди") ||
                                  textLower.StartsWith("выдели") ||
                                  textLower.StartsWith("удали") ||
                                  textLower.StartsWith("запусти") ||
                                  textLower.StartsWith("включи") ||
                                  textLower.StartsWith("выключи") ||
                                  textLower.StartsWith("перейди") ||
                                  textLower.StartsWith("создай") ||
                                  textLower.StartsWith("новый") ||
                                  textLower.StartsWith("новая") ||
                                  textLower.StartsWith("новое") ||
                                  textLower.StartsWith("в ") ||            // "в ворде напиши..."
                                  textLower.StartsWith("привет") ||
                                  textLower.StartsWith("здравствуй") ||
                                  textLower.StartsWith("help") ||
                                  textLower.StartsWith("помощь") ||
                                  textLower.StartsWith("справка");

            // Smart decision: Auto-type if in text editor and not a command
            if (isTextEditor && !looksLikeCommand)
            {
                // If speech already has "type" prefix, use original text
                // Otherwise, auto-add "type:" prefix
                if (speechHasTypePrefix)
                {
                    AddActivity($"✍️ Dictating to {targetWindowTitle}", true);
                    ProcessLocalCommand(text); // Process as-is
                }
                else
                {
                    AddActivity($"✍️ Auto-typing to {targetWindowTitle}", true);
                    ProcessLocalCommand($"type: {text}");
                }
            }
            else
            {
                // Regular command processing
                ProcessLocalCommand(text);
            }
        }

        private void OnListeningStateChanged(object? sender, bool isListening)
        {
            Dispatcher.Invoke(() =>
            {
                _isListening = isListening;
                UpdateListeningUI(isListening);
            });
        }

        private void ProcessLocalCommand(string text)
        {
            // If tutorial is active, let it process the command first
            // If tutorial handles the command, don't process it again
            if (_tutorialService.IsActive)
            {
                var handled = _tutorialService.ProcessCommand(text);
                if (handled)
                {
                    // Tutorial handled this command - don't double-process
                    return;
                }
            }

            var result = _commandProcessor.ProcessCommand(text);

            // Handle special tutorial markers
            if (result.SpeechResponse == "TUTORIAL_START")
            {
                _tutorialService.StartTutorial();
                AddActivity("🎓 Starting interactive tutorial...", true);
                UpdateAvatarExpression(true);
                return;
            }
            else if (result.SpeechResponse == "TUTORIAL_STOP")
            {
                _tutorialService.StopTutorial();
                return;
            }
            else if (result.SpeechResponse == "TUTORIAL_SKIP")
            {
                _tutorialService.SkipStep();
                return;
            }
            else if (result.SpeechResponse == "TUTORIAL_HINT")
            {
                _tutorialService.RequestHint();
                return;
            }

            AddActivity(result.Description, result.Success);
            UpdateAvatarExpression(result.Success);

            if (!string.IsNullOrEmpty(result.SpeechResponse))
            {
                _ = SpeakAsync(result.SpeechResponse);
            }

            UpdateStatus(result.Success ? "Done" : "Error", result.Success);

            var timer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(3)
            };
            timer.Tick += (s, e) =>
            {
                timer.Stop();
                UpdateStatus("Ready", true);
                ResetAvatarExpression();
            };
            timer.Start();
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
                // Reload settings after save
                AddActivity("Settings updated", true);
                _ = ReloadSettingsAsync();
            }
        }

        private async System.Threading.Tasks.Task ReloadSettingsAsync()
        {
            try
            {
                var config = await LoadVoiceConfigAsync();
                _elevenLabsApiKey = config.ElevenLabsApiKey;
                _elevenLabsVoiceId = config.ElevenLabsVoiceId;

                // Reconfigure through UnifiedVoiceManager
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
                // Detect target application
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
            // Update button appearance based on dictation state
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
                    System.Windows.MessageBox.Show("Please enter ElevenLabs API key", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                _elevenLabsApiKey = apiKey;
                _elevenLabsVoiceId = string.IsNullOrWhiteSpace(voiceIdBox.Text) ? _elevenLabsVoiceId : voiceIdBox.Text.Trim();

                var config = await LoadVoiceConfigAsync();
                config.ElevenLabsApiKey = _elevenLabsApiKey;
                config.ElevenLabsVoiceId = _elevenLabsVoiceId;
                config.ElevenLabsTtsEnabled = ttsCheck.IsChecked == true;
                config.ElevenLabsSttEnabled = sttCheck.IsChecked == true;
                await SaveVoiceConfigAsync(config);

                // Reconfigure through UnifiedVoiceManager
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
            {
                return config;
            }

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
            catch
            {
                // ignore parsing errors, keep defaults
            }

            return config;
        }

        private async System.Threading.Tasks.Task SaveVoiceConfigAsync(VoiceConfig config)
        {
            var configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "appsettings.json");

            var jsonObject = new
            {
                ElevenLabs = new
                {
                    ApiKey = config.ElevenLabsApiKey,
                    VoiceId = config.ElevenLabsVoiceId,
                    TTSEnabled = config.ElevenLabsTtsEnabled,
                    STTEnabled = config.ElevenLabsSttEnabled
                }
            };

            var json = JsonSerializer.Serialize(jsonObject, new JsonSerializerOptions { WriteIndented = true });
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
            // CRITICAL: Capture target window BEFORE we take focus for voice recognition
            _commandProcessor.CaptureTargetWindow();
            _dictationService?.CaptureTargetWindow();
            var targetTitle = _commandProcessor.GetTargetWindowTitle();
            UpdateActiveWindowStatus(targetTitle, IntPtr.Zero);

            StartListening();
        }

        private void MicrophoneButton_MouseUp(object sender, MouseButtonEventArgs e)
        {
            StopListening();
        }

        private void StartListening()
        {
            if (_unifiedVoiceManager == null)
            {
                AddActivity("Voice system not available.", false);
                return;
            }

            // Stop speech before recording to prevent feedback
            _unifiedVoiceManager.StopSpeaking();

            if (!_unifiedVoiceManager.IsInitialized)
            {
                AddActivity("Voice engine not initialized. Check API settings.", false);
                UpdateStatus("Error", false);
                return;
            }

            _isListening = true;
            UpdateListeningUI(true);

            var engineName = _unifiedVoiceManager.ActiveEngine.ToString();
            RecognizedText.Text = $"Listening ({engineName})... speak now";
            RecognizedText.FontStyle = FontStyles.Italic;
            RecognizedText.Foreground = (WpfBrush)FindResource("PrimaryBrush");

            // Start listening through unified manager - only ONE engine will be started
            _unifiedVoiceManager.StartListening();

            UpdateStatus("Listening", true);
        }

        private void StopListening()
        {
            _isListening = false;
            UpdateListeningUI(false);

            // Stop all engines through unified manager
            _unifiedVoiceManager?.StopListening();

            if (!string.IsNullOrWhiteSpace(_lastHypothesis) && RecognizedText.Text.EndsWith("..."))
            {
                HandleRecognizedText(_lastHypothesis, fromHypothesis: true);
            }
            else if (RecognizedText.Text.StartsWith("Listening"))
            {
                RecognizedText.Text = "Press and hold the microphone button to speak...";
                RecognizedText.FontStyle = FontStyles.Italic;
                RecognizedText.Foreground = (WpfBrush)FindResource("TextSecondaryBrush");
                UpdateStatus("Ready", true);
            }
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
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = command,
                        UseShellExecute = true
                    });

                    var friendlyName = tag switch
                    {
                        "word" => "Microsoft Word",
                        "notepad" => "Notepad",
                        "calc" => "Calculator",
                        "browser" => "Browser",
                        "explorer" => "File Explorer",
                        _ => tag
                    };

                    AddActivity($"Opened {friendlyName}", true);
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
            // Start the interactive tutorial when the tutorial button is clicked
            if (_tutorialService.IsActive)
            {
                // If tutorial is already running, ask if they want to stop
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
            // Show current target and remembered text editor
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
            // Set toast content
            ToastMessage.Text = message;
            ToastIcon.Text = isSuccess ? "✓" : "⚠";
            ToastBackground.Color = isSuccess
                ? WpfColor.FromRgb(0x34, 0xC7, 0x59)  // Green
                : WpfColor.FromRgb(0xFF, 0x3B, 0x30); // Red

            // Cancel previous timer if exists
            _toastTimer?.Stop();

            // Fade in
            var fadeIn = new DoubleAnimation(0, 1, TimeSpan.FromMilliseconds(200));
            ToastOverlay.BeginAnimation(OpacityProperty, fadeIn);

            // Auto-hide after 2.5 seconds
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
    }

    public class ActivityLogItem
    {
        public string Time { get; set; } = "";
        public string Message { get; set; } = "";
        public WpfBrush Background { get; set; } = WpfBrushes.White;
        public WpfBrush TextColor { get; set; } = WpfBrushes.Black;
    }
}
