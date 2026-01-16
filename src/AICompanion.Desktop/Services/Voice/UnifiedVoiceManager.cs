using System;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using AICompanion.Desktop.Models;

namespace AICompanion.Desktop.Services.Voice
{
    /// <summary>
    /// Unified Voice Manager - uses ElevenLabs for both STT and TTS.
    /// Simplified version - no Windows or Watson fallback.
    /// </summary>
    public class UnifiedVoiceManager : IDisposable
    {
        private readonly ILogger<UnifiedVoiceManager>? _logger;

        // ElevenLabs services only
        private readonly ElevenLabsService _elevenLabsTTS;
        private readonly ElevenLabsSpeechService _elevenLabsSTT;

        private bool _isInitialized;
        private bool _isListening;
        private bool _isDisposed;
        private bool _isSpeaking;

        // Configuration
        private VoiceSettings _settings = new();

        /// <summary>
        /// Returns true when TTS is playing, blocking recording to prevent feedback
        /// </summary>
        public bool IsRecordingBlocked => _isSpeaking;

        public event EventHandler<VoiceCommand>? CommandRecognized;
        public event EventHandler<string>? TextRecognized;
        public event EventHandler<bool>? ListeningStateChanged;
        public event EventHandler<string>? HypothesisGenerated;
        public event EventHandler<string>? RecognitionError;
        public event EventHandler? SpeakingStarted;
        public event EventHandler? SpeakingCompleted;

        public bool IsInitialized => _isInitialized;
        public bool IsListening => _isListening;
        public string ActiveEngine => "ElevenLabs";

        public UnifiedVoiceManager(
            ILogger<UnifiedVoiceManager>? logger,
            ElevenLabsService elevenLabsTTS,
            ElevenLabsSpeechService elevenLabsSTT)
        {
            _logger = logger;
            _elevenLabsTTS = elevenLabsTTS;
            _elevenLabsSTT = elevenLabsSTT;

            _logger?.LogInformation("[UnifiedVoice] Created - ElevenLabs only mode");

            // Subscribe to ElevenLabs events
            SubscribeToEvents();
        }

        private void SubscribeToEvents()
        {
            // ElevenLabs STT
            _elevenLabsSTT.CommandRecognized += (s, text) =>
            {
                TextRecognized?.Invoke(this, text);
                CommandRecognized?.Invoke(this, new VoiceCommand
                {
                    TranscribedText = text,
                    RecognitionConfidence = 0.9f,
                    CapturedAt = DateTime.UtcNow
                });
            };

            _elevenLabsSTT.ListeningStateChanged += (s, state) =>
            {
                ListeningStateChanged?.Invoke(this, state);
            };

            _elevenLabsSTT.HypothesisGenerated += (s, text) =>
            {
                HypothesisGenerated?.Invoke(this, text);
            };

            _elevenLabsSTT.RecognitionError += (s, err) =>
            {
                RecognitionError?.Invoke(this, err);
            };

            // ElevenLabs TTS
            _elevenLabsTTS.SpeakingStarted += (s, e) =>
            {
                SpeakingStarted?.Invoke(this, e);
            };

            _elevenLabsTTS.SpeakingCompleted += (s, e) =>
            {
                SpeakingCompleted?.Invoke(this, e);
            };
        }

        public void Configure(VoiceSettings settings)
        {
            _settings = settings;
            _logger?.LogInformation("[UnifiedVoice] Configuring ElevenLabs with API key: {HasKey}",
                !string.IsNullOrEmpty(settings.ElevenLabsApiKey));

            if (!string.IsNullOrEmpty(settings.ElevenLabsApiKey))
            {
                _elevenLabsTTS.Configure(settings.ElevenLabsApiKey, settings.ElevenLabsVoiceId);
                _elevenLabsSTT.Configure(settings.ElevenLabsApiKey);
            }
        }

        public async Task<bool> InitializeAsync()
        {
            _logger?.LogInformation("[UnifiedVoice] Initializing ElevenLabs...");

            // Stop any previous listening
            StopListening();

            bool initialized = false;

            try
            {
                var sttInit = await _elevenLabsSTT.InitializeAsync();
                var ttsInit = await _elevenLabsTTS.TestConnectionAsync();

                initialized = sttInit && ttsInit;

                if (initialized)
                {
                    _elevenLabsTTS.IsEnabled = true;
                    _elevenLabsSTT.IsEnabled = true;
                    _logger?.LogInformation("[UnifiedVoice] ElevenLabs initialized successfully (STT={Stt}, TTS={Tts})", sttInit, ttsInit);
                }
                else
                {
                    _logger?.LogError("[UnifiedVoice] ElevenLabs initialization failed (STT={Stt}, TTS={Tts})", sttInit, ttsInit);
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "[UnifiedVoice] ElevenLabs initialization error");
            }

            _isInitialized = initialized;
            return initialized;
        }

        public void StartListening()
        {
            if (_isListening)
            {
                _logger?.LogDebug("[UnifiedVoice] Already listening");
                return;
            }

            if (!_isInitialized)
            {
                _logger?.LogWarning("[UnifiedVoice] Cannot start - not initialized");
                RecognitionError?.Invoke(this, "ElevenLabs not initialized. Check API key in Settings.");
                return;
            }

            _logger?.LogInformation("[UnifiedVoice] Starting ElevenLabs listening");

            _elevenLabsSTT.IsEnabled = true;
            _elevenLabsSTT.StartListening();

            _isListening = true;
            ListeningStateChanged?.Invoke(this, true);
        }

        public void StopListening()
        {
            if (!_isListening)
                return;

            _logger?.LogInformation("[UnifiedVoice] Stopping listening");

            try { _elevenLabsSTT.StopListening(); } catch { }

            _isListening = false;
            ListeningStateChanged?.Invoke(this, false);
        }

        public async Task SpeakAsync(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return;

            StopSpeaking();
            await Task.Delay(50);

            _isSpeaking = true;
            SpeakingStarted?.Invoke(this, EventArgs.Empty);

            _logger?.LogInformation("[UnifiedVoice] Speaking via ElevenLabs: {Text}",
                text.Length > 50 ? text.Substring(0, 50) + "..." : text);

            try
            {
                if (_elevenLabsTTS != null && _elevenLabsTTS.IsInitialized)
                {
                    await _elevenLabsTTS.SpeakAsync(text);
                }
                else
                {
                    _logger?.LogWarning("[UnifiedVoice] ElevenLabs TTS not ready");
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "[UnifiedVoice] TTS error");
            }
            finally
            {
                _isSpeaking = false;
                SpeakingCompleted?.Invoke(this, EventArgs.Empty);
            }
        }

        public void StopSpeaking()
        {
            try { _elevenLabsTTS?.StopSpeaking(); } catch { }
        }

        public void Dispose()
        {
            if (_isDisposed) return;

            StopListening();
            StopSpeaking();

            _elevenLabsTTS?.Dispose();
            _elevenLabsSTT?.Dispose();

            _isDisposed = true;
            _logger?.LogInformation("[UnifiedVoice] Disposed");
        }
    }

    public class VoiceSettings
    {
        // ElevenLabs settings - HARDCODED for demo presentation
        public string ElevenLabsApiKey { get; set; } = "64014d81b4f9dba6da3bd751557ddcb8733fe721e386ac1d9dbe49989ff906bd";
        public string ElevenLabsVoiceId { get; set; } = "JBFqnCBsd6RMkjVDRZzb";
    }
}
