using System;
using System.IO;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using NAudio.Wave;
using Whisper.net;
using AICompanion.Desktop.Models;

namespace AICompanion.Desktop.Services.Voice
{
    /// <summary>
    /// Unified Voice Manager - uses ElevenLabs for both STT and TTS.
    /// Falls back to local Whisper.net STT when ElevenLabs is unreachable (offline mode).
    /// NFR: Availability — ensures voice recognition works without internet.
    /// </summary>
    public class UnifiedVoiceManager : IDisposable
    {
        private readonly ILogger<UnifiedVoiceManager>? _logger;

        // ElevenLabs services (cloud)
        private readonly ElevenLabsService _elevenLabsTTS;
        private readonly ElevenLabsSpeechService _elevenLabsSTT;

        // Whisper.net offline fallback (local)
        private WhisperProcessor? _whisperProcessor;
        private WhisperFactory? _whisperFactory;
        private bool _isWhisperInitialized;
        private WaveInEvent? _whisperWaveIn;
        private MemoryStream? _whisperAudioBuffer;
        private bool _isWhisperRecording;

        private const string WhisperModelFileName = "ggml-tiny.en.bin";
        private const int HealthCheckTimeoutMs = 2000;

        private bool _isInitialized;
        private bool _isListening;
        private bool _isDisposed;
        private bool _isSpeaking;
        private string _activeEngine = "ElevenLabs";

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
        public string ActiveEngine => _activeEngine;

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
            _logger?.LogInformation("[UnifiedVoice] Initializing — performing ElevenLabs health check...");

            // Stop any previous listening
            StopListening();

            bool cloudAvailable = await PerformElevenLabsHealthCheckAsync();

            if (cloudAvailable)
            {
                // Cloud path: ElevenLabs is reachable
                _activeEngine = "ElevenLabs";
                _elevenLabsTTS.IsEnabled = true;
                _elevenLabsSTT.IsEnabled = true;
                _isInitialized = true;
                _logger?.LogInformation("[UnifiedVoice] ElevenLabs online — cloud STT/TTS active");
            }
            else
            {
                // Offline fallback: initialize Whisper.net for local STT
                _logger?.LogWarning("[UnifiedVoice] ElevenLabs unreachable — switching to Whisper.net offline STT");
                bool whisperOk = await InitializeWhisperAsync();

                if (whisperOk)
                {
                    _activeEngine = "Whisper";
                    _isWhisperInitialized = true;
                    _isInitialized = true;
                    _logger?.LogInformation("[UnifiedVoice] Whisper.net initialized — offline STT active");
                }
                else
                {
                    _activeEngine = "None";
                    _isInitialized = false;
                    _logger?.LogError("[UnifiedVoice] Both ElevenLabs and Whisper.net failed — no STT available");
                }
            }

            return _isInitialized;
        }

        /// <summary>
        /// Health check: pings ElevenLabs API with a 2-second timeout.
        /// Returns false on timeout, network error, or non-success HTTP status.
        /// </summary>
        private async Task<bool> PerformElevenLabsHealthCheckAsync()
        {
            try
            {
                using var cts = new CancellationTokenSource(TimeSpan.FromMilliseconds(HealthCheckTimeoutMs));
                using var httpClient = new HttpClient { Timeout = TimeSpan.FromMilliseconds(HealthCheckTimeoutMs) };
                httpClient.DefaultRequestHeaders.Add("xi-api-key", _settings.ElevenLabsApiKey);

                var response = await httpClient.GetAsync("https://api.elevenlabs.io/v1/user", cts.Token);
                bool ok = response.IsSuccessStatusCode;
                _logger?.LogInformation("[UnifiedVoice] ElevenLabs health check: {Status} ({Code})",
                    ok ? "OK" : "FAILED", response.StatusCode);
                return ok;
            }
            catch (TaskCanceledException)
            {
                _logger?.LogWarning("[UnifiedVoice] ElevenLabs health check timed out (>{Timeout}ms)", HealthCheckTimeoutMs);
                return false;
            }
            catch (HttpRequestException ex)
            {
                _logger?.LogWarning(ex, "[UnifiedVoice] ElevenLabs health check network error");
                return false;
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "[UnifiedVoice] ElevenLabs health check unexpected error");
                return false;
            }
        }

        /// <summary>
        /// Initializes Whisper.net with the ggml-tiny.en.bin model for offline STT.
        /// Model file is expected next to the executable.
        /// </summary>
        private async Task<bool> InitializeWhisperAsync()
        {
            try
            {
                var modelPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, WhisperModelFileName);

                if (!File.Exists(modelPath))
                {
                    _logger?.LogWarning("[UnifiedVoice] Whisper model not found at {Path} — attempting to use bundled model", modelPath);
                    // Try alternative location (project root)
                    var altPath = Path.Combine(Directory.GetCurrentDirectory(), WhisperModelFileName);
                    if (File.Exists(altPath))
                        modelPath = altPath;
                    else
                    {
                        _logger?.LogError("[UnifiedVoice] Whisper model '{Model}' not found. Download from https://huggingface.co/ggerganov/whisper.cpp", WhisperModelFileName);
                        return false;
                    }
                }

                _whisperFactory = WhisperFactory.FromPath(modelPath);
                _whisperProcessor = _whisperFactory.CreateBuilder()
                    .WithLanguage("en")
                    .Build();

                _logger?.LogInformation("[UnifiedVoice] Whisper.net loaded model: {Model}", modelPath);
                return await Task.FromResult(true);
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "[UnifiedVoice] Failed to initialize Whisper.net");
                return false;
            }
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
                RecognitionError?.Invoke(this, "Voice engine not initialized. Check Settings or internet connection.");
                return;
            }

            if (_activeEngine == "Whisper")
            {
                _logger?.LogInformation("[UnifiedVoice] Starting Whisper.net offline listening");
                StartWhisperListening();
            }
            else
            {
                _logger?.LogInformation("[UnifiedVoice] Starting ElevenLabs listening");
                _elevenLabsSTT.IsEnabled = true;
                _elevenLabsSTT.StartListening();
            }

            _isListening = true;
            ListeningStateChanged?.Invoke(this, true);
        }

        public void StopListening()
        {
            if (!_isListening)
                return;

            _logger?.LogInformation("[UnifiedVoice] Stopping listening ({Engine})", _activeEngine);

            try
            {
                if (_activeEngine == "Whisper")
                    StopWhisperListening();
                else
                    _elevenLabsSTT.StopListening();
            }
            catch { }

            _isListening = false;
            ListeningStateChanged?.Invoke(this, false);
        }

        // ── Whisper.net local recording ──

        private void StartWhisperListening()
        {
            try
            {
                _whisperAudioBuffer = new MemoryStream();
                _whisperWaveIn = new WaveInEvent
                {
                    WaveFormat = new WaveFormat(16000, 16, 1),
                    BufferMilliseconds = 100
                };

                _whisperWaveIn.DataAvailable += (s, e) =>
                {
                    _whisperAudioBuffer?.Write(e.Buffer, 0, e.BytesRecorded);
                };

                _whisperWaveIn.RecordingStopped += async (s, e) =>
                {
                    await ProcessWhisperRecordingAsync();
                };

                _whisperWaveIn.StartRecording();
                _isWhisperRecording = true;
                HypothesisGenerated?.Invoke(this, "Listening (Whisper offline)...");
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "[UnifiedVoice] Whisper microphone start failed");
                RecognitionError?.Invoke(this, $"Microphone error: {ex.Message}");
            }
        }

        private void StopWhisperListening()
        {
            if (!_isWhisperRecording) return;
            _whisperWaveIn?.StopRecording();
            _isWhisperRecording = false;
        }

        private async Task ProcessWhisperRecordingAsync()
        {
            if (_whisperAudioBuffer == null || _whisperAudioBuffer.Length < 1000 || _whisperProcessor == null)
                return;

            try
            {
                HypothesisGenerated?.Invoke(this, "Processing speech (offline)...");

                // Convert raw PCM to float samples for Whisper
                var rawBytes = _whisperAudioBuffer.ToArray();
                var samples = new float[rawBytes.Length / 2];
                for (int i = 0; i < samples.Length; i++)
                {
                    samples[i] = BitConverter.ToInt16(rawBytes, i * 2) / 32768f;
                }

                // Run Whisper inference
                var fullText = string.Empty;
                await foreach (var segment in _whisperProcessor.ProcessAsync(new MemoryStream(rawBytes)))
                {
                    fullText += segment.Text;
                }

                fullText = fullText.Trim();

                if (!string.IsNullOrWhiteSpace(fullText))
                {
                    _logger?.LogInformation("[UnifiedVoice] Whisper recognized: {Text}", fullText);
                    TextRecognized?.Invoke(this, fullText);
                    CommandRecognized?.Invoke(this, new VoiceCommand
                    {
                        TranscribedText = fullText,
                        RecognitionConfidence = 0.85f,
                        CapturedAt = DateTime.UtcNow
                    });
                }
                else
                {
                    HypothesisGenerated?.Invoke(this, "No speech detected (offline)");
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "[UnifiedVoice] Whisper processing error");
                RecognitionError?.Invoke(this, $"Offline recognition error: {ex.Message}");
            }
            finally
            {
                _whisperAudioBuffer?.Dispose();
                _whisperAudioBuffer = null;
                _whisperWaveIn?.Dispose();
                _whisperWaveIn = null;
            }
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

            // Dispose Whisper resources
            _whisperProcessor?.Dispose();
            _whisperFactory?.Dispose();
            _whisperWaveIn?.Dispose();
            _whisperAudioBuffer?.Dispose();

            _isDisposed = true;
            _logger?.LogInformation("[UnifiedVoice] Disposed (engine was {Engine})", _activeEngine);
        }
    }

    public class VoiceSettings
    {
        /// <summary>Loaded at runtime from DPAPI (SecureApiKeyManager). Never hard-code here.</summary>
        public string ElevenLabsApiKey { get; set; } = "";
        public string ElevenLabsVoiceId { get; set; } = "JBFqnCBsd6RMkjVDRZzb";
    }
}
