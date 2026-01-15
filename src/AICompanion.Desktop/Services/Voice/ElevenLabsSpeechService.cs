using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using NAudio.Wave;

namespace AICompanion.Desktop.Services.Voice
{
    public class ElevenLabsSpeechService : IDisposable
    {
        private readonly ILogger<ElevenLabsSpeechService>? _logger;
        private readonly HttpClient _httpClient;
        private WaveInEvent? _waveIn;
        private MemoryStream? _audioBuffer;
        private bool _isRecording;
        private string _apiKey = "64014d81b4f9dba6da3bd751557ddcb8733fe721e386ac1d9dbe49989ff906bd";
        private bool _isInitialized;

        public bool IsInitialized => _isInitialized && !string.IsNullOrEmpty(_apiKey);
        public bool IsListening => _isRecording;
        
        // CRITICAL: Flag to prevent service from processing when not the active engine
        private bool _isEnabled = false;
        public bool IsEnabled 
        { 
            get => _isEnabled; 
            set 
            { 
                _isEnabled = value;
                if (!value && _isRecording)
                {
                    StopListening();
                }
                _logger?.LogInformation("[ElevenLabs STT] Enabled = {Enabled}", value);
            } 
        }

        public event EventHandler<string>? CommandRecognized;
        public event EventHandler<bool>? ListeningStateChanged;
        public event EventHandler<string>? HypothesisGenerated;
        public event EventHandler<string>? RecognitionError;

        public ElevenLabsSpeechService(ILogger<ElevenLabsSpeechService>? logger = null)
        {
            _logger = logger;
            _httpClient = new HttpClient
            {
                BaseAddress = new Uri("https://api.elevenlabs.io/"),
                Timeout = TimeSpan.FromSeconds(30)
            };
        }

        public void Configure(string apiKey)
        {
            _apiKey = apiKey;
            _httpClient.DefaultRequestHeaders.Clear();
            _httpClient.DefaultRequestHeaders.Add("xi-api-key", _apiKey);
            _isInitialized = !string.IsNullOrEmpty(_apiKey);
            _logger?.LogInformation("[ElevenLabs STT] Configured with API key");
        }

        public async Task<bool> InitializeAsync()
        {
            if (string.IsNullOrEmpty(_apiKey))
            {
                _logger?.LogWarning("[ElevenLabs STT] No API key configured");
                return false;
            }

            try
            {
                // Test connection
                var response = await _httpClient.GetAsync("v1/user");
                if (response.IsSuccessStatusCode)
                {
                    _isInitialized = true;
                    _logger?.LogInformation("[ElevenLabs STT] Initialized successfully");
                    return true;
                }
                else
                {
                    var error = await response.Content.ReadAsStringAsync();
                    _logger?.LogError("[ElevenLabs STT] Init failed: {Error}", error);
                    return false;
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "[ElevenLabs STT] Init exception");
                return false;
            }
        }

        public void StartListening()
        {
            // CRITICAL: Don't start if not enabled (another engine is active)
            if (!_isEnabled)
            {
                _logger?.LogWarning("[ElevenLabs STT] StartListening blocked - service is DISABLED");
                return;
            }

            if (_isRecording) return;
            if (!_isInitialized)
            {
                RecognitionError?.Invoke(this, "ElevenLabs not initialized");
                return;
            }

            try
            {
                _audioBuffer = new MemoryStream();
                _waveIn = new WaveInEvent
                {
                    WaveFormat = new WaveFormat(16000, 16, 1), // 16kHz, 16-bit, mono
                    BufferMilliseconds = 100
                };

                _waveIn.DataAvailable += (s, e) =>
                {
                    _audioBuffer?.Write(e.Buffer, 0, e.BytesRecorded);
                };

                _waveIn.RecordingStopped += async (s, e) =>
                {
                    await ProcessRecordingAsync();
                };

                _waveIn.StartRecording();
                _isRecording = true;
                ListeningStateChanged?.Invoke(this, true);
                HypothesisGenerated?.Invoke(this, "Listening (ElevenLabs)...");
                _logger?.LogInformation("[ElevenLabs STT] Started listening");
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "[ElevenLabs STT] Failed to start recording");
                RecognitionError?.Invoke(this, $"Microphone error: {ex.Message}");
            }
        }

        public void StopListening()
        {
            if (!_isRecording) return;

            try
            {
                _waveIn?.StopRecording();
                _isRecording = false;
                ListeningStateChanged?.Invoke(this, false);
                _logger?.LogInformation("[ElevenLabs STT] Stopped listening");
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "[ElevenLabs STT] Failed to stop recording");
            }
        }

        private async Task ProcessRecordingAsync()
        {
            if (_audioBuffer == null || _audioBuffer.Length < 1000)
            {
                _logger?.LogWarning("[ElevenLabs STT] Audio too short");
                return;
            }

            try
            {
                HypothesisGenerated?.Invoke(this, "Processing speech...");

                // Convert to WAV format
                var wavData = ConvertToWav(_audioBuffer.ToArray());

                // Send to ElevenLabs Speech-to-Text API
                using var content = new MultipartFormDataContent();
                var audioContent = new ByteArrayContent(wavData);
                audioContent.Headers.ContentType = new MediaTypeHeaderValue("audio/wav");
                content.Add(audioContent, "file", "recording.wav");

                // ElevenLabs uses Scribe model for STT
                content.Add(new StringContent("scribe_v2"), "model_id");
                var response = await _httpClient.PostAsync("v1/speech-to-text", content);

                if (response.IsSuccessStatusCode)
                {
                    var json = await response.Content.ReadAsStringAsync();
                    var result = JsonSerializer.Deserialize<ElevenLabsSTTResponse>(json);
                    
                    if (!string.IsNullOrWhiteSpace(result?.text))
                    {
                        _logger?.LogInformation("[ElevenLabs STT] Recognized: {Text}", result.text);
                        CommandRecognized?.Invoke(this, result.text);
                    }
                    else
                    {
                        HypothesisGenerated?.Invoke(this, "No speech detected");
                    }
                }
                else
                {
                    var error = await response.Content.ReadAsStringAsync();
                    _logger?.LogError("[ElevenLabs STT] API error: {Error}", error);
                    RecognitionError?.Invoke(this, $"Recognition failed: {response.StatusCode}. {error}");
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "[ElevenLabs STT] Processing error");
                RecognitionError?.Invoke(this, $"Processing error: {ex.Message}");
            }
            finally
            {
                _audioBuffer?.Dispose();
                _audioBuffer = null;
                _waveIn?.Dispose();
                _waveIn = null;
            }
        }

        private byte[] ConvertToWav(byte[] rawAudio)
        {
            using var ms = new MemoryStream();
            using var writer = new BinaryWriter(ms);

            // WAV header (ASCII bytes)
            var sampleRate = 16000;
            var channels = 1;
            var bitsPerSample = 16;
            var byteRate = sampleRate * channels * bitsPerSample / 8;
            var blockAlign = channels * bitsPerSample / 8;

            writer.Write(new byte[] { (byte)'R', (byte)'I', (byte)'F', (byte)'F' });
            writer.Write(36 + rawAudio.Length);
            writer.Write(new byte[] { (byte)'W', (byte)'A', (byte)'V', (byte)'E' });
            writer.Write(new byte[] { (byte)'f', (byte)'m', (byte)'t', (byte)' ' });
            writer.Write(16); // Subchunk1Size
            writer.Write((short)1); // AudioFormat (PCM)
            writer.Write((short)channels);
            writer.Write(sampleRate);
            writer.Write(byteRate);
            writer.Write((short)blockAlign);
            writer.Write((short)bitsPerSample);
            writer.Write(new byte[] { (byte)'d', (byte)'a', (byte)'t', (byte)'a' });
            writer.Write(rawAudio.Length);
            writer.Write(rawAudio);

            return ms.ToArray();
        }

        public void Dispose()
        {
            StopListening();
            _waveIn?.Dispose();
            _audioBuffer?.Dispose();
            _httpClient.Dispose();
        }

        private class ElevenLabsSTTResponse
        {
            public string? text { get; set; }
            public string? language_code { get; set; }
        }
    }
}
