using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using NAudio.Wave;

namespace AICompanion.Desktop.Services.Voice
{
    public class ElevenLabsService : IDisposable
    {
        private readonly ILogger<ElevenLabsService>? _logger;
        private readonly HttpClient _httpClient;
        private WaveOutEvent? _waveOut;
        private MemoryStream? _audioStream;
        
        private string _apiKey = "64014d81b4f9dba6da3bd751557ddcb8733fe721e386ac1d9dbe49989ff906bd";
        private string _voiceId = "JBFqnCBsd6RMkjVDRZzb"; // Default voice
        private string _modelId = "eleven_monolingual_v1";
        private bool _isInitialized = false;

        public bool IsInitialized => _isInitialized;
        public bool IsEnabled { get; set; } = false;

        public event EventHandler? SpeakingStarted;
        public event EventHandler? SpeakingCompleted;
        public event EventHandler<string>? Error;

        // Available voices
        public static readonly (string Id, string Name)[] AvailableVoices = new[]
        {
            ("21m00Tcm4TlvDq8ikWAM", "Rachel (Female, Calm)"),
            ("EXAVITQu4vr4xnSDxMaL", "Bella (Female, Soft)"),
            ("ErXwobaYiN019PkySvjV", "Antoni (Male, Warm)"),
            ("VR6AewLTigWG4xSOukaG", "Arnold (Male, Deep)"),
            ("pNInz6obpgDQGcFmaJgB", "Adam (Male, Clear)"),
            ("yoZ06aMxZJJ28mfd3POQ", "Sam (Male, Raspy)"),
        };

        public ElevenLabsService(ILogger<ElevenLabsService>? logger = null)
        {
            _logger = logger;
            _httpClient = new HttpClient
            {
                BaseAddress = new Uri("https://api.elevenlabs.io/"),
                Timeout = TimeSpan.FromSeconds(30)
            };
        }

        public void Configure(string apiKey, string? voiceId = null, string? modelId = null)
        {
            _apiKey = apiKey;
            if (!string.IsNullOrEmpty(voiceId)) _voiceId = voiceId;
            if (!string.IsNullOrEmpty(modelId)) _modelId = modelId;
            
            _httpClient.DefaultRequestHeaders.Clear();
            _httpClient.DefaultRequestHeaders.Add("xi-api-key", _apiKey);
            
            _isInitialized = !string.IsNullOrEmpty(_apiKey);
            _logger?.LogInformation("[ElevenLabs] Configured with voice: {VoiceId}", _voiceId);
        }

        public async Task<bool> TestConnectionAsync()
        {
            if (string.IsNullOrEmpty(_apiKey))
            {
                _logger?.LogWarning("[ElevenLabs] No API key configured");
                return false;
            }

            try
            {
                var response = await _httpClient.GetAsync("v1/user");
                if (response.IsSuccessStatusCode)
                {
                    _logger?.LogInformation("[ElevenLabs] Connection test successful");
                    _isInitialized = true;
                    return true;
                }
                else
                {
                    var error = await response.Content.ReadAsStringAsync();
                    _logger?.LogError("[ElevenLabs] Connection test failed: {Error}", error);
                    Error?.Invoke(this, $"API Error: {response.StatusCode}");
                    return false;
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "[ElevenLabs] Connection test exception");
                Error?.Invoke(this, $"Connection failed: {ex.Message}");
                return false;
            }
        }

        public async Task SpeakAsync(string text)
        {
            if (!_isInitialized || string.IsNullOrEmpty(_apiKey))
            {
                _logger?.LogWarning("[ElevenLabs] Not initialized, cannot speak");
                Error?.Invoke(this, "ElevenLabs not configured");
                return;
            }

            try
            {
                StopSpeaking();
                SpeakingStarted?.Invoke(this, EventArgs.Empty);

                var requestBody = new
                {
                    text = text,
                    model_id = _modelId,
                    voice_settings = new
                    {
                        stability = 0.5,
                        similarity_boost = 0.75,
                        style = 0.0,
                        use_speaker_boost = true
                    }
                };

                var json = JsonSerializer.Serialize(requestBody);
                var content = new StringContent(json, Encoding.UTF8, "application/json");
                content.Headers.ContentType = new MediaTypeHeaderValue("application/json");

                var response = await _httpClient.PostAsync($"v1/text-to-speech/{_voiceId}", content);

                if (response.IsSuccessStatusCode)
                {
                    var audioBytes = await response.Content.ReadAsByteArrayAsync();
                    await PlayAudioAsync(audioBytes);
                    _logger?.LogInformation("[ElevenLabs] Successfully spoke: {Text}", text.Substring(0, Math.Min(50, text.Length)));
                }
                else
                {
                    var error = await response.Content.ReadAsStringAsync();
                    _logger?.LogError("[ElevenLabs] TTS failed: {Error}", error);
                    Error?.Invoke(this, $"TTS Error: {response.StatusCode}");
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "[ElevenLabs] Speak exception");
                Error?.Invoke(this, $"Speak failed: {ex.Message}");
            }
            finally
            {
                SpeakingCompleted?.Invoke(this, EventArgs.Empty);
            }
        }

        private async Task PlayAudioAsync(byte[] audioBytes)
        {
            try
            {
                _audioStream = new MemoryStream(audioBytes);
                
                using var mp3Reader = new Mp3FileReader(_audioStream);
                _waveOut = new WaveOutEvent();
                _waveOut.Init(mp3Reader);
                
                var tcs = new TaskCompletionSource<bool>();
                _waveOut.PlaybackStopped += (s, e) => tcs.TrySetResult(true);
                
                _waveOut.Play();
                await tcs.Task;
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "[ElevenLabs] Audio playback error");
                throw;
            }
        }

        public void StopSpeaking()
        {
            try
            {
                _waveOut?.Stop();
                _waveOut?.Dispose();
                _waveOut = null;
                _audioStream?.Dispose();
                _audioStream = null;
            }
            catch { }
        }

        public void SetVoice(string voiceId)
        {
            _voiceId = voiceId;
            _logger?.LogInformation("[ElevenLabs] Voice changed to: {VoiceId}", voiceId);
        }

        public string GetCurrentVoiceId() => _voiceId;

        public void Dispose()
        {
            StopSpeaking();
            _httpClient.Dispose();
        }
    }
}
