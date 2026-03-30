using System;
using System.IO;
using System.Net.Http;
using System.Net.WebSockets;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using NAudio.Wave;

namespace AICompanion.Desktop.Services.Voice
{
    /// <summary>
    /// Real-time speech-to-text via ElevenLabs WebSocket API (scribe_v2_realtime).
    /// Streams PCM audio chunks live; receives partial and committed transcripts.
    /// </summary>
    public class ElevenLabsSpeechService : IDisposable
    {
        private readonly ILogger<ElevenLabsSpeechService>? _logger;

        private string _apiKey = "";   // Set via Configure() — never hard-coded
        private bool _isInitialized;
        private bool _isEnabled;
        private bool _isRecording;
        private bool _isDisposed;

        // Active WebSocket session
        private ClientWebSocket? _ws;
        private WaveInEvent? _waveIn;
        private CancellationTokenSource? _sessionCts;
        private TaskCompletionSource<bool>? _committedTcs;

        // Serialises WebSocket sends: DataAvailable fires on the audio thread
        private readonly SemaphoreSlim _sendLock = new(1, 1);

        private const string WssBase =
            "wss://api.elevenlabs.io/v1/speech-to-text/realtime" +
            "?model_id=scribe_v2_realtime" +
            "&commit_strategy=manual" +
            "&audio_format=pcm_16000";

        // ── Public API ─────────────────────────────────────────────────────────

        public bool IsInitialized => _isInitialized && !string.IsNullOrEmpty(_apiKey);
        public bool IsListening   => _isRecording;

        public bool IsEnabled
        {
            get => _isEnabled;
            set
            {
                _isEnabled = value;
                if (!value && _isRecording) StopListening();
                _logger?.LogInformation("[ElevenLabs STT] Enabled = {Enabled}", value);
            }
        }

        public event EventHandler<string>? CommandRecognized;
        public event EventHandler<bool>?   ListeningStateChanged;
        public event EventHandler<string>? HypothesisGenerated;
        public event EventHandler<string>? RecognitionError;

        // ── Construction ───────────────────────────────────────────────────────

        public ElevenLabsSpeechService(ILogger<ElevenLabsSpeechService>? logger = null)
        {
            _logger = logger;
        }

        // ── Configuration ──────────────────────────────────────────────────────

        public void Configure(string apiKey)
        {
            _apiKey = apiKey;
            _isInitialized = !string.IsNullOrEmpty(apiKey);
            _logger?.LogInformation("[ElevenLabs STT] Configured with API key");
        }

        public async Task<bool> InitializeAsync()
        {
            if (string.IsNullOrEmpty(_apiKey)) return false;
            try
            {
                using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(5) };
                http.DefaultRequestHeaders.Add("xi-api-key", _apiKey);
                var r = await http.GetAsync("https://api.elevenlabs.io/v1/user").ConfigureAwait(false);
                _isInitialized = r.IsSuccessStatusCode;
                _logger?.LogInformation("[ElevenLabs STT] Init check: {Status}", r.StatusCode);
                return _isInitialized;
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "[ElevenLabs STT] Init failed");
                _isInitialized = false;
                return false;
            }
        }

        // ── Listening ──────────────────────────────────────────────────────────

        public void StartListening()
        {
            if (!_isEnabled)
            {
                _logger?.LogWarning("[ElevenLabs STT] StartListening blocked — service disabled");
                return;
            }
            if (_isRecording) return;
            if (!_isInitialized)
            {
                RecognitionError?.Invoke(this, "ElevenLabs not initialized");
                return;
            }

            _isRecording = true;
            _sessionCts  = new CancellationTokenSource();
            _            = RunSessionAsync(_sessionCts.Token);
        }

        /// <summary>Open WebSocket, start microphone, stream until cancelled.</summary>
        private async Task RunSessionAsync(CancellationToken ct)
        {
            try
            {
                _ws = new ClientWebSocket();
                _ws.Options.SetRequestHeader("xi-api-key", _apiKey);

                _logger?.LogInformation("[ElevenLabs STT] Connecting WebSocket...");
                await _ws.ConnectAsync(new Uri(WssBase), ct).ConfigureAwait(false);
                _logger?.LogInformation("[ElevenLabs STT] WebSocket connected");

                ListeningStateChanged?.Invoke(this, true);
                HypothesisGenerated?.Invoke(this, "Listening...");

                // Start microphone AFTER the socket is open
                _waveIn = new WaveInEvent
                {
                    WaveFormat         = new WaveFormat(16000, 16, 1),
                    BufferMilliseconds = 100
                };
                _waveIn.DataAvailable += OnAudioData;
                _waveIn.StartRecording();

                await ReceiveLoopAsync(ct).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                _logger?.LogInformation("[ElevenLabs STT] Session cancelled");
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "[ElevenLabs STT] Session error");
                RecognitionError?.Invoke(this, $"Voice error: {ex.Message}");
            }
            finally
            {
                CleanupSession();
            }
        }

        /// <summary>Called on the audio thread; sends each PCM chunk as base64 JSON.</summary>
        private async void OnAudioData(object? sender, WaveInEventArgs e)
        {
            if (_ws?.State != WebSocketState.Open) return;

            await _sendLock.WaitAsync().ConfigureAwait(false);
            try
            {
                var chunk = new byte[e.BytesRecorded];
                Buffer.BlockCopy(e.Buffer, 0, chunk, 0, e.BytesRecorded);

                var msg = JsonSerializer.Serialize(new
                {
                    message_type  = "input_audio_chunk",
                    audio_base_64 = Convert.ToBase64String(chunk),
                    commit        = false,
                    sample_rate   = 16000
                });

                await _ws.SendAsync(
                    Encoding.UTF8.GetBytes(msg),
                    WebSocketMessageType.Text, true,
                    _sessionCts?.Token ?? CancellationToken.None).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                _logger?.LogDebug("[ElevenLabs STT] Audio send error: {Msg}", ex.Message);
            }
            finally
            {
                _sendLock.Release();
            }
        }

        /// <summary>Reads server messages until the WebSocket closes or ct fires.</summary>
        private async Task ReceiveLoopAsync(CancellationToken ct)
        {
            var buf = new ArraySegment<byte>(new byte[64 * 1024]);

            while (_ws?.State == WebSocketState.Open && !ct.IsCancellationRequested)
            {
                using var ms = new MemoryStream();
                WebSocketReceiveResult result;
                do
                {
                    result = await _ws.ReceiveAsync(buf, ct).ConfigureAwait(false);
                    if (result.MessageType == WebSocketMessageType.Close) return;
                    ms.Write(buf.Array!, 0, result.Count);
                }
                while (!result.EndOfMessage);

                HandleServerMessage(Encoding.UTF8.GetString(ms.ToArray()));
            }
        }

        private void HandleServerMessage(string json)
        {
            try
            {
                using var doc = JsonDocument.Parse(json);
                if (!doc.RootElement.TryGetProperty("message_type", out var tp)) return;
                var type = tp.GetString() ?? "";

                _logger?.LogDebug("[ElevenLabs STT] ← {Type}", type);

                switch (type)
                {
                    case "session_started":
                        _logger?.LogInformation("[ElevenLabs STT] Realtime session ready");
                        break;

                    case "partial_transcript":
                        var partial = doc.RootElement.TryGetProperty("text", out var p)
                            ? p.GetString() ?? "" : "";
                        if (!string.IsNullOrWhiteSpace(partial))
                            HypothesisGenerated?.Invoke(this, partial);
                        break;

                    case "committed_transcript":
                        var text = doc.RootElement.TryGetProperty("text", out var t)
                            ? t.GetString() ?? "" : "";
                        _logger?.LogInformation("[ElevenLabs STT] Committed: \"{Text}\"", text);
                        if (!string.IsNullOrWhiteSpace(text))
                            CommandRecognized?.Invoke(this, text);
                        HypothesisGenerated?.Invoke(this, "Ready");
                        _committedTcs?.TrySetResult(true);
                        break;

                    default:
                        // Any *_error message type
                        if (type.Contains("error", StringComparison.OrdinalIgnoreCase))
                        {
                            var errMsg = doc.RootElement.TryGetProperty("error", out var e2)
                                ? e2.GetString() ?? type : type;
                            _logger?.LogWarning("[ElevenLabs STT] API error: {Err}", errMsg);
                            RecognitionError?.Invoke(this, errMsg);
                            _committedTcs?.TrySetResult(false);
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                _logger?.LogDebug("[ElevenLabs STT] Message parse error: {Msg}", ex.Message);
            }
        }

        public void StopListening()
        {
            if (!_isRecording) return;

            _logger?.LogInformation("[ElevenLabs STT] Stopped — sending commit to ElevenLabs");
            _waveIn?.StopRecording();
            _ = CommitAndCloseAsync();
        }

        private async Task CommitAndCloseAsync()
        {
            _committedTcs = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);

            // Allow last audio buffers to flush
            await Task.Delay(120).ConfigureAwait(false);

            if (_ws?.State == WebSocketState.Open)
            {
                await _sendLock.WaitAsync().ConfigureAwait(false);
                try
                {
                    var msg = JsonSerializer.Serialize(new
                    {
                        message_type  = "input_audio_chunk",
                        audio_base_64 = "",
                        commit        = true,
                        sample_rate   = 16000
                    });
                    await _ws.SendAsync(
                        Encoding.UTF8.GetBytes(msg),
                        WebSocketMessageType.Text, true,
                        CancellationToken.None).ConfigureAwait(false);
                    _logger?.LogInformation("[ElevenLabs STT] Commit sent — waiting for transcript...");
                }
                catch (Exception ex)
                {
                    _logger?.LogWarning(ex, "[ElevenLabs STT] Commit send failed");
                    _committedTcs.TrySetResult(false);
                }
                finally { _sendLock.Release(); }
            }
            else
            {
                _committedTcs.TrySetResult(false);
            }

            // Wait up to 8 s for ElevenLabs to return the committed transcript
            await Task.WhenAny(_committedTcs.Task, Task.Delay(8000)).ConfigureAwait(false);
            _logger?.LogInformation("[ElevenLabs STT] Commit wait complete");

            // Cancel the receive loop → triggers RunSessionAsync finally → CleanupSession
            _sessionCts?.Cancel();
        }

        private void CleanupSession()
        {
            _isRecording = false;
            ListeningStateChanged?.Invoke(this, false);

            try { _waveIn?.StopRecording(); } catch { }
            _waveIn?.Dispose();
            _waveIn = null;

            try
            {
                if (_ws != null && _ws.State == WebSocketState.Open)
                {
                    _ws.CloseAsync(WebSocketCloseStatus.NormalClosure, "Done", CancellationToken.None)
                       .GetAwaiter().GetResult();
                }
            }
            catch { }
            _ws?.Dispose();
            _ws = null;

            _sessionCts?.Dispose();
            _sessionCts = null;
        }

        // ── IDisposable ────────────────────────────────────────────────────────

        public void Dispose()
        {
            if (_isDisposed) return;
            _isDisposed = true;

            _sessionCts?.Cancel();
            CleanupSession();
            _sendLock.Dispose();
        }
    }
}
