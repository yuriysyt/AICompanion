using System;
using System.Threading.Tasks;
using Grpc.Net.Client;
using AICompanion.Desktop.Grpc;
using AICompanion.Desktop.Models;
using Microsoft.Extensions.Logging;
using Google.Protobuf;

namespace AICompanion.Desktop.Services.Communication
{
    /*
        AIEngineClient manages communication with the Python AI engine via gRPC.
        
        This client establishes a connection to the Python backend running on
        localhost:50051 and sends voice commands along with screen context for
        AI processing. The Python engine runs the IBM Granite 3.3-2B model and
        returns structured action responses.
        
        gRPC was chosen over REST APIs for several reasons:
        - Binary protocol is 3-5x faster than JSON serialization
        - Strongly-typed message contracts prevent integration errors
        - Bidirectional streaming enables real-time conversation
        - Built-in code generation for both C# and Python
        
        The connection is maintained for the application lifetime to avoid
        reconnection overhead on each command.
        
        Reference: https://grpc.io/docs/languages/csharp/
    */
    public class AIEngineClient : IDisposable
    {
        private readonly ILogger<AIEngineClient> _logger;
        private GrpcChannel? _channel;
        private AIService.AIServiceClient? _client;
        private bool _isConnected;
        private bool _isDisposed;

        /*
            Default address for the local Python AI engine.
            The engine runs as a separate process on the same machine.
        */
        private const string DefaultServerAddress = "http://localhost:50051";

        /*
            Timeout for individual requests in seconds.
            Most AI processing completes within 2 seconds; this allows
            extra time for complex commands.
        */
        private const int RequestTimeoutSeconds = 10;

        /*
            Event raised when the connection state changes.
            UI components use this to show connection status indicators.
        */
        public event EventHandler<bool>? ConnectionStateChanged;

        public AIEngineClient(ILogger<AIEngineClient> logger)
        {
            _logger = logger;
        }

        /*
            Establishes a connection to the Python AI engine.
            
            This method creates a gRPC channel to the specified address
            and verifies connectivity by requesting the engine status.
            The connection remains open until Dispose is called.
        */
        public async Task<bool> ConnectAsync(string? serverAddress = null)
        {
            try
            {
                var address = serverAddress ?? DefaultServerAddress;
                _logger.LogInformation("Connecting to AI engine at {Address}", address);

                /*
                    Create the gRPC channel with default options.
                    The channel manages connection pooling and reconnection.
                */
                _channel = GrpcChannel.ForAddress(address, new GrpcChannelOptions
                {
                    MaxReceiveMessageSize = 50 * 1024 * 1024,
                    MaxSendMessageSize = 50 * 1024 * 1024
                });

                _client = new AIService.AIServiceClient(_channel);

                /*
                    Verify connectivity by requesting engine status.
                */
                var statusResponse = await GetEngineStatusAsync();
                
                if (statusResponse != null && statusResponse.IsReady)
                {
                    _isConnected = true;
                    _logger.LogInformation("Connected to AI engine. Model: {Model}", statusResponse.ModelName);
                    ConnectionStateChanged?.Invoke(this, true);
                    return true;
                }
                else
                {
                    _logger.LogWarning("AI engine not ready: {Error}", statusResponse?.ErrorMessage ?? "Unknown error");
                    return false;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to connect to AI engine");
                _isConnected = false;
                ConnectionStateChanged?.Invoke(this, false);
                return false;
            }
        }

        /*
            Sends a voice command to the AI engine for processing.
            
            This is the primary method for interacting with the AI. It packages
            the voice command text and screen context into a gRPC request,
            sends it to the Python engine, and returns the structured response.
        */
        public async Task<CommandResponse?> ProcessCommandAsync(VoiceCommand command, ScreenContext screenContext)
        {
            if (_client == null || !_isConnected)
            {
                _logger.LogWarning("Cannot process command: not connected to AI engine");
                return null;
            }

            try
            {
                _logger.LogDebug("Sending command to AI engine: {Command}", command.TranscribedText);
                var startTime = DateTime.UtcNow;

                /*
                    Build the gRPC request message with command text and screenshot.
                */
                var request = new CommandRequest
                {
                    CommandText = command.TranscribedText,
                    ScreenshotData = ByteString.CopyFrom(screenContext.ScreenshotData),
                    RequestId = command.CommandId,
                    TimestampMs = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds()
                };

                /*
                    Add conversation history for context understanding.
                    This helps the AI understand references like "do that again".
                */
                if (!string.IsNullOrEmpty(command.PreviousCommandId))
                {
                    request.ConversationHistory.Add(command.PreviousCommandId);
                }

                /*
                    Make the gRPC call with timeout.
                */
                var deadline = DateTime.UtcNow.AddSeconds(RequestTimeoutSeconds);
                var response = await _client.ProcessCommandAsync(request, deadline: deadline);

                var processingTime = (DateTime.UtcNow - startTime).TotalMilliseconds;
                _logger.LogDebug("AI response received in {Time:F0}ms, action: {Action}", 
                    processingTime, response.ActionType);

                return response;
            }
            catch (global::Grpc.Core.RpcException ex) when (ex.StatusCode == global::Grpc.Core.StatusCode.DeadlineExceeded)
            {
                _logger.LogWarning("AI engine request timed out after {Timeout}s", RequestTimeoutSeconds);
                return CreateTimeoutResponse(command.CommandId);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error processing command through AI engine");
                return CreateErrorResponse(command.CommandId, ex.Message);
            }
        }

        /*
            Requests the current status of the AI engine.
            
            Returns information about model loading state, memory usage,
            and processing statistics. Used for health monitoring and
            displaying status in the UI.
        */
        public async Task<StatusResponse?> GetEngineStatusAsync()
        {
            if (_client == null)
            {
                return null;
            }

            try
            {
                var request = new StatusRequest { IncludeDiagnostics = true };
                var response = await _client.GetSystemStatusAsync(request);
                return response;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting AI engine status");
                return null;
            }
        }

        /*
            Checks if the AI engine is currently available.
        */
        public bool IsConnected => _isConnected;

        /*
            Creates a response indicating the request timed out.
        */
        private CommandResponse CreateTimeoutResponse(string requestId)
        {
            return new CommandResponse
            {
                ActionType = ActionType.ActionError,
                RequestId = requestId,
                ResponseText = "I am taking too long to process your request. Please try again.",
                ConfidenceScore = 0
            };
        }

        /*
            Creates a response for connection or processing errors.
        */
        private CommandResponse CreateErrorResponse(string requestId, string errorMessage)
        {
            return new CommandResponse
            {
                ActionType = ActionType.ActionError,
                RequestId = requestId,
                ResponseText = "I had trouble understanding that. Could you try again?",
                ConfidenceScore = 0
            };
        }

        public void Dispose()
        {
            if (_isDisposed)
            {
                return;
            }

            _channel?.Dispose();
            _isConnected = false;
            _isDisposed = true;
            
            _logger.LogInformation("AI engine client disposed");
        }
    }
}
