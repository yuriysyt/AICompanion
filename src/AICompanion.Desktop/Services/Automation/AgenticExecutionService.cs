using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;

namespace AICompanion.Desktop.Services.Automation
{
    /// <summary>
    /// OpenClaw-inspired Agentic Execution Engine.
    /// Receives multi-step plans from Python /api/plan and physically executes
    /// each step on the Windows desktop using WindowAutomationHelper.
    /// 
    /// State Machine: Idle → Planning → Executing → StepComplete → (loop) → Done/Error
    /// </summary>
    public class AgenticExecutionService
    {
        private readonly ILogger<AgenticExecutionService>? _logger;
        private readonly WindowAutomationHelper _automation;
        private readonly HttpClient _httpClient;

        private const string PlanEndpoint = "http://localhost:8000/api/plan";
        private const string IntentEndpoint = "http://localhost:8000/api/intent";
        private const string ContextEndpoint = "http://localhost:8000/api/context";
        private const int HttpTimeoutSeconds = 30;

        // Execution state
        private AgenticState _state = AgenticState.Idle;
        private AgenticPlanDto? _currentPlan;
        private int _currentStepIndex;

        // Phase 6: Track the target window across plan steps to prevent
        // typing into the wrong window (e.g. AI Companion itself).
        private IntPtr _planTargetWindow = IntPtr.Zero;
        private string _planTargetWindowTitle = "";

        // Events for UI feedback
        public event EventHandler<AgenticStateChangedArgs>? StateChanged;
        public event EventHandler<StepExecutedArgs>? StepExecuted;
        public event EventHandler<string>? ExecutionError;
        public event EventHandler<AgenticPlanDto>? PlanReceived;

        public AgenticState CurrentState => _state;

        public AgenticExecutionService(
            WindowAutomationHelper automation,
            ILogger<AgenticExecutionService>? logger = null)
        {
            _automation = automation;
            _logger = logger;
            _httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(HttpTimeoutSeconds) };
        }

        /// <summary>
        /// Main entry point: sends command text to Python /api/plan,
        /// receives a multi-step plan, and executes it step by step.
        /// Returns a summary result string.
        /// </summary>
        public async Task<AgenticResult> ExecuteCommandAsync(string commandText, CancellationToken ct = default)
        {
            if (string.IsNullOrWhiteSpace(commandText))
                return AgenticResult.Fail("Empty command");

            // Phase 7: Concurrency guard — reject if already busy
            if (_state != AgenticState.Idle && _state != AgenticState.Done)
            {
                _logger?.LogWarning("[AGENT] Rejected command while busy (state={State}): '{Text}'",
                    _state, commandText);
                return AgenticResult.Fail($"Agent is busy ({_state}). Please wait for the current plan to finish.");
            }

            var sw = Stopwatch.StartNew();
            SetState(AgenticState.Planning);

            // Phase 7 FIX: Do NOT blindly reset the tracked window!
            // Instead, validate whether the existing tracked window is still alive.
            // This is critical for multi-command flows like:
            //   Command 1: "open notepad"  →  _planTargetWindow = notepad handle
            //   Command 2: "write hello"   →  must REUSE the notepad handle!
            if (_planTargetWindow != IntPtr.Zero)
            {
                var existingTitle = _automation.GetWindowTitle(_planTargetWindow);
                if (string.IsNullOrEmpty(existingTitle) || existingTitle == "(none)")
                {
                    _logger?.LogInformation("[AGENT] Previously tracked window is gone — clearing");
                    _planTargetWindow = IntPtr.Zero;
                    _planTargetWindowTitle = "";
                }
                else
                {
                    _logger?.LogInformation("[AGENT] Reusing tracked window from previous command: '{Title}' (Handle: {H})",
                        existingTitle, _planTargetWindow);
                    _planTargetWindowTitle = existingTitle;
                }
            }

            try
            {
                // Push current window context to Python backend before planning
                await PushWindowContextAsync();

                // Step 1: Get plan from Python backend
                _logger?.LogInformation("[AGENT] Requesting plan from Python backend for: '{Text}'", commandText);
                _currentPlan = await RequestPlanAsync(commandText, ct);
                if (_currentPlan == null || _currentPlan.Steps == null || _currentPlan.Steps.Count == 0)
                {
                    SetState(AgenticState.Error);
                    return AgenticResult.Fail("Backend returned empty plan");
                }

                _logger?.LogInformation("[AGENT] Plan received: {Id}, {N} steps, reasoning: {R}",
                    _currentPlan.PlanId, _currentPlan.TotalSteps, _currentPlan.Reasoning);
                PlanReceived?.Invoke(this, _currentPlan);

                // Step 2: Execute each step sequentially
                SetState(AgenticState.Executing);
                _currentStepIndex = 0;
                var results = new List<string>();

                foreach (var step in _currentPlan.Steps)
                {
                    if (ct.IsCancellationRequested)
                    {
                        SetState(AgenticState.Idle);
                        return AgenticResult.Fail("Execution cancelled by user");
                    }

                    _currentStepIndex = step.StepNumber;
                    _logger?.LogInformation(
                        "[AGENT] Step {N}/{Total}: {Action} → target='{Target}' | tracked window='{Win}' (Handle: {H})",
                        step.StepNumber, _currentPlan.TotalSteps, step.Action,
                        step.Target ?? "(none)", _planTargetWindowTitle, _planTargetWindow);

                    var stepResult = await ExecuteStepAsync(step);
                    results.Add($"Step {step.StepNumber} ({step.Action}): {(stepResult ? "OK" : "FAIL")}");

                    StepExecuted?.Invoke(this, new StepExecutedArgs
                    {
                        StepNumber = step.StepNumber,
                        TotalSteps = _currentPlan.TotalSteps,
                        Action = step.Action,
                        Success = stepResult,
                        Description = step.Description
                    });

                    if (!stepResult)
                    {
                        _logger?.LogWarning("[AGENT] Step {N} failed, continuing...", step.StepNumber);
                    }
                }

                SetState(AgenticState.Done);
                var elapsed = sw.ElapsedMilliseconds;
                var summary = $"Plan '{_currentPlan.PlanId}' completed: {_currentPlan.TotalSteps} steps in {elapsed}ms";
                _logger?.LogInformation("[AGENT] {Summary}", summary);

                return new AgenticResult
                {
                    Success = true,
                    Summary = summary,
                    StepResults = results,
                    ElapsedMs = (int)elapsed
                };
            }
            catch (HttpRequestException ex)
            {
                _logger?.LogWarning("[AGENT] Backend unreachable: {E}", ex.Message);
                SetState(AgenticState.Error);
                return AgenticResult.Fail($"Python backend unreachable: {ex.Message}");
            }
            catch (TaskCanceledException)
            {
                SetState(AgenticState.Error);
                return AgenticResult.Fail("Request timed out");
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "[AGENT] Execution error");
                SetState(AgenticState.Error);
                ExecutionError?.Invoke(this, ex.Message);
                return AgenticResult.Fail($"Execution error: {ex.Message}");
            }
            finally
            {
                // Always return to idle after execution
                if (_state != AgenticState.Idle)
                    SetState(AgenticState.Idle);
            }
        }

        /// <summary>
        /// Tries to get a quick intent from /api/intent (single action, no plan).
        /// Returns the action and target, or null if backend is unavailable.
        /// </summary>
        public async Task<IntentResultDto?> GetIntentAsync(string text)
        {
            try
            {
                var payload = JsonSerializer.Serialize(new { text });
                var content = new StringContent(payload, Encoding.UTF8, "application/json");
                var response = await _httpClient.PostAsync(IntentEndpoint, content).ConfigureAwait(false);

                if (!response.IsSuccessStatusCode) return null;

                var json = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                return JsonSerializer.Deserialize<IntentResultDto>(json, new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                });
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Checks if the Python backend is reachable.
        /// </summary>
        public async Task<bool> IsBackendAvailableAsync()
        {
            try
            {
                var response = await _httpClient.GetAsync("http://localhost:8000/api/health").ConfigureAwait(false);
                return response.IsSuccessStatusCode;
            }
            catch
            {
                return false;
            }
        }

        // ====== Private: Plan Request ======

        private async Task<AgenticPlanDto?> RequestPlanAsync(string text, CancellationToken ct)
        {
            // Phase 6: Include window context in plan request
            var fgTitle = _automation.GetWindowTitle(_automation.GetCurrentForegroundWindow());
            var payload = JsonSerializer.Serialize(new
            {
                text,
                max_steps = 8,
                window_title = fgTitle,
                window_process = "(desktop)"
            });
            var content = new StringContent(payload, Encoding.UTF8, "application/json");

            _logger?.LogInformation("[AGENT] POST /api/plan: '{Text}' (window: {Win})", text, fgTitle);
            var response = await _httpClient.PostAsync(PlanEndpoint, content, ct).ConfigureAwait(false);
            response.EnsureSuccessStatusCode();

            var json = await response.Content.ReadAsStringAsync(ct).ConfigureAwait(false);
            _logger?.LogDebug("[AGENT] Plan response: {Json}", json);

            return JsonSerializer.Deserialize<AgenticPlanDto>(json, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            });
        }

        // ====== Private: Step Execution (State Machine) ======

        private async Task<bool> ExecuteStepAsync(ActionStepDto step)
        {
            try
            {
                return step.Action.ToLowerInvariant() switch
                {
                    "open_app" => await ExecuteOpenApp(step),
                    "close_app" => await ExecuteCloseApp(step),
                    "type_text" => await ExecuteTypeText(step),
                    "format_bold" => await ExecuteShortcut(step, "^b"),
                    "format_italic" => await ExecuteShortcut(step, "^i"),
                    "format_underline" => await ExecuteShortcut(step, "^u"),
                    "save_document" => await ExecuteShortcut(step, "^s"),
                    "select_all" => await ExecuteShortcut(step, "^a"),
                    "copy" => await ExecuteShortcut(step, "^c"),
                    "paste" => await ExecuteShortcut(step, "^v"),
                    "cut" => await ExecuteShortcut(step, "^x"),
                    "undo" => await ExecuteShortcut(step, "^z"),
                    "redo" => await ExecuteShortcut(step, "^y"),
                    "new_line" => await ExecuteShortcut(step, "{ENTER}"),
                    "new_tab" => await ExecuteShortcut(step, "^t"),
                    "new_window" => await ExecuteShortcut(step, "^n"),
                    "close_tab" => await ExecuteShortcut(step, "^w"),
                    "refresh" => await ExecuteShortcut(step, "{F5}"),
                    "go_back" => await ExecuteShortcut(step, "%{LEFT}"),
                    "go_forward" => await ExecuteShortcut(step, "%{RIGHT}"),
                    "delete" => await ExecuteShortcut(step, "{DELETE}"),
                    "scroll_down" => await ExecuteShortcut(step, "{PGDN}"),
                    "scroll_up" => await ExecuteShortcut(step, "{PGUP}"),
                    "screenshot" => await ExecuteScreenshot(step),
                    "create_document" => await ExecuteShortcut(step, "^n"),
                    "minimize" => await ExecuteMinimize(step),
                    "wait" => await ExecuteWait(step),
                    "search_web" => await ExecuteSearchWeb(step),
                    _ => await ExecuteUnknown(step)
                };
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "[AGENT] Step {N} ({Action}) threw exception",
                    step.StepNumber, step.Action);
                return false;
            }
        }

        private async Task<bool> ExecuteOpenApp(ActionStepDto step)
        {
            var appName = WindowAutomationHelper.ResolveAppName(step.Target ?? "");
            if (string.IsNullOrEmpty(appName))
            {
                _logger?.LogWarning("[AGENT] open_app: no target specified");
                return false;
            }

            _logger?.LogInformation("[AGENT] open_app: launching '{App}'...", appName);

            var hwnd = await _automation.LaunchAndWaitForWindowAsync(appName, 5000);
            if (hwnd != IntPtr.Zero)
            {
                _automation.ForceFocusWindow(hwnd);
                // Track this as the plan's target window
                _planTargetWindow = hwnd;
                _planTargetWindowTitle = _automation.GetWindowTitle(hwnd);
                _logger?.LogInformation("[AGENT] open_app: SUCCESS — tracking window '{Title}' (Handle: {H})",
                    _planTargetWindowTitle, hwnd);
                return true;
            }

            // App might already be open
            hwnd = _automation.FindWindowByProcessName(appName);
            if (hwnd != IntPtr.Zero)
            {
                _automation.ForceFocusWindow(hwnd);
                _planTargetWindow = hwnd;
                _planTargetWindowTitle = _automation.GetWindowTitle(hwnd);
                _logger?.LogInformation("[AGENT] open_app: found existing window '{Title}' (Handle: {H})",
                    _planTargetWindowTitle, hwnd);
                return true;
            }

            // Phase 7: Try to open as a file (user said "open file final report")
            var originalTarget = step.Target ?? appName;
            _logger?.LogInformation("[AGENT] open_app: trying as file search for '{Target}'...", originalTarget);
            var filePath = TryFindFile(originalTarget);
            if (filePath != null)
            {
                _logger?.LogInformation("[AGENT] open_app: found file '{Path}', opening...", filePath);
                try
                {
                    Process.Start(new ProcessStartInfo { FileName = filePath, UseShellExecute = true });
                    await Task.Delay(2000).ConfigureAwait(false);
                    // Try to find the window that opened
                    var fgAfter = _automation.GetCurrentForegroundWindow();
                    var fgTitle = _automation.GetWindowTitle(fgAfter);
                    if (!string.IsNullOrEmpty(fgTitle) && !fgTitle.Contains("AI Companion", StringComparison.OrdinalIgnoreCase))
                    {
                        _planTargetWindow = fgAfter;
                        _planTargetWindowTitle = fgTitle;
                        _logger?.LogInformation("[AGENT] open_app: file opened — tracking '{Title}' (Handle: {H})",
                            fgTitle, fgAfter);
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    _logger?.LogWarning(ex, "[AGENT] open_app: failed to open file '{Path}'", filePath);
                }
            }

            _logger?.LogWarning("[AGENT] open_app: FAILED to find window for '{App}'", appName);
            return false;
        }

        /// <summary>
        /// Search common locations for a file matching the given name.
        /// </summary>
        private string? TryFindFile(string query)
        {
            var searchDirs = new[]
            {
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                Environment.GetFolderPath(Environment.SpecialFolder.Recent),
            };
            var keywords = query.ToLowerInvariant()
                .Replace("file", "").Replace("document", "").Replace("doc", "")
                .Trim().Split(' ', StringSplitOptions.RemoveEmptyEntries);

            if (keywords.Length == 0) return null;

            foreach (var dir in searchDirs)
            {
                if (!Directory.Exists(dir)) continue;
                try
                {
                    var files = Directory.GetFiles(dir, "*.*", SearchOption.TopDirectoryOnly);
                    foreach (var file in files)
                    {
                        var fn = Path.GetFileNameWithoutExtension(file).ToLowerInvariant();
                        if (keywords.All(k => fn.Contains(k)))
                        {
                            return file;
                        }
                    }
                }
                catch { /* skip dirs we can't read */ }
            }
            return null;
        }

        private async Task<bool> ExecuteCloseApp(ActionStepDto step)
        {
            var hwnd = GetBestTargetWindow();
            if (hwnd != IntPtr.Zero)
            {
                _logger?.LogInformation("[AGENT] close_app: closing '{Title}' (Handle: {H})",
                    _automation.GetWindowTitle(hwnd), hwnd);
                var result = await _automation.SendShortcutToWindowAsync(hwnd, "%{F4}");
                _planTargetWindow = IntPtr.Zero;
                _planTargetWindowTitle = "";
                return result;
            }
            return false;
        }

        private async Task<bool> ExecuteTypeText(ActionStepDto step)
        {
            var text = step.Target ?? "";
            if (string.IsNullOrEmpty(text)) return false;

            // Phase 6 FIX: Use tracked target window instead of blind foreground
            var hwnd = GetBestTargetWindow();
            if (hwnd == IntPtr.Zero)
            {
                _logger?.LogWarning("[AGENT] type_text: no target window available");
                return false;
            }

            _logger?.LogInformation("[AGENT] type_text: typing {Len} chars into '{Title}' (Handle: {H})",
                text.Length, _automation.GetWindowTitle(hwnd), hwnd);

            return await _automation.TypeTextIntoWindowAsync(hwnd, text);
        }

        private async Task<bool> ExecuteShortcut(ActionStepDto step, string keys)
        {
            // Phase 6 FIX: Use tracked target window
            var hwnd = GetBestTargetWindow();
            if (hwnd == IntPtr.Zero)
            {
                _logger?.LogWarning("[AGENT] shortcut: no target window available");
                return false;
            }

            _logger?.LogInformation("[AGENT] shortcut: sending '{Keys}' to '{Title}' (Handle: {H})",
                keys, _automation.GetWindowTitle(hwnd), hwnd);

            return await _automation.SendShortcutToWindowAsync(hwnd, keys);
        }

        private async Task<bool> ExecuteMinimize(ActionStepDto step)
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = "explorer.exe",
                    Arguments = "shell:::{3080F90D-D7AD-11D9-BD98-0000947B0257}",
                    UseShellExecute = true
                };
                Process.Start(psi);
                await Task.Delay(150).ConfigureAwait(false); // Phase 6: reduced from 300ms
                return true;
            }
            catch
            {
                return false;
            }
        }

        private async Task<bool> ExecuteWait(ActionStepDto step)
        {
            var ms = 500; // Phase 6: reduced default from 1000ms
            if (step.Parameters != null && step.Parameters.TryGetValue("ms", out var msVal))
            {
                if (msVal.ValueKind == JsonValueKind.Number)
                    ms = msVal.GetInt32();
            }

            _logger?.LogDebug("[AGENT] Waiting {Ms}ms", ms);
            await Task.Delay(Math.Min(ms, 10000)).ConfigureAwait(false); // cap at 10s
            return true;
        }

        private async Task<bool> ExecuteSearchWeb(ActionStepDto step)
        {
            var query = step.Target ?? "";
            if (string.IsNullOrEmpty(query)) return false;

            try
            {
                var url = $"https://www.google.com/search?q={Uri.EscapeDataString(query)}";
                Process.Start(new ProcessStartInfo { FileName = url, UseShellExecute = true });
                await Task.Delay(200).ConfigureAwait(false); // Phase 6: reduced from 500ms
                return true;
            }
            catch
            {
                return false;
            }
        }

        private async Task<bool> ExecuteScreenshot(ActionStepDto step)
        {
            try
            {
                var bytes = _automation.CaptureDesktopScreenshot();
                if (bytes.Length > 0)
                {
                    _logger?.LogInformation("[AGENT] screenshot: captured {Kb}KB", bytes.Length / 1024);
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                _logger?.LogWarning(ex, "[AGENT] screenshot failed");
                return false;
            }
        }

        private Task<bool> ExecuteUnknown(ActionStepDto step)
        {
            _logger?.LogWarning("[AGENT] Unknown action: {Action}", step.Action);
            return Task.FromResult(false);
        }

        // ====== Helpers ======

        /// <summary>
        /// Phase 6: Returns the best target window for the current plan.
        /// Priority: tracked plan window > current foreground (if not AI Companion).
        /// </summary>
        private IntPtr GetBestTargetWindow()
        {
            // If we tracked a window during this plan (e.g. from open_app), use it
            if (_planTargetWindow != IntPtr.Zero)
            {
                if (_automation.IsWindowValid(_planTargetWindow))
                {
                    // Re-focus the tracked window to make sure it's in front
                    _automation.ForceFocusWindow(_planTargetWindow, 2);
                    _logger?.LogInformation("[AGENT] Using tracked window: '{Title}' (Handle: {H})",
                        _planTargetWindowTitle, _planTargetWindow);
                    return _planTargetWindow;
                }
                // Window closed or invalid, clear tracking
                _logger?.LogWarning("[AGENT] Tracked window is no longer valid — clearing");
                _planTargetWindow = IntPtr.Zero;
                _planTargetWindowTitle = "";
            }

            // Fallback: use foreground, but skip if it's AI Companion
            var fg = _automation.GetCurrentForegroundWindow();
            var fgTitle = _automation.GetWindowTitle(fg);
            if (fgTitle.Contains("AI Companion", StringComparison.OrdinalIgnoreCase))
            {
                _logger?.LogWarning("[AGENT] Foreground is AI Companion — cannot use as target");
                return IntPtr.Zero;
            }

            return fg;
        }

        /// <summary>
        /// Push current window context + screenshot to Python backend for LLM awareness.
        /// </summary>
        private async Task PushWindowContextAsync()
        {
            try
            {
                var fg = _automation.GetCurrentForegroundWindow();
                var title = _automation.GetWindowTitle(fg);

                // Capture screenshot for visual context
                string screenshotB64 = "";
                try
                {
                    var screenshotBytes = _automation.CaptureDesktopScreenshot();
                    if (screenshotBytes.Length > 0)
                    {
                        screenshotB64 = Convert.ToBase64String(screenshotBytes);
                        _logger?.LogInformation("[AGENT] Screenshot captured: {Kb}KB",
                            screenshotBytes.Length / 1024);
                    }
                }
                catch (Exception ex)
                {
                    _logger?.LogDebug("[AGENT] Screenshot failed: {E}", ex.Message);
                }

                // Also include the tracked window info for better context
                var trackedInfo = _planTargetWindow != IntPtr.Zero
                    ? _planTargetWindowTitle
                    : "(none)";

                var payload = JsonSerializer.Serialize(new
                {
                    active_window = title,
                    active_process = "(desktop)",
                    tracked_window = trackedInfo,
                    screenshot_b64 = screenshotB64.Length > 100 ? screenshotB64.Substring(0, 100) + "..." : ""
                });
                var content = new StringContent(payload, Encoding.UTF8, "application/json");
                await _httpClient.PostAsync(ContextEndpoint, content).ConfigureAwait(false);
                _logger?.LogInformation("[AGENT] Pushed context: fg='{FG}', tracked='{Tracked}'",
                    title, trackedInfo);
            }
            catch (Exception ex)
            {
                _logger?.LogDebug("[AGENT] Could not push window context: {E}", ex.Message);
            }
        }

        private void SetState(AgenticState newState)
        {
            var old = _state;
            _state = newState;
            _logger?.LogInformation("[AGENT] State: {Old} → {New}", old, newState);
            StateChanged?.Invoke(this, new AgenticStateChangedArgs { OldState = old, NewState = newState });
        }
    }

    // ====== Enums and DTOs ======

    public enum AgenticState
    {
        Idle,
        Planning,
        Executing,
        Done,
        Error
    }

    /// <summary>
    /// Mirrors Python AgenticPlan pydantic model.
    /// </summary>
    public class AgenticPlanDto
    {
        [JsonPropertyName("request")]
        public string Request { get; set; } = "";

        [JsonPropertyName("plan_id")]
        public string PlanId { get; set; } = "";

        [JsonPropertyName("steps")]
        public List<ActionStepDto> Steps { get; set; } = new();

        [JsonPropertyName("total_steps")]
        public int TotalSteps { get; set; }

        [JsonPropertyName("estimated_duration_ms")]
        public int EstimatedDurationMs { get; set; }

        [JsonPropertyName("reasoning")]
        public string Reasoning { get; set; } = "";

        [JsonPropertyName("created_at")]
        public string CreatedAt { get; set; } = "";
    }

    /// <summary>
    /// Mirrors Python ActionStep pydantic model.
    /// </summary>
    public class ActionStepDto
    {
        [JsonPropertyName("step_number")]
        public int StepNumber { get; set; }

        [JsonPropertyName("action")]
        public string Action { get; set; } = "";

        [JsonPropertyName("target")]
        public string? Target { get; set; }

        [JsonPropertyName("parameters")]
        public Dictionary<string, JsonElement>? Parameters { get; set; }

        [JsonPropertyName("description")]
        public string Description { get; set; } = "";

        [JsonPropertyName("depends_on")]
        public int? DependsOn { get; set; }
    }

    /// <summary>
    /// Mirrors Python IntentResponse.
    /// </summary>
    public class IntentResultDto
    {
        [JsonPropertyName("action")]
        public string Action { get; set; } = "unknown";

        [JsonPropertyName("target")]
        public string? Target { get; set; }

        [JsonPropertyName("parameters")]
        public Dictionary<string, JsonElement>? Parameters { get; set; }

        [JsonPropertyName("confidence")]
        public float Confidence { get; set; }
    }

    public class AgenticResult
    {
        public bool Success { get; set; }
        public string Summary { get; set; } = "";
        public List<string> StepResults { get; set; } = new();
        public int ElapsedMs { get; set; }

        public static AgenticResult Fail(string reason) => new()
        {
            Success = false,
            Summary = reason
        };
    }

    public class AgenticStateChangedArgs : EventArgs
    {
        public AgenticState OldState { get; set; }
        public AgenticState NewState { get; set; }
    }

    public class StepExecutedArgs : EventArgs
    {
        public int StepNumber { get; set; }
        public int TotalSteps { get; set; }
        public string Action { get; set; } = "";
        public bool Success { get; set; }
        public string Description { get; set; } = "";
    }
}
