using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;

namespace AICompanion.Desktop.Services.Automation
{
    // State Machine: Idle → Planning → Executing → Done/Error → Idle
    public class AgenticExecutionService
    {
        // ====== Skills Dispatcher ======

        private static readonly Dictionary<string, Func<AgenticExecutionService, ActionStepDto, CancellationToken, Task<bool>>> _skills = new()
        {
            ["open_app"]          = (svc, s, ct) => svc.ExecuteOpenApp(s),
            ["focus_window"]      = (svc, s, ct) => svc.ExecuteFocusWindow(s),
            ["close_app"]         = (svc, s, ct) => svc.ExecuteCloseApp(s),
            ["minimize"]          = (svc, s, ct) => svc.ExecuteMinimize(s),
            ["maximize"]          = (svc, s, ct) => svc.ExecuteShortcut(s, "% "),
            ["switch_window"]     = (svc, s, ct) => svc.ExecuteShortcut(s, "%{TAB}"),
            ["type_text"]         = (svc, s, ct) => svc.ExecuteTypeText(s),
            ["type_filename"]     = (svc, s, ct) => svc.ExecuteTypeFilename(s),
            ["new_line"]          = (svc, s, ct) => svc.ExecuteShortcut(s, "{ENTER}"),
            ["select_all"]        = (svc, s, ct) => svc.ExecuteShortcut(s, "^a"),
            ["clear_all"]         = (svc, s, ct) => svc.ExecuteClearAll(s),
            ["delete"]            = (svc, s, ct) => svc.ExecuteShortcut(s, "{DELETE}"),
            ["backspace"]         = (svc, s, ct) => svc.ExecuteShortcut(s, "{BACKSPACE}"),
            ["delete_word"]       = (svc, s, ct) => svc.ExecuteShortcut(s, "^{BACKSPACE}"),
            ["delete_line"]       = (svc, s, ct) => svc.ExecuteDeleteLine(s),
            ["copy"]              = (svc, s, ct) => svc.ExecuteShortcut(s, "^c"),
            ["paste"]             = (svc, s, ct) => svc.ExecuteShortcut(s, "^v"),
            ["cut"]               = (svc, s, ct) => svc.ExecuteShortcut(s, "^x"),
            ["format_bold"]       = (svc, s, ct) => svc.ExecuteShortcut(s, "^b"),
            ["format_italic"]     = (svc, s, ct) => svc.ExecuteShortcut(s, "^i"),
            ["format_underline"]  = (svc, s, ct) => svc.ExecuteShortcut(s, "^u"),
            ["save_document"]     = (svc, s, ct) => svc.ExecuteSaveDocument(s),
            ["save_document_as"]  = (svc, s, ct) => svc.ExecuteShortcut(s, "^+s"),
            ["undo"]              = (svc, s, ct) => svc.ExecuteShortcut(s, "^z"),
            ["redo"]              = (svc, s, ct) => svc.ExecuteShortcut(s, "^y"),
            ["go_to_start"]       = (svc, s, ct) => svc.ExecuteShortcut(s, "^{HOME}"),
            ["go_to_end"]         = (svc, s, ct) => svc.ExecuteShortcut(s, "^{END}"),
            ["scroll_down"]       = (svc, s, ct) => svc.ExecuteShortcut(s, "{PGDN}"),
            ["scroll_up"]         = (svc, s, ct) => svc.ExecuteShortcut(s, "{PGUP}"),
            ["go_back"]           = (svc, s, ct) => svc.ExecuteShortcut(s, "%{LEFT}"),
            ["go_forward"]        = (svc, s, ct) => svc.ExecuteShortcut(s, "%{RIGHT}"),
            ["search_web"]        = (svc, s, ct) => svc.ExecuteSearchWeb(s),
            ["new_tab"]           = (svc, s, ct) => svc.ExecuteShortcut(s, "^t"),
            ["close_tab"]         = (svc, s, ct) => svc.ExecuteShortcut(s, "^w"),
            ["new_window"]        = (svc, s, ct) => svc.ExecuteShortcut(s, "^n"),
            ["refresh"]           = (svc, s, ct) => svc.ExecuteShortcut(s, "{F5}"),
            ["address_bar"]       = (svc, s, ct) => svc.ExecuteShortcut(s, "^l"),
            ["press_escape"]      = (svc, s, ct) => svc.ExecuteShortcut(s, "{ESC}"),
            ["press_tab"]         = (svc, s, ct) => svc.ExecuteShortcut(s, "{TAB}"),
            ["press_enter"]       = (svc, s, ct) => svc.ExecuteShortcut(s, "{ENTER}"),
            ["press_key"]         = (svc, s, ct) => svc.ExecuteShortcut(s, s.Target ?? "{ENTER}"),
            ["find_replace"]      = (svc, s, ct) => svc.ExecuteShortcut(s, "^h"),
            ["zoom_in"]           = (svc, s, ct) => svc.ExecuteShortcut(s, "^{ADD}"),
            ["zoom_out"]          = (svc, s, ct) => svc.ExecuteShortcut(s, "^{SUBTRACT}"),
            ["screenshot"]        = (svc, s, ct) => svc.ExecuteScreenshot(s),
            ["create_document"]   = (svc, s, ct) => svc.ExecuteNewDocument(s),
            ["new_document"]      = (svc, s, ct) => svc.ExecuteNewDocument(s),
            ["wait"]              = (svc, s, ct) => svc.ExecuteWait(s),
            ["message_user"]      = (svc, s, ct) => Task.FromResult(svc.ExecuteMessageUser(s)),
            ["ask_user"]          = (svc, s, ct) => Task.FromResult(svc.ExecuteAskUser(s)),
            ["click_button"]      = (svc, s, ct) => svc.ExecuteClickButton(s),
            ["mouse_click"]       = (svc, s, ct) => svc.ExecuteMouseClick(s),
            ["dismiss_dialog"]    = (svc, s, ct) => svc.ExecuteDismissDialog(s),
            // Navigation / browser actions that IBM Granite commonly returns
            ["navigate_url"]      = (svc, s, ct) => svc.ExecuteNavigateUrl(s),
            ["navigate"]          = (svc, s, ct) => svc.ExecuteNavigateUrl(s),
            ["open_url"]          = (svc, s, ct) => svc.ExecuteNavigateUrl(s),
            ["browse_to"]         = (svc, s, ct) => svc.ExecuteNavigateUrl(s),
            ["goto_url"]          = (svc, s, ct) => svc.ExecuteNavigateUrl(s),
            // Alias actions
            ["click"]             = (svc, s, ct) => svc.ExecuteMouseClick(s),
            ["left_click"]        = (svc, s, ct) => svc.ExecuteMouseClick(s),
            ["click_at"]          = (svc, s, ct) => svc.ExecuteMouseClick(s),
            ["right_click"]       = (svc, s, ct) => svc.ExecuteRightClick(s),
            ["double_click"]      = (svc, s, ct) => svc.ExecuteDoubleClick(s),
            ["find_and_click"]    = (svc, s, ct) => svc.ExecuteFindAndClick(s),
            ["click_element"]     = (svc, s, ct) => svc.ExecuteFindAndClick(s),
            ["scroll"]            = (svc, s, ct) => svc.ExecuteShortcut(s, "{PGDN}"),
            ["open_browser"]      = (svc, s, ct) => svc.ExecuteOpenApp(s),
            ["describe_screen"]   = (svc, s, ct) => svc.ExecuteDescribeScreen(s),
            ["hotkey"]            = (svc, s, ct) => svc.ExecuteHotkey(s),
            ["key_combo"]         = (svc, s, ct) => svc.ExecuteHotkey(s),
            ["mouse_move"]        = (svc, s, ct) => svc.ExecuteMouseMove(s),
            ["drag_to"]           = (svc, s, ct) => svc.ExecuteDragTo(s),
            ["scroll_at"]         = (svc, s, ct) => svc.ExecuteScrollAt(s),
            ["open_file_path"]    = (svc, s, ct) => svc.ExecuteOpenFilePath(s),
            ["open_file"]         = (svc, s, ct) => svc.ExecuteOpenFilePath(s),
        };

        public static IReadOnlyList<string> AvailableSkills => _skills.Keys.ToList();

        // ====== P/Invoke ======

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out WindowAutomationHelper.RECT lpRect);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        private const int SW_RESTORE = 9;

        // ====== Fields ======

        private readonly ILogger<AgenticExecutionService>? _logger;
        private readonly WindowAutomationHelper _automation;
        private readonly HttpClient _httpClient;

        private const string PlanEndpoint        = "http://localhost:8000/api/smart_command";
        private const string IntentEndpoint      = "http://localhost:8000/api/intent";
        private const string ContextEndpoint     = "http://localhost:8000/api/context";
        private const string OclExecuteEndpoint  = "http://localhost:8000/api/execute_step";
        private const int HttpTimeoutSeconds = 30;

        private AgenticState _state = AgenticState.Idle;
        private AgenticPlanDto? _currentPlan;
        private int _currentStepIndex;

        // OpenClaw memory: which window is being worked with across steps
        private IntPtr _planTargetWindow = IntPtr.Zero;
        private string _planTargetWindowTitle = "";

        // ====== Events ======

        public event EventHandler<AgenticStateChangedArgs>? StateChanged;
        public event EventHandler<StepExecutedArgs>? StepExecuted;
        public event EventHandler<string>? ExecutionError;
        public event EventHandler<AgenticPlanDto>? PlanReceived;
        public event EventHandler<string>? StatusMessage;

        // ====== Public Properties ======

        public AgenticState CurrentState => _state;

        /// <summary>
        /// When set, ExecuteCommandAsync pauses after the plan is received and awaits
        /// user confirmation. Return true to proceed, false to cancel.
        /// </summary>
        public Func<AgenticPlanDto, Task<bool>>? ConfirmationRequired { get; set; }

        /// <summary>
        /// Optional session context JSON injected by the caller before each plan request.
        /// Included in the /api/plan payload so the AI backend maintains cross-command memory.
        /// </summary>
        public string? SessionContext { get; set; }

        // ====== Constructor ======

        public AgenticExecutionService(
            WindowAutomationHelper automation,
            ILogger<AgenticExecutionService>? logger = null)
        {
            _automation = automation;
            _logger = logger;
            _httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(HttpTimeoutSeconds) };
        }

        // ====== Heartbeat (OpenClaw pattern) ======

        public Task<bool> HeartbeatAsync() =>
            Task.FromResult(_state == AgenticState.Idle || _state == AgenticState.Done);

        // ====== Public API ======

        /// <summary>
        /// Main entry point: sends command text to Python /api/plan,
        /// receives a multi-step plan, and executes it step by step.
        /// </summary>
        public async Task<AgenticResult> ExecuteCommandAsync(string commandText, CancellationToken ct = default)
        {
            if (string.IsNullOrWhiteSpace(commandText))
                return AgenticResult.Fail("Empty command");

            if (_state != AgenticState.Idle && _state != AgenticState.Done)
            {
                _logger?.LogWarning("[AGENT] Rejected command while busy (state={State}): '{Text}'",
                    _state, commandText);
                return AgenticResult.Fail($"Agent is busy ({_state}). Please wait for the current plan to finish.");
            }

            var sw = Stopwatch.StartNew();
            SetState(AgenticState.Planning);

            // Validate tracked window is still alive
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
                    _logger?.LogInformation("[AGENT] Reusing tracked window: '{Title}' (Handle: {H})",
                        existingTitle, _planTargetWindow);
                    _planTargetWindowTitle = existingTitle;
                }
            }

            try
            {
                await PushWindowContextAsync();

                const int MaxPlanRetries = 3;
                string    failureSummary = "";
                AgenticResult? lastResult = null;

                for (int planAttempt = 1; planAttempt <= MaxPlanRetries; planAttempt++)
                {
                    // On retry, append failure context so Granite can plan around it
                    var planText = planAttempt == 1
                        ? commandText
                        : $"{commandText} [RETRY {planAttempt}: previous attempt failed — {failureSummary}. Active window now: '{_automation.GetWindowTitle(_automation.GetCurrentForegroundWindow())}']";

                    _logger?.LogInformation("[AGENT] Plan attempt {Attempt}/{Max}: '{Text}'", planAttempt, MaxPlanRetries, planText);
                    if (planAttempt > 1)
                        StatusMessage?.Invoke(this, $"🔄 Re-planning (attempt {planAttempt}/{MaxPlanRetries})...");

                    _currentPlan = await RequestPlanAsync(planText, ct);
                    if (_currentPlan == null || _currentPlan.Steps == null || _currentPlan.Steps.Count == 0)
                    {
                        if (planAttempt < MaxPlanRetries)
                        {
                            StatusMessage?.Invoke(this, "⚠️ Backend returned empty plan — retrying...");
                            failureSummary = "backend returned empty plan";
                            continue;
                        }
                        SetState(AgenticState.Error);
                        return AgenticResult.Fail("Backend returned empty plan after retries");
                    }

                    _logger?.LogInformation("[AGENT] Plan received: {Id}, {N} steps, reasoning: {R}",
                        _currentPlan.PlanId, _currentPlan.TotalSteps, _currentPlan.Reasoning);
                    PlanReceived?.Invoke(this, _currentPlan);

                    // Pause for user confirmation only on the first attempt
                    if (planAttempt == 1 && ConfirmationRequired != null)
                    {
                        StatusMessage?.Invoke(this, "⏳ Waiting for confirmation...");
                        bool confirmed = await ConfirmationRequired(_currentPlan).ConfigureAwait(false);
                        if (!confirmed)
                        {
                            SetState(AgenticState.Idle);
                            return AgenticResult.Fail("Plan cancelled by user");
                        }
                        StatusMessage?.Invoke(this, "✅ Confirmed — executing plan...");
                    }

                    SetState(AgenticState.Executing);
                    _currentStepIndex = 0;
                    var results      = new List<string>();
                    var criticalFails = new List<string>();

                    foreach (var step in _currentPlan.Steps)
                    {
                        if (ct.IsCancellationRequested)
                        {
                            SetState(AgenticState.Idle);
                            return AgenticResult.Fail("Execution cancelled by user");
                        }

                        _currentStepIndex = step.StepNumber;
                        _logger?.LogInformation(
                            "[AGENT] Step {N}/{Total}: {Action} → target='{Target}' | tracked='{Win}' (Handle: {H})",
                            step.StepNumber, _currentPlan.TotalSteps, step.Action,
                            step.Target ?? "(none)", _planTargetWindowTitle, _planTargetWindow);

                        StatusMessage?.Invoke(this, $"▶ Step {step.StepNumber}/{_currentPlan.TotalSteps}: {step.Action}{(step.Target != null ? " → " + step.Target : "")}");

                        var stepResult = await ExecuteStepAsync(step, ct);

                        if (!stepResult && step.Action != "message_user" && step.Action != "wait")
                        {
                            _logger?.LogWarning("[AGENT] Step {N} failed (1st attempt), retrying...", step.StepNumber);
                            StatusMessage?.Invoke(this, $"⚠️ Step {step.StepNumber} ({step.Action}) failed — retrying...");

                            await Task.Delay(500, ct).ConfigureAwait(false);
                            stepResult = await ExecuteStepAsync(step, ct);

                            if (!stepResult)
                            {
                                _logger?.LogWarning("[AGENT] Step {N} failed (2nd attempt) — trying OCL Python fallback...", step.StepNumber);
                                StatusMessage?.Invoke(this, $"🐍 Step {step.StepNumber}: trying Python fallback...");

                                stepResult = await TryOclExecutionAsync(step);

                                if (stepResult)
                                {
                                    _logger?.LogInformation("[AGENT] Step {N} succeeded via OCL Python!", step.StepNumber);
                                    StatusMessage?.Invoke(this, $"✅ Step {step.StepNumber} succeeded via Python executor.");
                                }
                                else
                                {
                                    _logger?.LogWarning("[AGENT] Step {N} failed all methods.", step.StepNumber);
                                    StatusMessage?.Invoke(this, $"❌ Step {step.StepNumber} ({step.Action}) failed all methods.");

                                    // Track critical failures that warrant re-planning
                                    bool isCritical = step.Action is "type_text" or "new_document"
                                        or "open_app" or "find_and_click" or "focus_window";
                                    if (isCritical)
                                        criticalFails.Add($"{step.Action}({step.Target ?? step.Params ?? ""}) failed");
                                }
                            }
                            else
                            {
                                _logger?.LogInformation("[AGENT] Step {N} succeeded on retry!", step.StepNumber);
                                StatusMessage?.Invoke(this, $"✅ Step {step.StepNumber} ({step.Action}) succeeded on retry.");
                            }
                        }

                        results.Add($"Step {step.StepNumber} ({step.Action}): {(stepResult ? "OK" : "FAIL")}");

                        StepExecuted?.Invoke(this, new StepExecutedArgs
                        {
                            StepNumber  = step.StepNumber,
                            TotalSteps  = _currentPlan.TotalSteps,
                            Action      = step.Action,
                            Success     = stepResult,
                            Description = step.Description
                        });

                        if (stepResult && step.Action != "wait" && step.Action != "message_user")
                            await HandleUnexpectedDialogAsync().ConfigureAwait(false);

                        _ = PushWindowContextAsync();
                    }

                    SetState(AgenticState.Done);
                    var elapsed = sw.ElapsedMilliseconds;
                    var summary = $"Plan '{_currentPlan.PlanId}' attempt {planAttempt} completed: {_currentPlan.TotalSteps} steps in {elapsed}ms";
                    _logger?.LogInformation("[AGENT] {Summary}", summary);

                    lastResult = new AgenticResult
                    {
                        Success     = true,
                        Summary     = summary,
                        StepResults = results,
                        ElapsedMs   = (int)elapsed
                    };

                    // If no critical failures, verify the plan actually worked
                    if (criticalFails.Count == 0)
                    {
                        // Post-execution verification: check if expected app is now in foreground
                        var expectedApp = _currentPlan.Steps?
                            .FirstOrDefault(s => s.Action is "open_app" or "new_document" or "focus_window")?.Target;
                        if (!string.IsNullOrEmpty(expectedApp))
                        {
                            var verified = await VerifyExecutionAsync(expectedApp).ConfigureAwait(false);
                            if (verified.Success)
                                StatusMessage?.Invoke(this, $"✅ Verified: {verified.Description}");
                            else if (!string.IsNullOrEmpty(verified.Suggestion))
                                StatusMessage?.Invoke(this, $"⚠️ {verified.Suggestion}");
                        }
                        break;
                    }

                    // Still have attempts left — re-plan with failure context
                    if (planAttempt < MaxPlanRetries)
                    {
                        failureSummary = string.Join("; ", criticalFails);
                        _logger?.LogWarning("[AGENT] Critical step failures: {Fails} — re-planning...", failureSummary);
                        StatusMessage?.Invoke(this, $"⚠️ {criticalFails.Count} step(s) failed — re-planning...");
                        SetState(AgenticState.Planning);
                    }
                    else
                    {
                        _logger?.LogWarning("[AGENT] Max retries reached with failures: {Fails}", string.Join("; ", criticalFails));
                    }
                }

                return lastResult ?? AgenticResult.Fail("No plan was executed");
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
                if (_state != AgenticState.Idle)
                    SetState(AgenticState.Idle);
            }
        }

        public async Task<IntentResultDto?> GetIntentAsync(string text)
        {
            try
            {
                var payload = JsonSerializer.Serialize(new { text });
                var content = new StringContent(payload, Encoding.UTF8, "application/json");
                var response = await _httpClient.PostAsync(IntentEndpoint, content).ConfigureAwait(false);
                if (!response.IsSuccessStatusCode) return null;

                var json = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                return JsonSerializer.Deserialize<IntentResultDto>(json,
                    new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
            }
            catch { return null; }
        }

        public async Task<bool> IsBackendAvailableAsync()
        {
            try
            {
                var response = await _httpClient.GetAsync("http://localhost:8000/api/health").ConfigureAwait(false);
                return response.IsSuccessStatusCode;
            }
            catch { return false; }
        }

        // ====== Plan Request ======

        /// <summary>
        /// Strips file names and user paths from window titles before sending to the AI backend.
        /// "MySecret.docx - Microsoft Word" → "Microsoft Word"
        /// </summary>
        private static string SanitizeWindowTitle(string title)
        {
            if (string.IsNullOrWhiteSpace(title)) return "(unknown)";

            if (title.Contains(" - "))
            {
                var parts = title.Split(new[] { " - " }, StringSplitOptions.RemoveEmptyEntries);
                return parts[^1].Trim();
            }

            if (System.Text.RegularExpressions.Regex.IsMatch(title, @"(?:[A-Za-z]:\\|\\\\)"))
                return "Document Editor";

            var lower = title.ToLowerInvariant();
            if (lower.Contains(".docx") || lower.Contains(".doc") || lower.Contains(".txt") ||
                lower.Contains(".pdf") || lower.Contains(".xlsx") || lower.Contains(".csv"))
                return "Document Editor";

            return title;
        }

        private async Task<AgenticPlanDto?> RequestPlanAsync(string text, CancellationToken ct)
        {
            var rawTitle = _automation.GetWindowTitle(_automation.GetCurrentForegroundWindow());
            var fgTitle  = SanitizeWindowTitle(rawTitle);
            var openWindows = new List<string>();
            foreach (var pn in new[] { "WINWORD", "notepad", "EXCEL", "POWERPNT", "msedge", "chrome", "opera", "explorer", "mspaint", "code", "devenv" })
                if (System.Diagnostics.Process.GetProcessesByName(pn).Length > 0)
                    openWindows.Add(pn);
            // Also include window titles from EnumerateOpenWindows for richer context
            foreach (var wt in EnumerateOpenWindows())
                if (!openWindows.Contains(wt, StringComparer.OrdinalIgnoreCase))
                    openWindows.Add(wt);

            var payload = JsonSerializer.Serialize(new
            {
                text,
                max_steps        = 8,
                window_title     = fgTitle,
                window_process   = "(desktop)",
                session_context  = SessionContext,
                open_windows     = openWindows
            });
            var content = new StringContent(payload, Encoding.UTF8, "application/json");

            _logger?.LogInformation("[AGENT] POST /api/plan: '{Text}' (window: {Win})", text, fgTitle);
            var response = await _httpClient.PostAsync(PlanEndpoint, content, ct).ConfigureAwait(false);
            response.EnsureSuccessStatusCode();

            var json = await response.Content.ReadAsStringAsync(ct).ConfigureAwait(false);
            _logger?.LogDebug("[AGENT] Plan response: {Json}", json);

            return JsonSerializer.Deserialize<AgenticPlanDto>(json,
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
        }

        // ====== Step Execution ======

        private async Task<bool> ExecuteStepAsync(ActionStepDto step, CancellationToken ct)
        {
            try
            {
                var action = step.Action.ToLowerInvariant();
                if (_skills.TryGetValue(action, out var handler))
                    return await handler(this, step, ct);

                return await ExecuteUnknown(step);
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "[AGENT] Step {N} ({Action}) threw exception",
                    step.StepNumber, step.Action);
                StatusMessage?.Invoke(this, $"⚠️ Step {step.StepNumber} failed: {ex.Message}");
                return false;
            }
        }

        // ====== OpenClaw Python Fallback ======

        /// <summary>
        /// Calls the Python /api/execute_step endpoint as a fallback when C# automation fails.
        /// Uses pyautogui + win32gui + subprocess — multiple methods per action.
        /// </summary>
        private record VerifyResult(bool Success, string Description, string Suggestion);

        private async Task<VerifyResult> VerifyExecutionAsync(string expectedApp)
        {
            try
            {
                var payload = JsonSerializer.Serialize(new
                {
                    expected_app    = expectedApp,
                    window_before   = _planTargetWindowTitle
                });
                var content  = new StringContent(payload, Encoding.UTF8, "application/json");
                var response = await _httpClient.PostAsync("http://localhost:8000/api/verify", content)
                                               .ConfigureAwait(false);
                if (!response.IsSuccessStatusCode)
                    return new VerifyResult(true, "", "");  // Can't verify, assume ok

                var json = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;
                bool success = root.TryGetProperty("success", out var sv) && sv.GetBoolean();
                string desc  = root.TryGetProperty("description", out var dv) ? dv.GetString() ?? "" : "";
                string sugg  = root.TryGetProperty("suggestion",  out var sgv) ? sgv.GetString() ?? "" : "";
                _logger?.LogInformation("[AGENT] Verify: success={S} desc={D}", success, desc);
                return new VerifyResult(success, desc, sugg);
            }
            catch
            {
                return new VerifyResult(true, "", "");  // If verify fails, assume ok
            }
        }

        private async Task<bool> TryOclExecutionAsync(ActionStepDto step)
        {
            try
            {
                _logger?.LogInformation("[OCL] Python fallback: {Action} target='{Target}'",
                    step.Action, step.Target ?? "(none)");

                var payload = JsonSerializer.Serialize(new
                {
                    action        = step.Action,
                    target        = step.Target,
                    @params       = step.Params,
                    window_title  = _planTargetWindowTitle
                });
                var content  = new StringContent(payload, Encoding.UTF8, "application/json");
                var response = await _httpClient.PostAsync(OclExecuteEndpoint, content)
                                               .ConfigureAwait(false);

                if (!response.IsSuccessStatusCode)
                {
                    _logger?.LogWarning("[OCL] HTTP {Status}", response.StatusCode);
                    return false;
                }

                var json = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                using var doc     = JsonDocument.Parse(json);
                var root          = doc.RootElement;
                bool success      = root.TryGetProperty("success", out var sv) && sv.GetBoolean();
                string method     = root.TryGetProperty("method",  out var mv) ? mv.GetString() ?? "" : "";

                if (success)
                {
                    _logger?.LogInformation("[OCL] Succeeded via method='{Method}'", method);
                    // Give Python actions time to complete before next step
                    await Task.Delay(1500).ConfigureAwait(false);
                    // Re-capture foreground in case a new window appeared
                    var newFg = _automation.GetCurrentForegroundWindow();
                    var newTitle = _automation.GetWindowTitle(newFg);
                    if (!string.IsNullOrEmpty(newTitle) && !newTitle.Contains("AI Companion"))
                    {
                        _planTargetWindow      = newFg;
                        _planTargetWindowTitle = newTitle;
                        _logger?.LogInformation("[OCL] Updated target window: '{Title}'", newTitle);
                    }
                }
                else
                {
                    var reason = root.TryGetProperty("reason", out var rv) ? rv.GetString() : "";
                    _logger?.LogWarning("[OCL] Failed: {Reason}", reason);
                }

                return success;
            }
            catch (Exception ex)
            {
                _logger?.LogWarning(ex, "[OCL] Python fallback threw exception");
                return false;
            }
        }

        // ====== Action Handlers ======

        private async Task<bool> ExecuteOpenApp(ActionStepDto step)
        {
            var appName = WindowAutomationHelper.ResolveAppName(step.Target ?? "");
            if (string.IsNullOrEmpty(appName))
            {
                _logger?.LogWarning("[AGENT] open_app: no target specified");
                return false;
            }

            _logger?.LogInformation("[AGENT] open_app: requested '{App}' (resolved from '{Raw}')", appName, step.Target);

            var hwnd = _automation.FindWindowByProcessName(appName);
            if (hwnd == IntPtr.Zero)
                hwnd = _automation.FindWindowByTitleContains(appName);

            if (hwnd != IntPtr.Zero)
            {
                _automation.ForceFocusWindow(hwnd);
                _planTargetWindow = hwnd;
                _planTargetWindowTitle = _automation.GetWindowTitle(hwnd);
                _logger?.LogInformation("[AGENT] open_app: reused existing window '{Title}' (Handle: {H})",
                    _planTargetWindowTitle, hwnd);
                return true;
            }

            _logger?.LogInformation("[AGENT] open_app: not found, launching '{App}'...", appName);
            hwnd = await _automation.LaunchAndWaitForWindowAsync(appName, 5000);
            if (hwnd != IntPtr.Zero)
            {
                _automation.ForceFocusWindow(hwnd);
                _planTargetWindow = hwnd;
                _planTargetWindowTitle = _automation.GetWindowTitle(hwnd);
                _logger?.LogInformation("[AGENT] open_app: launched — tracking '{Title}' (Handle: {H})",
                    _planTargetWindowTitle, hwnd);
                return true;
            }

            var originalTarget = step.Target ?? appName;
            _logger?.LogInformation("[AGENT] open_app: trying file search for '{Target}'...", originalTarget);
            var filePath = TryFindFile(originalTarget);
            if (filePath != null)
            {
                _logger?.LogInformation("[AGENT] open_app: found file '{Path}', opening...", filePath);
                try
                {
                    Process.Start(new ProcessStartInfo { FileName = filePath, UseShellExecute = true });
                    await Task.Delay(2000).ConfigureAwait(false);
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

            // Calculator in Windows 10/11 is a UWP app — calc.exe is just a launcher stub.
            // The real window is owned by Calculator.exe / ApplicationFrameHost; try ms-calculator: URI.
            if (appName.Equals("calc", StringComparison.OrdinalIgnoreCase))
            {
                _logger?.LogInformation("[AGENT] open_app: Calculator not found as process — trying ms-calculator: URI");
                try
                {
                    Process.Start(new ProcessStartInfo { FileName = "ms-calculator:", UseShellExecute = true });
                    await Task.Delay(2500).ConfigureAwait(false);
                    var calcHwnd = _automation.FindWindowByTitleContains("Calculator");
                    if (calcHwnd == IntPtr.Zero)
                        calcHwnd = _automation.FindWindowByTitleContains("Калькулятор");
                    if (calcHwnd != IntPtr.Zero)
                    {
                        _automation.ForceFocusWindow(calcHwnd);
                        _planTargetWindow = calcHwnd;
                        _planTargetWindowTitle = _automation.GetWindowTitle(calcHwnd);
                        _logger?.LogInformation("[AGENT] open_app: Calculator opened via URI, tracking '{Title}'",
                            _planTargetWindowTitle);
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    _logger?.LogWarning(ex, "[AGENT] open_app: ms-calculator: URI failed");
                }
            }

            _logger?.LogWarning("[AGENT] open_app: FAILED to find window for '{App}'", appName);
            StatusMessage?.Invoke(this, $"⚠️ Could not open '{appName}'. Is it installed? Try saying 'open notepad' or 'open Word'.");
            return false;
        }

        private async Task<bool> ExecuteFocusWindow(ActionStepDto step)
        {
            var target = step.Target ?? "";
            if (string.IsNullOrWhiteSpace(target)) return false;

            var appName = WindowAutomationHelper.ResolveAppName(target);
            var hwnd = _automation.FindWindowByProcessName(appName);
            if (hwnd == IntPtr.Zero)
                hwnd = _automation.FindWindowByTitleContains(appName);
            if (hwnd == IntPtr.Zero)
                hwnd = _automation.FindWindowByTitleContains(target);

            if (hwnd == IntPtr.Zero)
            {
                _logger?.LogWarning("[AGENT] focus_window: '{Target}' not found — falling back to open_app", target);
                return await ExecuteOpenApp(step);
            }

            ShowWindow(hwnd, SW_RESTORE);
            _automation.ForceFocusWindow(hwnd);
            _planTargetWindow = hwnd;
            _planTargetWindowTitle = _automation.GetWindowTitle(hwnd);
            _logger?.LogInformation("[AGENT] focus_window: focused '{Title}' (Handle: {H})",
                _planTargetWindowTitle, hwnd);
            await Task.Delay(200).ConfigureAwait(false);
            return true;
        }

        private async Task<bool> ExecuteNewDocument(ActionStepDto step)
        {
            var appName = WindowAutomationHelper.ResolveAppName(step.Target ?? "");
            _logger?.LogInformation("[AGENT] new_document: target='{App}'", appName);
            StatusMessage?.Invoke(this, "📄 Creating new blank document...");

            // ── Notepad: launch a fresh instance with an explicit temp file to bypass
            //            Windows 11 session restore (which reopens the last closed file)
            if (appName.Equals("notepad", StringComparison.OrdinalIgnoreCase))
            {
                var tempPath = Path.Combine(Path.GetTempPath(),
                    $"note_{DateTime.Now:yyyyMMdd_HHmmss}.txt");
                File.WriteAllText(tempPath, "");
                _logger?.LogInformation("[AGENT] new_document: launching Notepad with temp file '{Path}'", tempPath);

                Process.Start(new ProcessStartInfo
                {
                    FileName = "notepad.exe",
                    Arguments = $"\"{tempPath}\"",
                    UseShellExecute = true
                });

                // Poll up to 5 seconds for Notepad to appear as foreground
                IntPtr nbHwnd = IntPtr.Zero;
                for (int poll = 0; poll < 25; poll++)
                {
                    await Task.Delay(200).ConfigureAwait(false);
                    var fg = _automation.GetCurrentForegroundWindow();
                    var fgT = _automation.GetWindowTitle(fg);
                    if (!string.IsNullOrEmpty(fgT) && !fgT.Contains("AI Companion", StringComparison.OrdinalIgnoreCase)
                        && (fgT.Contains("Notepad", StringComparison.OrdinalIgnoreCase)
                            || fgT.Contains("Блокнот", StringComparison.OrdinalIgnoreCase)
                            || fgT.Contains(".txt", StringComparison.OrdinalIgnoreCase)
                            || fgT.Contains("note_", StringComparison.OrdinalIgnoreCase)))
                    {
                        nbHwnd = fg;
                        break;
                    }
                    // Also try by process name as fallback
                    var byProc = _automation.FindWindowByProcessName("notepad");
                    if (byProc != IntPtr.Zero)
                    {
                        _automation.ForceFocusWindow(byProc);
                        await Task.Delay(150).ConfigureAwait(false);
                        nbHwnd = byProc;
                        break;
                    }
                }

                if (nbHwnd != IntPtr.Zero)
                {
                    _automation.ForceFocusWindow(nbHwnd);
                    _planTargetWindow      = nbHwnd;
                    _planTargetWindowTitle = _automation.GetWindowTitle(nbHwnd);
                    _logger?.LogInformation("[AGENT] new_document: Notepad ready '{Title}' (Handle: {H})",
                        _planTargetWindowTitle, nbHwnd);
                    StatusMessage?.Invoke(this, "✅ New Notepad file ready.");
                    return true;
                }

                _logger?.LogWarning("[AGENT] new_document: could not find Notepad window after launch");
                return false;
            }

            // ── Word: use /n flag to bypass the Start Screen (template picker)
            //   /n = open blank new document directly without showing the splash/start screen
            if (appName.Equals("WINWORD", StringComparison.OrdinalIgnoreCase))
            {
                var existingWord = _automation.FindWindowByProcessName("WINWORD");
                if (existingWord == IntPtr.Zero)
                {
                    // Launch WPS/Word — skip /n as WPS doesn't support it; instead send Ctrl+N after launch
                    _logger?.LogInformation("[AGENT] new_document: Word not running — launching WINWORD.EXE");
                    try
                    {
                        Process.Start(new ProcessStartInfo { FileName = "WINWORD.EXE", UseShellExecute = true });
                    }
                    catch
                    {
                        // Some systems need .EXE extension omitted
                        Process.Start(new ProcessStartInfo { FileName = "WINWORD", UseShellExecute = true });
                    }
                    await Task.Delay(4000).ConfigureAwait(false);

                    // Find the WPS/Word window (process name is WINWORD)
                    var wordHwnd = _automation.FindWindowByProcessName("WINWORD");
                    if (wordHwnd == IntPtr.Zero)
                    {
                        // Last try: grab whatever is foreground
                        wordHwnd = _automation.GetCurrentForegroundWindow();
                        var fgT = _automation.GetWindowTitle(wordHwnd);
                        if (string.IsNullOrEmpty(fgT) || fgT.Contains("AI Companion", StringComparison.OrdinalIgnoreCase))
                            wordHwnd = IntPtr.Zero;
                    }

                    if (wordHwnd != IntPtr.Zero)
                    {
                        _automation.ForceFocusWindow(wordHwnd);
                        _planTargetWindow = wordHwnd;
                        _planTargetWindowTitle = _automation.GetWindowTitle(wordHwnd);
                        await Task.Delay(300).ConfigureAwait(false);

                        // Send Ctrl+N to ensure we have a blank document (not Start Screen)
                        _logger?.LogInformation("[AGENT] new_document: sending Ctrl+N to ensure blank doc in '{Title}'", _planTargetWindowTitle);
                        await _automation.SendShortcutToWindowAsync(wordHwnd, "^n").ConfigureAwait(false);
                        await Task.Delay(2000).ConfigureAwait(false);
                        _automation.ForceFocusWindow(wordHwnd);
                        await Task.Delay(200).ConfigureAwait(false);

                        // Capture the new document window that Ctrl+N opened
                        var newDoc = _automation.GetCurrentForegroundWindow();
                        var newDocTitle = _automation.GetWindowTitle(newDoc);
                        if (!string.IsNullOrEmpty(newDocTitle) && !newDocTitle.Contains("AI Companion", StringComparison.OrdinalIgnoreCase))
                        {
                            _planTargetWindow = newDoc;
                            _planTargetWindowTitle = newDocTitle;
                        }

                        _logger?.LogInformation("[AGENT] new_document: Word blank doc ready '{Title}' (Handle: {H})", _planTargetWindowTitle, _planTargetWindow);
                        StatusMessage?.Invoke(this, "✅ New blank Word document ready.");
                        return true;
                    }

                    _logger?.LogWarning("[AGENT] new_document: could not open Word blank document");
                    return false;
                }

                // Word IS already running — send Ctrl+N for a new document
                _automation.ForceFocusWindow(existingWord);
                _planTargetWindow      = existingWord;
                _planTargetWindowTitle = _automation.GetWindowTitle(existingWord);
                await Task.Delay(300).ConfigureAwait(false);

                _logger?.LogInformation("[AGENT] new_document: Word running — sending Ctrl+N to '{Title}'", _planTargetWindowTitle);
                await _automation.SendShortcutToWindowAsync(existingWord, "^n").ConfigureAwait(false);
                await Task.Delay(2000).ConfigureAwait(false);
                _automation.ForceFocusWindow(existingWord);
                await Task.Delay(200).ConfigureAwait(false);

                var afterCtrlN  = _automation.GetCurrentForegroundWindow();
                var afterTitle  = _automation.GetWindowTitle(afterCtrlN);
                if (!string.IsNullOrEmpty(afterTitle) && !afterTitle.Contains("AI Companion", StringComparison.OrdinalIgnoreCase))
                {
                    _planTargetWindow = afterCtrlN;
                    _planTargetWindowTitle = afterTitle;
                }
                await _automation.SendShortcutToWindowAsync(_planTargetWindow, "^{END}").ConfigureAwait(false);
                StatusMessage?.Invoke(this, "✅ New blank Word document ready.");
                return true;
            }

            // ── Excel / other apps: ensure app is open, then Ctrl+N ──────────
            var existingHwnd = _automation.FindWindowByProcessName(appName);
            if (existingHwnd == IntPtr.Zero)
            {
                await ExecuteOpenApp(step);
                await Task.Delay(1000).ConfigureAwait(false);
            }
            else
            {
                _automation.ForceFocusWindow(existingHwnd);
                _planTargetWindow      = existingHwnd;
                _planTargetWindowTitle = _automation.GetWindowTitle(existingHwnd);
                await Task.Delay(300).ConfigureAwait(false);
            }

            var hwnd = GetBestTargetWindow();
            if (hwnd == IntPtr.Zero) return false;

            _logger?.LogInformation("[AGENT] new_document: Ctrl+N in '{Title}'", _automation.GetWindowTitle(hwnd));
            await _automation.SendShortcutToWindowAsync(hwnd, "^n").ConfigureAwait(false);
            await Task.Delay(2000).ConfigureAwait(false);

            var newFgHwnd  = _automation.GetCurrentForegroundWindow();
            var newFgTitle = _automation.GetWindowTitle(newFgHwnd);
            if (!string.IsNullOrEmpty(newFgTitle) && !newFgTitle.Contains("AI Companion", StringComparison.OrdinalIgnoreCase))
            {
                _planTargetWindow = newFgHwnd;
                _planTargetWindowTitle = newFgTitle;
                _logger?.LogInformation("[AGENT] new_document: foreground after Ctrl+N '{Title}'", newFgTitle);
            }
            StatusMessage?.Invoke(this, "✅ New blank document ready.");
            return true;
        }

        private async Task<bool> ExecuteCloseApp(ActionStepDto step)
        {
            var hwnd = GetBestTargetWindow();
            if (hwnd == IntPtr.Zero) return false;

            _logger?.LogInformation("[AGENT] close_app: closing '{Title}' (Handle: {H})",
                _automation.GetWindowTitle(hwnd), hwnd);
            var result = await _automation.SendShortcutToWindowAsync(hwnd, "%{F4}");
            _planTargetWindow = IntPtr.Zero;
            _planTargetWindowTitle = "";
            return result;
        }

        private async Task<bool> ExecuteTypeText(ActionStepDto step)
        {
            // Params from the step → Target → ContentGenerated at plan level (essay/research flow)
            var text = step.Params ?? step.Target ?? _currentPlan?.ContentGenerated ?? "";
            if (string.IsNullOrEmpty(text)) return false;

            var hwnd = GetBestTargetWindow();
            if (hwnd == IntPtr.Zero)
            {
                _logger?.LogWarning("[AGENT] type_text: no target window available");
                StatusMessage?.Invoke(this, "⚠️ No app focused for typing. Open an app first.");
                return false;
            }

            var windowTitle = _automation.GetWindowTitle(hwnd);
            _logger?.LogInformation("[AGENT] type_text: typing {Len} chars into '{Title}' (Handle: {H})",
                text.Length, windowTitle, hwnd);

            // For document apps, click in the document body first so the cursor lands in the right place
            var tl = windowTitle.ToLowerInvariant();
            bool isDocApp = tl.Contains("notepad") || tl.Contains("word") || tl.Contains("document") ||
                            tl.Contains("wordpad") || tl.Contains("блокнот") || tl.Contains("текст") ||
                            tl.Contains(".txt") || tl.Contains(".docx");

            if (isDocApp)
            {
                ClickInDocumentBody(hwnd);
                // WPS Word / MS Word document edit area is slow to accept cursor after click;
                // 800ms for any Word-like app, 400ms for Notepad.
                int docDelay = tl.Contains("word") || tl.Contains(".docx") ? 800 : 400;
                await Task.Delay(docDelay).ConfigureAwait(false);
            }

            bool success;
            try
            {
                success = await _automation.TypeTextIntoWindowAsync(hwnd, text).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                _logger?.LogWarning(ex, "[AGENT] type_text: first attempt threw exception");
                success = false;
            }

            if (!success)
            {
                _logger?.LogWarning("[AGENT] type_text: retrying after failure");
                StatusMessage?.Invoke(this, "⚠️ Typing failed — retrying...");
                await Task.Delay(500).ConfigureAwait(false);
                _automation.ForceFocusWindow(hwnd);
                await Task.Delay(150).ConfigureAwait(false);
                try
                {
                    success = await _automation.TypeTextIntoWindowAsync(hwnd, text).ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    _logger?.LogWarning(ex, "[AGENT] type_text: retry also threw exception");
                    success = false;
                }
            }

            if (success)
            {
                _logger?.LogInformation("[AGENT] type_text: bringing '{Title}' to front after typing", windowTitle);
                _automation.ForceFocusWindow(hwnd);
                StatusMessage?.Invoke(this, $"✏️ Typed {text.Length} characters.");
            }
            return success;
        }

        private async Task<bool> ExecuteSaveDocument(ActionStepDto step)
        {
            var hwnd = GetBestTargetWindow();
            if (hwnd == IntPtr.Zero)
            {
                _logger?.LogWarning("[AGENT] save_document: no target window available");
                StatusMessage?.Invoke(this, "⚠️ Cannot save — no app is focused. Please open an app first.");
                return false;
            }

            var filename = !string.IsNullOrWhiteSpace(step.Target) && step.Target.Length < 200
                ? step.Target.Trim()
                : $"Document_{DateTime.Now:yyyyMMdd_HHmmss}";

            foreach (var c in System.IO.Path.GetInvalidFileNameChars())
                filename = filename.Replace(c, '_');

            var titleBefore = _automation.GetWindowTitle(hwnd);
            _logger?.LogInformation("[AGENT] save_document: Ctrl+S → '{Title}', filename='{Name}'", titleBefore, filename);
            StatusMessage?.Invoke(this, $"💾 Saving '{titleBefore}'...");

            await _automation.SendShortcutToWindowAsync(hwnd, "^s").ConfigureAwait(false);

            IntPtr dialogHwnd = IntPtr.Zero;
            string dialogTitle = "";
            for (int attempt = 0; attempt < 12; attempt++)
            {
                await Task.Delay(250).ConfigureAwait(false);
                var (dlg, title) = _automation.DetectActiveDialog();
                if (dlg != IntPtr.Zero && (
                    title.Contains("Save", StringComparison.OrdinalIgnoreCase) ||
                    title.Contains("Сохран", StringComparison.OrdinalIgnoreCase) ||
                    title.Contains("Markdown", StringComparison.OrdinalIgnoreCase)))
                {
                    dialogHwnd = dlg;
                    dialogTitle = title;
                    _logger?.LogInformation("[AGENT] save_document: dialog detected at attempt {N}: '{Title}'", attempt + 1, title);
                    break;
                }
            }

            if (dialogHwnd == IntPtr.Zero)
            {
                _logger?.LogInformation("[AGENT] save_document: no dialog appeared — file saved in-place");
                StatusMessage?.Invoke(this, "✅ File saved.");
                return true;
            }

            StatusMessage?.Invoke(this, $"📝 Save dialog: '{dialogTitle}' — entering '{filename}'");

            if (dialogTitle.Contains("Markdown", StringComparison.OrdinalIgnoreCase))
            {
                _logger?.LogInformation("[AGENT] save_document: Markdown dialog — resolving format prompt");
                bool handled = _automation.ClickButtonInWindow(dialogHwnd, "text file")
                            || _automation.ClickButtonInWindow(dialogHwnd, "Keep Current Format")
                            || _automation.ClickButtonInWindow(dialogHwnd, "Don't Save")
                            || _automation.ClickButtonInWindow(dialogHwnd, "OK");
                if (!handled)
                {
                    _automation.ForceFocusWindow(dialogHwnd);
                    await Task.Delay(100).ConfigureAwait(false);
                    await _automation.SendShortcutToWindowAsync(dialogHwnd, "{ENTER}").ConfigureAwait(false);
                }
                await Task.Delay(500).ConfigureAwait(false);

                var (dlg2, title2) = _automation.DetectActiveDialog();
                if (dlg2 != IntPtr.Zero && title2.Contains("Save", StringComparison.OrdinalIgnoreCase))
                {
                    dialogHwnd = dlg2;
                    dialogTitle = title2;
                }
                else
                {
                    StatusMessage?.Invoke(this, $"✅ Saved as '{filename}'");
                    return true;
                }
            }

            var saved = await _automation.HandleSaveAsDialog(dialogHwnd, filename).ConfigureAwait(false);

            // Post-save polling: check for error dialogs
            for (int check = 0; check < 4; check++)
            {
                await Task.Delay(500).ConfigureAwait(false);
                var (postDlg, postTitle) = _automation.DetectActiveDialog();
                if (postDlg != IntPtr.Zero && (
                    postTitle.Contains("Error", StringComparison.OrdinalIgnoreCase) ||
                    postTitle.Contains("Warning", StringComparison.OrdinalIgnoreCase) ||
                    postTitle.Contains("Cannot save", StringComparison.OrdinalIgnoreCase) ||
                    postTitle.Contains("Cannot be saved", StringComparison.OrdinalIgnoreCase) ||
                    postTitle.Contains("нельзя сохранить", StringComparison.OrdinalIgnoreCase) ||
                    postTitle.Contains("не удается сохранить", StringComparison.OrdinalIgnoreCase)))
                {
                    _logger?.LogError("[AGENT] save_document: post-save error dialog: '{Title}'", postTitle);
                    StatusMessage?.Invoke(this, $"❌ Save failed — system reported: '{postTitle}'. Check file permissions or path.");
                    return false;
                }
            }

            StatusMessage?.Invoke(this, saved
                ? $"✅ Saved as '{filename}'"
                : $"⚠️ Save dialog found but could not complete save for '{filename}'");

            return saved;
        }

        private async Task<bool> ExecuteTypeFilename(ActionStepDto step)
        {
            var filename = step.Target ?? $"Document_{DateTime.Now:yyyyMMdd_HHmmss}";
            StatusMessage?.Invoke(this, $"📝 Looking for Save dialog to enter filename: {filename}");

            IntPtr dialogHwnd = IntPtr.Zero;
            string dialogTitle = "";
            for (int attempt = 0; attempt < 8; attempt++)
            {
                var (dlg, title) = _automation.DetectActiveDialog();
                if (dlg != IntPtr.Zero &&
                    (title.Contains("Save", StringComparison.OrdinalIgnoreCase) ||
                     title.Contains("Сохран", StringComparison.OrdinalIgnoreCase)))
                {
                    dialogHwnd = dlg;
                    dialogTitle = title;
                    break;
                }
                _logger?.LogDebug("[AGENT] type_filename: polling for Save dialog (attempt {N})", attempt + 1);
                await Task.Delay(400).ConfigureAwait(false);
            }

            if (dialogHwnd == IntPtr.Zero)
            {
                _logger?.LogInformation("[AGENT] type_filename: no Save dialog found — file may already be saved");
                StatusMessage?.Invoke(this, "✅ No Save dialog needed — file already saved.");
                return true;
            }

            _logger?.LogInformation("[AGENT] type_filename: found dialog '{Title}', setting filename '{Name}'",
                dialogTitle, filename);

            var saved = await _automation.HandleSaveAsDialog(dialogHwnd, filename);
            StatusMessage?.Invoke(this, saved
                ? $"✅ Saved as '{filename}'"
                : $"⚠️ Could not save as '{filename}' — dialog handling failed");
            return saved;
        }

        private async Task<bool> ExecuteClickButton(ActionStepDto step)
        {
            var buttonText = step.Target ?? "OK";
            _logger?.LogInformation("[AGENT] click_button: looking for button '{Text}'", buttonText);
            StatusMessage?.Invoke(this, $"🖱️ Clicking button '{buttonText}'...");

            var (dlgHwnd, dlgTitle) = _automation.DetectActiveDialog();
            if (dlgHwnd != IntPtr.Zero)
            {
                _logger?.LogInformation("[AGENT] click_button: found dialog '{Title}', clicking '{Text}'", dlgTitle, buttonText);
                if (_automation.ClickButtonInWindow(dlgHwnd, buttonText))
                {
                    await Task.Delay(300).ConfigureAwait(false);
                    return true;
                }
                var buttons = _automation.ListButtonsInWindow(dlgHwnd);
                _logger?.LogWarning("[AGENT] click_button: '{Text}' not found. Available: [{Buttons}]",
                    buttonText, string.Join(", ", buttons.Select(b => $"'{b.Text}'")));
            }

            var fg = _automation.GetCurrentForegroundWindow();
            if (fg != IntPtr.Zero && _automation.ClickButtonInWindow(fg, buttonText))
            {
                await Task.Delay(300).ConfigureAwait(false);
                return true;
            }

            var tracked = GetBestTargetWindow();
            if (tracked != IntPtr.Zero && tracked != fg && _automation.ClickButtonInWindow(tracked, buttonText))
            {
                await Task.Delay(300).ConfigureAwait(false);
                return true;
            }

            _logger?.LogWarning("[AGENT] click_button: button '{Text}' not found anywhere", buttonText);
            StatusMessage?.Invoke(this, $"⚠️ Button '{buttonText}' not found.");
            return false;
        }

        private async Task<bool> ExecuteMouseClick(ActionStepDto step)
        {
            var coords = step.Target ?? "";
            var parts = coords.Split(',');
            if (parts.Length != 2 ||
                !int.TryParse(parts[0].Trim(), out int x) ||
                !int.TryParse(parts[1].Trim(), out int y))
            {
                _logger?.LogWarning("[AGENT] mouse_click: invalid coordinates '{Coords}'", coords);
                return false;
            }

            _logger?.LogInformation("[AGENT] mouse_click: clicking at ({X}, {Y})", x, y);
            StatusMessage?.Invoke(this, $"🖱️ Clicking at ({x}, {y})...");
            _automation.ClickAtScreenPosition(x, y);
            await Task.Delay(200).ConfigureAwait(false);
            return true;
        }

        /// <summary>
        /// find_and_click — finds a UI element by accessible name (via UIAutomation) and invokes it.
        /// Falls back to ClickButtonInWindow for classic Win32 buttons.
        /// Params/Target = the button/element label to click (e.g. "Blank document", "Save", "OK")
        /// </summary>
        private async Task<bool> ExecuteFindAndClick(ActionStepDto step)
        {
            var elementName = step.Params ?? step.Target ?? "";
            if (string.IsNullOrEmpty(elementName)) return false;

            _logger?.LogInformation("[AGENT] find_and_click: looking for '{Name}'", elementName);
            StatusMessage?.Invoke(this, $"🖱️ Clicking '{elementName}'...");

            // 1. Try UIAutomation (works for UWP, XAML, and accessible Win32 apps)
            if (_automation.FindAndClickByName(elementName))
            {
                _logger?.LogInformation("[AGENT] find_and_click: UIAutomation clicked '{Name}'", elementName);
                await Task.Delay(500).ConfigureAwait(false);
                return true;
            }

            // 2. Fallback: Win32 BM_CLICK in the tracked / foreground window
            var hwnd = GetBestTargetWindow();
            if (hwnd != IntPtr.Zero && _automation.ClickButtonInWindow(hwnd, elementName))
            {
                _logger?.LogInformation("[AGENT] find_and_click: BM_CLICK clicked '{Name}'", elementName);
                await Task.Delay(300).ConfigureAwait(false);
                return true;
            }

            var fg = _automation.GetCurrentForegroundWindow();
            if (fg != IntPtr.Zero && _automation.ClickButtonInWindow(fg, elementName))
            {
                await Task.Delay(300).ConfigureAwait(false);
                return true;
            }

            _logger?.LogWarning("[AGENT] find_and_click: '{Name}' not found", elementName);
            StatusMessage?.Invoke(this, $"⚠️ Element '{elementName}' not found.");
            return false;
        }

        private async Task<bool> ExecuteRightClick(ActionStepDto step)
        {
            var coords = step.Params ?? step.Target ?? "";
            var parts  = coords.Split(',');
            if (parts.Length >= 2 &&
                int.TryParse(parts[0].Trim(), out int x) &&
                int.TryParse(parts[1].Trim(), out int y))
            {
                _automation.RightClickAtScreenPosition(x, y);
                await Task.Delay(300).ConfigureAwait(false);
                return true;
            }
            // Right-click center of tracked window
            var hwnd = GetBestTargetWindow();
            if (hwnd == IntPtr.Zero) return false;
            _automation.RightClickAtWindowCenter(hwnd);
            await Task.Delay(300).ConfigureAwait(false);
            return true;
        }

        private async Task<bool> ExecuteDoubleClick(ActionStepDto step)
        {
            var coords = step.Params ?? step.Target ?? "";
            var parts  = coords.Split(',');
            if (parts.Length >= 2 &&
                int.TryParse(parts[0].Trim(), out int x) &&
                int.TryParse(parts[1].Trim(), out int y))
            {
                _automation.DoubleClickAtScreenPosition(x, y);
                await Task.Delay(300).ConfigureAwait(false);
                return true;
            }
            var hwnd = GetBestTargetWindow();
            if (hwnd == IntPtr.Zero) return false;
            ClickInDocumentBody(hwnd);
            _automation.DoubleClickAtScreenPosition(0, 0); // center already set by ClickInDocumentBody
            await Task.Delay(300).ConfigureAwait(false);
            return true;
        }

        /// <summary>
        /// describe_screen — builds a text description of the current screen state
        /// (open windows, foreground app, visible buttons) and speaks/displays it.
        /// Also sends this description to the backend as context for the next plan.
        /// </summary>
        private async Task<bool> ExecuteDescribeScreen(ActionStepDto step)
        {
            var desc = BuildScreenDescription();
            _logger?.LogInformation("[AGENT] describe_screen:\n{Desc}", desc);
            StatusMessage?.Invoke(this, $"🖥️ {desc}");
            return true;
        }

        private string BuildScreenDescription()
        {
            var sb = new StringBuilder();
            var fg     = _automation.GetCurrentForegroundWindow();
            var fgTitle = _automation.GetWindowTitle(fg);
            sb.AppendLine($"Active window: {fgTitle}");

            var buttons = _automation.ListButtonsInWindow(fg);
            if (buttons.Count > 0)
                sb.AppendLine($"Visible buttons: {string.Join(", ", buttons.Take(10).Select(b => $"[{b.Text}]"))}");

            var wins = EnumerateOpenWindows();
            if (wins.Count > 0)
                sb.AppendLine($"Open windows: {string.Join(", ", wins.Take(8))}");

            if (_planTargetWindow != IntPtr.Zero)
                sb.AppendLine($"Tracked document: {_planTargetWindowTitle}");

            return sb.ToString().TrimEnd();
        }

        private async Task<bool> ExecuteDismissDialog(ActionStepDto step)
        {
            var (dlgHwnd, dlgTitle) = _automation.DetectActiveDialog();
            if (dlgHwnd == IntPtr.Zero)
            {
                _logger?.LogInformation("[AGENT] dismiss_dialog: no active dialog found");
                return true;
            }

            _logger?.LogInformation("[AGENT] dismiss_dialog: dismissing '{Title}'", dlgTitle);
            StatusMessage?.Invoke(this, $"❌ Dismissing dialog: '{dlgTitle}'");

            if (_automation.ClickButtonInWindow(dlgHwnd, "Cancel") ||
                _automation.ClickButtonInWindow(dlgHwnd, "Отмена") ||
                _automation.ClickButtonInWindow(dlgHwnd, "No") ||
                _automation.ClickButtonInWindow(dlgHwnd, "Нет"))
            {
                await Task.Delay(300).ConfigureAwait(false);
                return true;
            }

            _automation.ForceFocusWindow(dlgHwnd);
            await Task.Delay(100).ConfigureAwait(false);
            await _automation.SendShortcutToWindowAsync(dlgHwnd, "{ESC}");
            await Task.Delay(300).ConfigureAwait(false);
            return true;
        }

        private async Task<bool> ExecuteClearAll(ActionStepDto step)
        {
            var hwnd = GetBestTargetWindow();
            if (hwnd == IntPtr.Zero) return false;

            _logger?.LogInformation("[AGENT] clear_all: selecting all and deleting in '{Title}'",
                _automation.GetWindowTitle(hwnd));
            StatusMessage?.Invoke(this, "🗑️ Clearing all text...");

            await _automation.SendShortcutToWindowAsync(hwnd, "^a");
            await Task.Delay(100).ConfigureAwait(false);
            await _automation.SendShortcutToWindowAsync(hwnd, "{DELETE}");
            return true;
        }

        private async Task<bool> ExecuteDeleteLine(ActionStepDto step)
        {
            var hwnd = GetBestTargetWindow();
            if (hwnd == IntPtr.Zero) return false;

            _logger?.LogInformation("[AGENT] delete_line: deleting current line");
            await _automation.SendShortcutToWindowAsync(hwnd, "{HOME}");
            await Task.Delay(50).ConfigureAwait(false);
            await _automation.SendShortcutToWindowAsync(hwnd, "+{END}");
            await Task.Delay(50).ConfigureAwait(false);
            await _automation.SendShortcutToWindowAsync(hwnd, "{DELETE}");
            await _automation.SendShortcutToWindowAsync(hwnd, "{DELETE}");
            return true;
        }

        private bool ExecuteMessageUser(ActionStepDto step)
        {
            var message = step.Target ?? step.Description ?? "";
            if (!string.IsNullOrEmpty(message))
            {
                _logger?.LogInformation("[AGENT] message_user: '{Msg}'", message);
                StatusMessage?.Invoke(this, $"🤖 {message}");
            }
            return true;
        }

        private bool ExecuteAskUser(ActionStepDto step)
        {
            var question = step.Target ?? step.Description ?? "What should I do next?";
            _logger?.LogInformation("[AGENT] ask_user: '{Q}'", question);
            StatusMessage?.Invoke(this, $"❓ {question}");
            return true;
        }

        private async Task<bool> ExecuteShortcut(ActionStepDto step, string keys)
        {
            var hwnd = GetBestTargetWindow();
            if (hwnd == IntPtr.Zero)
            {
                _logger?.LogWarning("[AGENT] shortcut: no target window available");
                StatusMessage?.Invoke(this, $"⚠️ No window to send '{step.Action}' to.");
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
                Process.Start(new ProcessStartInfo
                {
                    FileName       = "explorer.exe",
                    Arguments      = "shell:::{3080F90D-D7AD-11D9-BD98-0000947B0257}",
                    UseShellExecute = true
                });
                await Task.Delay(150).ConfigureAwait(false);
                return true;
            }
            catch { return false; }
        }

        private async Task<bool> ExecuteWait(ActionStepDto step)
        {
            // Accept ms parameter OR params="3" meaning 3 seconds
            int ms = 500;
            if (step.Parameters != null && step.Parameters.TryGetValue("ms", out var msVal) && msVal.ValueKind == JsonValueKind.Number)
                ms = msVal.GetInt32();
            else if (!string.IsNullOrEmpty(step.Params) && double.TryParse(step.Params, out double secs))
                ms = (int)(secs * 1000);

            _logger?.LogDebug("[AGENT] Waiting {Ms}ms", ms);
            StatusMessage?.Invoke(this, $"⏳ Waiting {ms/1000.0:F1}s...");
            await Task.Delay(Math.Min(ms, 30000)).ConfigureAwait(false);
            return true;
        }

        private async Task<bool> ExecuteHotkey(ActionStepDto step)
        {
            var combo = step.Params ?? step.Target ?? "";
            if (string.IsNullOrWhiteSpace(combo)) return false;

            _logger?.LogInformation("[AGENT] hotkey: '{Combo}'", combo);
            StatusMessage?.Invoke(this, $"⌨️ Hotkey: {combo}");

            // Convert "ctrl+s" → SendKeys "^s", "alt+f4" → "%{F4}", etc.
            var sendKeys = combo.ToLowerInvariant()
                .Replace("ctrl+",    "^")
                .Replace("control+", "^")
                .Replace("shift+",   "+")
                .Replace("alt+",     "%")
                .Replace("win+",     "^{ESC}")  // Win key approximate
                .Replace("f4",       "{F4}")
                .Replace("f5",       "{F5}")
                .Replace("f11",      "{F11}")
                .Replace("delete",   "{DEL}")
                .Replace("backspace","{BS}")
                .Replace("enter",    "{ENTER}")
                .Replace("tab",      "{TAB}")
                .Replace("escape",   "{ESC}");

            var hwnd = GetBestTargetWindow();
            if (hwnd != IntPtr.Zero)
                await _automation.SendShortcutToWindowAsync(hwnd, sendKeys).ConfigureAwait(false);
            else
                _automation.SendKeysRaw(sendKeys);

            await Task.Delay(200).ConfigureAwait(false);
            return true;
        }

        private async Task<bool> ExecuteMouseMove(ActionStepDto step)
        {
            var coords = step.Target ?? step.Params ?? "";
            var parts  = coords.Replace(" ", "").Split(',');
            if (parts.Length < 2 || !int.TryParse(parts[0], out int x) || !int.TryParse(parts[1], out int y))
            {
                _logger?.LogWarning("[AGENT] mouse_move: invalid coords '{C}'", coords);
                return false;
            }
            _logger?.LogInformation("[AGENT] mouse_move → ({X},{Y})", x, y);
            StatusMessage?.Invoke(this, $"🖱️ Moving to ({x},{y})...");
            _automation.MouseMove(x, y);
            await Task.Delay(150).ConfigureAwait(false);
            return true;
        }

        private async Task<bool> ExecuteDragTo(ActionStepDto step)
        {
            var coords = step.Target ?? step.Params ?? "";
            var parts  = coords.Replace(" ", "").Split(',');
            if (parts.Length < 2 || !int.TryParse(parts[0], out int x) || !int.TryParse(parts[1], out int y))
                return false;
            _logger?.LogInformation("[AGENT] drag_to → ({X},{Y})", x, y);
            StatusMessage?.Invoke(this, $"🖱️ Dragging to ({x},{y})...");
            _automation.MouseDragTo(x, y);
            await Task.Delay(300).ConfigureAwait(false);
            return true;
        }

        private async Task<bool> ExecuteScrollAt(ActionStepDto step)
        {
            // params: "x,y,direction,amount" e.g. "960,540,down,3"
            var parts = (step.Params ?? "").Replace(" ","").Split(',');
            if (parts.Length >= 4 &&
                int.TryParse(parts[0], out int x) && int.TryParse(parts[1], out int y) &&
                int.TryParse(parts[3], out int amount))
            {
                _automation.MouseMove(x, y);
                await Task.Delay(100).ConfigureAwait(false);
                var dir = parts[2].ToLowerInvariant();
                _automation.MouseScroll(dir == "up" ? amount : -amount);
                await Task.Delay(200).ConfigureAwait(false);
                return true;
            }
            // Fallback: just scroll in active window
            var hwnd = GetBestTargetWindow();
            if (hwnd != IntPtr.Zero)
                await _automation.SendShortcutToWindowAsync(hwnd, "{PGDN}").ConfigureAwait(false);
            return true;
        }

        private async Task<bool> ExecuteOpenFilePath(ActionStepDto step)
        {
            var path = step.Params ?? step.Target ?? "";
            if (string.IsNullOrWhiteSpace(path)) return false;
            try
            {
                _logger?.LogInformation("[AGENT] open_file_path: '{Path}'", path);
                StatusMessage?.Invoke(this, $"📂 Opening: {System.IO.Path.GetFileName(path)}...");
                Process.Start(new ProcessStartInfo { FileName = path, UseShellExecute = true });
                await Task.Delay(1500).ConfigureAwait(false);
                return true;
            }
            catch (Exception ex)
            {
                _logger?.LogWarning("[AGENT] open_file_path failed: {E}", ex.Message);
                return false;
            }
        }

        private async Task<bool> ExecuteNavigateUrl(ActionStepDto step)
        {
            var url = step.Target ?? step.Params ?? "";
            if (string.IsNullOrWhiteSpace(url)) return false;

            // Ensure scheme so ShellExecute can open it
            if (!url.StartsWith("http://", StringComparison.OrdinalIgnoreCase) &&
                !url.StartsWith("https://", StringComparison.OrdinalIgnoreCase) &&
                !url.StartsWith("ftp://", StringComparison.OrdinalIgnoreCase))
                url = "https://" + url;

            try
            {
                _logger?.LogInformation("[AGENT] navigate_url → {Url}", url);
                Process.Start(new ProcessStartInfo { FileName = url, UseShellExecute = true });
                await Task.Delay(300).ConfigureAwait(false);
                return true;
            }
            catch (Exception ex)
            {
                _logger?.LogWarning(ex, "[AGENT] navigate_url failed for {Url}", url);
                return false;
            }
        }

        private async Task<bool> ExecuteSearchWeb(ActionStepDto step)
        {
            // Backend may put the query in either "target" or "params"
            var query = step.Target ?? step.Params ?? "";
            if (string.IsNullOrEmpty(query)) return false;

            try
            {
                var url = $"https://www.google.com/search?q={Uri.EscapeDataString(query)}";
                Process.Start(new ProcessStartInfo { FileName = url, UseShellExecute = true });
                await Task.Delay(200).ConfigureAwait(false);
                return true;
            }
            catch { return false; }
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

        private void ClickInDocumentBody(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero) return;
            try
            {
                if (GetWindowRect(hwnd, out var rect))
                {
                    int cx = (rect.Left + rect.Right) / 2;
                    // Click at 65% height — below toolbar/ribbon in Word and tab bar in Notepad
                    int cy = rect.Top + (int)((rect.Bottom - rect.Top) * 0.65);
                    _automation.ClickAtScreenPosition(cx, cy);
                    Thread.Sleep(100);
                    _logger?.LogDebug("[AGENT] ClickInDocumentBody: ({X}, {Y}) in '{Title}'", cx, cy,
                        _automation.GetWindowTitle(hwnd));
                }
            }
            catch (Exception ex)
            {
                _logger?.LogWarning(ex, "[AGENT] ClickInDocumentBody failed");
            }
        }

        private IntPtr GetBestTargetWindow()
        {
            if (_planTargetWindow != IntPtr.Zero)
            {
                if (_automation.IsWindowValid(_planTargetWindow))
                {
                    _automation.ForceFocusWindow(_planTargetWindow, 2);
                    _logger?.LogInformation("[AGENT] Using tracked window: '{Title}' (Handle: {H})",
                        _planTargetWindowTitle, _planTargetWindow);
                    return _planTargetWindow;
                }
                _logger?.LogWarning("[AGENT] Tracked window is no longer valid — clearing");
                _planTargetWindow = IntPtr.Zero;
                _planTargetWindowTitle = "";
            }

            var fg = _automation.GetCurrentForegroundWindow();
            var fgTitle = _automation.GetWindowTitle(fg);
            if (fgTitle.Contains("AI Companion", StringComparison.OrdinalIgnoreCase))
            {
                _logger?.LogWarning("[AGENT] Foreground is AI Companion — scanning for text editor process");
                // Fallback: scan for known text editor / document app processes
                foreach (var procName in new[] { "WINWORD", "notepad", "wordpad" })
                {
                    var editorHwnd = _automation.FindWindowByProcessName(procName);
                    if (editorHwnd != IntPtr.Zero)
                    {
                        _logger?.LogInformation("[AGENT] GetBestTargetWindow: found editor via process '{P}' (Handle: {H})", procName, editorHwnd);
                        return editorHwnd;
                    }
                }
                return IntPtr.Zero;
            }

            return fg;
        }

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
                            return file;
                    }
                }
                catch { }
            }
            return null;
        }

        private async Task PushWindowContextAsync()
        {
            try
            {
                var fg    = _automation.GetCurrentForegroundWindow();
                var title = _automation.GetWindowTitle(fg);

                string screenshotB64 = "";
                try
                {
                    var screenshotBytes = _automation.CaptureDesktopScreenshot();
                    if (screenshotBytes.Length > 0)
                    {
                        screenshotB64 = Convert.ToBase64String(screenshotBytes);
                        _logger?.LogInformation("[AGENT] Screenshot captured: {Kb}KB", screenshotBytes.Length / 1024);
                    }
                }
                catch (Exception ex)
                {
                    _logger?.LogDebug("[AGENT] Screenshot failed: {E}", ex.Message);
                }

                var trackedInfo = _planTargetWindow != IntPtr.Zero ? _planTargetWindowTitle : "(none)";

                var payload = JsonSerializer.Serialize(new
                {
                    active_window  = title,
                    active_process = "(desktop)",
                    tracked_window = trackedInfo,
                    screenshot_b64 = screenshotB64.Length > 100 ? screenshotB64[..100] + "..." : ""
                });
                var content = new StringContent(payload, Encoding.UTF8, "application/json");
                await _httpClient.PostAsync(ContextEndpoint, content).ConfigureAwait(false);
                _logger?.LogInformation("[AGENT] Pushed context: fg='{FG}', tracked='{Tracked}'", title, trackedInfo);
            }
            catch (Exception ex)
            {
                _logger?.LogDebug("[AGENT] Could not push window context: {E}", ex.Message);
            }
        }

        private async Task HandleUnexpectedDialogAsync()
        {
            try
            {
                var (dlgHwnd, dlgTitle) = _automation.DetectActiveDialog();
                if (dlgHwnd == IntPtr.Zero) return;

                _logger?.LogInformation("[AGENT] Unexpected dialog detected: '{Title}'", dlgTitle);
                var lo = dlgTitle.ToLowerInvariant();

                if (lo.Contains("do you want to save") || lo.Contains("save changes") || lo.Contains("сохранить изменения"))
                {
                    _logger?.LogInformation("[AGENT] Auto-handling 'save changes' dialog → Yes");
                    StatusMessage?.Invoke(this, $"💾 Handling save prompt: '{dlgTitle}' → Yes");
                    var clicked = _automation.ClickButtonInWindow(dlgHwnd, "Yes")
                               || _automation.ClickButtonInWindow(dlgHwnd, "Да")
                               || _automation.ClickButtonInWindow(dlgHwnd, "Save");
                    if (!clicked)
                    {
                        _automation.ForceFocusWindow(dlgHwnd);
                        await Task.Delay(100).ConfigureAwait(false);
                        await _automation.SendShortcutToWindowAsync(dlgHwnd, "{ENTER}").ConfigureAwait(false);
                    }
                    await Task.Delay(400).ConfigureAwait(false);
                }
                else if (lo.Contains("already exists") || lo.Contains("replace") || lo.Contains("overwrite"))
                {
                    _logger?.LogInformation("[AGENT] Auto-handling 'replace/overwrite' dialog → Yes");
                    StatusMessage?.Invoke(this, $"🔄 Confirming overwrite: '{dlgTitle}'");
                    var clicked = _automation.ClickButtonInWindow(dlgHwnd, "Yes")
                               || _automation.ClickButtonInWindow(dlgHwnd, "Replace")
                               || _automation.ClickButtonInWindow(dlgHwnd, "Да");
                    _logger?.LogDebug("[AGENT] overwrite click: {C}", clicked);
                    await Task.Delay(300).ConfigureAwait(false);
                }
                else if (lo.Contains("error") || lo.Contains("warning") || lo.Contains("ошибка"))
                {
                    _logger?.LogInformation("[AGENT] Auto-handling error/warning dialog → OK");
                    StatusMessage?.Invoke(this, $"⚠️ Handling error dialog: '{dlgTitle}'");
                    var clicked = _automation.ClickButtonInWindow(dlgHwnd, "OK")
                               || _automation.ClickButtonInWindow(dlgHwnd, "Close")
                               || _automation.ClickButtonInWindow(dlgHwnd, "Закрыть");
                    _logger?.LogDebug("[AGENT] error dialog click: {C}", clicked);
                    await Task.Delay(300).ConfigureAwait(false);
                }
            }
            catch (Exception ex)
            {
                _logger?.LogDebug("[AGENT] HandleUnexpectedDialogAsync error: {E}", ex.Message);
            }
        }

        private List<string> EnumerateOpenWindows()
        {
            var windows = new List<string>();
            foreach (var p in Process.GetProcesses())
            {
                try
                {
                    if (!string.IsNullOrEmpty(p.MainWindowTitle) && p.MainWindowHandle != IntPtr.Zero)
                    {
                        var t = p.MainWindowTitle;
                        if (!t.Contains("AI Companion") && t != "Program Manager" && t != "Task Manager")
                            windows.Add(t);
                    }
                }
                catch { }
            }
            return windows;
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

        [JsonPropertyName("content_generated")]
        public string? ContentGenerated { get; set; }
    }

    public class ActionStepDto
    {
        [JsonPropertyName("step_number")]
        public int StepNumber { get; set; }

        [JsonPropertyName("action")]
        public string Action { get; set; } = "";

        [JsonPropertyName("target")]
        public string? Target { get; set; }

        /// <summary>
        /// Text payload from Python backend (e.g. the text for type_text action).
        /// Python sends this as "params" (a string), not "parameters" (a dict).
        /// </summary>
        [JsonPropertyName("params")]
        public string? Params { get; set; }

        [JsonPropertyName("parameters")]
        public Dictionary<string, JsonElement>? Parameters { get; set; }

        [JsonPropertyName("description")]
        public string Description { get; set; } = "";

        [JsonPropertyName("depends_on")]
        public int? DependsOn { get; set; }
    }

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
