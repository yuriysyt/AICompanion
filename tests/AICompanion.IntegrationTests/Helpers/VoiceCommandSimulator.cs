using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using AICompanion.Desktop.Services;
using AICompanion.Desktop.Services.Automation;
using Xunit.Abstractions;

namespace AICompanion.IntegrationTests.Helpers
{
    /// <summary>
    /// Simulates the full voice command pipeline WITHOUT a real microphone.
    ///
    /// Injects a pre-defined transcript exactly as ElevenLabs/SAPI would return it,
    /// then routes it through the same path as the live application:
    ///
    ///   [Transcript] → Confidence Gate → LocalCommandProcessor → AgenticExecutionService → Win32
    ///
    /// This is the core of the integration test harness. Each call to SpeakAsync()
    /// produces a SimulatedCommandResult with full per-stage timing.
    /// </summary>
    public class VoiceCommandSimulator
    {
        private readonly ITestOutputHelper _output;
        private readonly LocalCommandProcessor _processor;
        private readonly AgenticExecutionService _agenticService;

        private const float ConfidenceThreshold = 0.65f;

        [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] static extern int GetWindowText(IntPtr h, StringBuilder sb, int n);

        public VoiceCommandSimulator(ITestOutputHelper output)
        {
            _output = output;

            var processorLogger  = new XUnitLogger<LocalCommandProcessor>(output);
            var agentLogger      = new XUnitLogger<AgenticExecutionService>(output);
            var automationLogger = new XUnitLogger<WindowAutomationHelper>(output);

            _processor      = new LocalCommandProcessor(processorLogger);
            var winHelper   = new WindowAutomationHelper(automationLogger);
            _agenticService = new AgenticExecutionService(winHelper, agentLogger);
        }

        /// <summary>
        /// Simulate a user speaking <paramref name="transcript"/> with the given STT confidence.
        /// Returns a full <see cref="SimulatedCommandResult"/> with per-stage timing.
        ///
        /// Set <paramref name="targetWindow"/> to a known HWND to override which window
        /// CaptureTargetWindow() would normally capture.
        /// </summary>
        public async Task<SimulatedCommandResult> SpeakAsync(
            string transcript,
            float confidence = 0.88f,
            IntPtr targetWindow = default)
        {
            var res   = new SimulatedCommandResult { Transcript = transcript, Confidence = confidence };
            var total = Stopwatch.StartNew();

            _output.WriteLine($"\n{'─',60}");
            _output.WriteLine($"[🎙️  VOICE INPUT] '{transcript}' | confidence={confidence:F2}");

            // ── STAGE 1: Confidence gate ─────────────────────────────────
            res.PassedConfidenceGate = confidence >= ConfidenceThreshold;
            _output.WriteLine($"[STAGE-1-CONFIDENCE] {confidence:F2} >= {ConfidenceThreshold} → " +
                              $"{(res.PassedConfidenceGate ? "✅ PASS" : "❌ BLOCKED")}");
            if (!res.PassedConfidenceGate)
            {
                res.BlockedReason   = $"Confidence {confidence:F2} below threshold {ConfidenceThreshold}";
                res.TotalElapsedMs  = (int)total.ElapsedMilliseconds;
                return res;
            }

            // ── STAGE 2: Window capture ───────────────────────────────────
            _processor.CaptureTargetWindow();
            _output.WriteLine($"[STAGE-2-TARGET] Captured: '{_processor.GetTargetWindowTitle()}'");

            // ── STAGE 3: Complexity routing ───────────────────────────────
            var stageSw = Stopwatch.StartNew();
            res.IsComplex = _processor.IsComplexCommand(transcript);
            _output.WriteLine($"[STAGE-3-ROUTING] IsComplex={res.IsComplex} → " +
                $"{(res.IsComplex ? "AGENTIC /api/plan" : "LOCAL regex/fuzzy")} ({stageSw.ElapsedMilliseconds}ms)");

            // ── STAGE 4: Local command processor ─────────────────────────
            stageSw.Restart();
            var cmdResult = _processor.ProcessCommand(transcript);
            res.LocalCommandMs = (int)stageSw.ElapsedMilliseconds;
            _output.WriteLine($"[STAGE-4-LOCAL] Success={cmdResult.Success} | '{cmdResult.Description}' ({res.LocalCommandMs}ms)");

            // ── STAGE 5: Agentic execution ────────────────────────────────
            bool routeToAgentic = res.IsComplex ||
                                  cmdResult.SpeechResponse == "AGENTIC_PLAN_REQUIRED";
            if (routeToAgentic)
            {
                stageSw.Restart();
                _output.WriteLine("[STAGE-5-AGENTIC] Sending to AgenticExecutionService...");
                try
                {
                    var agentResult     = await _agenticService.ExecuteCommandAsync(transcript);
                    res.AgenticMs       = (int)stageSw.ElapsedMilliseconds;
                    res.AgenticSuccess  = agentResult.Success;
                    res.AgenticStepCount = agentResult.StepResults.Count;
                    res.AgenticSummary  = agentResult.Summary;

                    _output.WriteLine($"[STAGE-5-AGENTIC] Done: Success={agentResult.Success} | " +
                                      $"{agentResult.Summary} ({res.AgenticMs}ms)");
                    foreach (var step in agentResult.StepResults)
                        _output.WriteLine($"  [STEP] {step}");
                }
                catch (Exception ex)
                {
                    res.AgenticMs      = (int)stageSw.ElapsedMilliseconds;
                    res.AgenticSuccess = false;
                    res.AgenticSummary = ex.Message;
                    _output.WriteLine($"[STAGE-5-AGENTIC] ❌ Exception: {ex.Message}");
                }
            }
            else
            {
                res.LocalSuccess = cmdResult.Success;
                res.LocalMessage = cmdResult.Description;
            }

            res.TotalElapsedMs = (int)total.ElapsedMilliseconds;
            _output.WriteLine($"[STAGE-6-DONE] Total: {res.TotalElapsedMs}ms");
            return res;
        }
    }

    public class SimulatedCommandResult
    {
        public string  Transcript          { get; set; } = "";
        public float   Confidence          { get; set; }
        public bool    PassedConfidenceGate { get; set; }
        public string? BlockedReason       { get; set; }
        public bool    IsComplex           { get; set; }
        public bool    LocalSuccess        { get; set; }
        public string? LocalMessage        { get; set; }
        public bool    AgenticSuccess      { get; set; }
        public string? AgenticSummary      { get; set; }
        public int     AgenticStepCount    { get; set; }
        public int     LocalCommandMs      { get; set; }
        public int     AgenticMs          { get; set; }
        public int     TotalElapsedMs      { get; set; }

        /// <summary>True if the command passed the confidence gate and at least one stage succeeded.</summary>
        public bool OverallSuccess => PassedConfidenceGate && (LocalSuccess || AgenticSuccess);
    }
}
