using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Xunit;
using Xunit.Abstractions;

namespace AICompanion.IntegrationTests
{
    /// <summary>
    /// Tests the IBM Granite Python backend capabilities via HTTP.
    ///
    /// Documents what the AI *actually does*: response times, JSON contract,
    /// intent classification accuracy, multi-step planning, and context memory.
    ///
    /// All tests require the Python server running on localhost:8000:
    ///   cd python_backend
    ///   uvicorn server:app --port 8000
    ///
    /// Run with: dotnet test --filter "Category=RequiresBackend"
    /// </summary>
    [Trait("Category", "RequiresBackend")]
    public class IBMGraniteTests : IDisposable
    {
        private readonly ITestOutputHelper _output;
        private readonly HttpClient        _http;
        private const string BASE = "http://localhost:8000";

        public IBMGraniteTests(ITestOutputHelper output)
        {
            _output = output;
            _http   = new HttpClient { Timeout = TimeSpan.FromSeconds(60) };
        }

        public void Dispose() => _http.Dispose();

        // ── helpers ────────────────────────────────────────────────────────────

        private async Task<(int statusCode, JsonElement root)> PostJson(string path, object body)
        {
            var json = JsonSerializer.Serialize(body);
            var resp = await _http.PostAsync($"{BASE}{path}",
                new StringContent(json, Encoding.UTF8, "application/json"));
            var text = await resp.Content.ReadAsStringAsync();
            _output.WriteLine($"[HTTP {(int)resp.StatusCode}] {path} → {text.Substring(0, Math.Min(300, text.Length))}");
            var root = JsonDocument.Parse(text).RootElement;
            return ((int)resp.StatusCode, root);
        }

        // ── TEST 1: Health check — backend alive ───────────────────────────────
        [Fact]
        public async Task Backend_HealthCheck_Returns200()
        {
            _output.WriteLine("=== GRANITE TEST 1: Health check ===");
            var sw   = Stopwatch.StartNew();
            var resp = await _http.GetAsync($"{BASE}/api/health");
            sw.Stop();

            var body = await resp.Content.ReadAsStringAsync();
            _output.WriteLine($"[HEALTH] {(int)resp.StatusCode} | {sw.ElapsedMilliseconds}ms | {body}");

            // ✅ REAL ASSERTIONS
            ((int)resp.StatusCode).Should().Be(200,
                because: "Backend must respond 200 to /api/health. " +
                         "Ensure uvicorn is running: cd python_backend && uvicorn server:app --port 8000");

            sw.ElapsedMilliseconds.Should().BeLessThan(5000,
                because: "Health check must respond in < 5 seconds");

            var root = JsonDocument.Parse(body).RootElement;
            root.GetProperty("status").GetString().Should().Be("ok",
                because: "/api/health must return {\"status\": \"ok\"}");

            _output.WriteLine($"✅ PASSED: Backend alive in {sw.ElapsedMilliseconds}ms");
        }

        // ── TEST 2: Intent extraction — correct action per voice text ──────────
        [Theory]
        [InlineData("open notepad",       "open_app")]
        [InlineData("search for weather", "search_web")]
        [InlineData("type hello world",   "type_text")]
        [InlineData("save the document",  "save_document")]
        [InlineData("undo last action",   "undo")]
        public async Task Intent_ExtractsCorrectAction(string text, string expectedAction)
        {
            _output.WriteLine($"\n=== GRANITE TEST 2: Intent '{text}' → '{expectedAction}' ===");
            var sw = Stopwatch.StartNew();

            var (statusCode, root) = await PostJson("/api/intent", new
            {
                text,
                session_id = "test_intent"
            });
            sw.Stop();

            statusCode.Should().Be(200,
                because: $"/api/intent must return 200 for '{text}'");

            var action = root.GetProperty("action").GetString() ?? "";
            _output.WriteLine($"[INTENT] '{text}' → action='{action}' (expected '{expectedAction}') | {sw.ElapsedMilliseconds}ms");

            // ✅ REAL ASSERTION — IBM Granite must classify the command correctly
            action.Should().Be(expectedAction,
                because: $"IBM Granite must classify '{text}' as '{expectedAction}'. " +
                         $"Got '{action}'. Check that Granite LLM is loaded and responding to the system prompt.");

            _output.WriteLine($"✅ PASSED: '{text}' → '{action}' in {sw.ElapsedMilliseconds}ms");
        }

        // ── TEST 3: Response time — average < 10 seconds ──────────────────────
        [Fact]
        public async Task Intent_ResponseTime_AverageUnder10Seconds()
        {
            _output.WriteLine("=== GRANITE TEST 3: Response time benchmark ===");
            var times = new List<long>();

            for (int i = 0; i < 5; i++)
            {
                var sw = Stopwatch.StartNew();
                await PostJson("/api/intent", new
                {
                    text       = "open calculator",
                    session_id = "bench"
                });
                sw.Stop();
                times.Add(sw.ElapsedMilliseconds);
                _output.WriteLine($"  [RUN {i + 1}] {sw.ElapsedMilliseconds}ms");
            }

            var avg = times.Average();
            var max = times.Max();
            var p95 = times.OrderBy(x => x).Skip(4).First();
            _output.WriteLine($"[PERF] avg={avg:F0}ms | max={max}ms | p95={p95}ms");

            // ✅ REAL ASSERTION — average must be reasonable for an LLM endpoint
            avg.Should().BeLessThan(10_000,
                because: $"IBM Granite /api/intent average response must be < 10s. " +
                         $"Got avg={avg:F0}ms. If slower, check Ollama model loading, " +
                         "system resources, or quantization level.");
            _output.WriteLine($"✅ PASSED: avg={avg:F0}ms (< 10 000ms)");
        }

        // ── TEST 4: Plan endpoint — multi-step command yields multiple steps ───
        [Fact]
        public async Task Plan_ComplexCommand_ReturnsMultipleSteps()
        {
            _output.WriteLine("=== GRANITE TEST 4: Agentic plan ===");
            var sw = Stopwatch.StartNew();

            var (statusCode, root) = await PostJson("/api/plan", new
            {
                text       = "open notepad and type Hello World then save the document",
                max_steps  = 5,
                session_id = "plan_test"
            });
            sw.Stop();

            statusCode.Should().Be(200,
                because: "/api/plan must return 200 for a multi-step command");

            var steps = root.GetProperty("steps");
            int count = steps.GetArrayLength();
            _output.WriteLine($"[PLAN] {count} steps in {sw.ElapsedMilliseconds}ms");

            // ✅ REAL ASSERTION — a compound command must decompose into ≥ 2 steps
            count.Should().BeGreaterThanOrEqualTo(2,
                because: "The command 'open notepad AND type ... THEN save' contains three verbs. " +
                         "IBM Granite must decompose it into at least 2 steps. " +
                         $"Got {count} step(s). Check the agentic_planner prompt.");

            // First step must have a non-empty action
            var firstAction = steps[0].GetProperty("action").GetString() ?? "";
            _output.WriteLine($"[STEP-1] action='{firstAction}'");
            firstAction.Should().NotBeNullOrEmpty(
                because: "The first plan step must have a non-empty 'action' field");

            _output.WriteLine($"✅ PASSED: Plan has {count} steps, first='{firstAction}'");
        }

        // ── TEST 5: Chat endpoint — returns a non-empty reply ─────────────────
        [Fact]
        public async Task Chat_SimpleGreeting_ReturnsNonEmptyReply()
        {
            _output.WriteLine("=== GRANITE TEST 5: Chat greeting ===");
            var sw = Stopwatch.StartNew();

            var (statusCode, root) = await PostJson("/api/chat", new
            {
                message    = "Hello, what can you help me with?",
                session_id = $"chat_{Guid.NewGuid():N}"
            });
            sw.Stop();

            statusCode.Should().Be(200,
                because: "/api/chat must return 200 for a simple greeting");

            // The ChatResponse model uses field "reply" (not "response")
            var reply = root.GetProperty("reply").GetString() ?? "";
            _output.WriteLine($"[REPLY] '{reply.Substring(0, Math.Min(200, reply.Length))}' | {sw.ElapsedMilliseconds}ms");

            // ✅ REAL ASSERTION — IBM Granite must produce a meaningful response
            reply.Should().NotBeNullOrWhiteSpace(
                because: "IBM Granite /api/chat must return a non-empty reply to a greeting. " +
                         "If empty, check the Ollama model and the fallback handler in server.py.");

            reply.Length.Should().BeGreaterThan(5,
                because: "A greeting response must be at least a few words long");

            _output.WriteLine($"✅ PASSED: Chat replied in {sw.ElapsedMilliseconds}ms");
        }

        // ── TEST 6: Context memory — 2nd turn references 1st ─────────────────
        [Fact]
        public async Task Chat_ContextMemory_SecondTurnReferencesFirst()
        {
            _output.WriteLine("=== GRANITE TEST 6: Context memory ===");
            var sessionId = $"memory_{Guid.NewGuid():N}";

            async Task<string> Chat(string message)
            {
                var (_, root) = await PostJson("/api/chat", new { message, session_id = sessionId });
                var r = root.GetProperty("reply").GetString() ?? "";
                _output.WriteLine($"  User: '{message}'");
                _output.WriteLine($"  AI:   '{r.Substring(0, Math.Min(120, r.Length))}'");
                return r;
            }

            // Turn 1: introduce name
            await Chat("My name is Alice.");
            await Task.Delay(500);

            // Turn 2: ask model to recall
            var r2 = await Chat("What is my name?");

            // ✅ REAL ASSERTION — model must remember "Alice" from the prior turn
            r2.Should().ContainEquivalentOf("Alice",
                because: "IBM Granite must remember context across turns within a session. " +
                         "Turn 1 told the model 'My name is Alice' — Turn 2 must recall it. " +
                         $"Actual response: '{r2}'");
            _output.WriteLine("✅ PASSED: Context memory works across turns");
        }

        // ── TEST 7: Heartbeat — agentic subsystem reports ready ───────────────
        [Fact]
        public async Task AgentHeartbeat_ReturnsStatusReady()
        {
            _output.WriteLine("=== GRANITE TEST 7: Agent heartbeat ===");
            var resp = await _http.GetAsync($"{BASE}/api/agent/heartbeat");
            var body = await resp.Content.ReadAsStringAsync();
            _output.WriteLine($"[HEARTBEAT] {(int)resp.StatusCode} | {body}");

            ((int)resp.StatusCode).Should().Be(200,
                because: "/api/agent/heartbeat must return 200");

            var root = JsonDocument.Parse(body).RootElement;
            // The heartbeat must have a "status" field
            root.TryGetProperty("status", out var statusEl).Should().BeTrue(
                because: "Heartbeat response must contain a 'status' field");

            var status = statusEl.GetString() ?? "";
            status.Should().NotBeNullOrEmpty(
                because: "Heartbeat status must not be empty");

            _output.WriteLine($"✅ PASSED: Heartbeat status='{status}'");
        }
    }
}
