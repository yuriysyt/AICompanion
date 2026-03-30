using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json.Serialization;
using System.Net.Http;
using System.Net.Http.Json;
using System.Threading;
using System.Threading.Tasks;
using FluentAssertions;
using Xunit;
using Xunit.Abstractions;

// ═══════════════════════════════════════════════════════════════════════════
//  AiResearchToNotepadTests
//
//  Full pipeline test: real EXE + real Granite model
//
//  AR-1  Multi-turn conversation context — model remembers previous turns
//  AR-2  Research-to-document workflow:
//          1. Launch AI Companion EXE (via AppExeFixture)
//          2. Click the Notepad quick-launch button — Notepad opens
//          3. Ask Granite to summarise a topic (3 paragraphs)
//          4. Type the AI reply into Notepad via clipboard paste
//          5. Read Notepad content back and verify it matches
//  AR-3  Conversation clear — after /api/chat/clear the model forgets
//
//  Prerequisites:
//    • Python backend running on :8000 (BackendFixture handles this)
//    • Ollama + granite3-dense:latest available
//    • AICompanion.Desktop.exe built (AppExeFixture builds it)
// ═══════════════════════════════════════════════════════════════════════════

namespace AICompanion.IntegrationTests.RealAppTests
{
    // ── DTOs ──────────────────────────────────────────────────────────────────

    public sealed class ChatHistoryItem
    {
        [JsonPropertyName("role")]    public string Role    { get; set; } = "";
        [JsonPropertyName("content")] public string Content { get; set; } = "";
    }

    public sealed class ChatWithHistoryRequest
    {
        [JsonPropertyName("message")]    public string                 Message   { get; set; } = "";
        [JsonPropertyName("session_id")] public string                 SessionId { get; set; } = "test";
        [JsonPropertyName("history")]    public List<ChatHistoryItem>  History   { get; set; } = new();
    }

    public sealed class ClearHistoryRequest
    {
        [JsonPropertyName("session_id")] public string SessionId { get; set; } = "test";
    }

    // ════════════════════════════════════════════════════════════════════════
    //  TEST CLASS — collection reuses AppExeFixture + BackendFixture
    // ════════════════════════════════════════════════════════════════════════

    [Collection("AppExe")]
    public class AiResearchToNotepadTests : IClassFixture<BackendFixture>
    {
        private readonly ITestOutputHelper _out;
        private readonly AppExeFixture     _app;
        private readonly HttpClient        _http;

        // ── Win32 ─────────────────────────────────────────────────────────────
        [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr h);
        [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
        [DllImport("user32.dll")] static extern int  GetWindowText(IntPtr h, StringBuilder sb, int n);
        [DllImport("user32.dll")] static extern bool EnumWindows(EnumWindowsProc cb, IntPtr lp);
        [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr h, out RECT r);
        [DllImport("user32.dll")] static extern bool SetCursorPos(int x, int y);
        [DllImport("user32.dll")] static extern void mouse_event(uint f, int x, int y, uint d, UIntPtr e);
        [DllImport("user32.dll")] static extern void keybd_event(byte vk, byte scan, uint flags, UIntPtr extra);
        [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr p, EnumChildProc cb, IntPtr lp);
        [DllImport("user32.dll")] static extern int  GetClassName(IntPtr h, StringBuilder sb, int n);

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT { public int L, T, R, B; }
        private delegate bool EnumWindowsProc(IntPtr h, IntPtr lp);
        private delegate bool EnumChildProc(IntPtr h, IntPtr lp);

        private const uint KEYUP              = 0x0002;
        private const uint MOUSEEVENTF_LEFT   = 0x0002;
        private const uint MOUSEEVENTF_LEFTUP = 0x0004;
        private const byte VK_CONTROL         = 0x11;
        private const byte VK_V               = 0x56;
        private const byte VK_A               = 0x41;
        private const byte VK_C               = 0x43;
        private const byte VK_ALT             = 0x12;
        private const byte VK_F4              = 0x73;

        public AiResearchToNotepadTests(AppExeFixture app, BackendFixture backend, ITestOutputHelper output)
        {
            _app  = app;
            _http = BackendFixture.Http;
            _out  = output;
        }

        // ── Helpers ───────────────────────────────────────────────────────────

        private void Banner(string title)
        {
            _out.WriteLine("");
            _out.WriteLine(new string('═', 64));
            _out.WriteLine($"  {title}");
            _out.WriteLine(new string('═', 64));
        }

        private void ClickAt(int x, int y)
        {
            SetCursorPos(x, y); Thread.Sleep(80);
            mouse_event(MOUSEEVENTF_LEFT,   x, y, 0, UIntPtr.Zero); Thread.Sleep(40);
            mouse_event(MOUSEEVENTF_LEFTUP, x, y, 0, UIntPtr.Zero); Thread.Sleep(150);
        }

        private void ClickCenter(IntPtr hwnd)
        {
            GetWindowRect(hwnd, out var r);
            ClickAt((r.L + r.R) / 2, (r.T + r.B) / 2);
        }

        private IntPtr WaitForWindow(string titleFragment, int timeoutMs = 10000)
        {
            var deadline = DateTime.UtcNow.AddMilliseconds(timeoutMs);
            IntPtr hwnd  = IntPtr.Zero;
            while (DateTime.UtcNow < deadline && hwnd == IntPtr.Zero)
            {
                Thread.Sleep(400);
                EnumWindows((h, _) =>
                {
                    if (!IsWindowVisible(h)) return true;
                    var sb = new StringBuilder(256);
                    GetWindowText(h, sb, 256);
                    if (sb.ToString().Contains(titleFragment, StringComparison.OrdinalIgnoreCase))
                    { hwnd = h; return false; }
                    return true;
                }, IntPtr.Zero);
            }
            return hwnd;
        }

        /// Snapshot all existing windows matching titleFragment so we can detect NEW ones later.
        private HashSet<IntPtr> SnapshotWindows(string titleFragment)
        {
            var set = new HashSet<IntPtr>();
            EnumWindows((h, _) =>
            {
                if (!IsWindowVisible(h)) return true;
                var sb = new StringBuilder(256);
                GetWindowText(h, sb, 256);
                if (sb.ToString().Contains(titleFragment, StringComparison.OrdinalIgnoreCase))
                    set.Add(h);
                return true;
            }, IntPtr.Zero);
            return set;
        }

        /// Wait for a window that was NOT in <paramref name="existing"/> snapshot.
        private IntPtr WaitForNewWindow(string titleFragment, HashSet<IntPtr> existing, int timeoutMs = 12000)
        {
            var deadline = DateTime.UtcNow.AddMilliseconds(timeoutMs);
            IntPtr hwnd  = IntPtr.Zero;
            while (DateTime.UtcNow < deadline && hwnd == IntPtr.Zero)
            {
                Thread.Sleep(400);
                EnumWindows((h, _) =>
                {
                    if (!IsWindowVisible(h)) return true;
                    if (existing.Contains(h))  return true;   // already existed
                    var sb = new StringBuilder(256);
                    GetWindowText(h, sb, 256);
                    if (sb.ToString().Contains(titleFragment, StringComparison.OrdinalIgnoreCase))
                    { hwnd = h; return false; }
                    return true;
                }, IntPtr.Zero);
            }
            return hwnd;
        }

        private void CloseWindow(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero) return;
            SetForegroundWindow(hwnd); Thread.Sleep(200);
            keybd_event(VK_ALT, 0, 0, UIntPtr.Zero); Thread.Sleep(20);
            keybd_event(VK_F4,  0, 0, UIntPtr.Zero); Thread.Sleep(30);
            keybd_event(VK_F4,  0, KEYUP, UIntPtr.Zero); Thread.Sleep(20);
            keybd_event(VK_ALT, 0, KEYUP, UIntPtr.Zero); Thread.Sleep(400);
        }

        /// Paste text into Notepad via clipboard (works on Russian keyboard layout).
        /// No ClickCenter — SetForegroundWindow alone gives keyboard focus to Notepad's
        /// text area without accidentally clicking the Win11 toolbar/tab strip.
        private void TypeViaClipboard(IntPtr hwnd, string text)
        {
            SetForegroundWindow(hwnd); Thread.Sleep(350);
            // Click in the lower 60 % of the Notepad window to land in the text area,
            // well below Win11's tab bar / toolbar (≈ top 80 px).
            GetWindowRect(hwnd, out var r);
            int cx = (r.L + r.R) / 2;
            int cy = r.T + (int)((r.B - r.T) * 0.65);
            ClickAt(cx, cy);           Thread.Sleep(200);

            var sta = new Thread(() => System.Windows.Forms.Clipboard.SetText(text));
            sta.SetApartmentState(ApartmentState.STA); sta.Start(); sta.Join();
            Thread.Sleep(120);
            keybd_event(VK_CONTROL, 0, 0,     UIntPtr.Zero);
            keybd_event(VK_V,       0, 0,     UIntPtr.Zero); Thread.Sleep(60);
            keybd_event(VK_V,       0, KEYUP, UIntPtr.Zero);
            keybd_event(VK_CONTROL, 0, KEYUP, UIntPtr.Zero);
            Thread.Sleep(500);
        }

        /// Select-all + copy, then return clipboard text.
        private string ReadNotepadContent(IntPtr hwnd)
        {
            // Re-focus in the text area (same 65 % heuristic as TypeViaClipboard)
            SetForegroundWindow(hwnd); Thread.Sleep(250);
            GetWindowRect(hwnd, out var r);
            int cx = (r.L + r.R) / 2;
            int cy = r.T + (int)((r.B - r.T) * 0.65);
            ClickAt(cx, cy);           Thread.Sleep(200);

            // Ctrl+A — select all text
            keybd_event(VK_CONTROL, 0, 0,     UIntPtr.Zero);
            keybd_event(VK_A,       0, 0,     UIntPtr.Zero); Thread.Sleep(50);
            keybd_event(VK_A,       0, KEYUP, UIntPtr.Zero);
            keybd_event(VK_CONTROL, 0, KEYUP, UIntPtr.Zero); Thread.Sleep(250);
            // Ctrl+C — copy to clipboard
            keybd_event(VK_CONTROL, 0, 0,     UIntPtr.Zero);
            keybd_event(VK_C,       0, 0,     UIntPtr.Zero); Thread.Sleep(50);
            keybd_event(VK_C,       0, KEYUP, UIntPtr.Zero);
            keybd_event(VK_CONTROL, 0, KEYUP, UIntPtr.Zero); Thread.Sleep(600);

            string text = "";
            var t = new Thread(() => { try { text = System.Windows.Forms.Clipboard.GetText(); } catch { } });
            t.SetApartmentState(ApartmentState.STA); t.Start(); t.Join(2000);
            return text;
        }

        /// Click the quick-launch Notepad button using the same heuristic as AppExeE2ETests.
        private void ClickQuickNotepad()
        {
            SetForegroundWindow(_app.MainHwnd); Thread.Sleep(300);

            // Try to find the Notepad child button first
            IntPtr btnHwnd = IntPtr.Zero;
            EnumChildWindows(_app.MainHwnd, (h, _) =>
            {
                var cls = new StringBuilder(64);
                GetClassName(h, cls, 64);
                if (!cls.ToString().Equals("Button", StringComparison.OrdinalIgnoreCase)) return true;
                var sb = new StringBuilder(256);
                GetWindowText(h, sb, 256);
                if (sb.ToString().Contains("Notepad", StringComparison.OrdinalIgnoreCase))
                { btnHwnd = h; return false; }
                return true;
            }, IntPtr.Zero);

            if (btnHwnd != IntPtr.Zero)
            {
                GetWindowRect(btnHwnd, out var br);
                _out.WriteLine($"  Found Notepad button at ({br.L},{br.T})");
                ClickAt((br.L + br.R) / 2, (br.T + br.B) / 2);
            }
            else
            {
                // Heuristic fallback (same layout as AppExeE2ETests)
                GetWindowRect(_app.MainHwnd, out var winRect);
                int x = winRect.L + 35 + 1 * 58;  // index 1 = Notepad
                int y = winRect.T + 450;
                _out.WriteLine($"  Heuristic Notepad click at ({x},{y})");
                ClickAt(x, y);
            }
        }

        // ════════════════════════════════════════════════════════════════════
        //  AR-1 — Conversation context: model remembers previous turns
        // ════════════════════════════════════════════════════════════════════

        [Fact]
        public async Task Ai_ConversationContext_RemembersPreviousTurns()
        {
            Banner("TEST AR-1: Conversation context — model remembers previous turns");

            const string sessionId = "ar1_context_test";

            // Turn 1 — introduce a custom fact
            var r1 = await _http.PostAsJsonAsync($"{BackendFixture.BackendUrl}/api/chat",
                new ChatWithHistoryRequest
                {
                    Message   = "Remember: my favourite programming language is Rust.",
                    SessionId = sessionId,
                });
            r1.EnsureSuccessStatusCode();
            var reply1 = (await r1.Content.ReadFromJsonAsync<ChatResponse>())!;
            _out.WriteLine($"  Turn 1 reply ({reply1.LatencyMs} ms): {reply1.Reply}");

            // Turn 2 — test whether the model retained the fact
            var r2 = await _http.PostAsJsonAsync($"{BackendFixture.BackendUrl}/api/chat",
                new ChatWithHistoryRequest
                {
                    Message   = "What is my favourite programming language?",
                    SessionId = sessionId,
                });
            r2.EnsureSuccessStatusCode();
            var reply2 = (await r2.Content.ReadFromJsonAsync<ChatResponse>())!;
            _out.WriteLine($"  Turn 2 reply ({reply2.LatencyMs} ms): {reply2.Reply}");

            reply2.Reply.Should().NotBeNullOrWhiteSpace(
                "the model must produce a reply to a follow-up question");

            // The word "Rust" should appear — it was stated in turn 1
            reply2.Reply.Should().ContainEquivalentOf("Rust",
                because: "the model must remember the fact from the previous turn");

            _out.WriteLine("  Context memory works across turns ✓");

            // Clean up session so it doesn't pollute other tests
            await _http.PostAsJsonAsync($"{BackendFixture.BackendUrl}/api/chat/clear",
                new ClearHistoryRequest { SessionId = sessionId });

            _out.WriteLine("✅ AR-1 PASSED");
        }

        // ════════════════════════════════════════════════════════════════════
        //  AR-2 — Research-to-document: AI summarises topic → types into Notepad
        // ════════════════════════════════════════════════════════════════════

        [Fact]
        public async Task Ai_ResearchAndWriteToNotepad_FullPipeline()
        {
            Banner("TEST AR-2: AI research summary → typed into Notepad");

            // ── Step 1: ensure AI Companion is running ────────────────────────
            _app.MainHwnd.Should().NotBe(IntPtr.Zero,
                "AI Companion must be running before this test");
            _out.WriteLine($"  AI Companion hwnd = 0x{_app.MainHwnd:X}");

            // ── Step 2: click the Notepad quick-launch button ─────────────────
            // Snapshot existing Notepad windows BEFORE clicking so we can identify the NEW one
            var existingNotepads = SnapshotWindows("Notepad");
            existingNotepads.UnionWith(SnapshotWindows("Блокнот"));
            _out.WriteLine($"  Existing Notepad windows: {existingNotepads.Count}");

            _out.WriteLine("  Clicking Notepad quick-launch button…");
            ClickQuickNotepad();

            // Wait for a NEW Notepad window (not one that was open before)
            var notepadHwnd = WaitForNewWindow("Notepad", existingNotepads, 12000);
            if (notepadHwnd == IntPtr.Zero)
                notepadHwnd = WaitForNewWindow("Блокнот", existingNotepads, 3000);

            notepadHwnd.Should().NotBe(IntPtr.Zero,
                "clicking the Notepad button must open a NEW Notepad window");
            _out.WriteLine($"  Notepad hwnd = 0x{notepadHwnd:X}");

            try
            {
                // ── Step 3: ask Granite to summarise a topic ──────────────────
                const string topic     = "artificial intelligence";
                const string sessionId = "ar2_research_test";

                _out.WriteLine($"  Asking Granite to summarise: \"{topic}\"…");

                var chatResp = await _http.PostAsJsonAsync(
                    $"{BackendFixture.BackendUrl}/api/chat",
                    new ChatWithHistoryRequest
                    {
                        Message   = $"Write a concise 3-paragraph summary about {topic} " +
                                    "suitable for a student. No headers, plain text only.",
                        SessionId = sessionId,
                    });
                chatResp.EnsureSuccessStatusCode();
                var chat = (await chatResp.Content.ReadFromJsonAsync<ChatResponse>())!;

                _out.WriteLine($"  Granite replied in {chat.LatencyMs} ms ({chat.Reply.Length} chars):");
                foreach (var line in chat.Reply.Split('\n'))
                    _out.WriteLine($"    {line}");

                chat.Reply.Should().NotBeNullOrWhiteSpace(
                    "Granite must return a non-empty summary");
                chat.Reply.Length.Should().BeGreaterThan(80,
                    "a 3-paragraph summary must be more than 80 characters");

                // ── Step 4: type the AI reply into Notepad ────────────────────
                _out.WriteLine("  Typing AI summary into Notepad via clipboard paste…");
                TypeViaClipboard(notepadHwnd, chat.Reply);
                Thread.Sleep(500);

                // ── Step 5: read Notepad content and verify ───────────────────
                _out.WriteLine("  Reading Notepad content back…");
                var notepadContent = ReadNotepadContent(notepadHwnd);

                _out.WriteLine($"  Notepad content ({notepadContent.Length} chars):");
                foreach (var line in notepadContent.Split('\n'))
                    _out.WriteLine($"    {line}");

                notepadContent.Should().NotBeNullOrWhiteSpace(
                    "Notepad must contain the pasted text");

                // At least the first 40 characters of the reply should appear in Notepad
                var expectedStart = chat.Reply.Substring(0, Math.Min(40, chat.Reply.Length)).Trim();
                notepadContent.Should().Contain(expectedStart,
                    because: "the AI summary should be faithfully pasted into Notepad");

                _out.WriteLine("  Content verified in Notepad ✓");

                // Clean up session
                await _http.PostAsJsonAsync($"{BackendFixture.BackendUrl}/api/chat/clear",
                    new ClearHistoryRequest { SessionId = sessionId });
            }
            finally
            {
                // ── Cleanup: close Notepad (discard changes) ──────────────────
                _out.WriteLine("  Closing Notepad (no save)…");
                CloseWindow(notepadHwnd);
                // Dismiss "do you want to save?" dialog if it appears
                Thread.Sleep(600);
                var saveDialog = WaitForWindow("Notepad", 1500);
                if (saveDialog != IntPtr.Zero && saveDialog != notepadHwnd)
                {
                    // Tab to "Don't Save" and press Enter, or just close
                    CloseWindow(saveDialog);
                }
            }

            _out.WriteLine("✅ AR-2 PASSED");
        }

        // ════════════════════════════════════════════════════════════════════
        //  AR-3 — Session clear: history is forgotten after /api/chat/clear
        // ════════════════════════════════════════════════════════════════════

        [Fact]
        public async Task Ai_ConversationClear_ForgetsHistory()
        {
            Banner("TEST AR-3: /api/chat/clear — model forgets previous turns");

            const string sessionId = "ar3_clear_test";

            // Establish a fact
            var r1 = await _http.PostAsJsonAsync($"{BackendFixture.BackendUrl}/api/chat",
                new ChatWithHistoryRequest
                {
                    Message   = "My secret code word is BANANA.",
                    SessionId = sessionId,
                });
            r1.EnsureSuccessStatusCode();
            var reply1 = (await r1.Content.ReadFromJsonAsync<ChatResponse>())!;
            _out.WriteLine($"  Turn 1: {reply1.Reply}");

            // Clear history
            var clearResp = await _http.PostAsJsonAsync(
                $"{BackendFixture.BackendUrl}/api/chat/clear",
                new ClearHistoryRequest { SessionId = sessionId });
            clearResp.EnsureSuccessStatusCode();
            _out.WriteLine("  History cleared ✓");

            // Ask without re-supplying history — model must not recall the fact
            var r2 = await _http.PostAsJsonAsync($"{BackendFixture.BackendUrl}/api/chat",
                new ChatWithHistoryRequest
                {
                    Message   = "What was my secret code word?",
                    SessionId = sessionId,   // same session, but history was cleared
                });
            r2.EnsureSuccessStatusCode();
            var reply2 = (await r2.Content.ReadFromJsonAsync<ChatResponse>())!;
            _out.WriteLine($"  Turn 2 (post-clear): {reply2.Reply}");

            reply2.Reply.Should().NotBeNullOrWhiteSpace();
            // After clearing, the model should NOT know "BANANA"
            reply2.Reply.Should().NotContainEquivalentOf("BANANA",
                because: "history was cleared so the model cannot know the secret word");

            // Final cleanup
            await _http.PostAsJsonAsync($"{BackendFixture.BackendUrl}/api/chat/clear",
                new ClearHistoryRequest { SessionId = sessionId });

            _out.WriteLine("✅ AR-3 PASSED");
        }
    }
}
