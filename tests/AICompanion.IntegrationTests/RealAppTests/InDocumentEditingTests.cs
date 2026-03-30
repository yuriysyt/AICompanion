using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using AICompanion.IntegrationTests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace AICompanion.IntegrationTests.RealAppTests
{
    /// <summary>
    /// REAL E2E TESTS: In-document precision editing using WordPad
    /// (always present on Windows — no install check needed).
    ///
    /// Verifies: word deletion, select-and-replace, copy-paste, multi-step undo.
    /// Every test reads back actual document content via clipboard to assert real state.
    ///
    /// Tests 1–4: require Python backend for "type …" voice commands (/api/intent).
    ///
    /// ⚠️  Opens a WordPad window on screen. Do NOT touch the desktop while tests run.
    /// </summary>
    [Collection("RealApp")]
    public class InDocumentEditingTests : IAsyncLifetime
    {
        private readonly ITestOutputHelper     _output;
        private readonly AppLauncher           _launcher;
        private readonly VoiceCommandSimulator _sim;
        private bool _wordPadAvailable;

        // ── P/Invoke ──────────────────────────────────────────────────────────
        [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr h);
        [DllImport("user32.dll")] static extern void keybd_event(byte vk, byte scan, uint flags, UIntPtr extra);

        private const uint KEYUP = 0x0002;

        private const byte VK_CONTROL = 0x11;
        private const byte VK_A       = 0x41;
        private const byte VK_C       = 0x43;
        private const byte VK_V       = 0x56;
        private const byte VK_Z       = 0x5A;
        private const byte VK_DELETE  = 0x2E;
        private const byte VK_BACK    = 0x08;
        private const byte VK_END     = 0x23;

        public InDocumentEditingTests(ITestOutputHelper output)
        {
            _output   = output;
            _launcher = new AppLauncher(output);
            _sim      = new VoiceCommandSimulator(output);
        }

        public async Task InitializeAsync()
        {
            _output.WriteLine("=== TEST SETUP: Launching WordPad ===");
            _wordPadAvailable = await _launcher.LaunchAsync("wordpad", 8000);
            if (!_wordPadAvailable)
            {
                _output.WriteLine("[SKIP] WordPad not available on this machine (removed in Windows 11 22H2+) — tests will skip");
                return;
            }
            await Task.Delay(1500);
            _launcher.Focus();
        }

        public Task DisposeAsync()
        {
            _output.WriteLine("=== TEST TEARDOWN: Closing WordPad ===");
            _launcher.Dispose();
            return Task.CompletedTask;
        }

        private bool Skip() { if (!_wordPadAvailable) _output.WriteLine("[SKIP] WordPad not available"); return !_wordPadAvailable; }

        // ── helpers ────────────────────────────────────────────────────────────

        private void Combo(byte mod, byte vk)
        {
            keybd_event(mod, 0, 0, UIntPtr.Zero);
            Thread.Sleep(20);
            keybd_event(vk, 0, 0, UIntPtr.Zero);
            Thread.Sleep(30);
            keybd_event(vk, 0, KEYUP, UIntPtr.Zero);
            Thread.Sleep(20);
            keybd_event(mod, 0, KEYUP, UIntPtr.Zero);
            Thread.Sleep(50);
        }

        private void Press(byte vk)
        {
            keybd_event(vk, 0, 0, UIntPtr.Zero);
            Thread.Sleep(30);
            keybd_event(vk, 0, KEYUP, UIntPtr.Zero);
            Thread.Sleep(30);
        }

        /// <summary>Ctrl+A → Ctrl+C → STA clipboard read.</summary>
        private string ReadContent()
        {
            _output.WriteLine("[READ] Reading WordPad content via clipboard");
            string text = string.Empty;
            for (int attempt = 0; attempt < 3 && text.Length == 0; attempt++)
            {
                if (attempt > 0)
                {
                    _output.WriteLine($"[READ] Retry {attempt}/2");
                    Thread.Sleep(400);
                }
                SetForegroundWindow(_launcher.WindowHandle);
                Thread.Sleep(300);
                Combo(VK_CONTROL, VK_A);
                Thread.Sleep(200);
                Combo(VK_CONTROL, VK_C);
                Thread.Sleep(700);

                var t = new Thread(() =>
                {
                    try   { text = System.Windows.Forms.Clipboard.GetText(); }
                    catch { text = string.Empty; }
                });
                t.SetApartmentState(ApartmentState.STA);
                t.Start();
                t.Join(2000);
            }
            _output.WriteLine($"[READ] {text.Length} chars: '{text.TrimEnd().Substring(0, Math.Min(120, text.TrimEnd().Length))}'");
            return text;
        }

        /// <summary>Clear document, verify it's empty before each test.</summary>
        private async Task ClearDocument()
        {
            SetForegroundWindow(_launcher.WindowHandle);
            Thread.Sleep(200);
            Combo(VK_CONTROL, VK_A);
            Thread.Sleep(100);
            Press(VK_DELETE);
            await Task.Delay(300);
        }

        private static int CountOccurrences(string text, string pattern)
        {
            int count = 0, idx = 0;
            while ((idx = text.IndexOf(pattern, idx, StringComparison.OrdinalIgnoreCase)) >= 0)
            { count++; idx += pattern.Length; }
            return count;
        }

        // ── TEST 1: Delete last word (Ctrl+Backspace) ──────────────────────────
        [Fact]
        [Trait("Category", "RequiresBackend")]
        public async Task InDocument_DeleteLastWord_WordRemovedFromText()
        {
            if (Skip()) return;
            _output.WriteLine("\n=== IN-DOC TEST 1: Delete last word (Ctrl+Backspace) ===");
            await ClearDocument();

            // Type a sentence with a distinct last word
            await _sim.SpeakAsync("type Hello World DeleteThisWord");
            await Task.Delay(600);

            var before = ReadContent();
            _output.WriteLine($"[BEFORE] '{before.TrimEnd()}'");
            before.Should().Contain("DeleteThisWord",
                because: "typed sentence must appear in document before deletion");

            // Go to end of document, then Ctrl+Backspace to delete the last word
            SetForegroundWindow(_launcher.WindowHandle);
            Thread.Sleep(150);
            Combo(VK_CONTROL, VK_END);  // move caret to end
            Thread.Sleep(100);

            // Ctrl+Backspace = delete previous word
            keybd_event(VK_CONTROL, 0, 0, UIntPtr.Zero);
            Thread.Sleep(20);
            Press(VK_BACK);
            keybd_event(VK_CONTROL, 0, KEYUP, UIntPtr.Zero);
            await Task.Delay(400);

            var after = ReadContent();
            _output.WriteLine($"[AFTER] '{after.TrimEnd()}'");

            // ✅ REAL ASSERTIONS
            after.Should().NotContain("DeleteThisWord",
                because: "Ctrl+Backspace must delete the last word 'DeleteThisWord'. " +
                         "If it still appears, Ctrl+Backspace did not execute in WordPad.");
            after.Should().Contain("Hello",
                because: "The preceding words must remain after deleting only the last word.");
            _output.WriteLine("✅ PASSED: Last word deleted correctly");
        }

        // ── TEST 2: Select all → replace content ──────────────────────────────
        [Fact]
        [Trait("Category", "RequiresBackend")]
        public async Task InDocument_SelectAllReplace_ContentCompletelyChanged()
        {
            if (Skip()) return;
            _output.WriteLine("\n=== IN-DOC TEST 2: Select All → Replace ===");
            await ClearDocument();

            // Type original content
            await _sim.SpeakAsync("type Original content that will be replaced");
            await Task.Delay(500);

            // Select all and delete
            SetForegroundWindow(_launcher.WindowHandle);
            Thread.Sleep(100);
            Combo(VK_CONTROL, VK_A);
            Thread.Sleep(100);
            Press(VK_DELETE);
            await Task.Delay(300);

            // Type replacement
            await _sim.SpeakAsync("type Completely new replacement text");
            await Task.Delay(500);

            var content = ReadContent();
            _output.WriteLine($"[CONTENT] '{content.TrimEnd()}'");

            // ✅ REAL ASSERTIONS
            content.Should().NotContain("Original content",
                because: "Original text must be gone after Select All + Delete. " +
                         "If present, the deletion did not execute.");
            content.Should().Contain("replacement",
                because: "New replacement text must be present after typing it.");
            _output.WriteLine("✅ PASSED: Select All → Replace worked");
        }

        // ── TEST 3: Copy-paste doubles content ────────────────────────────────
        [Fact]
        [Trait("Category", "RequiresBackend")]
        public async Task InDocument_CopyPaste_ContentDoubled()
        {
            if (Skip()) return;
            _output.WriteLine("\n=== IN-DOC TEST 3: Copy-Paste ===");
            await ClearDocument();

            const string word = "CopyMe";
            await _sim.SpeakAsync($"type {word}");
            await Task.Delay(500);

            // Select all, copy
            SetForegroundWindow(_launcher.WindowHandle);
            Thread.Sleep(100);
            Combo(VK_CONTROL, VK_A);
            Thread.Sleep(150);
            Combo(VK_CONTROL, VK_C);
            Thread.Sleep(200);

            // Move to end, paste
            Combo(VK_CONTROL, VK_END);
            Thread.Sleep(100);
            Combo(VK_CONTROL, VK_V);
            await Task.Delay(400);

            var content = ReadContent();
            _output.WriteLine($"[CONTENT] '{content.TrimEnd()}'");
            int count = CountOccurrences(content, word);
            _output.WriteLine($"[ASSERT] '{word}' appears {count} times (need >= 2)");

            // ✅ REAL ASSERTION
            count.Should().BeGreaterThanOrEqualTo(2,
                because: $"After copy-paste, '{word}' must appear at least twice. " +
                         $"Found {count} time(s) in: '{content.TrimEnd()}'");
            _output.WriteLine("✅ PASSED: Copy-paste doubled the content");
        }

        // ── TEST 4: Multiple undos restore previous states ────────────────────
        [Fact]
        [Trait("Category", "RequiresBackend")]
        public async Task InDocument_MultipleUndos_EachRestoresPreviousState()
        {
            if (Skip()) return;
            _output.WriteLine("\n=== IN-DOC TEST 4: Multiple undos ===");
            await ClearDocument();

            // Build three distinct additions
            await _sim.SpeakAsync("type First");
            await Task.Delay(300);
            await _sim.SpeakAsync("type Second");
            await Task.Delay(300);
            await _sim.SpeakAsync("type Third");
            await Task.Delay(300);

            var state3 = ReadContent();
            _output.WriteLine($"[STATE-3] '{state3.TrimEnd()}'");
            state3.Should().Contain("Third",
                because: "All three additions must be present before undo");

            // Undo once
            SetForegroundWindow(_launcher.WindowHandle);
            Thread.Sleep(100);
            Combo(VK_CONTROL, VK_Z);
            await Task.Delay(400);
            var state2 = ReadContent();
            _output.WriteLine($"[UNDO-1] '{state2.TrimEnd()}'");
            state2.Length.Should().BeLessThan(state3.Length,
                because: "After 1 undo, content must be shorter than the 3-addition state");

            // Undo again
            SetForegroundWindow(_launcher.WindowHandle);
            Thread.Sleep(100);
            Combo(VK_CONTROL, VK_Z);
            await Task.Delay(400);
            var state1 = ReadContent();
            _output.WriteLine($"[UNDO-2] '{state1.TrimEnd()}'");
            state1.Length.Should().BeLessThan(state2.Length,
                because: "After 2 undos, content must be shorter than after 1 undo");

            _output.WriteLine("✅ PASSED: Undo history works across multiple steps");
        }
    }
}
