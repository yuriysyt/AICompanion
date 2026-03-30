using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using AICompanion.IntegrationTests.Helpers;
using FluentAssertions;
using Xunit;
using Xunit.Abstractions;

namespace AICompanion.IntegrationTests.RealAppTests
{
    /// <summary>
    /// VOICE COMMAND E2E TESTS — simulate real user voice input without a microphone.
    ///
    /// Each test passes a text transcript through VoiceCommandSimulator, which routes it
    /// through the exact same pipeline as a live voice command:
    ///
    ///   [Text Input] → Confidence Gate → LocalCommandProcessor → AgenticExecutionService → Win32
    ///
    /// Test categories:
    ///   A. Confidence gate — low-confidence transcripts are blocked
    ///   B. Command routing — simple vs. complex command detection
    ///   C. App launch phrase variations — many ways to say "open notepad"
    ///   D. Russian language commands
    ///   E. Case sensitivity — UPPERCASE, Mixed Case, lowercase all work
    ///   F. Real app verification — voice command actually opens a window
    ///   G. Word document workflow — voice-driven document creation
    ///   H. Multi-phrase combinations — "open X and type Y" routing
    ///
    /// ⚠️  Tests in category F–G open real windows. Do NOT touch the desktop while they run.
    /// </summary>
    [Collection("RealApp")]
    public class VoiceCommandE2ETests : IAsyncLifetime
    {
        private readonly ITestOutputHelper    _output;
        private readonly VoiceCommandSimulator _sim;
        private readonly WindowVerifier        _verifier;

        // P/Invoke for cleanup
        [DllImport("user32.dll")] static extern bool EnumWindows(EnumWindowsProc cb, IntPtr lp);
        [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
        [DllImport("user32.dll")] static extern int  GetWindowText(IntPtr h, StringBuilder sb, int n);
        [DllImport("user32.dll")] static extern bool PostMessage(IntPtr h, uint msg, IntPtr w, IntPtr l);
        [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr h);
        [DllImport("user32.dll")] static extern void keybd_event(byte vk, byte scan, uint flags, UIntPtr extra);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lp);

        private const uint WM_CLOSE        = 0x0010;
        private const uint KEYEVENTF_KEYUP = 0x0002;
        private const byte VK_ALT          = 0x12;
        private const byte VK_F4           = 0x73;

        public VoiceCommandE2ETests(ITestOutputHelper output)
        {
            _output   = output;
            _sim      = new VoiceCommandSimulator(output);
            _verifier = new WindowVerifier(output);
        }

        public Task InitializeAsync()
        {
            // Release any stuck modifier keys left by previous test classes
            keybd_event(0x11 /* Ctrl  */, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
            keybd_event(VK_ALT,           0, KEYEVENTF_KEYUP, UIntPtr.Zero);
            keybd_event(0x10 /* Shift  */, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
            Thread.Sleep(50);
            return Task.CompletedTask;
        }

        public Task DisposeAsync() => Task.CompletedTask;

        // ── helpers ────────────────────────────────────────────────────────────

        /// <summary>
        /// Poll for a window containing <paramref name="partialTitle"/> using both
        /// UIAutomation and direct P/Invoke EnumWindows (for WinUI3 apps like Win11 Notepad
        /// that may not appear in UIAutomation immediately).
        /// </summary>
        private async Task<bool> WaitForWindowAsync(string partialTitle, int timeoutMs = 8000)
        {
            var deadline = DateTime.UtcNow.AddMilliseconds(timeoutMs);
            while (DateTime.UtcNow < deadline)
            {
                // Strategy 1: UIAutomation
                try
                {
                    if (_verifier.IsWindowOpen(partialTitle))
                    {
                        _output.WriteLine($"[WAIT] ✅ Window '{partialTitle}' found via UIAutomation");
                        return true;
                    }
                }
                catch { /* transient error — try next strategy */ }

                // Strategy 2: P/Invoke EnumWindows (works for WinUI3 that UIAutomation misses)
                if (FindWindowByTitleDirect(partialTitle))
                {
                    _output.WriteLine($"[WAIT] ✅ Window '{partialTitle}' found via EnumWindows");
                    return true;
                }

                await Task.Delay(500);
            }
            _output.WriteLine($"[WAIT] ❌ Window '{partialTitle}' not found after {timeoutMs}ms");
            return false;
        }

        private bool FindWindowByTitleDirect(string partialTitle)
        {
            bool found = false;
            EnumWindows((hWnd, _) =>
            {
                if (!IsWindowVisible(hWnd)) return true;
                var sb = new StringBuilder(256);
                GetWindowText(hWnd, sb, 256);
                var title = sb.ToString();
                if (!string.IsNullOrWhiteSpace(title) &&
                    title.Contains(partialTitle, StringComparison.OrdinalIgnoreCase) &&
                    !title.Contains("AI Companion", StringComparison.OrdinalIgnoreCase))
                {
                    _output.WriteLine($"[ENUMWND] Found: '{title}'");
                    found = true;
                    return false; // stop enumeration
                }
                return true;
            }, IntPtr.Zero);
            return found;
        }

        /// <summary>
        /// Close any visible window whose title contains <paramref name="partialTitle"/>.
        /// Uses WM_CLOSE so the app can handle unsaved-changes dialogs safely.
        /// </summary>
        private void CloseWindowsMatching(string partialTitle)
        {
            EnumWindows((hWnd, _) =>
            {
                if (!IsWindowVisible(hWnd)) return true;
                var sb = new StringBuilder(256);
                GetWindowText(hWnd, sb, 256);
                if (sb.ToString().Contains(partialTitle, StringComparison.OrdinalIgnoreCase))
                {
                    _output.WriteLine($"[CLEANUP] Closing window: '{sb}'");
                    PostMessage(hWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
                    Thread.Sleep(300);

                    // Dismiss any "save changes?" dialog that may have appeared
                    DismissSaveDialog();
                }
                return true;
            }, IntPtr.Zero);
        }

        /// <summary>Sends Alt+F4 to the foreground window (dismisses save dialogs).</summary>
        private void DismissSaveDialog()
        {
            Thread.Sleep(200);
            var fg = _verifier.GetForegroundTitle();
            if (fg.Contains("save", StringComparison.OrdinalIgnoreCase) ||
                fg.Contains("сохран", StringComparison.OrdinalIgnoreCase))
            {
                // Press N (Don't Save) or Escape
                keybd_event(0x4E /* N */, 0, 0, UIntPtr.Zero);
                Thread.Sleep(30);
                keybd_event(0x4E, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
                Thread.Sleep(200);
            }
        }

        // ═══════════════════════════════════════════════════════════════════════
        // A. CONFIDENCE GATE
        // ═══════════════════════════════════════════════════════════════════════

        /// <summary>
        /// Transcripts with confidence below 0.65 are filtered out before any
        /// command processing begins. This simulates background noise or mumbling.
        /// </summary>
        [Fact]
        public async Task Voice_LowConfidence_BlockedBeforeProcessing()
        {
            _output.WriteLine("\n=== A1: Low confidence transcript must be blocked ===");

            var result = await _sim.SpeakAsync("open notepad", confidence: 0.40f);

            result.PassedConfidenceGate.Should().BeFalse(
                because: "confidence 0.40 is below the 0.65 threshold");
            result.BlockedReason.Should().NotBeNullOrEmpty(
                because: "blocked result must include a reason");
            result.OverallSuccess.Should().BeFalse();

            _output.WriteLine($"✅ PASSED: Blocked at confidence gate — reason: '{result.BlockedReason}'");
        }

        /// <summary>
        /// Transcripts exactly at the threshold (0.65) must be accepted.
        /// </summary>
        [Fact]
        public async Task Voice_ThresholdConfidence_PassesGate()
        {
            _output.WriteLine("\n=== A2: Threshold confidence (0.65) must pass gate ===");

            var result = await _sim.SpeakAsync("help", confidence: 0.65f);

            result.PassedConfidenceGate.Should().BeTrue(
                because: "confidence 0.65 equals the threshold and should pass");

            _output.WriteLine($"✅ PASSED: Confidence gate passed at exactly 0.65");
        }

        // ═══════════════════════════════════════════════════════════════════════
        // B. COMMAND ROUTING — simple vs. complex
        // ═══════════════════════════════════════════════════════════════════════

        [Theory]
        [InlineData("open notepad")]
        [InlineData("help")]
        [InlineData("please open notepad")]       // polite prefix + simple → local
        [InlineData("can you open calculator")]   // polite prefix + simple → local
        public async Task Voice_SimpleCommands_RoutedLocally(string phrase)
        {
            _output.WriteLine($"\n=== B1: Simple command must route locally: '{phrase}' ===");

            var result = await _sim.SpeakAsync(phrase);

            result.PassedConfidenceGate.Should().BeTrue();
            result.IsComplex.Should().BeFalse(
                because: $"'{phrase}' is a simple command and must NOT go to agentic planner");

            _output.WriteLine($"✅ PASSED: '{phrase}' → LOCAL (IsComplex=false)");
        }

        [Theory]
        [InlineData("open notepad and type hello world")]
        [InlineData("start word then write my name is John")]
        [InlineData("open calculator and then close it")]
        [InlineData("create a new word document and type hello")]
        [InlineData("i want to open word and write something")]   // conversational + complex
        public async Task Voice_ComplexCommands_RoutedToAgentic(string phrase)
        {
            _output.WriteLine($"\n=== B2: Complex command must route to agentic: '{phrase}' ===");

            var result = await _sim.SpeakAsync(phrase);

            result.PassedConfidenceGate.Should().BeTrue();
            result.IsComplex.Should().BeTrue(
                because: $"'{phrase}' has conjunction/multiple verbs and must go to agentic planner");

            _output.WriteLine($"✅ PASSED: '{phrase}' → AGENTIC (IsComplex=true)");
        }

        // ═══════════════════════════════════════════════════════════════════════
        // C. APP LAUNCH PHRASE VARIATIONS
        // Every phrase below should produce Success=true when processed locally.
        // ═══════════════════════════════════════════════════════════════════════

        [Theory]
        [InlineData("open notepad")]
        [InlineData("start notepad")]
        [InlineData("launch notepad")]
        [InlineData("open the notepad")]
        [InlineData("open note pad")]        // two-word synonym
        [InlineData("open text editor")]     // synonym mapped to notepad
        public async Task Voice_OpenNotepad_PhraseVariations_AllSucceed(string phrase)
        {
            _output.WriteLine($"\n=== C1: Notepad open phrase variation: '{phrase}' ===");

            var result = await _sim.SpeakAsync(phrase);

            result.PassedConfidenceGate.Should().BeTrue();
            result.IsComplex.Should().BeFalse();
            result.LocalSuccess.Should().BeTrue(
                because: $"'{phrase}' should map to notepad and open successfully");

            _output.WriteLine($"✅ PASSED: '{phrase}' → Success=true, msg='{result.LocalMessage}'");

            // Give Notepad time to open, then clean up
            await Task.Delay(800);
            CloseWindowsMatching("Notepad");
            await Task.Delay(400);
        }

        [Theory]
        [InlineData("open calculator")]
        [InlineData("start calculator")]
        [InlineData("launch calculator")]
        [InlineData("open the calculator")]
        [InlineData("open calc")]
        public async Task Voice_OpenCalculator_PhraseVariations_AllSucceed(string phrase)
        {
            _output.WriteLine($"\n=== C2: Calculator open phrase variation: '{phrase}' ===");

            var result = await _sim.SpeakAsync(phrase);

            result.PassedConfidenceGate.Should().BeTrue();
            result.IsComplex.Should().BeFalse();
            result.LocalSuccess.Should().BeTrue(
                because: $"'{phrase}' should map to calculator and open successfully");

            _output.WriteLine($"✅ PASSED: '{phrase}' → Success=true, msg='{result.LocalMessage}'");

            await Task.Delay(800);
            CloseWindowsMatching("Calculator");
            await Task.Delay(400);
        }

        [Theory]
        [InlineData("open word")]
        [InlineData("start word")]
        [InlineData("launch word")]
        [InlineData("open microsoft word")]
        [InlineData("open ms word")]
        [InlineData("open word document")]
        public async Task Voice_OpenWord_PhraseVariations_AllSucceed(string phrase)
        {
            _output.WriteLine($"\n=== C3: Word open phrase variation: '{phrase}' ===");

            // Word may not be installed; check first and skip if absent
            bool wordAvailable = System.IO.File.Exists(
                @"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE") ||
                System.IO.File.Exists(
                @"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE");

            if (!wordAvailable)
            {
                _output.WriteLine("[SKIP] Microsoft Word not found — skipping Word launch test");
                return;
            }

            var result = await _sim.SpeakAsync(phrase);

            result.PassedConfidenceGate.Should().BeTrue();
            result.IsComplex.Should().BeFalse();
            result.LocalSuccess.Should().BeTrue(
                because: $"'{phrase}' should map to WINWORD and open successfully");

            _output.WriteLine($"✅ PASSED: '{phrase}' → Success=true");

            await Task.Delay(3000); // Word takes time to start
            CloseWindowsMatching("Word");
            await Task.Delay(500);
        }

        // ═══════════════════════════════════════════════════════════════════════
        // D. RUSSIAN LANGUAGE COMMANDS
        // ═══════════════════════════════════════════════════════════════════════

        [Theory]
        [InlineData("открой блокнот")]      // open notepad
        [InlineData("запусти блокнот")]     // launch notepad
        public async Task Voice_Russian_OpenNotepad_Succeeds(string phrase)
        {
            _output.WriteLine($"\n=== D1: Russian notepad command: '{phrase}' ===");

            var result = await _sim.SpeakAsync(phrase);

            result.PassedConfidenceGate.Should().BeTrue();
            result.LocalSuccess.Should().BeTrue(
                because: $"Russian command '{phrase}' must open Notepad");

            _output.WriteLine($"✅ PASSED: '{phrase}' → Success=true");

            await Task.Delay(800);
            CloseWindowsMatching("Notepad");
            CloseWindowsMatching("Блокнот");
            await Task.Delay(400);
        }

        [Theory]
        [InlineData("открой калькулятор")]  // open calculator
        [InlineData("запусти калькулятор")] // launch calculator
        public async Task Voice_Russian_OpenCalculator_Succeeds(string phrase)
        {
            _output.WriteLine($"\n=== D2: Russian calculator command: '{phrase}' ===");

            var result = await _sim.SpeakAsync(phrase);

            result.PassedConfidenceGate.Should().BeTrue();
            result.LocalSuccess.Should().BeTrue(
                because: $"Russian command '{phrase}' must open Calculator");

            _output.WriteLine($"✅ PASSED: '{phrase}' → Success=true");

            await Task.Delay(800);
            CloseWindowsMatching("Calculator");
            CloseWindowsMatching("Калькулятор");
            await Task.Delay(400);
        }

        [Theory]
        [InlineData("помощь")]              // help
        public async Task Voice_Russian_HelpCommand_Succeeds(string phrase)
        {
            _output.WriteLine($"\n=== D3: Russian help command: '{phrase}' ===");

            var result = await _sim.SpeakAsync(phrase);

            result.PassedConfidenceGate.Should().BeTrue();
            result.LocalSuccess.Should().BeTrue(
                because: $"Russian help command '{phrase}' must return help text");
            result.LocalMessage.Should().NotBeNullOrEmpty();

            _output.WriteLine($"✅ PASSED: '{phrase}' → Success=true, msg='{result.LocalMessage}'");
        }

        // ═══════════════════════════════════════════════════════════════════════
        // E. CASE SENSITIVITY — commands work in any capitalization
        // ═══════════════════════════════════════════════════════════════════════

        [Theory]
        [InlineData("OPEN NOTEPAD")]
        [InlineData("Open Notepad")]
        [InlineData("open notepad")]
        [InlineData("OPEN CALCULATOR")]
        [InlineData("Open Calculator")]
        [InlineData("HELP")]
        [InlineData("Help")]
        public async Task Voice_CaseVariations_AllSucceed(string phrase)
        {
            _output.WriteLine($"\n=== E1: Case variation: '{phrase}' ===");

            var result = await _sim.SpeakAsync(phrase);

            result.PassedConfidenceGate.Should().BeTrue();
            result.LocalSuccess.Should().BeTrue(
                because: $"Command processing must be case-insensitive: '{phrase}'");

            _output.WriteLine($"✅ PASSED: '{phrase}' → Success=true");

            // Close any window that may have opened
            await Task.Delay(600);
            CloseWindowsMatching("Notepad");
            CloseWindowsMatching("Calculator");
            await Task.Delay(300);
        }

        // ═══════════════════════════════════════════════════════════════════════
        // F. REAL APP VERIFICATION — voice command actually opens a window
        // ═══════════════════════════════════════════════════════════════════════

        /// <summary>
        /// "open notepad" must result in a visible Notepad window on screen.
        /// Uses WindowVerifier to confirm the window actually appeared.
        /// </summary>
        [Fact]
        public async Task Voice_OpenNotepad_WindowActuallyAppears()
        {
            _output.WriteLine("\n=== F1: 'open notepad' → Notepad window must appear ===");

            // Ensure no Notepad is already open
            CloseWindowsMatching("Notepad");
            await Task.Delay(500);

            var beforeCount = _verifier.CountOpenWindows("Notepad");

            var result = await _sim.SpeakAsync("open notepad");

            result.PassedConfidenceGate.Should().BeTrue();
            result.LocalSuccess.Should().BeTrue();

            // Poll until Notepad window appears (handles WinUI3 slower startup)
            var appeared = await WaitForWindowAsync("Notepad", timeoutMs: 6000);

            var afterCount = _verifier.CountOpenWindows("Notepad");
            _output.WriteLine($"[VERIFY] Notepad windows before={beforeCount}, after={afterCount}, appeared={appeared}");

            appeared.Should().BeTrue(
                because: "the 'open notepad' voice command must open a new Notepad window");

            _output.WriteLine($"✅ PASSED: Notepad window confirmed open");

            CloseWindowsMatching("Notepad");
            await Task.Delay(600);
        }

        /// <summary>
        /// "launch calculator" must result in a visible Calculator window.
        /// </summary>
        [Fact]
        public async Task Voice_LaunchCalculator_WindowActuallyAppears()
        {
            _output.WriteLine("\n=== F2: 'launch calculator' → Calculator window must appear ===");

            CloseWindowsMatching("Calculator");
            await Task.Delay(500);

            var beforeCount = _verifier.CountOpenWindows("Calculator");

            var result = await _sim.SpeakAsync("launch calculator");

            result.PassedConfidenceGate.Should().BeTrue();
            result.LocalSuccess.Should().BeTrue();

            var appeared = await WaitForWindowAsync("Calculator", timeoutMs: 6000);
            var afterCount = _verifier.CountOpenWindows("Calculator");
            _output.WriteLine($"[VERIFY] Calculator windows before={beforeCount}, after={afterCount}, appeared={appeared}");

            appeared.Should().BeTrue(
                because: "the 'launch calculator' voice command must open a Calculator window");

            _output.WriteLine($"✅ PASSED: Calculator window confirmed open");

            CloseWindowsMatching("Calculator");
            await Task.Delay(600);
        }

        /// <summary>
        /// Russian "открой блокнот" (open notepad) must also produce a real window.
        /// </summary>
        [Fact]
        public async Task Voice_RussianOpenNotepad_WindowActuallyAppears()
        {
            _output.WriteLine("\n=== F3: Russian 'открой блокнот' → Notepad window must appear ===");

            CloseWindowsMatching("Notepad");
            CloseWindowsMatching("Блокнот");
            await Task.Delay(500);

            var result = await _sim.SpeakAsync("открой блокнот");

            result.PassedConfidenceGate.Should().BeTrue();
            result.LocalSuccess.Should().BeTrue();

            // Poll — either English "Notepad" or Russian "Блокнот" title is acceptable
            var appeared = await WaitForWindowAsync("Notepad", timeoutMs: 6000) ||
                           await WaitForWindowAsync("Блокнот", timeoutMs: 1000);

            appeared.Should().BeTrue(
                because: "Russian command 'открой блокнот' must open a Notepad window");

            _output.WriteLine($"✅ PASSED: Notepad window opened via Russian voice command");

            CloseWindowsMatching("Notepad");
            CloseWindowsMatching("Блокнот");
            await Task.Delay(600);
        }

        // ═══════════════════════════════════════════════════════════════════════
        // G. WORD DOCUMENT WORKFLOW
        // ═══════════════════════════════════════════════════════════════════════

        /// <summary>
        /// "open word" must open Microsoft Word if it is installed.
        /// Tests the primary user scenario: voice-driven document creation.
        /// </summary>
        [Fact]
        public async Task Voice_OpenWord_WindowActuallyAppears()
        {
            _output.WriteLine("\n=== G1: 'open word' → Word window must appear ===");

            bool wordAvailable =
                System.IO.File.Exists(@"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE") ||
                System.IO.File.Exists(@"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE");

            if (!wordAvailable)
            {
                _output.WriteLine("[SKIP] Microsoft Word not installed — skipping");
                return;
            }

            CloseWindowsMatching("Word");
            await Task.Delay(500);

            var result = await _sim.SpeakAsync("open word");

            result.PassedConfidenceGate.Should().BeTrue();
            result.LocalSuccess.Should().BeTrue(
                because: "'open word' must successfully launch WINWORD.EXE");

            var isOpen = await WaitForWindowAsync("Word", timeoutMs: 10000);
            isOpen.Should().BeTrue(because: "Word window must be visible after 'open word' voice command");

            _output.WriteLine($"✅ PASSED: Word window confirmed open via voice command");

            CloseWindowsMatching("Word");
            await Task.Delay(500);
        }

        /// <summary>
        /// "open microsoft word" — full product name variation — must also work.
        /// </summary>
        [Fact]
        public async Task Voice_OpenMicrosoftWord_FullName_WindowActuallyAppears()
        {
            _output.WriteLine("\n=== G2: 'open microsoft word' (full name) → Word window must appear ===");

            bool wordAvailable =
                System.IO.File.Exists(@"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE") ||
                System.IO.File.Exists(@"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE");

            if (!wordAvailable)
            {
                _output.WriteLine("[SKIP] Microsoft Word not installed — skipping");
                return;
            }

            CloseWindowsMatching("Word");
            await Task.Delay(500);

            var result = await _sim.SpeakAsync("open microsoft word");

            result.PassedConfidenceGate.Should().BeTrue();
            result.LocalSuccess.Should().BeTrue(
                because: "'open microsoft word' (full product name) must open Word");

            (await WaitForWindowAsync("Word", timeoutMs: 10000)).Should().BeTrue();

            _output.WriteLine($"✅ PASSED: Full name 'open microsoft word' confirmed working");

            CloseWindowsMatching("Word");
            await Task.Delay(500);
        }

        /// <summary>
        /// Multi-step Word document command "open word and type hello world" must be
        /// recognized as a COMPLEX command and routed to the agentic planner.
        /// This verifies the routing logic for the primary Word document use case.
        /// </summary>
        [Fact]
        public async Task Voice_OpenWordAndType_RoutesToAgentic()
        {
            _output.WriteLine("\n=== G3: 'open word and type hello world' → must route to agentic ===");

            var result = await _sim.SpeakAsync("open word and type hello world");

            result.PassedConfidenceGate.Should().BeTrue();
            result.IsComplex.Should().BeTrue(
                because: "'open word and type hello world' has a conjunction and two action verbs — " +
                         "must be routed to AgenticExecutionService for multi-step execution");

            _output.WriteLine(
                $"✅ PASSED: Complex Word+type command correctly identified as AGENTIC. " +
                $"AgenticSuccess={result.AgenticSuccess}, Summary='{result.AgenticSummary}'");
        }

        /// <summary>
        /// "create a new word document and write my name is John" — conversational phrasing.
        /// Tests a realistic user request using natural language.
        /// </summary>
        [Fact]
        public async Task Voice_CreateWordDocument_ConversationalPhrase_RoutesToAgentic()
        {
            _output.WriteLine("\n=== G4: Conversational 'create word document' → must route to agentic ===");

            var phrases = new[]
            {
                "create a new word document and write my name is John",
                "open word then type hello this is a test",
                "start word and write some text for me",
                "i want to open word and type a letter",
            };

            foreach (var phrase in phrases)
            {
                var result = await _sim.SpeakAsync(phrase);

                result.PassedConfidenceGate.Should().BeTrue(
                    because: $"Phrase '{phrase}' must pass confidence gate");
                result.IsComplex.Should().BeTrue(
                    because: $"'{phrase}' is conversational/multi-step and must route to agentic");

                _output.WriteLine($"  ✅ '{phrase}' → AGENTIC (IsComplex=true)");
            }

            _output.WriteLine($"✅ PASSED: All {phrases.Length} conversational Word phrases routed to agentic");
        }

        // ═══════════════════════════════════════════════════════════════════════
        // H. EDGE CASES
        // ═══════════════════════════════════════════════════════════════════════

        [Fact]
        public async Task Voice_EmptyTranscript_ReturnsFailure()
        {
            _output.WriteLine("\n=== H1: Empty transcript must fail gracefully ===");

            var result = await _sim.SpeakAsync("", confidence: 0.95f);

            // Empty command: either blocked by confidence or returns failure from processor
            (result.OverallSuccess == false || !result.LocalSuccess).Should().BeTrue(
                because: "empty transcript must not succeed");

            _output.WriteLine($"✅ PASSED: Empty transcript handled gracefully");
        }

        [Fact]
        public async Task Voice_Gibberish_DoesNotCrash()
        {
            _output.WriteLine("\n=== H2: Gibberish transcript must not crash ===");

            // Simulate speech recognition returning garbage
            var result = await _sim.SpeakAsync("xyzzy blorp fnord", confidence: 0.90f);

            // Should complete without throwing — result may be failure but no exception
            result.Should().NotBeNull(because: "pipeline must handle unrecognized input gracefully");

            _output.WriteLine(
                $"✅ PASSED: Gibberish handled gracefully. Success={result.OverallSuccess}");
        }

        [Fact]
        public async Task Voice_HelpCommand_ReturnsHelpText()
        {
            _output.WriteLine("\n=== H3: 'help' command must return help text ===");

            var result = await _sim.SpeakAsync("help");

            result.PassedConfidenceGate.Should().BeTrue();
            result.LocalSuccess.Should().BeTrue();
            result.LocalMessage.Should().NotBeNullOrEmpty(
                because: "help command must return a non-empty description");

            _output.WriteLine($"✅ PASSED: help → '{result.LocalMessage}'");
        }

        [Fact]
        public async Task Voice_TimeCommand_PassesConfidenceGate()
        {
            _output.WriteLine("\n=== H4: 'what time is it' must pass the confidence gate ===");

            // "what time is it" is 4 words → IsComplexCommand=true → routed to agentic planner.
            // Final success depends on the Python backend being available.
            // This test only verifies the transcript passes the confidence gate and
            // reaches the processing pipeline (backend availability is not required).
            var result = await _sim.SpeakAsync("what time is it");

            result.PassedConfidenceGate.Should().BeTrue(
                because: "'what time is it' is a valid transcript and must pass the 0.65 confidence gate");
            result.IsComplex.Should().BeTrue(
                because: "4-word queries are classified as complex and routed to the agentic planner");

            _output.WriteLine(
                $"✅ PASSED: confidence gate passed, routed to agentic " +
                $"(AgenticSuccess={result.AgenticSuccess}, backend status: " +
                $"{(result.AgenticSuccess ? "OK" : result.AgenticSummary ?? "unavailable")})");
        }

        [Fact]
        public async Task Voice_SearchCommand_ReturnsSuccess()
        {
            _output.WriteLine("\n=== H5: 'search IBM Watson' must succeed ===");

            var result = await _sim.SpeakAsync("search IBM Watson");

            result.PassedConfidenceGate.Should().BeTrue();
            result.LocalSuccess.Should().BeTrue(
                because: "search command must execute successfully");

            _output.WriteLine($"✅ PASSED: search → '{result.LocalMessage}'");

            // Close any browser window that may have opened
            await Task.Delay(1000);
            CloseWindowsMatching("msedge");
            await Task.Delay(300);
        }
    }
}
