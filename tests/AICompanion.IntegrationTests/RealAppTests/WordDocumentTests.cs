using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;
using AICompanion.IntegrationTests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace AICompanion.IntegrationTests.RealAppTests
{
    /// <summary>
    /// REAL E2E TESTS: Opens Microsoft Word with a pre-created RTF file, types via
    /// hardware injection, saves with Ctrl+S (no dialog — file already exists),
    /// and verifies results by reading disk or clipboard.
    ///
    /// All tests skip gracefully if Word is not installed.
    /// No backend required — all operations use direct key injection.
    ///
    /// ⚠️  Opens a Word window on screen. Do NOT touch the desktop while tests run.
    /// </summary>
    [Collection("RealApp")]
    public class WordDocumentTests : IAsyncLifetime
    {
        private readonly ITestOutputHelper _output;
        private bool                       _wordAvailable;

        // ── P/Invoke ──────────────────────────────────────────────────────────
        [DllImport("user32.dll")] static extern bool   SetForegroundWindow(IntPtr h);
        [DllImport("user32.dll")] static extern bool   IsWindow(IntPtr h);
        [DllImport("user32.dll")] static extern bool   IsWindowVisible(IntPtr h);
        [DllImport("user32.dll")] static extern void   keybd_event(byte vk, byte scan, uint flags, UIntPtr extra);
        [DllImport("user32.dll")] static extern int    GetWindowText(IntPtr h, StringBuilder sb, int n);
        [DllImport("user32.dll")] static extern IntPtr FindWindowEx(IntPtr parent, IntPtr after, string? cls, string? title);
        [DllImport("user32.dll")] static extern bool   EnumChildWindows(IntPtr parent, EnumChildProc cb, IntPtr lp);
        [DllImport("user32.dll")] static extern bool   EnumWindows(EnumWindowsProc cb, IntPtr lp);
        [DllImport("user32.dll")] static extern int    GetClassName(IntPtr h, StringBuilder sb, int n);
        [DllImport("user32.dll")] static extern short  VkKeyScan(char ch);
        [DllImport("user32.dll")] static extern bool   GetWindowRect(IntPtr h, out RECT r);
        [DllImport("user32.dll")] static extern bool   SetCursorPos(int x, int y);
        [DllImport("user32.dll")] static extern void   mouse_event(uint flags, int x, int y, uint data, UIntPtr extra);

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT { public int L, T, R, B; }

        private delegate bool EnumChildProc(IntPtr h, IntPtr lp);
        private delegate bool EnumWindowsProc(IntPtr h, IntPtr lp);

        private const uint KEYUP              = 0x0002;
        private const uint MOUSEEVENTF_LEFT   = 0x0002;
        private const uint MOUSEEVENTF_LEFTUP = 0x0004;

        // VK constants
        private const byte VK_CONTROL = 0x11;
        private const byte VK_SHIFT   = 0x10;
        private const byte VK_A       = 0x41;
        private const byte VK_C       = 0x43;
        private const byte VK_S       = 0x53;
        private const byte VK_V       = 0x56;
        private const byte VK_Z       = 0x5A;
        private const byte VK_HOME    = 0x24;
        private const byte VK_END     = 0x23;
        private const byte VK_DELETE  = 0x2E;
        private const byte VK_RETURN  = 0x0D;
        private const byte VK_ESCAPE  = 0x1B;

        private static readonly string[] WordPaths =
        {
            @"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
            @"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE",
            @"C:\Program Files\Microsoft Office\Office16\WINWORD.EXE",
            @"C:\Program Files\Microsoft Office\Office15\WINWORD.EXE",
        };

        public WordDocumentTests(ITestOutputHelper output) => _output = output;

        public async Task InitializeAsync()
        {
            // Each test opens Word itself via OpenWordWithFile().
            // Only check that Word is installed here — do NOT launch Word.
            string? wordExe = WordPaths.FirstOrDefault(File.Exists);
            _wordAvailable = wordExe != null;
            if (!_wordAvailable)
                _output.WriteLine("[SETUP] Microsoft Word not found — all Word tests will skip");
            else
                _output.WriteLine($"[SETUP] Word found: {wordExe}");
            await Task.CompletedTask;
        }

        public Task DisposeAsync() => Task.CompletedTask;

        // ── Self-contained Word session ────────────────────────────────────────

        /// <summary>
        /// Wraps a Word process opened against a specific RTF file.
        /// Dispose() kills the process; tests delete the file in their finally blocks.
        /// </summary>
        private sealed class WordSession : IDisposable
        {
            public IntPtr WindowHandle { get; }
            public string FilePath     { get; }
            private readonly System.Diagnostics.Process? _proc;

            public WordSession(IntPtr hwnd, System.Diagnostics.Process? proc, string path)
            { WindowHandle = hwnd; _proc = proc; FilePath = path; }

            public void Dispose()
            {
                try { _proc?.Kill(); }    catch { }
                try { _proc?.Dispose(); } catch { }
            }
        }

        /// <summary>
        /// Creates a minimal RTF file containing <paramref name="initialText"/>,
        /// opens it in Word, waits for the document window, dismisses any splash,
        /// and returns a <see cref="WordSession"/>.
        /// Returns null if Word is not installed or the window is not found.
        /// </summary>
        private async Task<WordSession?> OpenWordWithFile(string initialText)
        {
            if (!_wordAvailable) return null;
            var wordExe = WordPaths.FirstOrDefault(File.Exists)!;

            // Use MyDocuments — Office trusts this folder; Desktop may trigger Protected View
            var path = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                $"WordTest_{DateTime.Now:HHmmssff}.rtf");

            // Minimal RTF envelope — Word opens .rtf natively without conversion dialogs
            var rtf = $@"{{\rtf1\ansi\deff0 {{\fonttbl{{\f0 Times New Roman;}}}} \f0\fs24 {EscapeRtf(initialText)}}}";
            File.WriteAllText(path, rtf, Encoding.ASCII);
            _output.WriteLine($"[WORD FILE] Created: {path}");

            System.Diagnostics.Process? proc = null;
            try
            {
                proc = System.Diagnostics.Process.Start(
                    new System.Diagnostics.ProcessStartInfo
                    {
                        FileName        = wordExe,
                        Arguments       = $"\"{path}\"",
                        UseShellExecute = false
                    });
            }
            catch (Exception ex)
            {
                _output.WriteLine($"[SKIP] Cannot start Word: {ex.Message}");
                try { File.Delete(path); } catch { }
                return null;
            }
            await Task.Delay(4000); // Word needs time to open the document

            // Find the document window (title contains file base-name or "Word")
            var filename = Path.GetFileNameWithoutExtension(path);
            var deadline = DateTime.UtcNow.AddSeconds(10);
            IntPtr hwnd  = IntPtr.Zero;

            while (DateTime.UtcNow < deadline && hwnd == IntPtr.Zero)
            {
                await Task.Delay(400);
                EnumWindows((h, _) =>
                {
                    if (!IsWindowVisible(h)) return true;
                    var sb = new StringBuilder(256);
                    GetWindowText(h, sb, 256);
                    var t = sb.ToString();
                    if (t.Contains(filename, StringComparison.OrdinalIgnoreCase) ||
                        t.Contains("Word",   StringComparison.OrdinalIgnoreCase))
                    { hwnd = h; return false; }
                    return true;
                }, IntPtr.Zero);
            }

            if (hwnd == IntPtr.Zero)
            {
                _output.WriteLine("[SKIP] Word document window not found");
                try { proc?.Kill(); } catch { }
                try { File.Delete(path); } catch { }
                return null;
            }

            var sb2 = new StringBuilder(256);
            GetWindowText(hwnd, sb2, 256);
            _output.WriteLine($"[WORD] hWnd=0x{hwnd:X} title='{sb2}'");

            // Dismiss splash / start screen with Escape
            SetForegroundWindow(hwnd);
            Thread.Sleep(300);
            keybd_event(VK_ESCAPE, 0, 0, UIntPtr.Zero);
            keybd_event(VK_ESCAPE, 0, KEYUP, UIntPtr.Zero);
            await Task.Delay(500);

            // Click window center to ensure editing canvas has focus
            ForceClickWindowCenter(hwnd);
            Thread.Sleep(200);

            // Ctrl+Home to go to start of document
            Combo(VK_CONTROL, VK_HOME);
            Thread.Sleep(300);

            return new WordSession(hwnd, proc, path);
        }

        private static string EscapeRtf(string text) =>
            text.Replace("\\", "\\\\").Replace("{", "\\{").Replace("}", "\\}");

        // ── Canvas-targeting type helper ───────────────────────────────────────

        /// <summary>
        /// Finds Word's _WwG editing canvas and types text using hardware keybd_event.
        /// Falls back to the main Word hwnd if _WwG is not found.
        /// </summary>
        /// <summary>
        /// Types text into the active Word document via clipboard+paste (Ctrl+V).
        /// VkKeyScan is NOT used: on a Russian keyboard layout VkKeyScan returns -1
        /// for Latin characters, silently dropping all typed content.
        /// Clipboard paste is keyboard-layout-agnostic.
        /// </summary>
        private void TypeIntoWord(IntPtr wordHwnd, string text)
        {
            var canvas = FindWordEditingCanvas(wordHwnd);
            SetForegroundWindow(wordHwnd);
            if (canvas != IntPtr.Zero) SetForegroundWindow(canvas);
            Thread.Sleep(150);

            var segments = text.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n');
            for (int i = 0; i < segments.Length; i++)
            {
                if (segments[i].Length > 0)
                {
                    var seg = segments[i];
                    var sta = new Thread(() => System.Windows.Forms.Clipboard.SetText(seg));
                    sta.SetApartmentState(ApartmentState.STA);
                    sta.Start();
                    sta.Join();
                    Thread.Sleep(100);

                    keybd_event(VK_CONTROL, 0, 0,     UIntPtr.Zero);
                    keybd_event(VK_V,       0, 0,     UIntPtr.Zero); Thread.Sleep(30);
                    keybd_event(VK_V,       0, KEYUP, UIntPtr.Zero);
                    keybd_event(VK_CONTROL, 0, KEYUP, UIntPtr.Zero);
                    Thread.Sleep(200);
                }
                if (i < segments.Length - 1)
                    Press(VK_RETURN);
            }
        }

        private IntPtr FindWordEditingCanvas(IntPtr wordHwnd)
        {
            IntPtr canvas = FindWindowEx(wordHwnd, IntPtr.Zero, "_WwG", null);
            if (canvas != IntPtr.Zero) return canvas;

            IntPtr found = IntPtr.Zero;
            EnumChildWindows(wordHwnd, (h, _) =>
            {
                var sb = new StringBuilder(64);
                GetClassName(h, sb, 64);
                if (sb.ToString() == "_WwG") { found = h; return false; }
                return true;
            }, IntPtr.Zero);
            return found;
        }

        // ── Key helpers ────────────────────────────────────────────────────────

        private void Combo(byte mod, byte vk)
        {
            keybd_event(mod, 0, 0, UIntPtr.Zero); Thread.Sleep(20);
            keybd_event(vk,  0, 0, UIntPtr.Zero); Thread.Sleep(30);
            keybd_event(vk,  0, KEYUP, UIntPtr.Zero); Thread.Sleep(20);
            keybd_event(mod, 0, KEYUP, UIntPtr.Zero); Thread.Sleep(50);
        }

        private void Press(byte vk)
        {
            keybd_event(vk, 0, 0, UIntPtr.Zero); Thread.Sleep(30);
            keybd_event(vk, 0, KEYUP, UIntPtr.Zero); Thread.Sleep(30);
        }

        // ── Click-to-focus helper ──────────────────────────────────────────────

        /// <summary>
        /// Physically clicks the document body area (offset +50 px below center
        /// to clear the ribbon/toolbar and land in the editing canvas).
        /// </summary>
        private void ForceClickWindowCenter(IntPtr hwnd)
        {
            GetWindowRect(hwnd, out var r);
            int cx = (r.L + r.R) / 2;
            int cy = (r.T + r.B) / 2 + 50; // offset down past ribbon
            SetCursorPos(cx, cy);
            Thread.Sleep(100);
            mouse_event(MOUSEEVENTF_LEFT,   cx, cy, 0, UIntPtr.Zero);
            Thread.Sleep(50);
            mouse_event(MOUSEEVENTF_LEFTUP, cx, cy, 0, UIntPtr.Zero);
            Thread.Sleep(200);
            _output.WriteLine($"[FOCUS] Clicked Word document area at ({cx},{cy})");
        }

        // ── Clipboard helper ───────────────────────────────────────────────────

        private string ReadClipboardFromWindow(IntPtr hwnd)
        {
            SetForegroundWindow(hwnd);
            Thread.Sleep(200);

            var clearT = new Thread(() => { try { System.Windows.Forms.Clipboard.Clear(); } catch { } });
            clearT.SetApartmentState(ApartmentState.STA);
            clearT.Start(); clearT.Join(1000);
            Thread.Sleep(100);

            Combo(VK_CONTROL, VK_A); Thread.Sleep(200); // Select All
            Combo(VK_CONTROL, VK_C); Thread.Sleep(800); // Copy

            string text = "";
            var readT = new Thread(() => { try { text = System.Windows.Forms.Clipboard.GetText(); } catch { } });
            readT.SetApartmentState(ApartmentState.STA);
            readT.Start(); readT.Join(2000);

            _output.WriteLine($"[CLIPBOARD] {text.Length} chars: '{text.Substring(0, Math.Min(80, text.Length)).TrimEnd()}'");
            return text;
        }

        // ── Format-dialog dismissal ────────────────────────────────────────────

        /// <summary>
        /// Word may show "Do you want to keep saving in this format?" when saving an RTF.
        /// Presses Enter (= Keep Format) on any new dialog that appears.
        /// </summary>
        private async Task DismissWordFormatDialog(IntPtr wordHwnd)
        {
            await Task.Delay(600);
            try
            {
                EnumWindows((h, _) =>
                {
                    if (!IsWindowVisible(h) || h == wordHwnd) return true;
                    var sb = new StringBuilder(256);
                    GetWindowText(h, sb, 256);
                    var t = sb.ToString();
                    if (t.Contains("Microsoft Word", StringComparison.OrdinalIgnoreCase) ||
                        t.Contains("format",         StringComparison.OrdinalIgnoreCase) ||
                        t.Contains("compatible",     StringComparison.OrdinalIgnoreCase) ||
                        t.Contains("keep",           StringComparison.OrdinalIgnoreCase))
                    {
                        _output.WriteLine($"[FORMAT DIALOG] Dismissing: '{t}'");
                        SetForegroundWindow(h); Thread.Sleep(100);
                        keybd_event(VK_RETURN, 0, 0, UIntPtr.Zero);
                        keybd_event(VK_RETURN, 0, KEYUP, UIntPtr.Zero);
                        return false;
                    }
                    return true;
                }, IntPtr.Zero);
            }
            catch { }
            await Task.Delay(500);
        }

        private bool ShouldSkip()
        {
            if (!_wordAvailable) _output.WriteLine("[SKIP] Word not installed — test skipped");
            return !_wordAvailable;
        }

        // ── TEST 1: Open RTF → type → Ctrl+S → mtime changed + text on disk ──
        [Fact]
        public async Task Word_OpenDocType_CtrlS_ContentOnDisk()
        {
            if (ShouldSkip()) return;
            _output.WriteLine("\n=== WORD TEST 1: Type + Ctrl+S + verify disk ===");

            var session = await OpenWordWithFile("ORIGINAL_WORD_CONTENT");
            if (session == null) return;
            try
            {
                const string newText = "TYPED_BY_TEST";
                SetForegroundWindow(session.WindowHandle); Thread.Sleep(300);
                ForceClickWindowCenter(session.WindowHandle);
                Combo(VK_CONTROL, VK_END); Thread.Sleep(200);
                Press(VK_RETURN); Thread.Sleep(100);
                TypeIntoWord(session.WindowHandle, newText);
                await Task.Delay(800);

                var modBefore = File.GetLastWriteTime(session.FilePath);
                _output.WriteLine($"[SAVE] Pressing Ctrl+S (mtime before: {modBefore:ss.fff})");
                Combo(VK_CONTROL, VK_S);
                await Task.Delay(3000);
                await DismissWordFormatDialog(session.WindowHandle);

                var modAfter = File.GetLastWriteTime(session.FilePath);
                _output.WriteLine($"[MTIME] {modBefore:ss.fff} → {modAfter:ss.fff}");

                // ✅ ASSERTION 1: file modification time changed
                modAfter.Should().BeAfter(modBefore,
                    "After Ctrl+S, Word must update file. If mtime unchanged, save did not happen.");

                // Close Word to release the file lock before reading disk content
                var savedPath1 = session.FilePath;
                session.Dispose();
                await Task.Delay(1500);

                // ✅ ASSERTION 2: text is on disk (RTF stores plain text directly)
                var diskRaw = File.ReadAllText(savedPath1);
                _output.WriteLine($"[DISK RAW] {diskRaw.Length} chars");
                diskRaw.Should().Contain(newText,
                    $"Word must save '{newText}' to the RTF file.");

                _output.WriteLine("✅ PASSED: Word Ctrl+S saved content to disk");
                try { File.Delete(savedPath1); } catch { }
            }
            finally
            {
                session.Dispose(); // no-op if already disposed
                await Task.Delay(200);
                try { File.Delete(session.FilePath); } catch { }
            }
        }

        // ── TEST 2: Multiple paragraphs — all saved to disk ───────────────────
        [Fact]
        public async Task Word_MultipleParagraphs_AllSavedToDisk()
        {
            if (ShouldSkip()) return;
            _output.WriteLine("\n=== WORD TEST 2: Multiple paragraphs saved ===");

            var session = await OpenWordWithFile("START");
            if (session == null) return;
            try
            {
                SetForegroundWindow(session.WindowHandle); Thread.Sleep(300);
                ForceClickWindowCenter(session.WindowHandle);
                Combo(VK_CONTROL, VK_END); Thread.Sleep(200);

                var paragraphs = new[] { "PARA_ONE", "PARA_TWO", "PARA_THREE" };
                foreach (var p in paragraphs)
                {
                    Press(VK_RETURN); Thread.Sleep(80);
                    TypeIntoWord(session.WindowHandle, p);
                    await Task.Delay(200);
                }

                Combo(VK_CONTROL, VK_S);
                await Task.Delay(3000);
                await DismissWordFormatDialog(session.WindowHandle);

                var savedPath2 = session.FilePath;
                session.Dispose();
                await Task.Delay(1500);

                var disk = File.ReadAllText(savedPath2);
                _output.WriteLine($"[DISK] {disk.Length} chars");

                // ✅ REAL ASSERTIONS — all 3 paragraphs must be on disk
                foreach (var p in paragraphs)
                {
                    disk.Should().Contain(p, $"Paragraph '{p}' must be on disk after save");
                    _output.WriteLine($"  ✅ '{p}' found on disk");
                }
                _output.WriteLine("✅ PASSED: All paragraphs saved");
                try { File.Delete(savedPath2); } catch { }
            }
            finally
            {
                session.Dispose();
                await Task.Delay(200);
                try { File.Delete(session.FilePath); } catch { }
            }
        }

        // ── TEST 3: Select All → Delete → type replacement → Ctrl+S → old gone
        [Fact]
        public async Task Word_SelectAllDelete_NewContent_SavedToDisk()
        {
            if (ShouldSkip()) return;
            _output.WriteLine("\n=== WORD TEST 3: SelectAll+Delete+Type+Save ===");

            var session = await OpenWordWithFile("OLD_WORD_CONTENT");
            if (session == null) return;
            try
            {
                SetForegroundWindow(session.WindowHandle); Thread.Sleep(300);
                ForceClickWindowCenter(session.WindowHandle);
                Combo(VK_CONTROL, VK_A); Thread.Sleep(150);
                Press(VK_DELETE); await Task.Delay(300);

                const string replacement = "REPLACED_CONTENT";
                TypeIntoWord(session.WindowHandle, replacement);
                await Task.Delay(500);

                Combo(VK_CONTROL, VK_S);
                await Task.Delay(3000);
                await DismissWordFormatDialog(session.WindowHandle);

                var savedPath3 = session.FilePath;
                session.Dispose();
                await Task.Delay(1500);

                var disk = File.ReadAllText(savedPath3);
                _output.WriteLine($"[DISK] '{disk.TrimEnd()}'");

                // ✅ REAL ASSERTIONS
                disk.Should().NotContain("OLD_WORD_CONTENT",
                    "Old content must be gone from disk after select-all-delete + save");
                disk.Should().Contain(replacement,
                    $"'{replacement}' must be on disk");

                _output.WriteLine("✅ PASSED: Word content replaced and saved");
                try { File.Delete(savedPath3); } catch { }
            }
            finally
            {
                session.Dispose();
                await Task.Delay(200);
                try { File.Delete(session.FilePath); } catch { }
            }
        }

        // ── TEST 4: Clipboard read-back verifies text IS in document ──────────
        [Fact]
        public async Task Word_TypeText_ClipboardReadback_VerifiesContent()
        {
            if (ShouldSkip()) return;
            _output.WriteLine("\n=== WORD TEST 4: Type → clipboard read-back ===");

            var session = await OpenWordWithFile("");
            if (session == null) return;
            try
            {
                SetForegroundWindow(session.WindowHandle); Thread.Sleep(300);
                ForceClickWindowCenter(session.WindowHandle);
                const string testSentence = "Word document test sentence for verification";
                TypeIntoWord(session.WindowHandle, testSentence);
                await Task.Delay(600);

                var content = ReadClipboardFromWindow(session.WindowHandle);

                // ✅ REAL ASSERTION — text IS in Word's document area
                content.Should().Contain("Word document",
                    "The typed text must appear in Word's document. " +
                    "If clipboard is empty, TypeIntoWord failed to inject into the _WwG canvas.");

                _output.WriteLine("✅ PASSED: Text confirmed in Word via clipboard");
            }
            finally
            {
                session.Dispose();
                await Task.Delay(500);
                try { File.Delete(session.FilePath); } catch { }
            }
        }

        // ── TEST 5: Type → Undo → content shortened ───────────────────────────
        [Fact]
        public async Task Word_TypeThenUndo_ContentReverted()
        {
            if (ShouldSkip()) return;
            _output.WriteLine("\n=== WORD TEST 5: Undo in Word ===");

            var session = await OpenWordWithFile("BASE_CONTENT");
            if (session == null) return;
            try
            {
                SetForegroundWindow(session.WindowHandle); Thread.Sleep(300);
                ForceClickWindowCenter(session.WindowHandle);
                Combo(VK_CONTROL, VK_END); Thread.Sleep(100);
                Press(VK_RETURN); Thread.Sleep(100);
                TypeIntoWord(session.WindowHandle, "UNDO_ME");
                await Task.Delay(400);

                var before = ReadClipboardFromWindow(session.WindowHandle);
                before.Should().Contain("UNDO_ME", "text must be present before undo");
                _output.WriteLine($"[BEFORE UNDO] {before.Length} chars");

                // Undo ×3
                for (int i = 0; i < 3; i++) { Combo(VK_CONTROL, VK_Z); Thread.Sleep(200); }

                var after = ReadClipboardFromWindow(session.WindowHandle);
                _output.WriteLine($"[AFTER UNDO] {after.Length} chars: '{after.TrimEnd()}'");

                // ✅ REAL ASSERTION
                after.Length.Should().BeLessThan(before.Length,
                    "After Ctrl+Z ×3, content must be shorter than before undo");

                _output.WriteLine("✅ PASSED: Word undo reverted typed text");
            }
            finally
            {
                session.Dispose();
                await Task.Delay(500);
                try { File.Delete(session.FilePath); } catch { }
            }
        }
    }
}
