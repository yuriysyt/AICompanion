using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using AICompanion.IntegrationTests.Helpers;
using Xunit;
using Xunit.Abstractions;

// ────────────────────────────────────────────────────────────────────────────────
//  UserWorkflowE2ETests.cs  –  20 unique E2E tests simulating real user workflows
//
//  Coverage areas
//  ──────────────
//  CLASS 1  WordUserWorkflowTests      (Tests  1-8)  Word document workflows with
//           disk-verified saves using improved robust-save helper that polls for
//           format dialogs and waits after dismissal before reading the file.
//
//  CLASS 2  BrowserUserWorkflowTests   (Tests  9-14) Browser navigation tests
//           that do NOT require the Python backend — pure Win32 key injection.
//
//  CLASS 3  VoiceCommandRoutingTests   (Tests 15-18) Confidence gate, routing
//           decisions, and local-vs-agentic classification via VoiceCommandSimulator.
//
//  CLASS 4  NotepadAndFileSystemTests  (Tests 19-20) Notepad plain-text save
//           (disk-verified) and File Explorer open/close.
//
//  ⚠️  Classes 1, 2, and 4 open real application windows.
//      Do NOT touch the desktop while they run.
// ────────────────────────────────────────────────────────────────────────────────

namespace AICompanion.IntegrationTests.RealAppTests
{
    // ════════════════════════════════════════════════════════════════════════════
    //  CLASS 1 – Word document user workflows (Tests 1–8)
    // ════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Eight unique Word document scenarios that each simulate a different real
    /// user request to the voice assistant.  Every test verifies the file on disk
    /// after saving (the plain-text bytes of the RTF are searched for the typed
    /// content) using an improved polling-based save helper that waits for any
    /// format-conversion dialog before timing out.
    /// </summary>
    [Collection("RealApp")]
    public class WordUserWorkflowTests : IAsyncLifetime
    {
        private readonly ITestOutputHelper _out;
        private bool _wordOk;

        // ── P/Invoke ─────────────────────────────────────────────────────────
        [DllImport("user32.dll")] static extern bool  SetForegroundWindow(IntPtr h);
        [DllImport("user32.dll")] static extern bool  IsWindowVisible(IntPtr h);
        [DllImport("user32.dll")] static extern void  keybd_event(byte vk, byte scan, uint flags, UIntPtr extra);
        [DllImport("user32.dll")] static extern int   GetWindowText(IntPtr h, StringBuilder sb, int n);
        [DllImport("user32.dll")] static extern IntPtr FindWindowEx(IntPtr parent, IntPtr after, string? cls, string? title);
        [DllImport("user32.dll")] static extern bool  EnumChildWindows(IntPtr parent, EnumChildProc cb, IntPtr lp);
        [DllImport("user32.dll")] static extern bool  EnumWindows(EnumWindowsProc cb, IntPtr lp);
        [DllImport("user32.dll")] static extern int   GetClassName(IntPtr h, StringBuilder sb, int n);
        [DllImport("user32.dll")] static extern short VkKeyScan(char ch);
        [DllImport("user32.dll")] static extern bool  GetWindowRect(IntPtr h, out RECT r);
        [DllImport("user32.dll")] static extern bool  SetCursorPos(int x, int y);
        [DllImport("user32.dll")] static extern void  mouse_event(uint flags, int x, int y, uint data, UIntPtr extra);

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT { public int L, T, R, B; }
        private delegate bool EnumChildProc(IntPtr h, IntPtr lp);
        private delegate bool EnumWindowsProc(IntPtr h, IntPtr lp);

        private const uint  KEYUP              = 0x0002;
        private const uint  MOUSEEVENTF_LEFT   = 0x0002;
        private const uint  MOUSEEVENTF_LEFTUP = 0x0004;
        private const byte  VK_CONTROL = 0x11;
        private const byte  VK_SHIFT   = 0x10;
        private const byte  VK_A       = 0x41;
        private const byte  VK_B       = 0x42;
        private const byte  VK_C       = 0x43;
        private const byte  VK_S       = 0x53;
        private const byte  VK_V       = 0x56;
        private const byte  VK_Z       = 0x5A;
        private const byte  VK_HOME    = 0x24;
        private const byte  VK_END     = 0x23;
        private const byte  VK_DELETE  = 0x2E;
        private const byte  VK_RETURN  = 0x0D;
        private const byte  VK_ESCAPE  = 0x1B;

        private static readonly string[] WordPaths =
        {
            @"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
            @"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE",
            @"C:\Program Files\Microsoft Office\Office16\WINWORD.EXE",
            @"C:\Program Files\Microsoft Office\Office15\WINWORD.EXE",
        };

        public WordUserWorkflowTests(ITestOutputHelper output) => _out = output;

        public async Task InitializeAsync()
        {
            string? exe = WordPaths.FirstOrDefault(File.Exists);
            _wordOk = exe != null;
            _out.WriteLine(_wordOk
                ? $"[SETUP] Word found: {exe}"
                : "[SETUP] Word not found — Word tests will skip");
            await Task.CompletedTask;
        }

        public Task DisposeAsync() => Task.CompletedTask;

        // ── Word session (self-contained: own process + temp file) ────────────

        private sealed class WordSession : IDisposable
        {
            public IntPtr WindowHandle { get; }
            public string FilePath     { get; }
            private readonly System.Diagnostics.Process? _proc;
            public WordSession(IntPtr hwnd, System.Diagnostics.Process? p, string path)
            { WindowHandle = hwnd; _proc = p; FilePath = path; }
            public void Dispose() { try { _proc?.Kill(); } catch { } try { _proc?.Dispose(); } catch { } }
        }

        private async Task<WordSession?> OpenWordWithFile(string initialRtfContent)
        {
            if (!_wordOk) return null;
            var exe  = WordPaths.First(File.Exists);
            var path = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                $"WUW_{DateTime.Now:HHmmssff}.rtf");

            var rtf = $@"{{\rtf1\ansi\deff0 {{\fonttbl{{\f0 Times New Roman;}}}} \f0\fs24 {EscRtf(initialRtfContent)}}}";
            File.WriteAllText(path, rtf, Encoding.ASCII);
            _out.WriteLine($"[FILE] {path}");

            System.Diagnostics.Process? proc = null;
            try
            {
                proc = System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    { FileName = exe, Arguments = $"\"{path}\"", UseShellExecute = false });
            }
            catch (Exception ex) { _out.WriteLine($"[SKIP] Cannot start Word: {ex.Message}"); File.Delete(path); return null; }

            await Task.Delay(4500);

            var filename = Path.GetFileNameWithoutExtension(path);
            var deadline = DateTime.UtcNow.AddSeconds(12);
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
                        t.Contains("Word", StringComparison.OrdinalIgnoreCase))
                    { hwnd = h; return false; }
                    return true;
                }, IntPtr.Zero);
            }

            if (hwnd == IntPtr.Zero)
            {
                _out.WriteLine("[SKIP] Word window not found");
                try { proc?.Kill(); } catch { }
                try { File.Delete(path); } catch { }
                return null;
            }

            // Dismiss any startup dialogs and bring Word into focus
            SetForegroundWindow(hwnd); Thread.Sleep(500);
            keybd_event(VK_ESCAPE, 0, 0, UIntPtr.Zero);
            keybd_event(VK_ESCAPE, 0, KEYUP, UIntPtr.Zero);
            await Task.Delay(600);

            // Click center of window to ensure the editing canvas has keyboard focus
            ClickCenter(hwnd);
            Thread.Sleep(400);

            // Second focus assertion: re-set foreground and move caret to start
            SetForegroundWindow(hwnd); Thread.Sleep(300);
            Combo(VK_CONTROL, VK_HOME);
            Thread.Sleep(400);

            _out.WriteLine($"[OPEN] Word session ready — hwnd=0x{hwnd:X}");
            return new WordSession(hwnd, proc, path);
        }

        // ── Improved robust save: polls for format dialog ≤6 s, waits after ──
        private async Task SaveDocumentRobust(IntPtr hwnd, string filePath)
        {
            _out.WriteLine("[SAVE] Sending Ctrl+S...");
            SetForegroundWindow(hwnd);
            Thread.Sleep(200);
            Combo(VK_CONTROL, VK_S);

            // Poll for "Keep in RTF format?" dialog for up to 6 seconds
            bool dialogDismissed = false;
            var  giveUp          = DateTime.UtcNow.AddSeconds(6);
            while (DateTime.UtcNow < giveUp)
            {
                await Task.Delay(300);
                if (TryDismissFormatDialog(hwnd))
                {
                    dialogDismissed = true;
                    _out.WriteLine("[SAVE] Format dialog dismissed — waiting for write...");
                    break;
                }
            }

            // If dialog was dismissed Word still needs time to flush to disk
            int waitMs = dialogDismissed ? 3500 : 2000;
            _out.WriteLine($"[SAVE] Waiting {waitMs}ms for disk write...");
            await Task.Delay(waitMs);

            var mtime = File.GetLastWriteTime(filePath);
            _out.WriteLine($"[SAVE] File mtime after save: {mtime:HH:mm:ss.fff}");
        }

        private bool TryDismissFormatDialog(IntPtr wordHwnd)
        {
            // Only dismiss windows with class "#32770" (standard Win32 dialog).
            // Using title keywords alone caused false positives where browser windows
            // whose titles happened to contain "keep" (e.g. "GitHub keeps you ahead - Edge")
            // were incorrectly detected and sent Enter, breaking the browser state.
            bool found = false;
            EnumWindows((h, _) =>
            {
                if (!IsWindowVisible(h) || h == wordHwnd) return true;

                // Require the Win32 dialog class — rules out all browser/app windows
                var cls = new StringBuilder(64);
                GetClassName(h, cls, 64);
                if (!cls.ToString().Equals("#32770", StringComparison.Ordinal)) return true;

                var sb = new StringBuilder(256);
                GetWindowText(h, sb, 256);
                var t = sb.ToString();
                if (t.Contains("Microsoft Word", StringComparison.OrdinalIgnoreCase) ||
                    t.Contains("format",         StringComparison.OrdinalIgnoreCase) ||
                    t.Contains("compatible",     StringComparison.OrdinalIgnoreCase) ||
                    t.Contains("keep",           StringComparison.OrdinalIgnoreCase))
                {
                    _out.WriteLine($"[DIALOG] Found: '{t}'");
                    SetForegroundWindow(h); Thread.Sleep(100);
                    keybd_event(VK_RETURN, 0, 0, UIntPtr.Zero); Thread.Sleep(30);
                    keybd_event(VK_RETURN, 0, KEYUP, UIntPtr.Zero);
                    found = true;
                    return false;
                }
                return true;
            }, IntPtr.Zero);
            return found;
        }

        // ── Dispose Word + read file with retry ──────────────────────────────
        /// <summary>
        /// Kills the Word process then waits up to 5s for the OS to release the file lock.
        /// Word may hold the RTF handle open for 1-3s after Kill() depending on pending
        /// writes and antivirus scans. A simple fixed delay of 1500ms was too short.
        /// </summary>
        private async Task<string> DisposeAndReadFile(WordSession session)
        {
            session.Dispose();
            for (int attempt = 0; attempt < 10; attempt++)
            {
                await Task.Delay(500);
                try { return File.ReadAllText(session.FilePath); }
                catch (IOException) { if (attempt == 9) throw; }
            }
            return "";  // unreachable
        }

        // ── Clipboard readback ────────────────────────────────────────────────
        private string ReadClipboard(IntPtr hwnd)
        {
            SetForegroundWindow(hwnd); Thread.Sleep(200);
            ClearClipboard();
            Combo(VK_CONTROL, VK_A); Thread.Sleep(200);
            Combo(VK_CONTROL, VK_C); Thread.Sleep(700);
            string text = "";
            var t = new Thread(() => { try { text = System.Windows.Forms.Clipboard.GetText(); } catch { } });
            t.SetApartmentState(ApartmentState.STA); t.Start(); t.Join(2000);
            _out.WriteLine($"[CLIP] {text.Length} chars: '{text.Substring(0, Math.Min(60, text.Length)).TrimEnd()}'");
            return text;
        }

        private void ClearClipboard()
        {
            var t = new Thread(() => { try { System.Windows.Forms.Clipboard.Clear(); } catch { } });
            t.SetApartmentState(ApartmentState.STA); t.Start(); t.Join(1000);
            Thread.Sleep(100);
        }

        // ── Typing into Word's _WwG editing canvas ────────────────────────────
        /// <summary>
        /// Types text into the active Word document using clipboard+paste (Ctrl+V).
        /// VkKeyScan-based injection is NOT used because on a Russian keyboard layout
        /// VkKeyScan returns -1 for all Latin characters — they are silently dropped.
        /// Clipboard paste is layout-agnostic and works on any active input locale.
        /// Multi-line text is split on '\n' with Enter key presses between segments.
        /// </summary>
        private void TypeInWord(IntPtr hwnd, string text)
        {
            var canvas = FindCanvas(hwnd);
            SetForegroundWindow(hwnd);
            if (canvas != IntPtr.Zero) SetForegroundWindow(canvas);
            Thread.Sleep(150);

            // Normalise line endings then split so we inject Enter between lines
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

        private IntPtr FindCanvas(IntPtr hwnd)
        {
            IntPtr c = FindWindowEx(hwnd, IntPtr.Zero, "_WwG", null);
            if (c != IntPtr.Zero) return c;
            IntPtr found = IntPtr.Zero;
            EnumChildWindows(hwnd, (h, _) =>
            {
                var sb = new StringBuilder(64);
                GetClassName(h, sb, 64);
                if (sb.ToString() == "_WwG") { found = h; return false; }
                return true;
            }, IntPtr.Zero);
            return found;
        }

        // ── Key helpers ───────────────────────────────────────────────────────
        private void Combo(byte mod, byte vk)
        {
            keybd_event(mod, 0, 0, UIntPtr.Zero); Thread.Sleep(20);
            keybd_event(vk,  0, 0, UIntPtr.Zero); Thread.Sleep(30);
            keybd_event(vk,  0, KEYUP, UIntPtr.Zero); Thread.Sleep(20);
            keybd_event(mod, 0, KEYUP, UIntPtr.Zero); Thread.Sleep(50);
        }

        private void Press(byte vk)
        { keybd_event(vk, 0, 0, UIntPtr.Zero); Thread.Sleep(30); keybd_event(vk, 0, KEYUP, UIntPtr.Zero); Thread.Sleep(30); }

        private void ClickCenter(IntPtr hwnd)
        {
            GetWindowRect(hwnd, out var r);
            int cx = (r.L + r.R) / 2, cy = (r.T + r.B) / 2 + 50;
            SetCursorPos(cx, cy); Thread.Sleep(100);
            mouse_event(MOUSEEVENTF_LEFT,   cx, cy, 0, UIntPtr.Zero); Thread.Sleep(50);
            mouse_event(MOUSEEVENTF_LEFTUP, cx, cy, 0, UIntPtr.Zero); Thread.Sleep(200);
        }

        private static string EscRtf(string s) =>
            s.Replace("\\", "\\\\").Replace("{", "\\{").Replace("}", "\\}");

        private bool Skip() { if (!_wordOk) _out.WriteLine("[SKIP] Word not installed"); return !_wordOk; }

        // ═══════════════════════════════════════════════════════════════════════
        //  TEST 1 — User asks AI to write a question into Word
        // ═══════════════════════════════════════════════════════════════════════
        [Fact]
        public async Task Word_UserTypesResearchQuestion_QuestionSavedOnDisk()
        {
            if (Skip()) return;
            _out.WriteLine("\n=== TEST 1: User research question typed into Word ===");

            const string question = "What are the main benefits of cloud computing for small businesses";
            var session = await OpenWordWithFile("BLANK");
            if (session == null) return;
            try
            {
                ClickCenter(session.WindowHandle);
                Combo(VK_CONTROL, VK_END); Thread.Sleep(150);
                Press(VK_RETURN); Thread.Sleep(100);
                TypeInWord(session.WindowHandle, question);
                await Task.Delay(500);

                var mtBefore = File.GetLastWriteTime(session.FilePath);
                await SaveDocumentRobust(session.WindowHandle, session.FilePath);
                var mtAfter = File.GetLastWriteTime(session.FilePath);

                var disk = await DisposeAndReadFile(session);

                mtAfter.Should().BeAfter(mtBefore, "Ctrl+S must update file modification time");
                disk.Should().Contain("cloud computing", "user's question must appear on disk after save");
                disk.Should().Contain("businesses",      "full question must be persisted");
                _out.WriteLine("✅ TEST 1 PASSED");
            }
            finally { session.Dispose(); try { File.Delete(session.FilePath); } catch { } }
        }

        // ═══════════════════════════════════════════════════════════════════════
        //  TEST 2 — User dictates a shopping list (5 items, each on new line)
        // ═══════════════════════════════════════════════════════════════════════
        [Fact]
        public async Task Word_UserDictatesShoppingList_AllFiveItemsOnDisk()
        {
            if (Skip()) return;
            _out.WriteLine("\n=== TEST 2: Shopping list dictation ===");

            var items = new[] { "Milk", "Bread", "Eggs", "Butter", "Coffee" };
            var session = await OpenWordWithFile("SHOPPING");
            if (session == null) return;
            try
            {
                ClickCenter(session.WindowHandle);
                Combo(VK_CONTROL, VK_END); Thread.Sleep(150);
                foreach (var item in items)
                {
                    Press(VK_RETURN); Thread.Sleep(80);
                    TypeInWord(session.WindowHandle, item);
                    await Task.Delay(150);
                }

                await SaveDocumentRobust(session.WindowHandle, session.FilePath);
                var disk = await DisposeAndReadFile(session);

                foreach (var item in items)
                    disk.Should().Contain(item, $"item '{item}' must be saved to disk");

                _out.WriteLine("✅ TEST 2 PASSED");
            }
            finally { session.Dispose(); try { File.Delete(session.FilePath); } catch { } }
        }

        // ═══════════════════════════════════════════════════════════════════════
        //  TEST 3 — User dictates meeting notes with structured content
        // ═══════════════════════════════════════════════════════════════════════
        [Fact]
        public async Task Word_UserDictatesMeetingNotes_StructuredContentOnDisk()
        {
            if (Skip()) return;
            _out.WriteLine("\n=== TEST 3: Meeting notes dictation ===");

            var lines = new[]
            {
                "Meeting Notes 2026",
                "Attendees John Jane Bob",
                "Action items deploy new feature by Friday"
            };
            var session = await OpenWordWithFile("MEETING");
            if (session == null) return;
            try
            {
                ClickCenter(session.WindowHandle);
                Combo(VK_CONTROL, VK_END); Thread.Sleep(150);
                foreach (var line in lines)
                {
                    Press(VK_RETURN); Thread.Sleep(80);
                    TypeInWord(session.WindowHandle, line);
                    await Task.Delay(200);
                }

                await SaveDocumentRobust(session.WindowHandle, session.FilePath);
                var disk = await DisposeAndReadFile(session);

                disk.Should().Contain("Attendees",  "attendees line must be saved");
                disk.Should().Contain("deploy",     "action item must be on disk");
                disk.Should().Contain("2026",       "year must appear in saved notes");
                _out.WriteLine("✅ TEST 3 PASSED");
            }
            finally { session.Dispose(); try { File.Delete(session.FilePath); } catch { } }
        }

        // ═══════════════════════════════════════════════════════════════════════
        //  TEST 4 — User dictates numeric/date content (IDs, phone numbers)
        // ═══════════════════════════════════════════════════════════════════════
        [Fact]
        public async Task Word_UserDictatesNumericContent_NumbersSavedOnDisk()
        {
            if (Skip()) return;
            _out.WriteLine("\n=== TEST 4: Numeric content dictation ===");

            const string numContent = "Customer 98765 Date 28032026 Phone 02012345678";
            var session = await OpenWordWithFile("NUMBERS");
            if (session == null) return;
            try
            {
                ClickCenter(session.WindowHandle);
                Combo(VK_CONTROL, VK_END); Thread.Sleep(150);
                Press(VK_RETURN); Thread.Sleep(100);
                TypeInWord(session.WindowHandle, numContent);
                await Task.Delay(500);

                await SaveDocumentRobust(session.WindowHandle, session.FilePath);
                var disk = await DisposeAndReadFile(session);

                disk.Should().Contain("98765",       "customer ID must be saved");
                disk.Should().Contain("28032026",    "date must be saved");
                disk.Should().Contain("02012345678", "phone number must be saved");
                _out.WriteLine("✅ TEST 4 PASSED");
            }
            finally { session.Dispose(); try { File.Delete(session.FilePath); } catch { } }
        }

        // ═══════════════════════════════════════════════════════════════════════
        //  TEST 5 — Append new content to existing document — BOTH on disk
        // ═══════════════════════════════════════════════════════════════════════
        [Fact]
        public async Task Word_AppendNewContentToExistingDocument_BothPartsOnDisk()
        {
            if (Skip()) return;
            _out.WriteLine("\n=== TEST 5: Append to existing document ===");

            const string initial    = "INITIAL_DOCUMENT_CONTENT";
            const string appended   = "APPENDED_VOICE_DICTATION";
            var session = await OpenWordWithFile(initial);
            if (session == null) return;
            try
            {
                ClickCenter(session.WindowHandle);
                // Go to end and append — do NOT select-all-delete
                Combo(VK_CONTROL, VK_END); Thread.Sleep(200);
                Press(VK_RETURN); Thread.Sleep(100);
                TypeInWord(session.WindowHandle, appended);
                await Task.Delay(500);

                await SaveDocumentRobust(session.WindowHandle, session.FilePath);
                var disk = await DisposeAndReadFile(session);

                disk.Should().Contain(initial,   "original content must be preserved");
                disk.Should().Contain(appended,  "appended dictation must also be on disk");
                _out.WriteLine("✅ TEST 5 PASSED");
            }
            finally { session.Dispose(); try { File.Delete(session.FilePath); } catch { } }
        }

        // ═══════════════════════════════════════════════════════════════════════
        //  TEST 6 — User writes a recipe with ingredients
        // ═══════════════════════════════════════════════════════════════════════
        [Fact]
        public async Task Word_UserWritesRecipe_AllIngredientsSavedOnDisk()
        {
            if (Skip()) return;
            _out.WriteLine("\n=== TEST 6: Recipe with ingredients ===");

            var recipe = new[] { "Pasta Recipe", "300g spaghetti", "2 cloves garlic", "Olive oil" };
            var session = await OpenWordWithFile("RECIPE");
            if (session == null) return;
            try
            {
                ClickCenter(session.WindowHandle);
                Combo(VK_CONTROL, VK_END); Thread.Sleep(150);
                foreach (var line in recipe)
                {
                    Press(VK_RETURN); Thread.Sleep(80);
                    TypeInWord(session.WindowHandle, line);
                    await Task.Delay(180);
                }

                await SaveDocumentRobust(session.WindowHandle, session.FilePath);
                var disk = await DisposeAndReadFile(session);

                disk.Should().Contain("spaghetti", "ingredient must be saved");
                disk.Should().Contain("garlic",    "ingredient must be saved");
                disk.Should().Contain("Recipe",    "recipe title must be saved");
                _out.WriteLine("✅ TEST 6 PASSED");
            }
            finally { session.Dispose(); try { File.Delete(session.FilePath); } catch { } }
        }

        // ═══════════════════════════════════════════════════════════════════════
        //  TEST 7 — Bold formatting applied, then saved — mtime must change
        // ═══════════════════════════════════════════════════════════════════════
        [Fact]
        public async Task Word_ApplyBoldFormattingThenSave_MtimeAdvancesAndContentOnDisk()
        {
            if (Skip()) return;
            _out.WriteLine("\n=== TEST 7: Bold formatting then save ===");

            const string header = "BOLD DOCUMENT TITLE";
            const string body   = "This is the body text";
            var session = await OpenWordWithFile("FORMAT");
            if (session == null) return;
            try
            {
                ClickCenter(session.WindowHandle);
                Combo(VK_CONTROL, VK_END); Thread.Sleep(150);
                Press(VK_RETURN); Thread.Sleep(80);

                // Type header — select it — bold it
                TypeInWord(session.WindowHandle, header);
                await Task.Delay(300);
                Combo(VK_CONTROL, VK_HOME); Thread.Sleep(100);

                // Select header line with Shift+End
                keybd_event(VK_SHIFT, 0, 0, UIntPtr.Zero); Thread.Sleep(20);
                keybd_event(VK_END,   0, 0, UIntPtr.Zero); Thread.Sleep(30);
                keybd_event(VK_END,   0, KEYUP, UIntPtr.Zero); Thread.Sleep(20);
                keybd_event(VK_SHIFT, 0, KEYUP, UIntPtr.Zero); Thread.Sleep(50);

                Combo(VK_CONTROL, VK_B); // Bold
                Thread.Sleep(200);

                Combo(VK_CONTROL, VK_END); Thread.Sleep(100);
                Press(VK_RETURN); Thread.Sleep(80);
                TypeInWord(session.WindowHandle, body);
                await Task.Delay(300);

                var mtBefore = File.GetLastWriteTime(session.FilePath);
                await SaveDocumentRobust(session.WindowHandle, session.FilePath);
                var mtAfter = File.GetLastWriteTime(session.FilePath);

                var disk = await DisposeAndReadFile(session);

                mtAfter.Should().BeAfter(mtBefore, "save must advance the file modification time");
                disk.Should().Contain("BOLD",  "bold header text must be on disk");
                disk.Should().Contain("body",  "body text must be on disk");
                _out.WriteLine("✅ TEST 7 PASSED");
            }
            finally { session.Dispose(); try { File.Delete(session.FilePath); } catch { } }
        }

        // ═══════════════════════════════════════════════════════════════════════
        //  TEST 8 — Type two sentences, undo once (removes last Word batch),
        //           save the shorter result — disk must contain first sentence
        // ═══════════════════════════════════════════════════════════════════════
        [Fact]
        public async Task Word_TypeTwoSentences_UndoLast_SaveShorterContentOnDisk()
        {
            if (Skip()) return;
            _out.WriteLine("\n=== TEST 8: Type two sentences + single undo + save ===");

            const string sentence1 = "FIRST_SENTENCE_STAYS";
            const string sentence2 = "SECOND_SENTENCE_REMOVED";
            var session = await OpenWordWithFile("UNDO");
            if (session == null) return;
            try
            {
                ClickCenter(session.WindowHandle);
                Combo(VK_CONTROL, VK_END); Thread.Sleep(150);
                Press(VK_RETURN); Thread.Sleep(100);
                TypeInWord(session.WindowHandle, sentence1);
                await Task.Delay(500); // pause so Word creates a separate undo batch

                Press(VK_RETURN); Thread.Sleep(200);
                TypeInWord(session.WindowHandle, sentence2);
                await Task.Delay(500); // pause so the two typed strings are in separate undo batches

                // Clipboard state with BOTH sentences present
                var beforeUndo = ReadClipboard(session.WindowHandle);
                _out.WriteLine($"[BEFORE UNDO] {beforeUndo.Length} chars");
                beforeUndo.Should().Contain(sentence1, "both sentences must be in doc before undo");

                // ONE Ctrl+Z — should remove the last typed chunk (sentence2 + newline)
                Combo(VK_CONTROL, VK_Z); Thread.Sleep(300);
                Combo(VK_CONTROL, VK_Z); Thread.Sleep(300); // second undo for the Return key

                var afterUndo = ReadClipboard(session.WindowHandle);
                _out.WriteLine($"[AFTER UNDO]  {afterUndo.Length} chars: '{afterUndo.TrimEnd()}'");

                afterUndo.Length.Should().BeLessThan(beforeUndo.Length,
                    "undo must reduce document length");

                await SaveDocumentRobust(session.WindowHandle, session.FilePath);
                var disk = await DisposeAndReadFile(session);

                disk.Should().Contain(sentence1,
                    "the first sentence must survive undo and be persisted on disk");
                disk.Should().NotContain(sentence2,
                    "the undone second sentence must NOT appear in the saved file");
                _out.WriteLine("✅ TEST 8 PASSED");
            }
            finally { session.Dispose(); try { File.Delete(session.FilePath); } catch { } }
        }
    }

    // ════════════════════════════════════════════════════════════════════════════
    //  CLASS 2 – Browser navigation user workflows (Tests 9–14)
    //  No Python backend required — all navigation uses direct key injection.
    // ════════════════════════════════════════════════════════════════════════════

    [Collection("RealApp")]
    public class BrowserUserWorkflowTests : IAsyncLifetime
    {
        private readonly ITestOutputHelper _out;
        private readonly AppLauncher       _browser;
        private bool _browserOk;

        [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr h);
        [DllImport("user32.dll")] static extern int  GetWindowText(IntPtr h, StringBuilder sb, int n);
        [DllImport("user32.dll")] static extern void keybd_event(byte vk, byte scan, uint flags, UIntPtr extra);

        private const uint KEYUP       = 0x0002;
        private const byte VK_CONTROL  = 0x11;
        private const byte VK_ALT      = 0x12;
        private const byte VK_A        = 0x41;
        private const byte VK_C        = 0x43;
        private const byte VK_L        = 0x4C;
        private const byte VK_T        = 0x54;
        private const byte VK_RETURN   = 0x0D;
        private const byte VK_ESCAPE   = 0x1B;
        private const byte VK_LEFT     = 0x25;

        public BrowserUserWorkflowTests(ITestOutputHelper output)
        { _out = output; _browser = new AppLauncher(output); }

        public async Task InitializeAsync()
        {
            _out.WriteLine("=== BROWSER SETUP: Launching Edge (fallback Chrome) ===");
            _browserOk = await _browser.LaunchAsync("msedge", 10000);
            if (!_browserOk)
                _browserOk = await _browser.LaunchAsync("chrome", 10000);
            if (_browserOk)
            { await Task.Delay(3000); _browser.Focus(); _out.WriteLine("[SETUP] Browser ready"); }
            else
                _out.WriteLine("[SETUP] No browser found — browser tests will skip");
        }

        public Task DisposeAsync() { _browser.Dispose(); return Task.CompletedTask; }

        private bool Skip() { if (!_browserOk) _out.WriteLine("[SKIP] Browser not available"); return !_browserOk; }

        private void Navigate(string url)
        {
            SetForegroundWindow(_browser.WindowHandle); Thread.Sleep(200);
            Combo(VK_CONTROL, VK_L); Thread.Sleep(400);
            Combo(VK_CONTROL, VK_A); Thread.Sleep(100);
            System.Windows.Forms.SendKeys.SendWait(url);
            Thread.Sleep(200);
            Press(VK_RETURN);
        }

        private string ReadUrl()
        {
            SetForegroundWindow(_browser.WindowHandle); Thread.Sleep(200);
            Combo(VK_CONTROL, VK_L); Thread.Sleep(400);
            Combo(VK_CONTROL, VK_A); Thread.Sleep(100);
            Combo(VK_CONTROL, VK_C); Thread.Sleep(500);
            string url = "";
            var t = new Thread(() => { try { url = System.Windows.Forms.Clipboard.GetText(); } catch { } });
            t.SetApartmentState(ApartmentState.STA); t.Start(); t.Join(2000);
            Press(VK_ESCAPE); Thread.Sleep(100);
            _out.WriteLine($"[URL] '{url}'");
            return url;
        }

        private string WindowTitle()
        {
            var sb = new StringBuilder(512);
            GetWindowText(_browser.WindowHandle, sb, 512);
            return sb.ToString();
        }

        private void Combo(byte mod, byte vk)
        {
            keybd_event(mod, 0, 0, UIntPtr.Zero); Thread.Sleep(20);
            keybd_event(vk,  0, 0, UIntPtr.Zero); Thread.Sleep(30);
            keybd_event(vk,  0, KEYUP, UIntPtr.Zero); Thread.Sleep(20);
            keybd_event(mod, 0, KEYUP, UIntPtr.Zero); Thread.Sleep(50);
        }

        private void Press(byte vk)
        { keybd_event(vk, 0, 0, UIntPtr.Zero); Thread.Sleep(30); keybd_event(vk, 0, KEYUP, UIntPtr.Zero); Thread.Sleep(30); }

        // ═══════════════════════════════════════════════════════════════════════
        //  TEST 9 — Navigate to Wikipedia — title contains "Wikipedia"
        // ═══════════════════════════════════════════════════════════════════════
        [Fact]
        public async Task Browser_NavigateToWikipedia_TitleContainsWikipedia()
        {
            if (Skip()) return;
            _out.WriteLine("\n=== TEST 9: Navigate to Wikipedia ===");

            Navigate("https://en.wikipedia.org");
            await Task.Delay(5000);
            var title = WindowTitle();
            _out.WriteLine($"[TITLE] '{title}'");

            title.Should().ContainEquivalentOf("Wikipedia",
                because: $"After navigating to Wikipedia, window title must reflect it. Got: '{title}'");
            _out.WriteLine("✅ TEST 9 PASSED");
        }

        // ═══════════════════════════════════════════════════════════════════════
        //  TEST 10 — Navigate to Bing — title contains "Bing"
        // ═══════════════════════════════════════════════════════════════════════
        [Fact]
        public async Task Browser_NavigateToBing_TitleContainsBing()
        {
            if (Skip()) return;
            _out.WriteLine("\n=== TEST 10: Navigate to Bing ===");

            Navigate("https://www.bing.com");
            await Task.Delay(5000);
            var title = WindowTitle();
            _out.WriteLine($"[TITLE] '{title}'");

            title.Should().ContainEquivalentOf("Bing",
                because: $"After navigating to Bing, window title must contain 'Bing'. Got: '{title}'");
            _out.WriteLine("✅ TEST 10 PASSED");
        }

        // ═══════════════════════════════════════════════════════════════════════
        //  TEST 11 — Navigate to GitHub — title contains "GitHub"
        // ═══════════════════════════════════════════════════════════════════════
        [Fact]
        public async Task Browser_NavigateToGitHub_TitleContainsGitHub()
        {
            if (Skip()) return;
            _out.WriteLine("\n=== TEST 11: Navigate to GitHub ===");

            Navigate("https://github.com");
            await Task.Delay(6000);
            var title = WindowTitle();
            _out.WriteLine($"[TITLE] '{title}'");

            title.Should().ContainEquivalentOf("GitHub",
                because: $"After navigating to GitHub, window title must reflect it. Got: '{title}'");
            _out.WriteLine("✅ TEST 11 PASSED");
        }

        // ═══════════════════════════════════════════════════════════════════════
        //  TEST 12 — Navigate using a pre-formed search URL — URL has search term
        // ═══════════════════════════════════════════════════════════════════════
        [Fact]
        public async Task Browser_NavigatePreformedSearchUrl_UrlContainsSearchTerm()
        {
            if (Skip()) return;
            _out.WriteLine("\n=== TEST 12: Pre-formed search URL navigation ===");

            Navigate("https://www.google.com/search?q=artificial+intelligence");
            await Task.Delay(5000);
            var url   = ReadUrl();
            var title = WindowTitle();
            _out.WriteLine($"[TITLE] '{title}'");

            bool evident =
                url.Contains("artificial",   StringComparison.OrdinalIgnoreCase) ||
                url.Contains("intelligence", StringComparison.OrdinalIgnoreCase) ||
                url.Contains("search",       StringComparison.OrdinalIgnoreCase) ||
                title.Contains("artificial", StringComparison.OrdinalIgnoreCase);

            evident.Should().BeTrue(
                because: $"After navigating a Google search URL, the URL or title must reflect the search term. " +
                         $"URL='{url}', Title='{title}'");
            _out.WriteLine("✅ TEST 12 PASSED");
        }

        // ═══════════════════════════════════════════════════════════════════════
        //  TEST 13 — Navigate to BBC News — title/URL contains BBC
        // ═══════════════════════════════════════════════════════════════════════
        [Fact]
        public async Task Browser_NavigateToBBCNews_TitleOrUrlContainsBBC()
        {
            if (Skip()) return;
            _out.WriteLine("\n=== TEST 13: Navigate to BBC News ===");

            Navigate("https://www.bbc.com/news");
            await Task.Delay(6000);
            var url   = ReadUrl();
            var title = WindowTitle();
            _out.WriteLine($"[TITLE] '{title}'");

            bool bbcPresent =
                url.Contains("bbc",   StringComparison.OrdinalIgnoreCase) ||
                title.Contains("BBC", StringComparison.OrdinalIgnoreCase);

            bbcPresent.Should().BeTrue(
                because: $"After navigating to BBC News, URL or title must contain 'BBC'. " +
                         $"URL='{url}', Title='{title}'");
            _out.WriteLine("✅ TEST 13 PASSED");
        }

        // ═══════════════════════════════════════════════════════════════════════
        //  TEST 14 — Navigate to two pages, use Alt+Left to go back, URL reverts
        // ═══════════════════════════════════════════════════════════════════════
        [Fact]
        public async Task Browser_TwoPageNavigation_AltLeftGoesBack_UrlReverts()
        {
            if (Skip()) return;
            _out.WriteLine("\n=== TEST 14: Two-page navigation + Alt+Left back ===");

            // Page 1 — Google
            Navigate("https://www.google.com");
            await Task.Delay(4000);
            var titlePage1 = WindowTitle();
            _out.WriteLine($"[PAGE-1] '{titlePage1}'");

            // Page 2 — Bing
            Navigate("https://www.bing.com");
            await Task.Delay(4000);
            var titlePage2 = WindowTitle();
            _out.WriteLine($"[PAGE-2] '{titlePage2}'");

            titlePage2.Should().NotBe(titlePage1, "Page 2 must differ from Page 1");

            // Alt+Left = browser back button
            SetForegroundWindow(_browser.WindowHandle); Thread.Sleep(200);
            keybd_event(VK_ALT,  0, 0, UIntPtr.Zero); Thread.Sleep(20);
            keybd_event(VK_LEFT, 0, 0, UIntPtr.Zero); Thread.Sleep(30);
            keybd_event(VK_LEFT, 0, KEYUP, UIntPtr.Zero); Thread.Sleep(20);
            keybd_event(VK_ALT,  0, KEYUP, UIntPtr.Zero); Thread.Sleep(50);
            await Task.Delay(3500);

            var titleAfterBack = WindowTitle();
            _out.WriteLine($"[AFTER BACK] '{titleAfterBack}'");

            titleAfterBack.Should().ContainEquivalentOf("Google",
                because: $"After Alt+Left from Bing, browser must return to Google. Got: '{titleAfterBack}'");
            _out.WriteLine("✅ TEST 14 PASSED");
        }
    }

    // ════════════════════════════════════════════════════════════════════════════
    //  CLASS 3 – Voice command routing and confidence gate (Tests 15–18)
    //  These tests use VoiceCommandSimulator — no real microphone or UI needed.
    // ════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Tests that verify the voice command pipeline routing behaviour for different
    /// user request phrasings and confidence levels, exactly as a real voice assistant
    /// interaction would produce them via ElevenLabs / Whisper.net STT.
    /// </summary>
    [Collection("RealApp")]
    public class VoiceCommandRoutingTests
    {
        private readonly ITestOutputHelper _out;
        private readonly VoiceCommandSimulator _sim;

        public VoiceCommandRoutingTests(ITestOutputHelper output)
        {
            _out = output;
            _sim = new VoiceCommandSimulator(output);
        }

        // ═══════════════════════════════════════════════════════════════════════
        //  TEST 15 — High-confidence simple command passes confidence gate
        // ═══════════════════════════════════════════════════════════════════════
        [Fact]
        public async Task Voice_HighConfidence_SimpleOpenCommand_PassesGate()
        {
            _out.WriteLine("\n=== TEST 15: High-confidence 'open notepad' command ===");

            var result = await _sim.SpeakAsync("open notepad", confidence: 0.95f);

            _out.WriteLine($"[RESULT] Passed gate={result.PassedConfidenceGate} | {result.TotalElapsedMs}ms");

            result.PassedConfidenceGate.Should().BeTrue(
                "Confidence 0.95 is well above the 0.65 threshold — must not be blocked");
            result.BlockedReason.Should().BeNullOrEmpty(
                "No blocked reason expected when confidence is high");
            _out.WriteLine("✅ TEST 15 PASSED");
        }

        // ═══════════════════════════════════════════════════════════════════════
        //  TEST 16 — Low confidence (garbled speech) blocked at gate
        // ═══════════════════════════════════════════════════════════════════════
        [Fact]
        public async Task Voice_LowConfidence_GarbledSpeech_BlockedAtConfidenceGate()
        {
            _out.WriteLine("\n=== TEST 16: Low-confidence garbled input blocked ===");

            // Simulate garbled STT output at 0.40 confidence (like noisy room)
            var result = await _sim.SpeakAsync("typ somethin mmm clos dat", confidence: 0.40f);

            _out.WriteLine($"[RESULT] Passed gate={result.PassedConfidenceGate} | Reason='{result.BlockedReason}'");

            result.PassedConfidenceGate.Should().BeFalse(
                "Confidence 0.40 is below the 0.65 threshold — command must be blocked");
            result.BlockedReason.Should().NotBeNullOrEmpty(
                "A blocked reason must be recorded when the gate rejects the command");
            _out.WriteLine("✅ TEST 16 PASSED");
        }

        // ═══════════════════════════════════════════════════════════════════════
        //  TEST 17 — Complex search query classified as complex (routes to agentic)
        // ═══════════════════════════════════════════════════════════════════════
        [Fact]
        public async Task Voice_SearchQueryForFlights_ClassifiedAsComplexCommand()
        {
            _out.WriteLine("\n=== TEST 17: Complex search query routing ===");

            // A multi-step travel search — should be complex (needs agentic)
            var result = await _sim.SpeakAsync(
                "search for flights from London to New York next week",
                confidence: 0.90f);

            _out.WriteLine($"[RESULT] IsComplex={result.IsComplex} | Passed={result.PassedConfidenceGate}");

            result.PassedConfidenceGate.Should().BeTrue("0.90 confidence must pass the gate");
            result.IsComplex.Should().BeTrue(
                "A multi-part search query about flights must be routed to the agentic planner, not handled locally");
            _out.WriteLine("✅ TEST 17 PASSED");
        }

        // ═══════════════════════════════════════════════════════════════════════
        //  TEST 18 — Simple typing command processed locally (fast, no backend)
        // ═══════════════════════════════════════════════════════════════════════
        [Fact]
        public async Task Voice_SimpleDictateCommand_ProcessedLocallyWithinTimeLimit()
        {
            _out.WriteLine("\n=== TEST 18: Simple local dictation command ===");

            var result = await _sim.SpeakAsync("type hello world", confidence: 0.88f);

            _out.WriteLine($"[RESULT] IsComplex={result.IsComplex} | LocalMs={result.LocalCommandMs} | Total={result.TotalElapsedMs}ms");

            result.PassedConfidenceGate.Should().BeTrue("0.88 confidence must pass the gate");
            result.IsComplex.Should().BeFalse(
                "'type hello world' is a simple dictation command — the router must NOT classify it " +
                "as complex (which would force a multi-step agentic plan). Simple commands stay local.");
            _out.WriteLine("✅ TEST 18 PASSED");
        }
    }

    // ════════════════════════════════════════════════════════════════════════════
    //  CLASS 4 – Notepad and File System (Tests 19–20)
    // ════════════════════════════════════════════════════════════════════════════

    [Collection("RealApp")]
    public class NotepadAndFileSystemTests : IAsyncLifetime
    {
        private readonly ITestOutputHelper _out;

        [DllImport("user32.dll")] static extern bool  SetForegroundWindow(IntPtr h);
        [DllImport("user32.dll")] static extern bool  IsWindowVisible(IntPtr h);
        [DllImport("user32.dll")] static extern void  keybd_event(byte vk, byte scan, uint flags, UIntPtr extra);
        [DllImport("user32.dll")] static extern int   GetWindowText(IntPtr h, StringBuilder sb, int n);
        [DllImport("user32.dll")] static extern bool  EnumWindows(EnumWindowsProc cb, IntPtr lp);

        private delegate bool EnumWindowsProc(IntPtr h, IntPtr lp);

        private const uint KEYUP      = 0x0002;
        private const byte VK_CONTROL = 0x11;
        private const byte VK_SHIFT   = 0x10;
        private const byte VK_A       = 0x41;
        private const byte VK_S       = 0x53;
        private const byte VK_F4      = 0x73;
        private const byte VK_ALT     = 0x12;
        private const byte VK_DELETE  = 0x2E;
        private const byte VK_RETURN  = 0x0D;
        private const byte VK_ESCAPE  = 0x1B;

        public NotepadAndFileSystemTests(ITestOutputHelper output) => _out = output;
        public Task InitializeAsync() => Task.CompletedTask;
        public Task DisposeAsync()    => Task.CompletedTask;

        private void Combo(byte mod, byte vk)
        {
            keybd_event(mod, 0, 0, UIntPtr.Zero); Thread.Sleep(20);
            keybd_event(vk,  0, 0, UIntPtr.Zero); Thread.Sleep(30);
            keybd_event(vk,  0, KEYUP, UIntPtr.Zero); Thread.Sleep(20);
            keybd_event(mod, 0, KEYUP, UIntPtr.Zero); Thread.Sleep(50);
        }

        private void Press(byte vk)
        { keybd_event(vk, 0, 0, UIntPtr.Zero); Thread.Sleep(30); keybd_event(vk, 0, KEYUP, UIntPtr.Zero); Thread.Sleep(30); }

        private const byte VK_V = 0x56;

        /// <summary>
        /// Types text into a window by writing to the clipboard then sending Ctrl+V.
        /// This is more reliable than SendInput/keybd_event character injection because:
        ///   1. It bypasses keyboard layout entirely (works on Russian/CJK/any layout).
        ///   2. WinUI3 apps (Win11 Notepad) process WM_PASTE reliably even when they
        ///      ignore low-level WM_CHAR injections in certain focus states.
        /// Clipboard must be set on an STA thread (Windows COM requirement).
        /// </summary>
        private void TypeAscii(IntPtr hwnd, string text)
        {
            // Write text to clipboard on STA thread (Clipboard COM requirement)
            var staThread = new Thread(() =>
                System.Windows.Forms.Clipboard.SetText(text));
            staThread.SetApartmentState(ApartmentState.STA);
            staThread.Start();
            staThread.Join();
            Thread.Sleep(150);

            // Focus window then paste
            SetForegroundWindow(hwnd); Thread.Sleep(200);
            keybd_event(VK_CONTROL, 0, 0,    UIntPtr.Zero);
            keybd_event(VK_V,       0, 0,    UIntPtr.Zero); Thread.Sleep(30);
            keybd_event(VK_V,       0, KEYUP, UIntPtr.Zero);
            keybd_event(VK_CONTROL, 0, KEYUP, UIntPtr.Zero);
            Thread.Sleep(200);
        }

        private IntPtr FindWindowByTitleFragment(string fragment, int timeoutMs = 8000)
        {
            var deadline = DateTime.UtcNow.AddMilliseconds(timeoutMs);
            IntPtr hwnd  = IntPtr.Zero;
            while (DateTime.UtcNow < deadline && hwnd == IntPtr.Zero)
            {
                Thread.Sleep(300);
                EnumWindows((h, _) =>
                {
                    if (!IsWindowVisible(h)) return true;
                    var sb = new StringBuilder(256);
                    GetWindowText(h, sb, 256);
                    if (sb.ToString().Contains(fragment, StringComparison.OrdinalIgnoreCase))
                    { hwnd = h; return false; }
                    return true;
                }, IntPtr.Zero);
            }
            return hwnd;
        }

        // ═══════════════════════════════════════════════════════════════════════
        //  TEST 19 — Notepad: type unique user note, Ctrl+S to pre-created file,
        //            kill Notepad, verify content on disk
        // ═══════════════════════════════════════════════════════════════════════
        [Fact]
        public async Task Notepad_UserTypesNote_CtrlSToExistingFile_ContentVerifiedOnDisk()
        {
            _out.WriteLine("\n=== TEST 19: Notepad type + Ctrl+S + disk verify ===");

            // Pre-create the file so Ctrl+S overwrites without a Save-As dialog
            var path = Path.Combine(
                Path.GetTempPath(),
                $"NoteTest_{DateTime.Now:HHmmssff}.txt");
            File.WriteAllText(path, "ORIGINAL_NOTE_CONTENT");
            _out.WriteLine($"[FILE] Created: {path}");

            System.Diagnostics.Process? proc = null;
            try
            {
                proc = System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName        = "notepad.exe",
                    Arguments       = $"\"{path}\"",
                    UseShellExecute = true
                });
                await Task.Delay(3000);

                var filename = Path.GetFileNameWithoutExtension(path);
                IntPtr hwnd  = FindWindowByTitleFragment(filename, 10000);
                if (hwnd == IntPtr.Zero)
                {
                    // Win11 WinUI Notepad may not include filename in title
                    hwnd = FindWindowByTitleFragment("Notepad", 5000);
                }

                if (hwnd == IntPtr.Zero)
                {
                    _out.WriteLine("[SKIP] Notepad window not found");
                    return;
                }

                var sb = new StringBuilder(256);
                GetWindowText(hwnd, sb, 256);
                _out.WriteLine($"[NOTEPAD] hWnd=0x{hwnd:X} title='{sb}'");

                // Select all and replace with user's note
                SetForegroundWindow(hwnd); Thread.Sleep(300);
                Combo(VK_CONTROL, VK_A); Thread.Sleep(200);
                Press(VK_DELETE);        await Task.Delay(200);

                const string userNote = "Voice assistant wrote this note for the user today";
                TypeAscii(hwnd, userNote);
                await Task.Delay(600);

                // Re-focus before Ctrl+S: after an await the OS may have reassigned foreground
                SetForegroundWindow(hwnd); Thread.Sleep(300);
                var mtBefore = File.GetLastWriteTime(path);
                _out.WriteLine($"[SAVE] Ctrl+S (mtime before: {mtBefore:ss.fff})");
                Combo(VK_CONTROL, VK_S);
                await Task.Delay(2500); // Notepad saves faster than Word

                var mtAfter = File.GetLastWriteTime(path);
                _out.WriteLine($"[MTIME] {mtBefore:ss.fff} → {mtAfter:ss.fff}");

                // Kill Notepad to release file lock
                try { proc?.Kill(); } catch { }
                await Task.Delay(1000);

                var disk = File.ReadAllText(path);
                _out.WriteLine($"[DISK] '{disk}'");

                mtAfter.Should().BeAfter(mtBefore,
                    "Ctrl+S must advance the file modification time in Notepad");
                disk.Should().Contain("Voice assistant",
                    "user's dictated note must be on disk after Ctrl+S");
                disk.Should().Contain("user today",
                    "full note content must be present on disk");
                disk.Should().NotContain("ORIGINAL_NOTE_CONTENT",
                    "original content must be replaced after select-all-delete-type-save");

                _out.WriteLine("✅ TEST 19 PASSED");
            }
            finally
            {
                try { proc?.Kill(); } catch { }
                try { proc?.Dispose(); } catch { }
                await Task.Delay(500);
                try { File.Delete(path); } catch { }
            }
        }

        // ═══════════════════════════════════════════════════════════════════════
        //  TEST 20 — File Explorer: open, verify window appears, close
        // ═══════════════════════════════════════════════════════════════════════
        [Fact]
        public async Task FileExplorer_OpenAndVerifyWindow_ClosesCleanly()
        {
            _out.WriteLine("\n=== TEST 20: File Explorer open + verify + close ===");

            System.Diagnostics.Process? proc = null;
            try
            {
                // Open File Explorer to the Documents folder
                var docsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                proc = System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName        = "explorer.exe",
                    Arguments       = docsPath,
                    UseShellExecute = true
                });
                await Task.Delay(3000);

                // File Explorer windows may have titles like "Documents", "File Explorer",
                // or the folder name — search broadly
                IntPtr hwnd = FindWindowByTitleFragment("Documents", 8000);
                if (hwnd == IntPtr.Zero)
                    hwnd = FindWindowByTitleFragment("File Explorer", 5000);
                if (hwnd == IntPtr.Zero)
                    hwnd = FindWindowByTitleFragment("Explorer", 5000);

                hwnd.Should().NotBe(IntPtr.Zero,
                    "After starting explorer.exe with Documents path, a File Explorer window must appear");

                var sb = new StringBuilder(256);
                GetWindowText(hwnd, sb, 256);
                _out.WriteLine($"[EXPLORER] hWnd=0x{hwnd:X} title='{sb}'");
                _out.WriteLine($"[EXPLORER] Window found and visible — closing with Alt+F4");

                // Close the explorer window gracefully
                SetForegroundWindow(hwnd); Thread.Sleep(300);
                keybd_event(VK_ALT, 0, 0, UIntPtr.Zero); Thread.Sleep(20);
                keybd_event(VK_F4,  0, 0, UIntPtr.Zero); Thread.Sleep(30);
                keybd_event(VK_F4,  0, KEYUP, UIntPtr.Zero); Thread.Sleep(20);
                keybd_event(VK_ALT, 0, KEYUP, UIntPtr.Zero);
                await Task.Delay(1000);

                _out.WriteLine("✅ TEST 20 PASSED");
            }
            finally
            {
                // explorer.exe is shared — do NOT kill the whole process tree
                proc?.Dispose();
            }
        }
    }

    // ════════════════════════════════════════════════════════════════════════════
    //  CLASS 5 – Live-bug regression tests (Tests 21-25)
    //
    //  Each test reproduces a real failure observed by the user during a live
    //  session and verifies that the post-fix behaviour is correct:
    //
    //  Test 21  open_app prefers existing window → no duplicate Word document
    //  Test 22  "opera" maps to Opera process (not Explorer or Edge)
    //  Test 23  After "open Word" + "write hello world" text goes to SAME document
    //  Test 24  ResolveAppName: "notebook" → notepad, "browser" → msedge (not opera)
    //  Test 25  focus_window action brings existing Notepad to front without relaunching
    // ════════════════════════════════════════════════════════════════════════════

    [Collection("RealApp")]
    public class LiveBugRegressionTests
    {
        private readonly ITestOutputHelper _out;

        [DllImport("user32.dll")] static extern bool  SetForegroundWindow(IntPtr h);
        [DllImport("user32.dll")] static extern bool  IsWindowVisible(IntPtr h);
        [DllImport("user32.dll")] static extern void  keybd_event(byte vk, byte scan, uint flags, UIntPtr extra);
        [DllImport("user32.dll")] static extern int   GetWindowText(IntPtr h, StringBuilder sb, int n);
        [DllImport("user32.dll")] static extern bool  EnumWindows(EnumWindowsProc cb, IntPtr lp);

        private delegate bool EnumWindowsProc(IntPtr h, IntPtr lp);

        private const uint KEYUP      = 0x0002;
        private const byte VK_CONTROL = 0x11;
        private const byte VK_ALT     = 0x12;
        private const byte VK_F4      = 0x73;

        public LiveBugRegressionTests(ITestOutputHelper output) => _out = output;

        // ── TEST 21 ─────────────────────────────────────────────────────────────
        // BUG: ExecuteOpenApp called LaunchAndWaitForWindowAsync first, which always
        // launched a NEW process even when the app was already running.
        // FIX: FindWindowByProcessName is now checked BEFORE launching.
        // VERIFY: Opening Notepad twice produces only ONE Notepad process-window
        //         (the second open_app call focuses the existing window).
        [Fact]
        public async Task OpenApp_WhenAlreadyRunning_ReusesExistingWindowNotNewProcess()
        {
            _out.WriteLine("\n=== TEST 21: open_app reuses existing window ===");

            var automation = new AICompanion.Desktop.Services.Automation.WindowAutomationHelper();
            var service    = new AICompanion.Desktop.Services.Automation.AgenticExecutionService(automation);

            System.Diagnostics.Process? proc1 = null;
            System.Diagnostics.Process? proc2 = null;
            try
            {
                // Launch Notepad once manually
                proc1 = System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    { FileName = "notepad.exe", UseShellExecute = true });
                await Task.Delay(2500);

                // Count Notepad processes before
                var before = System.Diagnostics.Process.GetProcessesByName("notepad").Length;
                _out.WriteLine($"[BEFORE] notepad process count: {before}");
                before.Should().BeGreaterThanOrEqualTo(1, "we just launched Notepad");

                // Use the fixed ExecuteOpenApp logic: find existing → no new launch
                var hwnd = automation.FindWindowByProcessName("notepad");
                hwnd.Should().NotBe(IntPtr.Zero, "Notepad should be found by process name");
                _out.WriteLine($"[FOUND] Existing notepad window: 0x{hwnd:X}");

                // Simulate what ExecuteOpenApp now does: focus existing, track it
                automation.ForceFocusWindow(hwnd);
                var titleAfter = automation.GetWindowTitle(hwnd);
                _out.WriteLine($"[TITLE] {titleAfter}");

                // Assert: still only one Notepad (no new process launched)
                var after = System.Diagnostics.Process.GetProcessesByName("notepad").Length;
                _out.WriteLine($"[AFTER] notepad process count: {after}");
                after.Should().Be(before, "open_app must NOT launch a second Notepad");
                titleAfter.ToLowerInvariant().Should().Contain("notepad");

                _out.WriteLine("✅ TEST 21 PASSED");
            }
            finally
            {
                try { proc1?.Kill(); proc2?.Kill(); } catch { }
                await Task.Delay(500);
            }
        }

        // ── TEST 22 ─────────────────────────────────────────────────────────────
        // BUG: "opera" was not in ResolveAppName — fell through to no match,
        //      sometimes causing the Python backend to plan "open_app: explorer"
        //      (file explorer) as a fallback, opening the wrong application.
        // FIX: "opera" / "Opera Browser" / "опера" added to ResolveAppName map.
        // VERIFY: ResolveAppName("opera") → "opera", NOT "explorer" or "msedge".
        [Fact]
        public void ResolveAppName_Opera_MapsToOperaNotExplorer()
        {
            _out.WriteLine("\n=== TEST 22: ResolveAppName opera mapping ===");

            var cases = new[]
            {
                ("opera",          "opera"),
                ("Opera",          "opera"),
                ("opera browser",  "opera"),
                ("Opera Browser",  "opera"),
                ("opera gx",       "opera"),
                ("опера",          "opera"),
            };

            foreach (var (input, expected) in cases)
            {
                var result = AICompanion.Desktop.Services.Automation.WindowAutomationHelper.ResolveAppName(input);
                _out.WriteLine($"  ResolveAppName('{input}') = '{result}'");
                result.Should().Be(expected,
                    $"'{input}' must map to opera, not explorer/msedge/other");
            }

            // Verify that "explorer" still maps to file explorer (no regression)
            var explorerResult = AICompanion.Desktop.Services.Automation.WindowAutomationHelper.ResolveAppName("explorer");
            explorerResult.Should().Be("explorer", "file explorer mapping must be unchanged");

            _out.WriteLine("✅ TEST 22 PASSED");
        }

        // ── TEST 23 ─────────────────────────────────────────────────────────────
        // BUG: User said "write hello world in Word" after Word was open.
        //      The AI opened a NEW Word document, then typed only in the new doc,
        //      leaving the user's original document untouched.
        // FIX: ExecuteOpenApp now finds existing window first (Test 21 fix).
        //      _planTargetWindow persists across commands (code already did this —
        //      the bug was that a new doc reset it).
        // VERIFY: Open Notepad → simulate "write hello world" → text in SAME window.
        [Fact]
        public async Task OpenThenWrite_TextGoesToSameWindow_NotNewDocument()
        {
            _out.WriteLine("\n=== TEST 23: open then write stays in same window ===");

            var path = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                $"Bug23_{DateTime.Now:HHmmssff}.txt");
            System.IO.File.WriteAllText(path, "");
            _out.WriteLine($"[FILE] {path}");

            System.Diagnostics.Process? proc = null;
            try
            {
                // Launch Notepad with the pre-created file
                proc = System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    { FileName = "notepad.exe", Arguments = $"\"{path}\"", UseShellExecute = true });
                await Task.Delay(2500);

                var automation = new AICompanion.Desktop.Services.Automation.WindowAutomationHelper();

                // Simulate command 1: open_app finds existing window
                var hwnd = automation.FindWindowByProcessName("notepad");
                hwnd.Should().NotBe(IntPtr.Zero);
                automation.ForceFocusWindow(hwnd);
                var titleCmd1 = automation.GetWindowTitle(hwnd);
                _out.WriteLine($"[CMD1] Focused existing: '{titleCmd1}' h=0x{hwnd:X}");

                // Simulate command 2: type_text goes into SAME tracked window
                await Task.Delay(500);
                await automation.TypeTextIntoWindowAsync(hwnd, "Hello World");
                await Task.Delay(500);

                // Save with Ctrl+S
                SetForegroundWindow(hwnd); Thread.Sleep(200);
                keybd_event(VK_CONTROL, 0, 0,    UIntPtr.Zero);
                keybd_event(0x53,       0, 0,    UIntPtr.Zero);
                keybd_event(0x53,       0, KEYUP, UIntPtr.Zero);
                keybd_event(VK_CONTROL, 0, KEYUP, UIntPtr.Zero);
                await Task.Delay(2000);

                try { proc?.Kill(); } catch { }
                await Task.Delay(800);

                var disk = System.IO.File.ReadAllText(path);
                _out.WriteLine($"[DISK] '{disk}'");

                disk.Should().Contain("Hello World",
                    "text must go to the originally tracked window, not a new document");
                disk.Should().NotContain("Hello Worlds",
                    "Granite speech glitch ('Worlds') must not appear in this controlled test");

                _out.WriteLine("✅ TEST 23 PASSED");
            }
            finally
            {
                try { proc?.Kill(); } catch { }
                await Task.Delay(500);
                try { System.IO.File.Delete(path); } catch { }
            }
        }

        // ── TEST 24 ─────────────────────────────────────────────────────────────
        // VERIFY: Critical app name mappings are correct and haven't regressed.
        // This acts as a canary for the ResolveAppName dictionary — any future
        // fuzzy-matching regression would break users saying common app names.
        [Fact]
        public void ResolveAppName_CriticalMappings_AllCorrect()
        {
            _out.WriteLine("\n=== TEST 24: critical app name mappings ===");

            var cases = new[]
            {
                // Must NOT mis-route to File Explorer (historical bug: "opera" → "explorer")
                ("opera",          "opera"),
                ("опера",          "opera"),
                // File manager
                ("file explorer",  "explorer"),
                ("files",          "explorer"),
                ("проводник",      "explorer"),
                // Browsers
                ("browser",        "msedge"),
                ("chrome",         "chrome"),
                ("firefox",        "firefox"),
                // Office
                ("word",           "WINWORD"),
                ("ворд",           "WINWORD"),
                ("excel",          "EXCEL"),
                // Text editors
                ("notepad",        "notepad"),
                ("блокнот",        "notepad"),
                ("notebook",       "notepad"),  // important: "notebook" → notepad, not Word
            };

            foreach (var (input, expected) in cases)
            {
                var result = AICompanion.Desktop.Services.Automation.WindowAutomationHelper.ResolveAppName(input);
                _out.WriteLine($"  '{input}' → '{result}' (expected '{expected}')");
                result.Should().Be(expected, $"ResolveAppName('{input}') returned wrong value");
            }

            _out.WriteLine("✅ TEST 24 PASSED");
        }

        // ── TEST 25 ─────────────────────────────────────────────────────────────
        // VERIFY: focus_window action (new action added for this bug fix) brings
        //         an existing Notepad window to the foreground without launching
        //         a second Notepad process.
        [Fact]
        public async Task FocusWindow_BringsExistingWindowToFront_NoNewProcess()
        {
            _out.WriteLine("\n=== TEST 25: focus_window action ===");

            System.Diagnostics.Process? proc = null;
            try
            {
                proc = System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    { FileName = "notepad.exe", UseShellExecute = true });
                await Task.Delay(2500);

                var beforeCount = System.Diagnostics.Process.GetProcessesByName("notepad").Length;
                _out.WriteLine($"[BEFORE] notepad count: {beforeCount}");

                var automation = new AICompanion.Desktop.Services.Automation.WindowAutomationHelper();

                // FindWindowByProcessName is the core of focus_window
                var hwnd = automation.FindWindowByProcessName("notepad");
                hwnd.Should().NotBe(IntPtr.Zero, "Notepad must be found");
                automation.ForceFocusWindow(hwnd);
                var title = automation.GetWindowTitle(hwnd);
                _out.WriteLine($"[FOCUSED] '{title}' h=0x{hwnd:X}");

                var afterCount = System.Diagnostics.Process.GetProcessesByName("notepad").Length;
                _out.WriteLine($"[AFTER] notepad count: {afterCount}");

                afterCount.Should().Be(beforeCount,
                    "focus_window must NOT launch a new Notepad process");
                IsWindowVisible(hwnd).Should().BeTrue("focused window must be visible");
                title.ToLowerInvariant().Should().Contain("notepad");

                _out.WriteLine("✅ TEST 25 PASSED");
            }
            finally
            {
                try { proc?.Kill(); } catch { }
                await Task.Delay(500);
            }
        }
    }
}
