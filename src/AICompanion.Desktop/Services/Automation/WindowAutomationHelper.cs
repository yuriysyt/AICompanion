using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Extensions.Logging;

namespace AICompanion.Desktop.Services.Automation
{
    /// <summary>
    /// Reliable Win32-based window automation helper.
    /// Replaces broken SendKeys logic with robust SetForegroundWindow + thread attachment
    /// and proper keystroke simulation. This is the single source of truth for all
    /// window focus, text typing, and keyboard shortcut operations.
    /// </summary>
    public class WindowAutomationHelper
    {
        private readonly ILogger<WindowAutomationHelper>? _logger;

        // ====== Win32 P/Invoke ======
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

        [DllImport("user32.dll")]
        private static extern int GetWindowThreadProcessId(IntPtr hWnd, out int processId);

        [DllImport("user32.dll")]
        private static extern bool AttachThreadInput(int idAttach, int idAttachTo, bool fAttach);

        [DllImport("user32.dll")]
        private static extern bool IsWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool IsIconic(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool BringWindowToTop(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr SetFocus(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool AllowSetForegroundWindow(int dwProcessId);

        [DllImport("kernel32.dll")]
        private static extern int GetCurrentThreadId();

        // ====== Phase 9: Dialog + Mouse control P/Invoke ======
        [DllImport("user32.dll")]
        private static extern IntPtr FindWindowEx(IntPtr hWndParent, IntPtr hWndChildAfter, string? lpszClass, string? lpszWindow);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, string? lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, StringBuilder lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern bool SetCursorPos(int X, int Y);

        [DllImport("user32.dll")]
        private static extern void mouse_event(uint dwFlags, int dx, int dy, uint dwData, UIntPtr dwExtraInfo);

        [DllImport("user32.dll")]
        private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

        [DllImport("user32.dll")]
        private static extern IntPtr GetAncestor(IntPtr hWnd, uint gaFlags);

        private const uint GA_ROOT = 2;
        private const byte VK_CONTROL = 0x11;
        private const byte VK_V       = 0x56;
        private const uint KEYEVENTF_KEYUP = 0x0002;

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT { public int Left, Top, Right, Bottom; }

        // Win32 message constants
        private const uint WM_SETTEXT = 0x000C;
        private const uint WM_GETTEXT = 0x000D;
        private const uint WM_GETTEXTLENGTH = 0x000E;
        private const uint BM_CLICK = 0x00F5;
        private const uint WM_CLOSE = 0x0010;
        private const uint MOUSEEVENTF_LEFTDOWN  = 0x0002;
        private const uint MOUSEEVENTF_LEFTUP    = 0x0004;
        private const uint MOUSEEVENTF_RIGHTDOWN = 0x0008;
        private const uint MOUSEEVENTF_RIGHTUP   = 0x0010;
        private const uint MOUSEEVENTF_WHEEL     = 0x0800;

        private const int SW_RESTORE = 9;
        private const int SW_SHOW = 5;
        private const int SW_MINIMIZE = 6;

        public WindowAutomationHelper(ILogger<WindowAutomationHelper>? logger = null)
        {
            _logger = logger;
        }

        /// <summary>
        /// Gets the current foreground window handle via Win32.
        /// </summary>
        public IntPtr GetCurrentForegroundWindow()
        {
            return GetForegroundWindow();
        }

        /// <summary>
        /// Gets the window title for a given handle.
        /// </summary>
        public string GetWindowTitle(IntPtr hWnd)
        {
            if (hWnd == IntPtr.Zero) return "(none)";
            var sb = new StringBuilder(512);
            GetWindowText(hWnd, sb, 512);
            return sb.ToString();
        }

        /// <summary>
        /// Finds a window by process name (e.g., "notepad", "WINWORD").
        /// Returns the MainWindowHandle of the first matching process.
        /// </summary>
        public IntPtr FindWindowByProcessName(string processName)
        {
            try
            {
                var processes = Process.GetProcessesByName(processName);
                foreach (var p in processes)
                {
                    if (p.MainWindowHandle != IntPtr.Zero)
                    {
                        _logger?.LogDebug("[WIN] Found {Process} window: {Handle}", processName, p.MainWindowHandle);
                        return p.MainWindowHandle;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.LogWarning(ex, "[WIN] Error finding process: {Name}", processName);
            }
            return IntPtr.Zero;
        }

        /// <summary>
        /// Searches all visible top-level windows for one whose title contains
        /// <paramref name="fragment"/> (case-insensitive).
        /// Used as a fallback when the process name doesn't match (e.g. Opera browser
        /// whose process is "opera.exe" but the resolved name might vary).
        /// </summary>
        public IntPtr FindWindowByTitleContains(string fragment)
        {
            if (string.IsNullOrWhiteSpace(fragment)) return IntPtr.Zero;

            IntPtr found = IntPtr.Zero;
            var sb = new StringBuilder(256);
            EnumWindows((h, _) =>
            {
                if (!IsWindowVisible(h)) return true;
                sb.Clear();
                GetWindowText(h, sb, 256);
                var t = sb.ToString();
                if (t.Contains(fragment, StringComparison.OrdinalIgnoreCase))
                {
                    found = h;
                    return false;   // stop enumeration
                }
                return true;
            }, IntPtr.Zero);
            return found;
        }

        /// <summary>
        /// Forcefully brings a window to the foreground using Win32 thread attachment trick.
        /// This is the ONLY reliable way to steal focus on modern Windows.
        /// Returns true if focus was successfully acquired and verified.
        /// </summary>
        public bool ForceFocusWindow(IntPtr hWnd, int maxRetries = 3)
        {
            if (hWnd == IntPtr.Zero || !IsWindow(hWnd))
            {
                _logger?.LogWarning("[WIN] Invalid window handle");
                return false;
            }

            for (int attempt = 0; attempt < maxRetries; attempt++)
            {
                try
                {
                    // If window is minimized, restore it first
                    if (IsIconic(hWnd))
                    {
                        ShowWindow(hWnd, SW_RESTORE);
                        Thread.Sleep(200);
                    }

                    // Get thread IDs for attachment
                    int foregroundThreadId = GetWindowThreadProcessId(GetForegroundWindow(), out _);
                    int targetThreadId = GetWindowThreadProcessId(hWnd, out int targetProcessId);
                    int currentThreadId = GetCurrentThreadId();

                    // Allow the target process to set foreground
                    AllowSetForegroundWindow(targetProcessId);

                    // Attach our thread to foreground thread, set focus, then detach
                    bool attached = false;
                    if (foregroundThreadId != currentThreadId)
                    {
                        attached = AttachThreadInput(currentThreadId, foregroundThreadId, true);
                    }

                    try
                    {
                        BringWindowToTop(hWnd);
                        ShowWindow(hWnd, SW_SHOW);
                        SetForegroundWindow(hWnd);
                    }
                    finally
                    {
                        if (attached)
                        {
                            AttachThreadInput(currentThreadId, foregroundThreadId, false);
                        }
                    }

                    Thread.Sleep(100);

                    // Verify focus was acquired — also accept child window focus (WPS/Word
                    // foregrounds the document edit pane, not the main frame handle)
                    var fgAfter = GetForegroundWindow();
                    if (fgAfter == hWnd || GetAncestor(fgAfter, GA_ROOT) == hWnd)
                    {
                        _logger?.LogDebug("[WIN] Focus acquired on attempt {N}", attempt + 1);
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    _logger?.LogWarning(ex, "[WIN] Focus attempt {N} failed", attempt + 1);
                }

                Thread.Sleep(150);
            }

            _logger?.LogWarning("[WIN] Failed to acquire focus after {N} attempts", maxRetries);
            return false;
        }

        /// <summary>
        /// Verifies that the specified window currently has foreground focus.
        /// Also returns true when the foreground window is a child/descendant of expectedHwnd
        /// (e.g. Word's document edit pane is a child of the main frame).
        /// </summary>
        public bool VerifyFocus(IntPtr expectedHwnd)
        {
            var fg = GetForegroundWindow();
            if (fg == expectedHwnd) return true;
            // Accept a child window — GetAncestor(GA_ROOT) walks up to the root owner
            var root = GetAncestor(fg, GA_ROOT);
            return root == expectedHwnd;
        }

        /// <summary>
        /// Types text into the specified window via clipboard paste (Ctrl+V).
        /// Much more reliable than SendKeys — works with Russian keyboard layout,
        /// Unicode characters, long text, and special characters.
        /// </summary>
        public async Task<bool> TypeTextIntoWindowAsync(IntPtr hWnd, string text)
        {
            if (string.IsNullOrEmpty(text)) return false;

            if (!ForceFocusWindow(hWnd))
            {
                _logger?.LogWarning("[WIN] Cannot focus window for typing");
                return false;
            }

            await Task.Delay(200);

            if (!VerifyFocus(hWnd))
            {
                ForceFocusWindow(hWnd, 2);
                await Task.Delay(150);
                if (!VerifyFocus(hWnd))
                {
                    _logger?.LogWarning("[WIN] Lost focus — aborting type");
                    return false;
                }
            }

            try
            {
                // Set clipboard on STA thread (required for WinForms Clipboard API)
                bool clipboardSet = false;
                var setClip = new Thread(() =>
                {
                    try { Clipboard.SetText(text); clipboardSet = true; }
                    catch (Exception ex) { _logger?.LogWarning(ex, "[WIN] Clipboard.SetText failed"); }
                });
                setClip.SetApartmentState(ApartmentState.STA);
                setClip.Start();
                setClip.Join(2000);

                if (!clipboardSet)
                {
                    _logger?.LogWarning("[WIN] Clipboard unavailable — using SendKeys fallback");
                    SendKeys.SendWait(EscapeForSendKeys(text));
                    return true;
                }

                await Task.Delay(80);
                ForceFocusWindow(hWnd, 1);
                await Task.Delay(150);

                // Ctrl+V via keybd_event — layout-independent, works on Russian keyboards
                keybd_event(VK_CONTROL, 0, 0, UIntPtr.Zero);
                keybd_event(VK_V,       0, 0, UIntPtr.Zero);
                keybd_event(VK_V,       0, KEYEVENTF_KEYUP, UIntPtr.Zero);
                keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
                await Task.Delay(300);

                _logger?.LogInformation("[WIN] Clipboard-pasted {Len} chars into '{Title}'",
                    text.Length, GetWindowTitle(hWnd));
                return true;
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "[WIN] TypeText via clipboard failed");
                return false;
            }
        }

        /// <summary>
        /// Sends a keyboard shortcut (e.g., "^s" for Ctrl+S) to the specified window.
        /// </summary>
        public async Task<bool> SendShortcutToWindowAsync(IntPtr hWnd, string shortcutKeys)
        {
            if (!ForceFocusWindow(hWnd))
            {
                _logger?.LogWarning("[WIN] Cannot focus window for shortcut");
                return false;
            }

            await Task.Delay(80);

            if (!VerifyFocus(hWnd))
            {
                _logger?.LogWarning("[WIN] Lost focus before shortcut — retrying once");
                ForceFocusWindow(hWnd, 1);
                await Task.Delay(50);
                if (!VerifyFocus(hWnd))
                {
                    _logger?.LogWarning("[WIN] Lost focus after retry — aborting shortcut");
                    return false;
                }
            }

            try
            {
                SendKeys.SendWait(shortcutKeys);
                _logger?.LogInformation("[WIN] Sent shortcut '{Keys}' to {Title}",
                    shortcutKeys, GetWindowTitle(hWnd));
                return true;
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "[WIN] Shortcut failed: {Keys}", shortcutKeys);
                return false;
            }
        }

        /// <summary>
        /// Opens an application by process name or path. Returns the process or null.
        /// </summary>
        public Process? LaunchApplication(string appNameOrPath)
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = appNameOrPath,
                    UseShellExecute = true
                };
                var process = Process.Start(psi);
                _logger?.LogInformation("[WIN] Launched: {App}", appNameOrPath);
                return process;
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "[WIN] Failed to launch: {App}", appNameOrPath);
                return null;
            }
        }

        /// <summary>
        /// Launches an app and waits for its main window to appear.
        /// Returns the window handle or IntPtr.Zero.
        /// </summary>
        public async Task<IntPtr> LaunchAndWaitForWindowAsync(string processName, int timeoutMs = 5000)
        {
            // Try with and without .exe extension; also try WINWORD.EXE explicitly for WPS Office
            var process = LaunchApplication(processName)
                          ?? LaunchApplication(processName + ".EXE")
                          ?? (processName.Equals("WINWORD", StringComparison.OrdinalIgnoreCase)
                              ? LaunchApplication("WINWORD.EXE") : null);
            if (process == null)
            {
                _logger?.LogWarning("[WIN] Could not launch '{App}' — trying ShellExecute as URI", processName);
                try { Process.Start(new ProcessStartInfo { FileName = processName, UseShellExecute = true }); }
                catch { return IntPtr.Zero; }
                await Task.Delay(500).ConfigureAwait(false);
            }

            var sw = Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                await Task.Delay(150); // Phase 6: reduced from 300ms
                var hwnd = FindWindowByProcessName(processName);
                if (hwnd != IntPtr.Zero) return hwnd;

                // Also try refreshing the process
                try
                {
                    process.Refresh();
                    if (process.MainWindowHandle != IntPtr.Zero)
                        return process.MainWindowHandle;
                }
                catch { }
            }

            _logger?.LogWarning("[WIN] Timed out waiting for {App} window", processName);
            return IntPtr.Zero;
        }

        /// <summary>
        /// Escapes text for use with SendKeys.SendWait.
        /// SendKeys treats +, ^, %, ~, {, }, (, ) as special characters.
        /// </summary>
        public static string EscapeForSendKeys(string text)
        {
            var sb = new StringBuilder(text.Length + 16);
            foreach (var c in text)
            {
                switch (c)
                {
                    case '+': sb.Append("{+}"); break;
                    case '^': sb.Append("{^}"); break;
                    case '%': sb.Append("{%}"); break;
                    case '~': sb.Append("{~}"); break;
                    case '(': sb.Append("{(}"); break;
                    case ')': sb.Append("{)}"); break;
                    case '{': sb.Append("{{}"); break;
                    case '}': sb.Append("{}}"); break;
                    default: sb.Append(c); break;
                }
            }
            return sb.ToString();
        }

        /// <summary>
        /// Checks if a window handle is still valid (not closed/destroyed).
        /// </summary>
        public bool IsWindowValid(IntPtr hWnd)
        {
            return hWnd != IntPtr.Zero && IsWindow(hWnd);
        }

        /// <summary>
        /// Captures a screenshot of the primary monitor as PNG bytes.
        /// Used to send visual context to the Python AI backend.
        /// </summary>
        public byte[] CaptureDesktopScreenshot()
        {
            try
            {
                var bounds = System.Windows.Forms.Screen.PrimaryScreen?.Bounds
                    ?? new System.Drawing.Rectangle(0, 0, 1920, 1080);

                // Scale down to 960px wide to keep payload small
                int w = Math.Min(bounds.Width, 1920);
                int h = Math.Min(bounds.Height, 1080);

                using var bitmap = new System.Drawing.Bitmap(w, h);
                using var graphics = System.Drawing.Graphics.FromImage(bitmap);
                graphics.CopyFromScreen(bounds.Location, System.Drawing.Point.Empty,
                    new System.Drawing.Size(w, h));

                using var ms = new System.IO.MemoryStream();
                bitmap.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                var bytes = ms.ToArray();
                _logger?.LogInformation("[WIN] Screenshot captured: {W}x{H}, {Kb}KB",
                    w, h, bytes.Length / 1024);
                return bytes;
            }
            catch (Exception ex)
            {
                _logger?.LogWarning(ex, "[WIN] Screenshot capture failed");
                return Array.Empty<byte>();
            }
        }

        // =============================================================
        //  Phase 9: Dialog detection, button clicking, mouse control
        // =============================================================

        /// <summary>
        /// Finds a top-level dialog window whose title contains the given text.
        /// Scans ALL visible windows. Returns IntPtr.Zero if not found.
        /// </summary>
        public IntPtr FindDialogWindow(params string[] titleFragments)
        {
            IntPtr found = IntPtr.Zero;
            EnumWindows((hWnd, _) =>
            {
                if (!IsWindowVisible(hWnd)) return true;
                var title = GetWindowTitle(hWnd);
                if (string.IsNullOrEmpty(title)) return true;
                foreach (var frag in titleFragments)
                {
                    if (title.Contains(frag, StringComparison.OrdinalIgnoreCase))
                    {
                        found = hWnd;
                        _logger?.LogInformation("[WIN] Found dialog: '{Title}' (Handle: {H})", title, hWnd);
                        return false; // stop enumeration
                    }
                }
                return true;
            }, IntPtr.Zero);
            return found;
        }

        /// <summary>
        /// Gets the text of a child control via WM_GETTEXT (no focus needed).
        /// </summary>
        public string GetChildControlText(IntPtr hWnd)
        {
            int len = (int)SendMessage(hWnd, WM_GETTEXTLENGTH, IntPtr.Zero, IntPtr.Zero);
            if (len <= 0) return "";
            var sb = new StringBuilder(len + 2);
            SendMessage(hWnd, WM_GETTEXT, (IntPtr)(len + 1), sb);
            return sb.ToString();
        }

        /// <summary>
        /// Sets text in an edit control via WM_SETTEXT (no focus/SendKeys needed).
        /// This is the SAFE way to set filenames in Save As dialogs.
        /// </summary>
        public bool SetControlText(IntPtr hWnd, string text)
        {
            if (hWnd == IntPtr.Zero) return false;
            SendMessage(hWnd, WM_SETTEXT, IntPtr.Zero, text);
            _logger?.LogInformation("[WIN] SetControlText: '{Text}' → handle {H}", text, hWnd);
            return true;
        }

        /// <summary>
        /// Clicks a button inside a window by finding it by text and sending BM_CLICK.
        /// No mouse movement or focus needed — this is the RELIABLE way.
        /// Returns true if button was found and clicked.
        /// </summary>
        public bool ClickButtonInWindow(IntPtr parentHwnd, string buttonText)
        {
            if (parentHwnd == IntPtr.Zero) return false;

            bool clicked = false;
            EnumChildWindows(parentHwnd, (child, _) =>
            {
                var className = GetControlClassName(child);
                // Match Button class
                if (!className.Contains("Button", StringComparison.OrdinalIgnoreCase)) return true;

                var text = GetChildControlText(child);
                if (text.Contains(buttonText, StringComparison.OrdinalIgnoreCase))
                {
                    _logger?.LogInformation("[WIN] Clicking button '{Text}' (Handle: {H})", text, child);
                    SendMessage(child, BM_CLICK, IntPtr.Zero, IntPtr.Zero);
                    clicked = true;
                    return false; // stop
                }
                return true;
            }, IntPtr.Zero);

            if (!clicked)
                _logger?.LogWarning("[WIN] Button '{Text}' not found in window {H}", buttonText, parentHwnd);
            return clicked;
        }

        /// <summary>
        /// Finds the first Edit control (text field) inside a dialog window.
        /// This is used to find the filename field in Save As dialogs.
        /// </summary>
        public IntPtr FindEditControlInWindow(IntPtr parentHwnd)
        {
            IntPtr editHwnd = IntPtr.Zero;
            EnumChildWindows(parentHwnd, (child, _) =>
            {
                var className = GetControlClassName(child);
                if (className.Equals("Edit", StringComparison.OrdinalIgnoreCase) ||
                    className.Contains("RichEdit", StringComparison.OrdinalIgnoreCase))
                {
                    editHwnd = child;
                    return false; // found it
                }
                // Also check ComboBox children (Save As uses ComboBoxEx32 → ComboBox → Edit)
                if (className.Contains("ComboBox", StringComparison.OrdinalIgnoreCase))
                {
                    var innerEdit = FindWindowEx(child, IntPtr.Zero, "Edit", null);
                    if (innerEdit != IntPtr.Zero)
                    {
                        editHwnd = innerEdit;
                        return false;
                    }
                }
                return true;
            }, IntPtr.Zero);
            return editHwnd;
        }

        /// <summary>
        /// Gets the Win32 class name of a control (e.g., "Button", "Edit", "#32770").
        /// </summary>
        public string GetControlClassName(IntPtr hWnd)
        {
            var sb = new StringBuilder(256);
            GetClassName(hWnd, sb, 256);
            return sb.ToString();
        }

        /// <summary>
        /// Lists all visible buttons in a window (for debugging/smart dialog handling).
        /// Returns list of (handle, text) pairs.
        /// </summary>
        public List<(IntPtr Handle, string Text)> ListButtonsInWindow(IntPtr parentHwnd)
        {
            var buttons = new List<(IntPtr, string)>();
            if (parentHwnd == IntPtr.Zero) return buttons;

            EnumChildWindows(parentHwnd, (child, _) =>
            {
                var className = GetControlClassName(child);
                if (className.Contains("Button", StringComparison.OrdinalIgnoreCase))
                {
                    var text = GetChildControlText(child);
                    if (!string.IsNullOrWhiteSpace(text))
                        buttons.Add((child, text));
                }
                return true;
            }, IntPtr.Zero);
            return buttons;
        }

        /// <summary>
        /// Clicks at a specific screen position using mouse_event.
        /// </summary>
        public void ClickAtScreenPosition(int x, int y)
        {
            SetCursorPos(x, y);
            Thread.Sleep(50);
            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, UIntPtr.Zero);
            Thread.Sleep(30);
            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, UIntPtr.Zero);
            _logger?.LogInformation("[WIN] Mouse click at ({X}, {Y})", x, y);
        }

        /// <summary>Moves mouse cursor without clicking.</summary>
        public void MouseMove(int x, int y)
        {
            SetCursorPos(x, y);
            _logger?.LogInformation("[WIN] Mouse moved to ({X},{Y})", x, y);
        }

        /// <summary>Click-and-drag from current mouse position to (x,y).</summary>
        public void MouseDragTo(int x, int y)
        {
            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, UIntPtr.Zero);
            Thread.Sleep(60);
            SetCursorPos(x, y);
            Thread.Sleep(60);
            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, UIntPtr.Zero);
            _logger?.LogInformation("[WIN] Mouse drag to ({X},{Y})", x, y);
        }

        /// <summary>Scroll mouse wheel. Positive = up, negative = down.</summary>
        public void MouseScroll(int clicks)
        {
            // Windows wheel delta = 120 per click
            mouse_event(MOUSEEVENTF_WHEEL, 0, 0, (uint)(clicks * 120), UIntPtr.Zero);
            _logger?.LogInformation("[WIN] Mouse scroll {Clicks}", clicks);
        }

        /// <summary>Send raw SendKeys string (for hotkey executor).</summary>
        public void SendKeysRaw(string keys)
        {
            System.Windows.Forms.SendKeys.SendWait(keys);
        }

        /// <summary>
        /// Right-clicks at a specific screen position.
        /// </summary>
        public void RightClickAtScreenPosition(int x, int y)
        {
            SetCursorPos(x, y);
            Thread.Sleep(50);
            mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, UIntPtr.Zero);
            Thread.Sleep(30);
            mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, UIntPtr.Zero);
            _logger?.LogInformation("[WIN] Right-click at ({X}, {Y})", x, y);
        }

        /// <summary>
        /// Right-clicks the center of a window.
        /// </summary>
        public bool RightClickAtWindowCenter(IntPtr hWnd)
        {
            if (hWnd == IntPtr.Zero) return false;
            if (!GetWindowRect(hWnd, out RECT rect)) return false;
            int cx = (rect.Left + rect.Right) / 2;
            int cy = (rect.Top + rect.Bottom) / 2;
            RightClickAtScreenPosition(cx, cy);
            return true;
        }

        /// <summary>
        /// Double-clicks at a specific screen position.
        /// </summary>
        public void DoubleClickAtScreenPosition(int x, int y)
        {
            // x=0,y=0 means "at current cursor position" (caller already set it)
            if (x != 0 || y != 0) SetCursorPos(x, y);
            Thread.Sleep(50);
            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, UIntPtr.Zero);
            mouse_event(MOUSEEVENTF_LEFTUP,   0, 0, 0, UIntPtr.Zero);
            Thread.Sleep(50);
            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, UIntPtr.Zero);
            mouse_event(MOUSEEVENTF_LEFTUP,   0, 0, 0, UIntPtr.Zero);
            _logger?.LogInformation("[WIN] Double-click at ({X}, {Y})", x, y);
        }

        /// <summary>
        /// Uses Windows UIAutomation to find an element anywhere on the screen by its
        /// accessible Name property, then invokes (clicks) it.
        /// Works for UWP, WPF, Win32 accessible controls — including Word's Start Screen
        /// "Blank document" button which is an AccessibilityObject, not a classic BUTTON class.
        /// </summary>
        public bool FindAndClickByName(string elementName, IntPtr searchRoot = default)
        {
            try
            {
                var root = searchRoot != IntPtr.Zero
                    ? System.Windows.Automation.AutomationElement.FromHandle(searchRoot)
                    : System.Windows.Automation.AutomationElement.RootElement;

                var condition = new System.Windows.Automation.PropertyCondition(
                    System.Windows.Automation.AutomationElement.NameProperty,
                    elementName,
                    System.Windows.Automation.PropertyConditionFlags.IgnoreCase);

                var element = root.FindFirst(
                    System.Windows.Automation.TreeScope.Descendants, condition);

                if (element == null)
                {
                    _logger?.LogDebug("[WIN] FindAndClickByName: '{Name}' not found in accessibility tree", elementName);
                    return false;
                }

                // Try InvokePattern first (works for buttons, menu items, etc.)
                if (element.TryGetCurrentPattern(
                        System.Windows.Automation.InvokePattern.Pattern, out var inv))
                {
                    ((System.Windows.Automation.InvokePattern)inv).Invoke();
                    _logger?.LogInformation("[WIN] FindAndClickByName: InvokePattern on '{Name}'", elementName);
                    return true;
                }

                // Fallback: click the element's bounding rect centre
                var rect = element.Current.BoundingRectangle;
                if (!rect.IsEmpty)
                {
                    int cx = (int)(rect.Left + rect.Width  / 2);
                    int cy = (int)(rect.Top  + rect.Height / 2);
                    ClickAtScreenPosition(cx, cy);
                    _logger?.LogInformation("[WIN] FindAndClickByName: mouse-click on '{Name}' at ({X},{Y})", elementName, cx, cy);
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                _logger?.LogWarning(ex, "[WIN] FindAndClickByName failed for '{Name}'", elementName);
                return false;
            }
        }

        /// <summary>
        /// Clicks the center of a window/control using mouse_event.
        /// </summary>
        public bool ClickWindowCenter(IntPtr hWnd)
        {
            if (hWnd == IntPtr.Zero) return false;
            if (!GetWindowRect(hWnd, out RECT rect)) return false;
            int cx = (rect.Left + rect.Right) / 2;
            int cy = (rect.Top + rect.Bottom) / 2;
            ClickAtScreenPosition(cx, cy);
            return true;
        }

        /// <summary>
        /// Returns a list of all visible top-level window titles.
        /// Used for screen state monitoring.
        /// </summary>
        public List<string> GetAllVisibleWindowTitles()
        {
            var titles = new List<string>();
            EnumWindows((hWnd, _) =>
            {
                if (!IsWindowVisible(hWnd)) return true;
                var title = GetWindowTitle(hWnd);
                if (!string.IsNullOrWhiteSpace(title))
                    titles.Add(title);
                return true;
            }, IntPtr.Zero);
            return titles;
        }

        /// <summary>
        /// Returns all visible window titles as a single string for LLM context.
        /// </summary>
        public string GetScreenStateText()
        {
            var titles = GetAllVisibleWindowTitles();
            var fg = GetWindowTitle(GetCurrentForegroundWindow());
            var (dlg, dlgTitle) = DetectActiveDialog();

            var sb = new StringBuilder();
            sb.AppendLine($"Active foreground window: '{fg}'");
            if (dlg != IntPtr.Zero)
                sb.AppendLine($"Active dialog: '{dlgTitle}'");
            sb.AppendLine($"All visible windows: {string.Join(", ", titles.Take(10).Select(t => $"'{t}'"))}");
            return sb.ToString();
        }

        /// <summary>
        /// Detects if there's currently a dialog window open (Save As, Confirm, etc.)
        /// Returns the dialog handle and its title, or (Zero, "") if none found.
        /// </summary>
        public (IntPtr Handle, string Title) DetectActiveDialog()
        {
            string[] dialogKeywords = {
                "Save As", "Save this file", "Сохранить как", "Сохранение",
                "Save as Markdown", "Confirm Save As", "Replace",
                "already exists", "Do you want to save",
                "Confirm", "Warning", "Error",
                "Cannot save", "Cannot be saved",
                "нельзя сохранить", "не удается сохранить"
            };
            var dlg = FindDialogWindow(dialogKeywords);
            if (dlg != IntPtr.Zero)
            {
                var title = GetWindowTitle(dlg);
                return (dlg, title);
            }
            return (IntPtr.Zero, "");
        }

        /// <summary>
        /// HIGH-LEVEL: Sets filename in a Save As dialog and clicks Save.
        /// Uses WM_SETTEXT (no SendKeys!) to set the filename, then BM_CLICK to press Save.
        /// This is the SAFE, RELIABLE way to save files.
        /// </summary>
        public async Task<bool> HandleSaveAsDialog(IntPtr dialogHwnd, string filename)
        {
            if (dialogHwnd == IntPtr.Zero) return false;

            var dialogTitle = GetWindowTitle(dialogHwnd);
            _logger?.LogInformation("[WIN] HandleSaveAsDialog: dialog='{Title}', filename='{Name}'",
                dialogTitle, filename);

            // 1) Find the filename edit control
            var editHwnd = FindEditControlInWindow(dialogHwnd);
            if (editHwnd == IntPtr.Zero)
            {
                _logger?.LogWarning("[WIN] HandleSaveAsDialog: no edit control found — trying SendKeys fallback");
                ForceFocusWindow(dialogHwnd);
                await Task.Delay(200);
                await SendShortcutToWindowAsync(dialogHwnd, "^a");
                await Task.Delay(100);
                await TypeTextIntoWindowAsync(dialogHwnd, filename);
                await Task.Delay(200);
                await SendShortcutToWindowAsync(dialogHwnd, "{ENTER}");
                return true;
            }

            // 2) Set filename via WM_SETTEXT (SAFE — no typing into wrong window!)
            SetControlText(editHwnd, filename);
            await Task.Delay(200);

            // 3) Click the Save button
            bool saved = ClickButtonInWindow(dialogHwnd, "Save")
                      || ClickButtonInWindow(dialogHwnd, "Сохранить")
                      || ClickButtonInWindow(dialogHwnd, "&Save");
            if (!saved)
            {
                // Fallback: press Enter
                _logger?.LogWarning("[WIN] HandleSaveAsDialog: Save button not found — pressing Enter");
                ForceFocusWindow(dialogHwnd);
                await Task.Delay(100);
                await SendShortcutToWindowAsync(dialogHwnd, "{ENTER}");
            }

            await Task.Delay(500);

            // 4) Check for overwrite confirmation
            var (confirmDlg, confirmTitle) = DetectActiveDialog();
            if (confirmDlg != IntPtr.Zero && (
                confirmTitle.Contains("already exists", StringComparison.OrdinalIgnoreCase) ||
                confirmTitle.Contains("Confirm", StringComparison.OrdinalIgnoreCase) ||
                confirmTitle.Contains("Replace", StringComparison.OrdinalIgnoreCase)))
            {
                _logger?.LogInformation("[WIN] HandleSaveAsDialog: overwrite confirmation — clicking Yes");
                bool confirmed = ClickButtonInWindow(confirmDlg, "Yes")
                              || ClickButtonInWindow(confirmDlg, "Да")
                              || ClickButtonInWindow(confirmDlg, "&Yes");
                if (!confirmed)
                {
                    ForceFocusWindow(confirmDlg);
                    await Task.Delay(100);
                    await SendShortcutToWindowAsync(confirmDlg, "{ENTER}");
                }
                await Task.Delay(300);
            }

            // 5) Check for Markdown save dialog
            var (mdDlg, mdTitle) = DetectActiveDialog();
            if (mdDlg != IntPtr.Zero && mdTitle.Contains("Markdown", StringComparison.OrdinalIgnoreCase))
            {
                _logger?.LogInformation("[WIN] HandleSaveAsDialog: Markdown dialog — clicking 'Save as text file'");
                bool mdSaved = ClickButtonInWindow(mdDlg, "text file")
                            || ClickButtonInWindow(mdDlg, "plain text")
                            || ClickButtonInWindow(mdDlg, "Cancel"); // fallback
                if (!mdSaved)
                {
                    ForceFocusWindow(mdDlg);
                    await Task.Delay(100);
                    await SendShortcutToWindowAsync(mdDlg, "{ENTER}");
                }
                await Task.Delay(300);
            }

            // 6) Check for Error/Warning dialogs (e.g. "File cannot be saved", "Permission denied")
            var (errDlg, errTitle) = DetectActiveDialog();
            if (errDlg != IntPtr.Zero && (
                errTitle.Contains("Error", StringComparison.OrdinalIgnoreCase) ||
                errTitle.Contains("Warning", StringComparison.OrdinalIgnoreCase) ||
                errTitle.Contains("Cannot save", StringComparison.OrdinalIgnoreCase) ||
                errTitle.Contains("Cannot be saved", StringComparison.OrdinalIgnoreCase) ||
                errTitle.Contains("нельзя сохранить", StringComparison.OrdinalIgnoreCase) ||
                errTitle.Contains("не удается сохранить", StringComparison.OrdinalIgnoreCase)))
            {
                _logger?.LogError("[WIN] HandleSaveAsDialog: Error/Warning dialog detected: '{Title}' — dismissing", errTitle);
                // Try OK, then Cancel, then close via WM_CLOSE
                bool dismissed = ClickButtonInWindow(errDlg, "OK")
                              || ClickButtonInWindow(errDlg, "ОК")
                              || ClickButtonInWindow(errDlg, "Cancel")
                              || ClickButtonInWindow(errDlg, "Отмена")
                              || ClickButtonInWindow(errDlg, "Close");
                if (!dismissed)
                {
                    SendMessage(errDlg, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
                }
                await Task.Delay(200);
            }

            _logger?.LogInformation("[WIN] HandleSaveAsDialog: completed for '{Name}'", filename);
            return true;
        }

        /// <summary>
        /// Resolves friendly app names to Windows process names.
        /// Handles natural language variations like "the Word document", "my notepad", etc.
        /// </summary>
        public static string ResolveAppName(string input)
        {
            var map = new System.Collections.Generic.Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "notepad", "notepad" }, { "note pad", "notepad" }, { "notebook", "notepad" },
                { "text editor", "notepad" }, { "notepad++", "notepad++" }, { "note book", "notepad" },
                { "текстовый редактор", "notepad" }, { "текстовик", "notepad" },
                { "word", "WINWORD" }, { "microsoft word", "WINWORD" }, { "ms word", "WINWORD" },
                { "word document", "WINWORD" }, { "word app", "WINWORD" }, { "word processor", "WINWORD" },
                { "office word", "WINWORD" }, { "wps writer", "wps" },
                { "excel", "EXCEL" }, { "microsoft excel", "EXCEL" }, { "ms excel", "EXCEL" },
                { "spreadsheet", "EXCEL" }, { "excel sheet", "EXCEL" },
                { "chrome", "chrome" }, { "google chrome", "chrome" }, { "google", "chrome" },
                { "edge", "msedge" }, { "microsoft edge", "msedge" }, { "browser", "msedge" },
                { "internet explorer", "iexplore" },
                { "firefox", "firefox" }, { "mozilla", "firefox" },
                { "opera", "opera" }, { "opera browser", "opera" }, { "opera gx", "opera" },
                { "опера", "opera" }, { "оперу", "opera" }, { "оперой", "opera" },
                { "explorer", "explorer" }, { "file explorer", "explorer" }, { "files", "explorer" },
                { "windows explorer", "explorer" }, { "my computer", "explorer" },
                { "calculator", "calc" }, { "calc", "calc" }, { "калькулятор", "calc" },
                { "paint", "mspaint" }, { "ms paint", "mspaint" }, { "paintbrush", "mspaint" },
                { "terminal", "wt" }, { "windows terminal", "wt" }, { "cmd terminal", "wt" },
                { "command prompt", "cmd" }, { "cmd", "cmd" }, { "command line", "cmd" },
                { "powershell", "powershell" }, { "power shell", "powershell" },
                { "task manager", "taskmgr" }, { "taskmanager", "taskmgr" },
                { "settings", "ms-settings:" }, { "system settings", "ms-settings:" },
                { "control panel", "control" },
                { "snipping tool", "SnippingTool" }, { "snip", "SnippingTool" },
                // Russian
                { "блокнот", "notepad" }, { "ворд", "WINWORD" }, { "эксель", "EXCEL" },
                { "хром", "chrome" }, { "гугл хром", "chrome" }, { "эдж", "msedge" },
                { "проводник", "explorer" }, { "файловый менеджер", "explorer" },
                { "пейнт", "mspaint" }, { "терминал", "wt" },
                { "браузер", "msedge" }, { "интернет", "msedge" },
                { "командная строка", "cmd" }, { "командная строчка", "cmd" },
                { "диспетчер задач", "taskmgr" }, { "настройки", "ms-settings:" },
            };

            var cleaned = input.Trim();

            // Direct match
            if (map.TryGetValue(cleaned, out var directResult)) return directResult;

            // Strip common natural language words and retry
            var stripped = System.Text.RegularExpressions.Regex.Replace(
                cleaned.ToLowerInvariant(),
                @"\b(the|my|our|open|launch|start|a|an|this|that|программу|приложение|запусти|открой)\b",
                " "
            ).Trim();
            stripped = System.Text.RegularExpressions.Regex.Replace(stripped, @"\s+", " ").Trim();

            if (map.TryGetValue(stripped, out var strippedResult)) return strippedResult;

            // Try contains matching (longest match wins)
            string? bestMatch = null;
            int bestMatchLen = 0;
            var lo = cleaned.ToLowerInvariant();
            foreach (var key in map.Keys)
            {
                var klo = key.ToLowerInvariant();
                if (lo.Contains(klo) && klo.Length > bestMatchLen)
                {
                    bestMatch = map[key];
                    bestMatchLen = klo.Length;
                }
            }
            if (bestMatch != null) return bestMatch;

            return cleaned;
        }
    }
}
