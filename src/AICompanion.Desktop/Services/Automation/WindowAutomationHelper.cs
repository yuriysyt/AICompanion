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

                    // Verify focus was acquired
                    if (GetForegroundWindow() == hWnd)
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
        /// </summary>
        public bool VerifyFocus(IntPtr expectedHwnd)
        {
            return GetForegroundWindow() == expectedHwnd;
        }

        /// <summary>
        /// Types text into the specified window using SendKeys with proper focus management.
        /// Returns true on success.
        /// </summary>
        public async Task<bool> TypeTextIntoWindowAsync(IntPtr hWnd, string text)
        {
            if (string.IsNullOrEmpty(text)) return false;

            // Focus the window
            if (!ForceFocusWindow(hWnd))
            {
                _logger?.LogWarning("[WIN] Cannot focus window for typing");
                return false;
            }

            // Small delay to let window activate (Phase 6: reduced from 200ms)
            await Task.Delay(100);

            // Verify focus is still ours
            if (!VerifyFocus(hWnd))
            {
                _logger?.LogWarning("[WIN] Lost focus before typing — retrying once");
                // Phase 6: one retry instead of immediate fail
                ForceFocusWindow(hWnd, 1);
                await Task.Delay(50);
                if (!VerifyFocus(hWnd))
                {
                    _logger?.LogWarning("[WIN] Lost focus after retry — aborting type");
                    return false;
                }
            }

            try
            {
                // Escape special SendKeys characters
                var escaped = EscapeForSendKeys(text);
                SendKeys.SendWait(escaped);
                _logger?.LogInformation("[WIN] Typed {Len} chars into {Title}",
                    text.Length, GetWindowTitle(hWnd));
                return true;
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "[WIN] SendKeys failed");
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
            var process = LaunchApplication(processName);
            if (process == null) return IntPtr.Zero;

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

        /// <summary>
        /// Resolves friendly app names to Windows process names.
        /// </summary>
        public static string ResolveAppName(string input)
        {
            var map = new System.Collections.Generic.Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "notepad", "notepad" }, { "note pad", "notepad" }, { "notebook", "notepad" },
                { "text editor", "notepad" }, { "notepad++", "notepad++" }, { "note book", "notepad" },
                { "word", "WINWORD" }, { "microsoft word", "WINWORD" }, { "ms word", "WINWORD" },
                { "excel", "EXCEL" }, { "microsoft excel", "EXCEL" }, { "ms excel", "EXCEL" },
                { "chrome", "chrome" }, { "google chrome", "chrome" },
                { "edge", "msedge" }, { "microsoft edge", "msedge" }, { "browser", "msedge" },
                { "firefox", "firefox" },
                { "explorer", "explorer" }, { "file explorer", "explorer" },
                { "calculator", "calc" }, { "calc", "calc" },
                { "paint", "mspaint" }, { "ms paint", "mspaint" },
                { "terminal", "wt" }, { "windows terminal", "wt" },
                { "command prompt", "cmd" }, { "cmd", "cmd" },
                { "powershell", "powershell" },
                // Russian
                { "блокнот", "notepad" }, { "ворд", "WINWORD" }, { "эксель", "EXCEL" },
                { "калькулятор", "calc" }, { "браузер", "msedge" }, { "хром", "chrome" },
                { "проводник", "explorer" }, { "пейнт", "mspaint" }, { "терминал", "wt" },
            };

            var cleaned = input.Trim();
            return map.TryGetValue(cleaned, out var resolved) ? resolved : cleaned;
        }
    }
}
