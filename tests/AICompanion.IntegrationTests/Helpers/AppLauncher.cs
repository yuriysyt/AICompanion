using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Xunit.Abstractions;

namespace AICompanion.IntegrationTests.Helpers
{
    /// <summary>
    /// Launches real Windows applications for integration testing.
    /// Uses EnumWindows + process-ID matching to reliably find the window
    /// even when the test runner owns the foreground.
    /// </summary>
    public class AppLauncher : IDisposable
    {
        private readonly ITestOutputHelper _output;
        private Process? _process;

        public IntPtr WindowHandle { get; private set; }
        public string WindowTitle  { get; private set; } = "";

        // P/Invoke
        [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr h);
        [DllImport("user32.dll")] static extern bool IsWindow(IntPtr h);
        [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
        [DllImport("user32.dll")] static extern int  GetWindowText(IntPtr h, StringBuilder sb, int n);
        [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
        [DllImport("user32.dll")] static extern bool EnumWindows(EnumWindowsProc cb, IntPtr lp);
        [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        public AppLauncher(ITestOutputHelper output) => _output = output;

        /// <summary>
        /// Start <paramref name="exeName"/> and poll until a visible top-level window
        /// belonging to that process appears (or <paramref name="timeoutMs"/> expires).
        /// </summary>
        public async Task<bool> LaunchAsync(string exeName, int timeoutMs = 6000, string? arguments = null)
        {
            var launchDesc = arguments == null ? $"'{exeName}'" : $"'{exeName}' with args: {arguments}";
            _output.WriteLine($"[LAUNCH] Starting {launchDesc}...");
            var sw = Stopwatch.StartNew();

            try
            {
                _process = Process.Start(new ProcessStartInfo
                {
                    FileName        = exeName,
                    Arguments       = arguments ?? "",
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                _output.WriteLine($"[LAUNCH] ❌ Cannot start '{exeName}': {ex.Message}");
                return false;
            }

            if (_process == null)
            {
                _output.WriteLine("[LAUNCH] ❌ Process.Start returned null");
                return false;
            }

            // Collect all process IDs that might own the window (the direct process
            // PLUS any child processes it spawns, e.g. modern apps that re-launch themselves)
            var targetPids = new HashSet<uint> { (uint)_process.Id };

            var expectedTitle = GetExpectedTitlePart(exeName);
            var deadline      = DateTime.UtcNow.AddMilliseconds(timeoutMs);

            while (DateTime.UtcNow < deadline)
            {
                await Task.Delay(250);

                // Refresh child process list (e.g. calc.exe spawns ApplicationFrameHost)
                try
                {
                    _process.Refresh();
                    targetPids.Add((uint)_process.Id);
                }
                catch { }

                // Strategy 1: find by PID (works for classic Win32 apps)
                var found = FindWindowByPid(targetPids);
                if (found == IntPtr.Zero)
                {
                    // Strategy 2: title search (works for WinUI3 apps like Win11 Notepad/Calculator
                    // that re-launch in a different host process)
                    found = FindWindowByTitle(expectedTitle);
                }

                if (found != IntPtr.Zero)
                {
                    WindowHandle = found;
                    var sb2 = new StringBuilder(256);
                    GetWindowText(found, sb2, 256);
                    WindowTitle = sb2.ToString();
                    _output.WriteLine(
                        $"[LAUNCH] ✅ Window: '{WindowTitle}' hWnd=0x{found:X} ({sw.ElapsedMilliseconds}ms)");
                    return true;
                }
            }

            _output.WriteLine($"[LAUNCH] ❌ Window not found for '{exeName}' within {timeoutMs}ms");
            return false;
        }

        private static readonly string[] _ignoredTitles =
            { "Program Manager", "Default IME", "MSCTFIME UI", "GDI+ Window" };

        private IntPtr FindWindowByPid(HashSet<uint> targetPids)
        {
            IntPtr result = IntPtr.Zero;
            EnumWindows((hWnd, _) =>
            {
                if (!IsWindowVisible(hWnd)) return true;
                GetWindowThreadProcessId(hWnd, out var pid);
                if (!targetPids.Contains(pid)) return true;

                var sb = new StringBuilder(256);
                GetWindowText(hWnd, sb, 256);
                var title = sb.ToString();
                if (string.IsNullOrWhiteSpace(title)) return true;
                foreach (var ign in _ignoredTitles)
                    if (title.Equals(ign, StringComparison.OrdinalIgnoreCase)) return true;

                result = hWnd;
                return false;
            }, IntPtr.Zero);
            return result;
        }

        private IntPtr FindWindowByTitle(string partialTitle)
        {
            IntPtr result = IntPtr.Zero;
            EnumWindows((hWnd, _) =>
            {
                if (!IsWindowVisible(hWnd)) return true;
                var sb = new StringBuilder(256);
                GetWindowText(hWnd, sb, 256);
                var title = sb.ToString();
                if (title.Contains(partialTitle, StringComparison.OrdinalIgnoreCase) &&
                    !title.Contains("AI Companion", StringComparison.OrdinalIgnoreCase))
                {
                    result = hWnd;
                    return false;
                }
                return true;
            }, IntPtr.Zero);
            return result;
        }

        private static string GetExpectedTitlePart(string exeName)
        {
            // Support full paths like "C:\Windows\System32\notepad.exe" → "notepad"
            var baseName = Path.GetFileNameWithoutExtension(exeName).ToLowerInvariant();
            return baseName switch
            {
                "notepad" => "Notepad",
                "calc"    => "Calculator",
                "mspaint" => "Paint",
                "wordpad" => "WordPad",
                "winword" => "Word",
                "chrome"  => "Chrome",
                "msedge"  => "Edge",
                _         => baseName  // use extracted base name, not full path
            };
        }

        /// <summary>Bring the launched window to the foreground.</summary>
        public void Focus()
        {
            if (WindowHandle != IntPtr.Zero && IsWindow(WindowHandle))
            {
                SetForegroundWindow(WindowHandle);
                Thread.Sleep(250);
            }
        }

        public void Dispose()
        {
            try { _process?.CloseMainWindow(); } catch { }
            try { if (_process is { HasExited: false }) _process.Kill(entireProcessTree: true); }
            catch { }
            _process?.Dispose();
        }
    }
}
