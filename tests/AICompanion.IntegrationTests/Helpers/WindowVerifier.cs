using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Automation;
using Xunit.Abstractions;

namespace AICompanion.IntegrationTests.Helpers
{
    /// <summary>
    /// Reads the actual state of a live Windows application using UI Automation.
    ///
    /// Used to verify that a voice command really worked:
    ///   - Did text appear in Notepad?
    ///   - Is a specific window currently visible?
    ///   - How many windows matching a name are open?
    /// </summary>
    public class WindowVerifier
    {
        private readonly ITestOutputHelper _output;

        [DllImport("user32.dll")] static extern int GetWindowText(IntPtr h, StringBuilder sb, int n);
        [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();

        public WindowVerifier(ITestOutputHelper output) => _output = output;

        /// <summary>
        /// Read all text content visible in a window using UI Automation TextPattern.
        /// Returns null if the window does not expose a text pattern.
        /// </summary>
        public string? ReadWindowText(IntPtr hwnd)
        {
            try
            {
                if (hwnd == IntPtr.Zero) return null;
                var element = AutomationElement.FromHandle(hwnd);
                if (element == null) return null;

                // Strategy 1: TextPattern on the window itself (classic Notepad, Word)
                if (element.TryGetCurrentPattern(TextPattern.Pattern, out var tp))
                {
                    var text = ((TextPattern)tp).DocumentRange.GetText(-1);
                    _output.WriteLine($"[VERIFY] Read {text.Length} chars via TextPattern (root)");
                    return text;
                }

                // Strategy 2: Find any descendant that exposes TextPattern
                // (Win11 Notepad WinUI3 nests the edit control several levels deep)
                var subtreeCond = new OrCondition(
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Document),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit));
                var candidates = element.FindAll(TreeScope.Subtree, subtreeCond);
                foreach (AutomationElement candidate in candidates)
                {
                    if (candidate.TryGetCurrentPattern(TextPattern.Pattern, out var ctp))
                    {
                        var text = ((TextPattern)ctp).DocumentRange.GetText(-1);
                        _output.WriteLine($"[VERIFY] Read {text.Length} chars via TextPattern (descendant)");
                        return text;
                    }
                    if (candidate.TryGetCurrentPattern(ValuePattern.Pattern, out var cvp))
                    {
                        var text = ((ValuePattern)cvp).Current.Value;
                        _output.WriteLine($"[VERIFY] Read {text.Length} chars via ValuePattern (descendant)");
                        return text;
                    }
                }

                _output.WriteLine("[VERIFY] ⚠️ No text pattern accessible — returning empty string");
                return string.Empty; // return "" not null so tests can check Contains
            }
            catch (Exception ex)
            {
                _output.WriteLine($"[VERIFY] ❌ ReadWindowText error: {ex.Message}");
                return string.Empty;
            }
        }

        /// <summary>Returns true if any top-level window's title contains <paramref name="partialTitle"/>.</summary>
        public bool IsWindowOpen(string partialTitle)
        {
            try
            {
                var cond = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window);
                var windows = AutomationElement.RootElement.FindAll(TreeScope.Children, cond);
                foreach (AutomationElement w in windows)
                {
                    var name = w.Current.Name;
                    if (name.Contains(partialTitle, StringComparison.OrdinalIgnoreCase))
                    {
                        _output.WriteLine($"[VERIFY] ✅ Window found: '{name}'");
                        return true;
                    }
                }
                _output.WriteLine($"[VERIFY] ❌ No window matching '{partialTitle}'");
                return false;
            }
            catch (Exception ex)
            {
                _output.WriteLine($"[VERIFY] IsWindowOpen error: {ex.Message}");
                return false;
            }
        }

        /// <summary>Returns the title of the current foreground window.</summary>
        public string GetForegroundTitle()
        {
            var hwnd = GetForegroundWindow();
            var sb = new StringBuilder(256);
            GetWindowText(hwnd, sb, 256);
            return sb.ToString();
        }

        /// <summary>Count how many top-level windows match <paramref name="partialTitle"/>.</summary>
        public int CountOpenWindows(string partialTitle)
        {
            int count = 0;
            try
            {
                var cond = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window);
                var windows = AutomationElement.RootElement.FindAll(TreeScope.Children, cond);
                foreach (AutomationElement w in windows)
                    if (w.Current.Name.Contains(partialTitle, StringComparison.OrdinalIgnoreCase))
                        count++;
            }
            catch { }
            return count;
        }

        /// <summary>
        /// Searches all visible windows for one titled with <paramref name="partialTitle"/>
        /// and returns its handle.
        /// </summary>
        public IntPtr FindWindowHandle(string partialTitle)
        {
            try
            {
                var cond = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window);
                var windows = AutomationElement.RootElement.FindAll(TreeScope.Children, cond);
                foreach (AutomationElement w in windows)
                {
                    if (w.Current.Name.Contains(partialTitle, StringComparison.OrdinalIgnoreCase))
                        return (IntPtr)w.Current.NativeWindowHandle;
                }
            }
            catch { }
            return IntPtr.Zero;
        }
    }
}
