using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Extensions.Logging;
using AICompanion.Desktop.Services.Voice;

namespace AICompanion.Desktop.Services.Dictation
{
    /// <summary>
    /// Dictation service for typing text into Word, Notepad, or any focused application.
    /// Uses Windows API to send text to the active window without requiring focus on our app.
    /// </summary>
    public class DictationService : IDisposable
    {
        private readonly ILogger<DictationService>? _logger;
        private readonly UnifiedVoiceManager _voiceManager;

        private bool _isDictating;
        private IntPtr _targetWindow;
        private string _targetAppName = "";
        private DictationMode _mode = DictationMode.Disabled;

        public event EventHandler<string>? TextDictated;
        public event EventHandler<DictationMode>? ModeChanged;
        public event EventHandler<string>? DictationError;

        public bool IsDictating => _isDictating;
        public DictationMode CurrentMode => _mode;
        public string TargetApplication => _targetAppName;

        // Windows API imports
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll")]
        private static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, string lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        private const uint WM_CHAR = 0x0102;
        private const uint WM_SETTEXT = 0x000C;

        // Track last dictated words for deletion
        private readonly System.Collections.Generic.List<string> _dictatedWords = new();
        private const int MaxWordHistory = 50;

        // Captured target window (stored before AI Companion takes focus)
        private IntPtr _capturedTargetWindow = IntPtr.Zero;
        private string _capturedTargetTitle = "";

        public DictationService(ILogger<DictationService>? logger, UnifiedVoiceManager voiceManager)
        {
            _logger = logger;
            _voiceManager = voiceManager;

            // Subscribe to voice recognition
            _voiceManager.TextRecognized += OnTextRecognized;
        }

        /// <summary>
        /// Capture the current foreground window as the dictation target.
        /// Call this BEFORE AI Companion takes focus (e.g. when user presses mic button).
        /// </summary>
        public void CaptureTargetWindow()
        {
            var hwnd = GetForegroundWindow();
            if (hwnd == IntPtr.Zero) return;

            var title = GetWindowTitle(hwnd);
            // Don't capture AI Companion itself
            if (!string.IsNullOrEmpty(title) &&
                !title.Contains("AI Companion", StringComparison.OrdinalIgnoreCase))
            {
                _capturedTargetWindow = hwnd;
                _capturedTargetTitle = title;
                _logger?.LogInformation("[Dictation] Captured target window: '{Title}'", title);
            }
        }

        /// <summary>
        /// Get the best available target window, avoiding AI Companion's own window.
        /// Uses captured window first, then searches for a suitable window.
        /// </summary>
        private IntPtr GetBestTargetWindow()
        {
            // First try: use previously captured window
            if (_capturedTargetWindow != IntPtr.Zero)
            {
                var title = GetWindowTitle(_capturedTargetWindow);
                if (!string.IsNullOrEmpty(title) && title != "(none)")
                {
                    return _capturedTargetWindow;
                }
            }

            // Second try: current foreground if it's not AI Companion
            var foreground = GetForegroundWindow();
            if (foreground != IntPtr.Zero)
            {
                var fgTitle = GetWindowTitle(foreground);
                if (!string.IsNullOrEmpty(fgTitle) &&
                    !fgTitle.Contains("AI Companion", StringComparison.OrdinalIgnoreCase))
                {
                    return foreground;
                }
            }

            // Third try: find the most recent text editor window
            return FindRecentTextEditorWindow();
        }

        /// <summary>
        /// Find a recently used text editor window (Word, Notepad, etc.)
        /// </summary>
        private IntPtr FindRecentTextEditorWindow()
        {
            IntPtr found = IntPtr.Zero;
            EnumWindows((hwnd, _) =>
            {
                if (!IsWindowVisible(hwnd)) return true;

                var title = GetWindowTitle(hwnd);
                if (string.IsNullOrEmpty(title)) return true;
                if (title.Contains("AI Companion", StringComparison.OrdinalIgnoreCase)) return true;

                var lower = title.ToLowerInvariant();
                if (lower.Contains("word") || lower.Contains(".docx") || lower.Contains(".doc") ||
                    lower.Contains("notepad") || lower.Contains("блокнот") || lower.Contains(".txt") ||
                    lower.Contains("notepad++") || lower.Contains("code"))
                {
                    found = hwnd;
                    return false; // Stop enumeration
                }
                return true;
            }, IntPtr.Zero);

            return found;
        }

        private string GetWindowTitle(IntPtr hwnd)
        {
            var sb = new System.Text.StringBuilder(256);
            GetWindowText(hwnd, sb, 256);
            return sb.ToString();
        }

        /// <summary>
        /// Start dictation mode - text will be sent to the target application.
        /// </summary>
        public async Task<bool> StartDictationAsync(DictationMode mode = DictationMode.Auto)
        {
            if (_isDictating)
            {
                _logger?.LogWarning("[Dictation] Already dictating");
                return true;
            }

            _mode = mode;

            // Get target window based on mode
            switch (mode)
            {
                case DictationMode.Auto:
                case DictationMode.Custom:
                    _targetWindow = GetBestTargetWindow();
                    break;
                case DictationMode.Word:
                    _targetWindow = await FindWindowByProcessAsync("WINWORD");
                    if (_targetWindow == IntPtr.Zero)
                        _targetWindow = GetBestTargetWindow(); // Fallback
                    break;
                case DictationMode.Notepad:
                    _targetWindow = await FindWindowByProcessAsync("notepad");
                    if (_targetWindow == IntPtr.Zero)
                        _targetWindow = GetBestTargetWindow(); // Fallback
                    break;
                default:
                    _targetWindow = GetBestTargetWindow();
                    break;
            }

            if (_targetWindow == IntPtr.Zero)
            {
                var errorMsg = $"Could not find target application for dictation mode: {mode}";
                _logger?.LogError("[Dictation] {Error}", errorMsg);
                DictationError?.Invoke(this, errorMsg);
                return false;
            }

            // Get window title
            _targetAppName = GetWindowTitle(_targetWindow);

            _isDictating = true;
            ModeChanged?.Invoke(this, mode);
            _logger?.LogInformation("[Dictation] Started dictation to: {App} (Mode: {Mode})", _targetAppName, mode);

            return true;
        }

        /// <summary>
        /// Stop dictation mode.
        /// </summary>
        public void StopDictation()
        {
            if (!_isDictating) return;

            _isDictating = false;
            _targetWindow = IntPtr.Zero;
            _targetAppName = "";
            _mode = DictationMode.Disabled;
            ModeChanged?.Invoke(this, DictationMode.Disabled);
            _logger?.LogInformation("[Dictation] Stopped");
        }

        /// <summary>
        /// Toggle dictation mode on/off.
        /// </summary>
        public async Task<bool> ToggleDictationAsync(DictationMode mode = DictationMode.Auto)
        {
            if (_isDictating)
            {
                StopDictation();
                return false;
            }
            else
            {
                return await StartDictationAsync(mode);
            }
        }

        /// <summary>
        /// Manually send text to the target application.
        /// </summary>
        public async Task SendTextAsync(string text)
        {
            if (string.IsNullOrEmpty(text)) return;

            if (!_isDictating || _targetWindow == IntPtr.Zero)
            {
                _logger?.LogWarning("[Dictation] Not in dictation mode, cannot send text");
                return;
            }

            try
            {
                // Bring target window to front briefly
                SetForegroundWindow(_targetWindow);
                await Task.Delay(50);

                // Track words for deletion feature
                var words = text.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                foreach (var word in words)
                {
                    _dictatedWords.Add(word);
                    if (_dictatedWords.Count > MaxWordHistory)
                    {
                        _dictatedWords.RemoveAt(0);
                    }
                }

                // Send text using SendKeys (most reliable cross-application method)
                SendKeys.SendWait(EscapeForSendKeys(text));

                TextDictated?.Invoke(this, text);
                _logger?.LogInformation("[Dictation] Sent: {Text}", text.Length > 50 ? text[..50] + "..." : text);
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "[Dictation] Failed to send text");
                DictationError?.Invoke(this, $"Failed to send text: {ex.Message}");
            }
        }

        /// <summary>
        /// Delete the last word that was dictated.
        /// </summary>
        public async Task DeleteLastWordAsync()
        {
            await DeleteWordsAsync(1);
        }

        /// <summary>
        /// Delete the last N words that were dictated.
        /// </summary>
        public async Task DeleteWordsAsync(int count)
        {
            if (!_isDictating || _targetWindow == IntPtr.Zero)
            {
                _logger?.LogWarning("[Dictation] Not in dictation mode");
                return;
            }

            if (count <= 0) return;

            try
            {
                SetForegroundWindow(_targetWindow);
                await Task.Delay(50);

                // Calculate how many characters to delete
                int charsToDelete = 0;
                int wordsDeleted = 0;

                for (int i = _dictatedWords.Count - 1; i >= 0 && wordsDeleted < count; i--)
                {
                    charsToDelete += _dictatedWords[i].Length + 1; // +1 for space
                    _dictatedWords.RemoveAt(i);
                    wordsDeleted++;
                }

                // Send backspace keys
                for (int i = 0; i < charsToDelete; i++)
                {
                    SendKeys.SendWait("{BACKSPACE}");
                }

                _logger?.LogInformation("[Dictation] Deleted {Count} words ({Chars} chars)", wordsDeleted, charsToDelete);
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "[Dictation] Failed to delete words");
                DictationError?.Invoke(this, $"Failed to delete: {ex.Message}");
            }
        }

        /// <summary>
        /// Apply formatting to selected text (Bold, Italic, Underline).
        /// </summary>
        public async Task FormatTextAsync(TextFormat format)
        {
            if (!_isDictating || _targetWindow == IntPtr.Zero)
            {
                _logger?.LogWarning("[Dictation] Not in dictation mode");
                return;
            }

            try
            {
                SetForegroundWindow(_targetWindow);
                await Task.Delay(50);

                var keys = format switch
                {
                    TextFormat.Bold => "^b",
                    TextFormat.Italic => "^i",
                    TextFormat.Underline => "^u",
                    TextFormat.Strikethrough => "^-", // May not work in all apps
                    TextFormat.Subscript => "^=",     // Word-specific
                    TextFormat.Superscript => "^+",   // Word-specific
                    _ => ""
                };

                if (!string.IsNullOrEmpty(keys))
                {
                    SendKeys.SendWait(keys);
                    _logger?.LogInformation("[Dictation] Applied format: {Format}", format);
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "[Dictation] Failed to format text");
            }
        }

        /// <summary>
        /// Insert special punctuation or symbols.
        /// </summary>
        public async Task InsertPunctuationAsync(string punctuation)
        {
            if (!_isDictating || _targetWindow == IntPtr.Zero) return;

            try
            {
                SetForegroundWindow(_targetWindow);
                await Task.Delay(30);

                SendKeys.SendWait(EscapeForSendKeys(punctuation));
                _logger?.LogDebug("[Dictation] Inserted punctuation: {Punct}", punctuation);
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "[Dictation] Failed to insert punctuation");
            }
        }

        /// <summary>
        /// Auto-detect if Word or Notepad is the target window.
        /// Uses captured target or best available window (not AI Companion).
        /// </summary>
        public DictationMode DetectTargetApplication()
        {
            try
            {
                var targetHwnd = GetBestTargetWindow();
                if (targetHwnd == IntPtr.Zero) return DictationMode.Disabled;

                var title = GetWindowTitle(targetHwnd).ToLowerInvariant();

                // Detect Word
                if (title.Contains("word") || title.Contains(".docx") || title.Contains(".doc"))
                {
                    return DictationMode.Word;
                }

                // Detect Notepad
                if (title.Contains("notepad") || title.Contains("блокнот") || title.Contains(".txt"))
                {
                    return DictationMode.Notepad;
                }

                // Detect other text editors
                if (title.Contains("visual studio") || title.Contains("code") ||
                    title.Contains("sublime") || title.Contains("atom"))
                {
                    return DictationMode.Custom;
                }

                return DictationMode.Auto;
            }
            catch
            {
                return DictationMode.Disabled;
            }
        }

        /// <summary>
        /// Check if current foreground window is a text editor.
        /// </summary>
        public bool IsTextEditorFocused()
        {
            var mode = DetectTargetApplication();
            return mode == DictationMode.Word || mode == DictationMode.Notepad || mode == DictationMode.Custom;
        }

        /// <summary>
        /// Send special key commands (Enter, Tab, etc.)
        /// </summary>
        public async Task SendCommandAsync(DictationCommand command)
        {
            if (!_isDictating || _targetWindow == IntPtr.Zero)
            {
                _logger?.LogWarning("[Dictation] Not in dictation mode");
                return;
            }

            // Handle delete word commands specially
            if (command == DictationCommand.DeleteWord)
            {
                await DeleteLastWordAsync();
                return;
            }
            if (command == DictationCommand.DeleteWords)
            {
                await DeleteWordsAsync(3); // Default to 3 words
                return;
            }

            try
            {
                SetForegroundWindow(_targetWindow);
                await Task.Delay(50);

                var keys = command switch
                {
                    DictationCommand.Enter => "{ENTER}",
                    DictationCommand.Tab => "{TAB}",
                    DictationCommand.Backspace => "{BACKSPACE}",
                    DictationCommand.Delete => "{DELETE}",
                    DictationCommand.SelectAll => "^a",
                    DictationCommand.Copy => "^c",
                    DictationCommand.Paste => "^v",
                    DictationCommand.Cut => "^x",
                    DictationCommand.Undo => "^z",
                    DictationCommand.Redo => "^y",
                    DictationCommand.Save => "^s",
                    DictationCommand.Bold => "^b",
                    DictationCommand.Italic => "^i",
                    DictationCommand.Underline => "^u",
                    DictationCommand.NewLine => "{ENTER}",
                    DictationCommand.NewParagraph => "{ENTER}{ENTER}",
                    DictationCommand.Period => ".",
                    DictationCommand.Comma => ",",
                    DictationCommand.Question => "?",
                    DictationCommand.Exclamation => "!",
                    DictationCommand.Colon => ":",
                    DictationCommand.Semicolon => ";",
                    DictationCommand.OpenParenthesis => "{(}",
                    DictationCommand.CloseParenthesis => "{)}",
                    DictationCommand.Quote => "\"",
                    DictationCommand.Dash => "-",
                    DictationCommand.Ellipsis => "...",
                    _ => ""
                };

                if (!string.IsNullOrEmpty(keys))
                {
                    SendKeys.SendWait(keys);
                    _logger?.LogDebug("[Dictation] Sent command: {Command}", command);
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "[Dictation] Failed to send command {Command}", command);
            }
        }

        /// <summary>
        /// Parse voice command for dictation control.
        /// Returns true if command was handled as dictation control.
        /// </summary>
        public async Task<(bool handled, string? message)> ProcessVoiceCommandAsync(string command)
        {
            var lowerCommand = command.ToLowerInvariant().Trim();

            // Start dictation commands
            if (lowerCommand.Contains("start dictation") || 
                lowerCommand.Contains("начать диктовку") ||
                lowerCommand.Contains("dictate") ||
                lowerCommand.Contains("диктуй"))
            {
                DictationMode mode = DictationMode.Auto;
                
                if (lowerCommand.Contains("word") || lowerCommand.Contains("ворд"))
                    mode = DictationMode.Word;
                else if (lowerCommand.Contains("notepad") || lowerCommand.Contains("блокнот"))
                    mode = DictationMode.Notepad;

                var success = await StartDictationAsync(mode);
                return (true, success 
                    ? $"Dictation started to {_targetAppName}. Speak and I'll type for you."
                    : "Could not start dictation. Please open the target application first.");
            }

            // Stop dictation commands
            if (lowerCommand.Contains("stop dictation") || 
                lowerCommand.Contains("остановить диктовку") ||
                lowerCommand.Contains("end dictation") ||
                lowerCommand.Contains("прекратить"))
            {
                StopDictation();
                return (true, "Dictation stopped.");
            }

            // If dictating, process as text or command
            if (_isDictating)
            {
                // Check for dictation commands
                var dictCmd = ParseDictationCommand(lowerCommand);
                if (dictCmd != null)
                {
                    await SendCommandAsync(dictCmd.Value);
                    return (true, null);
                }

                // Otherwise, send as text
                await SendTextAsync(command + " ");
                return (true, null);
            }

            return (false, null);
        }

        private DictationCommand? ParseDictationCommand(string command)
        {
            return command switch
            {
                // Delete commands
                var c when c.Contains("delete word") || c.Contains("удали слово") || c.Contains("delete last word") => DictationCommand.DeleteWord,
                var c when c.Contains("delete words") || c.Contains("удали слова") => DictationCommand.DeleteWords,

                // Navigation/editing
                var c when c.Contains("new line") || c.Contains("новая строка") => DictationCommand.NewLine,
                var c when c.Contains("new paragraph") || c.Contains("новый абзац") => DictationCommand.NewParagraph,
                var c when c.Contains("enter") || c.Contains("ввод") => DictationCommand.Enter,
                var c when c.Contains("tab") || c.Contains("табуляция") => DictationCommand.Tab,
                var c when c.Contains("backspace") || c.Contains("удалить назад") => DictationCommand.Backspace,
                var c when c.Contains("delete") && !c.Contains("file") && !c.Contains("word") => DictationCommand.Delete,
                var c when c.Contains("select all") || c.Contains("выделить всё") => DictationCommand.SelectAll,
                var c when c.Contains("copy") || c.Contains("копировать") => DictationCommand.Copy,
                var c when c.Contains("paste") || c.Contains("вставить") => DictationCommand.Paste,
                var c when c.Contains("cut") || c.Contains("вырезать") => DictationCommand.Cut,
                var c when c.Contains("undo") || c.Contains("отменить") => DictationCommand.Undo,
                var c when c.Contains("redo") || c.Contains("повторить") => DictationCommand.Redo,
                var c when c.Contains("save") || c.Contains("сохранить") => DictationCommand.Save,

                // Formatting
                var c when c.Contains("bold") || c.Contains("жирный") => DictationCommand.Bold,
                var c when c.Contains("italic") || c.Contains("курсив") => DictationCommand.Italic,
                var c when c.Contains("underline") || c.Contains("подчеркнуть") => DictationCommand.Underline,

                // Punctuation
                var c when c == "period" || c == "точка" || c == "." => DictationCommand.Period,
                var c when c == "comma" || c == "запятая" || c == "," => DictationCommand.Comma,
                var c when c == "question mark" || c == "question" || c == "вопрос" || c == "?" => DictationCommand.Question,
                var c when c == "exclamation" || c == "exclamation mark" || c == "восклицательный" || c == "!" => DictationCommand.Exclamation,
                var c when c == "colon" || c == "двоеточие" => DictationCommand.Colon,
                var c when c == "semicolon" || c == "точка с запятой" => DictationCommand.Semicolon,
                var c when c.Contains("open paren") || c.Contains("открыть скобку") || c == "(" => DictationCommand.OpenParenthesis,
                var c when c.Contains("close paren") || c.Contains("закрыть скобку") || c == ")" => DictationCommand.CloseParenthesis,
                var c when c == "quote" || c == "кавычка" || c == "\"" => DictationCommand.Quote,
                var c when c == "dash" || c == "тире" || c == "-" => DictationCommand.Dash,
                var c when c == "ellipsis" || c == "многоточие" || c == "..." => DictationCommand.Ellipsis,

                _ => null
            };
        }

        private void OnTextRecognized(object? sender, string text)
        {
            if (_isDictating && !string.IsNullOrWhiteSpace(text))
            {
                // Fire and forget - don't await in event handler
                _ = SendTextAsync(text + " ");
            }
        }

        private async Task<IntPtr> FindWindowByProcessAsync(string processName)
        {
            return await Task.Run(() =>
            {
                var processes = Process.GetProcessesByName(processName);
                foreach (var process in processes)
                {
                    if (process.MainWindowHandle != IntPtr.Zero)
                    {
                        return process.MainWindowHandle;
                    }
                }
                return IntPtr.Zero;
            });
        }

        private static string EscapeForSendKeys(string text)
        {
            // Escape special SendKeys characters
            var escaped = new System.Text.StringBuilder();
            foreach (char c in text)
            {
                switch (c)
                {
                    case '+':
                    case '^':
                    case '%':
                    case '~':
                    case '(':
                    case ')':
                    case '{':
                    case '}':
                    case '[':
                    case ']':
                        escaped.Append('{').Append(c).Append('}');
                        break;
                    default:
                        escaped.Append(c);
                        break;
                }
            }
            return escaped.ToString();
        }

        public void Dispose()
        {
            StopDictation();
            _voiceManager.TextRecognized -= OnTextRecognized;
            _logger?.LogInformation("[Dictation] Disposed");
        }
    }

    public enum DictationMode
    {
        Disabled = 0,
        Auto = 1,      // Dictate to current foreground window
        Word = 2,      // Dictate to Microsoft Word
        Notepad = 3,   // Dictate to Notepad
        Custom = 4     // Dictate to specified window
    }

    public enum DictationCommand
    {
        Enter,
        Tab,
        Backspace,
        Delete,
        SelectAll,
        Copy,
        Paste,
        Cut,
        Undo,
        Redo,
        Save,
        Bold,
        Italic,
        Underline,
        NewLine,
        NewParagraph,
        Period,
        Comma,
        Question,
        Exclamation,
        DeleteWord,
        DeleteWords,
        Colon,
        Semicolon,
        OpenParenthesis,
        CloseParenthesis,
        Quote,
        Dash,
        Ellipsis
    }

    public enum TextFormat
    {
        Bold,
        Italic,
        Underline,
        Strikethrough,
        Subscript,
        Superscript
    }
}
