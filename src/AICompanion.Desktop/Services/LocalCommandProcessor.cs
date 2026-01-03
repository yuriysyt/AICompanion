using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using AICompanion.Desktop.Models;
using Microsoft.Extensions.Logging;

namespace AICompanion.Desktop.Services
{
    /*
        LocalCommandProcessor handles voice commands locally without AI.

        This service provides basic command recognition and execution for
        common tasks like opening applications, controlling windows, and
        system operations. It serves as a fallback when AI is unavailable
        and demonstrates the application's capabilities.
    */
    public class LocalCommandProcessor
    {
        private readonly ILogger<LocalCommandProcessor>? _logger;

        // P/Invoke for window focus management
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

        [DllImport("user32.dll")]
        private static extern int GetWindowThreadProcessId(IntPtr hWnd, out int processId);

        [DllImport("user32.dll")]
        private static extern bool AttachThreadInput(int idAttach, int idAttachTo, bool fAttach);

        [DllImport("kernel32.dll")]
        private static extern int GetCurrentThreadId();

        // Event for focus errors (UI can subscribe to show toast)
        public event EventHandler<string>? FocusError;
        public event EventHandler<string>? ActionExecuted;

        // Store the last target window (before AICompanion took focus)
        private IntPtr _lastTargetWindow = IntPtr.Zero;
        private string _lastTargetWindowTitle = "";

        // SMART MEMORY: Remember the last text editor window for easier dictation
        private IntPtr _lastTextEditorWindow = IntPtr.Zero;
        private string _lastTextEditorTitle = "";
        private string _lastTextEditorProcess = ""; // "WINWORD", "notepad", etc.

        public void CaptureTargetWindow()
        {
            // Call this BEFORE starting voice recognition to capture the window user wants to control
            var currentWindow = GetForegroundWindow();
            var title = GetActiveWindowTitle(currentWindow);

            // Don't capture AICompanion itself
            if (!title.Contains("AI Companion", StringComparison.OrdinalIgnoreCase))
            {
                _lastTargetWindow = currentWindow;
                _lastTargetWindowTitle = title;
                _logger?.LogInformation("[CAPTURE] Stored target window: '{Title}' (Handle: {Handle})", title, currentWindow);

                // SMART MEMORY: If this is a text editor, remember it for future dictation
                if (IsTextEditorWindow(title))
                {
                    _lastTextEditorWindow = currentWindow;
                    _lastTextEditorTitle = title;
                    _lastTextEditorProcess = DetectProcessType(title);
                    _logger?.LogInformation("[SMART] Remembered text editor: '{Title}' ({Process})", title, _lastTextEditorProcess);
                }
            }
        }

        private bool IsTextEditorWindow(string title)
        {
            var lower = title.ToLowerInvariant();
            return lower.Contains("word") ||
                   lower.Contains(".docx") ||
                   lower.Contains(".doc") ||
                   lower.Contains("notepad") ||
                   lower.Contains("блокнот") ||
                   lower.Contains(".txt") ||
                   lower.Contains("document") ||
                   lower.Contains("документ");
        }

        private string DetectProcessType(string title)
        {
            var lower = title.ToLowerInvariant();
            if (lower.Contains("word") || lower.Contains(".docx") || lower.Contains(".doc"))
                return "WINWORD";
            if (lower.Contains("notepad") || lower.Contains("блокнот") || lower.Contains(".txt"))
                return "notepad";
            return "unknown";
        }

        public IntPtr GetTargetWindow()
        {
            // Return the captured target window, or fall back to current foreground
            if (_lastTargetWindow != IntPtr.Zero)
            {
                // Verify window still exists
                var title = GetActiveWindowTitle(_lastTargetWindow);
                if (!string.IsNullOrEmpty(title) && title != "(none)")
                {
                    return _lastTargetWindow;
                }
            }
            return GetForegroundWindow();
        }

        public string GetTargetWindowTitle()
        {
            return !string.IsNullOrEmpty(_lastTargetWindowTitle) ? _lastTargetWindowTitle : "(none)";
        }

        /// <summary>
        /// Returns the last remembered text editor window title (Word, Notepad, etc.)
        /// This persists across voice commands for easier dictation.
        /// </summary>
        public string GetActiveTextEditorTitle()
        {
            if (_lastTextEditorWindow != IntPtr.Zero)
            {
                var title = GetActiveWindowTitle(_lastTextEditorWindow);
                if (!string.IsNullOrEmpty(title) && title != "(none)")
                {
                    return title;
                }
            }
            return "(no text editor)";
        }

        /// <summary>
        /// Clears the remembered text editor (user can say "forget editor" or similar)
        /// </summary>
        public void ClearTextEditorMemory()
        {
            _lastTextEditorWindow = IntPtr.Zero;
            _lastTextEditorTitle = "";
            _lastTextEditorProcess = "";
            _logger?.LogInformation("[SMART] Text editor memory cleared");
        }

        // Common application mappings with synonyms for fuzzy matching
        private readonly Dictionary<string, string> _appMappings = new(StringComparer.OrdinalIgnoreCase)
        {
            // Russian
            { "блокнот", "notepad" },
            { "калькулятор", "calc" },
            { "браузер", "msedge" },
            { "хром", "chrome" },
            { "проводник", "explorer" },
            { "ворд", "WINWORD" },
            { "эксель", "EXCEL" },
            { "пэинт", "mspaint" },
            { "пейнт", "mspaint" },
            { "командную строку", "cmd" },
            { "терминал", "wt" },
            { "настройки", "ms-settings:" },

            // English - primary names
            { "notepad", "notepad" },
            { "notebook", "notepad" },
            { "calculator", "calc" },
            { "browser", "msedge" },
            { "chrome", "chrome" },
            { "explorer", "explorer" },
            { "word", "WINWORD" },
            { "excel", "EXCEL" },
            { "paint", "mspaint" },
            { "command prompt", "cmd" },
            { "terminal", "wt" },
            { "settings", "ms-settings:" },
            { "edge", "msedge" },
            { "firefox", "firefox" },
            
            // Additional synonyms for better recognition
            { "note pad", "notepad" },
            { "note book", "notepad" },
            { "text editor", "notepad" },
            { "calc", "calc" },
            { "the browser", "msedge" },
            { "microsoft edge", "msedge" },
            { "google chrome", "chrome" },
            { "chrome browser", "chrome" },
            { "file explorer", "explorer" },
            { "files", "explorer" },
            { "microsoft word", "WINWORD" },
            { "ms word", "WINWORD" },
            { "microsoft excel", "EXCEL" },
            { "ms excel", "EXCEL" },
            { "spreadsheet", "EXCEL" },
            { "ms paint", "mspaint" },
            { "cmd", "cmd" },
            { "powershell", "powershell" },
            { "windows terminal", "wt" },
            
            // Additional synonyms for better Watson recognition
            { "the notepad", "notepad" },
            { "a notepad", "notepad" },
            { "my notepad", "notepad" },
            { "open notepad", "notepad" },
            { "the calculator", "calc" },
            { "a calculator", "calc" },
            { "the word", "WINWORD" },
            { "document", "WINWORD" },
            { "word document", "WINWORD" },
            { "the excel", "EXCEL" },
            { "spreadsheets", "EXCEL" },
            { "web browser", "msedge" },
            { "internet", "msedge" },
            { "the internet", "msedge" },
            { "my files", "explorer" },
            { "my documents", "explorer" },
            { "folder", "explorer" },
            { "folders", "explorer" },
        };

        // Phrases to strip from the beginning of commands for flexible matching
        private readonly string[] _prefixPhrases = new[]
        {
            "please", "can you", "could you", "would you", "i want to", "i'd like to",
            "i would like to", "kindly", "go ahead and", "just", "now",
            "пожалуйста", "можешь", "можешь ли ты", "хочу"
        };

        // Command patterns (Russian and English)
        private readonly List<(Regex Pattern, Func<Match, CommandResult> Handler)> _commandPatterns;

        public LocalCommandProcessor(ILogger<LocalCommandProcessor>? logger = null)
        {
            _logger = logger;
            _commandPatterns = InitializePatterns();
        }

        private List<(Regex Pattern, Func<Match, CommandResult> Handler)> InitializePatterns()
        {
            return new List<(Regex, Func<Match, CommandResult>)>
            {
                // Open application commands
                (new Regex(@"(?:открой|открыть|открывай|открывайте|запусти|запустить|запускай|open(?:s|ing)?|start(?:s|ing)?|launch(?:es|ing)?)\s+(.+)", RegexOptions.IgnoreCase),
                    m => ExecuteOpenApp(m.Groups[1].Value.Trim())),

                // Close commands
                (new Regex(@"(?:закрой|закрыть|close)\s+(?:это|this|окно|window)?", RegexOptions.IgnoreCase),
                    _ => ExecuteCloseWindow()),

                // Time command
                (new Regex(@"(?:который час|сколько времени|what time|current time|time)", RegexOptions.IgnoreCase),
                    _ => ExecuteGetTime()),

                // Date command
                (new Regex(@"(?:какой сегодня день|какая дата|what day|what date|today)", RegexOptions.IgnoreCase),
                    _ => ExecuteGetDate()),

                // Minimize command
                (new Regex(@"(?:сверни|свернуть|minimize)", RegexOptions.IgnoreCase),
                    _ => ExecuteMinimize()),

                // Help command
                (new Regex(@"(?:помощь|справка|help|что ты умеешь|what can you do)", RegexOptions.IgnoreCase),
                    _ => ExecuteHelp()),

                // Search command
                (new Regex(@"(?:найди|поиск|search|find)\s+(.+)", RegexOptions.IgnoreCase),
                    m => ExecuteSearch(m.Groups[1].Value.Trim())),

                // Hello/greeting
                (new Regex(@"(?:привет|здравствуй|hello|hi|hey)", RegexOptions.IgnoreCase),
                    _ => new CommandResult(true, "Hello! I'm your AI assistant. What can I help you with?",
                        "Hello! I'm ready to assist you.")),

                // ====== WINDOW-AWARE COMMANDS (Feature D) - MUST BE FIRST ======
                // Flexible patterns for speech recognition (handles "in a word", punctuation, quotes, etc.)

                // "in Word type..." - simplified and more robust
                (new Regex(@"in\s+(?:a\s+)?(?:microsoft\s+)?word[,\s]*(?:type|write|print)?[:\s,]*(.+)", RegexOptions.IgnoreCase),
                    m => ExecuteInTargetWindow("WINWORD", "Microsoft Word", CleanTextForTyping(m.Groups[1].Value))),

                // "in Notepad type..." - simplified
                (new Regex(@"in\s+(?:a\s+)?(?:the\s+)?notepad[,\s]*(?:type|write)?[:\s,]*(.+)", RegexOptions.IgnoreCase),
                    m => ExecuteInTargetWindow("notepad", "Notepad", CleanTextForTyping(m.Groups[1].Value))),

                // Russian: "в ворде напиши..."
                (new Regex(@"в\s+ворд[еу]?\s*(?:напиши|введи|напечатай)?[:\s]*(.+)", RegexOptions.IgnoreCase),
                    m => ExecuteInTargetWindow("WINWORD", "Microsoft Word", CleanTextForTyping(m.Groups[1].Value))),

                // Russian: "в блокноте напиши..."
                (new Regex(@"в\s+блокнот[еу]?\s*(?:напиши|введи)?[:\s]*(.+)", RegexOptions.IgnoreCase),
                    m => ExecuteInTargetWindow("notepad", "Notepad", CleanTextForTyping(m.Groups[1].Value))),

                // Simple: "word type hello"
                (new Regex(@"^(?:microsoft\s+)?word\s+(?:type|write)[:\s]+(.+)", RegexOptions.IgnoreCase),
                    m => ExecuteInTargetWindow("WINWORD", "Microsoft Word", CleanTextForTyping(m.Groups[1].Value))),

                // Simple: "notepad type hello"
                (new Regex(@"^notepad\s+(?:type|write)[:\s]+(.+)", RegexOptions.IgnoreCase),
                    m => ExecuteInTargetWindow("notepad", "Notepad", CleanTextForTyping(m.Groups[1].Value))),

                // "in Explorer delete..."
                (new Regex(@"in\s+(?:file\s+)?explorer\s+(?:delete|remove)[:\s]+(.+)", RegexOptions.IgnoreCase),
                    m => ExecuteDeleteInExplorer(m.Groups[1].Value.Trim())),

                // "in Chrome/Browser open..."
                (new Regex(@"in\s+(?:chrome|browser|edge)\s+(?:open|go to)[:\s]+(.+)", RegexOptions.IgnoreCase),
                    m => ExecuteOpenInBrowser(m.Groups[1].Value.Trim())),

                // ====== DOCUMENT INTERACTION COMMANDS (Feature A) ======
                
                // Type/dictate text into active application
                (new Regex(@"(?:type|write|напиши|введи|напечатай)[:\s]+(.+)", RegexOptions.IgnoreCase),
                    m => ExecuteTypeText(m.Groups[1].Value.Trim())),

                // Select all text
                (new Regex(@"(?:select all|выдели всё|выделить всё|select everything)", RegexOptions.IgnoreCase),
                    _ => ExecuteKeyboardShortcut("^a", "Selected all text", "I've selected all the text.")),

                // Copy text
                (new Regex(@"(?:copy|copy that|copy this|скопируй|копировать)", RegexOptions.IgnoreCase),
                    _ => ExecuteKeyboardShortcut("^c", "Copied to clipboard", "Copied to clipboard.")),

                // Paste text
                (new Regex(@"(?:paste|paste that|вставь|вставить)", RegexOptions.IgnoreCase),
                    _ => ExecuteKeyboardShortcut("^v", "Pasted from clipboard", "Pasted from clipboard.")),

                // Cut text
                (new Regex(@"(?:cut|cut that|вырежи|вырезать)", RegexOptions.IgnoreCase),
                    _ => ExecuteKeyboardShortcut("^x", "Cut to clipboard", "Cut to clipboard.")),

                // Undo last action
                (new Regex(@"(?:undo|undo that|отмени|отменить)", RegexOptions.IgnoreCase),
                    _ => ExecuteKeyboardShortcut("^z", "Undone", "I've undone the last action.")),

                // Redo action
                (new Regex(@"(?:redo|redo that|повтори|вернуть)", RegexOptions.IgnoreCase),
                    _ => ExecuteKeyboardShortcut("^y", "Redone", "I've redone the action.")),

                // Make text bold (Word, rich text editors)
                (new Regex(@"(?:make.*bold|bold|жирный|сделай жирным)", RegexOptions.IgnoreCase),
                    _ => ExecuteKeyboardShortcut("^b", "Applied bold formatting", "Made the text bold.")),

                // Make text italic
                (new Regex(@"(?:make.*italic|italic|курсив|сделай курсивом)", RegexOptions.IgnoreCase),
                    _ => ExecuteKeyboardShortcut("^i", "Applied italic formatting", "Made the text italic.")),

                // Underline text
                (new Regex(@"(?:make.*underline|underline|подчеркни|подчеркнуть)", RegexOptions.IgnoreCase),
                    _ => ExecuteKeyboardShortcut("^u", "Applied underline formatting", "Underlined the text.")),

                // Save document
                (new Regex(@"(?:save|save this|save document|сохрани|сохранить)", RegexOptions.IgnoreCase),
                    _ => ExecuteKeyboardShortcut("^s", "Document saved", "I've saved the document.")),

                // New line / Enter
                (new Regex(@"(?:new line|next line|enter|новая строка)", RegexOptions.IgnoreCase),
                    _ => ExecuteKeyboardShortcut("{ENTER}", "New line", "New line.")),

                // Delete selected text
                (new Regex(@"(?:delete|delete that|удали|удалить)", RegexOptions.IgnoreCase),
                    _ => ExecuteKeyboardShortcut("{DELETE}", "Deleted", "Deleted.")),

                // Select word (double-click simulation via Shift+Ctrl+Right)
                (new Regex(@"(?:select word|выдели слово)", RegexOptions.IgnoreCase),
                    _ => ExecuteKeyboardShortcut("^+{RIGHT}", "Selected word", "I've selected the word.")),

                // Go to start of document
                (new Regex(@"(?:go to start|go to beginning|в начало)", RegexOptions.IgnoreCase),
                    _ => ExecuteKeyboardShortcut("^{HOME}", "Moved to start", "Moved to the start of the document.")),

                // Go to end of document
                (new Regex(@"(?:go to end|в конец)", RegexOptions.IgnoreCase),
                    _ => ExecuteKeyboardShortcut("^{END}", "Moved to end", "Moved to the end of the document.")),

                // ====== TEACHING MODE COMMANDS (Feature C) ======
                
                // How do I queries
                (new Regex(@"(?:how do i|how can i|как мне|как я могу)\s+(.+)", RegexOptions.IgnoreCase),
                    m => ExecuteHowDoI(m.Groups[1].Value.Trim())),

                // What commands / list commands
                (new Regex(@"(?:what commands|list commands|show commands|какие команды|покажи команды)", RegexOptions.IgnoreCase),
                    _ => ExecuteShowCommands()),

                // Start tutorial command
                (new Regex(@"(?:start tutorial|begin tutorial|teach me|начать обучение|обучение)", RegexOptions.IgnoreCase),
                    _ => new CommandResult(true, "Starting tutorial", 
                        "TUTORIAL_START")),  // Special marker for UI to start tutorial

                // Stop tutorial command  
                (new Regex(@"(?:stop tutorial|end tutorial|exit tutorial|закончить обучение)", RegexOptions.IgnoreCase),
                    _ => new CommandResult(true, "Stopping tutorial",
                        "TUTORIAL_STOP")),  // Special marker for UI to stop tutorial

                // Skip tutorial step
                (new Regex(@"(?:skip|skip this|next step|пропустить)", RegexOptions.IgnoreCase),
                    _ => new CommandResult(true, "Skipping step",
                        "TUTORIAL_SKIP")),  // Special marker for UI to skip step

                // Request hint
                (new Regex(@"(?:give me a hint|hint|подсказка)", RegexOptions.IgnoreCase),
                    _ => new CommandResult(true, "Hint requested",
                        "TUTORIAL_HINT")),  // Special marker for UI to give hint

                // ====== FILE OPERATIONS (Feature B) ======
                
                // Open file dialog (Ctrl+O)
                (new Regex(@"(?:open file|open document|открыть файл)\s*(.+)?", RegexOptions.IgnoreCase),
                    m => ExecuteOpenFile(m.Groups[1].Value.Trim())),

                // Save As dialog (F12 or Ctrl+Shift+S)
                (new Regex(@"(?:save as|save file as|сохранить как)\s*(.+)?", RegexOptions.IgnoreCase),
                    m => ExecuteSaveAs(m.Groups[1].Value.Trim())),

                // Create new file/document (Ctrl+N)
                (new Regex(@"(?:new file|new document|create new|create file|создать файл|новый документ)", RegexOptions.IgnoreCase),
                    _ => ExecuteKeyboardShortcut("^n", "Created new document", "I've created a new document.")),

                // Close document (Ctrl+W)
                (new Regex(@"(?:close document|close file|закрыть документ|закрыть файл)", RegexOptions.IgnoreCase),
                    _ => ExecuteKeyboardShortcut("^w", "Closed document", "I've closed the document.")),

                // Print (Ctrl+P)
                (new Regex(@"(?:print|print this|печать|распечатать)", RegexOptions.IgnoreCase),
                    _ => ExecuteKeyboardShortcut("^p", "Opening print dialog", "Opening the print dialog.")),

                // ====== DIRECT TYPE TO CURRENT WINDOW ======
                // "type [text]" / "напиши [текст]" - types to current foreground window
                (new Regex(@"^(?:type|write|напиши|введи|печатай)[:\s]+(.+)", RegexOptions.IgnoreCase),
                    m => ExecuteTypeToCurrentWindow(m.Groups[1].Value.Trim())),
            };
        }

        public CommandResult ProcessCommand(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return new CommandResult(false, "Empty command", "");
            }

            _logger?.LogInformation("Processing local command: {Text}", text);

            // Strip common prefix phrases for more flexible matching
            // This allows "please open notepad" to work the same as "open notepad"
            var normalizedText = StripPrefixPhrases(text);
            _logger?.LogDebug("Normalized command: {NormalizedText}", normalizedText);

            // Try to match against known patterns with normalized text
            foreach (var (pattern, handler) in _commandPatterns)
            {
                var match = pattern.Match(normalizedText);
                if (match.Success)
                {
                    try
                    {
                        return handler(match);
                    }
                    catch (Exception ex)
                    {
                        _logger?.LogError(ex, "Error executing command");
                        return new CommandResult(false, $"Error: {ex.Message}", "");
                    }
                }
            }

            // Also try with original text in case normalization removed important parts
            if (normalizedText != text)
            {
                foreach (var (pattern, handler) in _commandPatterns)
                {
                    var match = pattern.Match(text);
                    if (match.Success)
                    {
                        try
                        {
                            return handler(match);
                        }
                        catch (Exception ex)
                        {
                            _logger?.LogError(ex, "Error executing command");
                            return new CommandResult(false, $"Error: {ex.Message}", "");
                        }
                    }
                }
            }

            // Unknown command
            return new CommandResult(false, $"Unknown command: {text}", "");
        }

        private string StripPrefixPhrases(string text)
        {
            var result = text.Trim();
            foreach (var prefix in _prefixPhrases)
            {
                if (result.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    result = result.Substring(prefix.Length).TrimStart();
                    break;
                }
            }
            return result;
        }

        private CommandResult ExecuteOpenApp(string appName)
        {
            string processName = appName;
            var cleanedAppName = appName.Trim().ToLowerInvariant();

            // Remove common articles and words that might be transcribed
            cleanedAppName = Regex.Replace(cleanedAppName, @"^(the|a|an)\s+", "");

            // First try exact match
            if (_appMappings.TryGetValue(cleanedAppName, out var mapped))
            {
                processName = mapped;
            }
            else
            {
                // Try fuzzy matching if exact match fails
                var bestMatch = FindBestAppMatch(cleanedAppName);
                if (bestMatch != null)
                {
                    processName = bestMatch;
                    _logger?.LogInformation("Fuzzy matched '{Input}' to '{Match}'", appName, processName);
                }
                else
                {
                    // Use original input as process name if no match found
                    processName = appName;
                }
            }

            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = processName,
                    UseShellExecute = true
                };

                Process.Start(psi);

                var friendlyName = GetFriendlyName(processName);
                return new CommandResult(true, $"Opened {friendlyName}",
                    $"Opening {friendlyName} for you!");
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Failed to open {App}", appName);
                return new CommandResult(false, $"Could not open {appName}",
                    $"Sorry, I couldn't find or open {appName}.");
            }
        }

        private string? FindBestAppMatch(string input)
        {
            // Find the application mapping with the smallest Levenshtein distance
            // Only accept matches with distance <= 4 (allows small typos and speech variations)
            const int maxDistance = 4;
            string? bestKey = null;
            int bestDistance = int.MaxValue;

            foreach (var kvp in _appMappings)
            {
                var distance = CalculateLevenshteinDistance(input, kvp.Key.ToLowerInvariant());
                if (distance < bestDistance && distance <= maxDistance)
                {
                    bestDistance = distance;
                    bestKey = kvp.Key;
                }
            }

            if (bestKey != null && _appMappings.TryGetValue(bestKey, out var processName))
            {
                return processName;
            }
            return null;
        }

        private static int CalculateLevenshteinDistance(string source, string target)
        {
            // Standard Levenshtein distance algorithm for fuzzy string matching
            // Returns the minimum number of single-character edits needed
            // to transform source into target
            if (string.IsNullOrEmpty(source))
                return string.IsNullOrEmpty(target) ? 0 : target.Length;
            if (string.IsNullOrEmpty(target))
                return source.Length;

            var sourceLength = source.Length;
            var targetLength = target.Length;
            var distanceMatrix = new int[sourceLength + 1, targetLength + 1];

            for (var i = 0; i <= sourceLength; i++)
                distanceMatrix[i, 0] = i;
            for (var j = 0; j <= targetLength; j++)
                distanceMatrix[0, j] = j;

            for (var i = 1; i <= sourceLength; i++)
            {
                for (var j = 1; j <= targetLength; j++)
                {
                    var cost = (target[j - 1] == source[i - 1]) ? 0 : 1;
                    distanceMatrix[i, j] = Math.Min(
                        Math.Min(
                            distanceMatrix[i - 1, j] + 1,      // deletion
                            distanceMatrix[i, j - 1] + 1),     // insertion
                        distanceMatrix[i - 1, j - 1] + cost);  // substitution
                }
            }

            return distanceMatrix[sourceLength, targetLength];
        }

        private CommandResult ExecuteCloseWindow()
        {
            try
            {
                // Send Alt+F4 to close current window
                System.Windows.Forms.SendKeys.SendWait("%{F4}");
                return new CommandResult(true, "Window closed", "Done!");
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Failed to close window");
                return new CommandResult(false, "Could not close window",
                    "I couldn't close the window.");
            }
        }

        private CommandResult ExecuteGetTime()
        {
            var time = DateTime.Now.ToString("HH:mm");
            return new CommandResult(true, $"Current time: {time}",
                $"The current time is {time}.");
        }

        private CommandResult ExecuteGetDate()
        {
            var date = DateTime.Now.ToString("dddd, MMMM d, yyyy");
            return new CommandResult(true, $"Today's date: {date}",
                $"Today is {date}.");
        }

        private CommandResult ExecuteMinimize()
        {
            try
            {
                // Minimize all windows
                var shell = new ProcessStartInfo
                {
                    FileName = "explorer.exe",
                    Arguments = "shell:::{3080F90D-D7AD-11D9-BD98-0000947B0257}",
                    UseShellExecute = true
                };
                Process.Start(shell);
                return new CommandResult(true, "Windows minimized", "Done!");
            }
            catch
            {
                return new CommandResult(false, "Could not minimize", "I couldn't minimize windows.");
            }
        }

        private CommandResult ExecuteSearch(string query)
        {
            try
            {
                var searchUrl = $"https://www.google.com/search?q={Uri.EscapeDataString(query)}";
                Process.Start(new ProcessStartInfo
                {
                    FileName = searchUrl,
                    UseShellExecute = true
                });
                return new CommandResult(true, $"Searching for: {query}",
                    $"Searching for '{query}' in your browser.");
            }
            catch
            {
                return new CommandResult(false, "Search failed",
                    "I couldn't open the search.");
            }
        }

        private CommandResult ExecuteHelp()
        {
            var helpText = @"Available commands:

📂 APPLICATION CONTROL:
• 'Open [app]' - Open notepad, calculator, browser, Word, Excel
• 'Close' - Close current window
• 'Minimize' - Minimize all windows

📝 DOCUMENT EDITING (works in Notepad, Word):
• 'Type: [text]' - Type text into active application
• 'Select all' - Select all text
• 'Copy' / 'Paste' / 'Cut' - Clipboard operations
• 'Undo' / 'Redo' - Undo or redo last action
• 'Bold' / 'Italic' / 'Underline' - Format selected text
• 'Save' - Save the document
• 'Delete' - Delete selected text

🎯 WINDOW-AWARE COMMANDS (NEW!):
• 'In Word write [text]' - Type directly into Word
• 'In Notepad write [text]' - Type directly into Notepad
• 'In browser open [url]' - Open URL in browser
• 'В ворде напиши [текст]' - Напишет в Word
• 'В блокноте напиши [текст]' - Напишет в Notepad

🔍 UTILITIES:
• 'Search [query]' - Search the web
• 'What time is it?' - Get current time
• 'What date?' - Get current date

❓ LEARNING:
• 'How do I [task]?' - Get help with a specific task
• 'What commands' - Show all available commands

Все команды работают и на русском языке!";

            return new CommandResult(true, "Help displayed", helpText);
        }

        private CommandResult ExecuteTypeText(string textToType)
        {
            // SMART TYPING: Find the best window to type into
            try
            {
                var targetWindow = FindBestTypingTarget();
                var windowTitle = GetActiveWindowTitle(targetWindow);
                _logger?.LogInformation("[TYPE] Smart target: '{Title}' (Handle: {Handle})", windowTitle, targetWindow);

                // If no valid window found, try to find or open a text editor
                if (targetWindow == IntPtr.Zero || windowTitle.Contains("AI Companion", StringComparison.OrdinalIgnoreCase))
                {
                    // Try to find any open text editor
                    targetWindow = FindAnyTextEditor();
                    windowTitle = GetActiveWindowTitle(targetWindow);

                    if (targetWindow == IntPtr.Zero)
                    {
                        // No text editor found - offer helpful message
                        var message = "No text editor found. Say 'open Word' or 'open Notepad' first, or click on a text window before speaking.";
                        _logger?.LogWarning(message);
                        FocusError?.Invoke(this, message);
                        return new CommandResult(false, "No text editor", message);
                    }
                }

                // Force focus on the TARGET window with retry
                if (!ForceFocusWindow(targetWindow))
                {
                    var error = $"Could not focus window: {windowTitle}";
                    _logger?.LogWarning(error);
                    FocusError?.Invoke(this, error);
                    return new CommandResult(false, "Focus failed", error);
                }

                // Increased delay after focus to ensure window is ready
                System.Threading.Thread.Sleep(300);

                // Use SendKeys to type the text
                SendKeys.SendWait(EscapeForSendKeys(textToType));

                // Update smart memory - this window is now our active text editor
                if (IsTextEditorWindow(windowTitle))
                {
                    _lastTextEditorWindow = targetWindow;
                    _lastTextEditorTitle = windowTitle;
                    _lastTextEditorProcess = DetectProcessType(windowTitle);
                }

                _logger?.LogInformation("[TYPE] Successfully typed to '{Title}': {Text}", windowTitle, textToType);
                ActionExecuted?.Invoke(this, $"✍️ Typed to {windowTitle}: {textToType}");

                return new CommandResult(true, $"✍️ Typed: {textToType}",
                    $"I've typed: {textToType}");
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Failed to type text: {Text}", textToType);
                FocusError?.Invoke(this, $"Error typing: {ex.Message}");
                return new CommandResult(false, "Could not type text",
                    "I couldn't type the text. Make sure a text field is focused.");
            }
        }

        private IntPtr FindBestTypingTarget()
        {
            // Priority 1: Use captured target window if it's a text editor
            if (_lastTargetWindow != IntPtr.Zero)
            {
                var title = GetActiveWindowTitle(_lastTargetWindow);
                if (!string.IsNullOrEmpty(title) && title != "(none)" && IsTextEditorWindow(title))
                {
                    _logger?.LogDebug("[SMART] Using captured target: {Title}", title);
                    return _lastTargetWindow;
                }
            }

            // Priority 2: Use remembered text editor if still open
            if (_lastTextEditorWindow != IntPtr.Zero)
            {
                var title = GetActiveWindowTitle(_lastTextEditorWindow);
                if (!string.IsNullOrEmpty(title) && title != "(none)" && !title.Contains("AI Companion"))
                {
                    _logger?.LogDebug("[SMART] Using remembered editor: {Title}", title);
                    return _lastTextEditorWindow;
                }
            }

            // Priority 3: Use any captured target window
            if (_lastTargetWindow != IntPtr.Zero)
            {
                var title = GetActiveWindowTitle(_lastTargetWindow);
                if (!string.IsNullOrEmpty(title) && title != "(none)" && !title.Contains("AI Companion"))
                {
                    return _lastTargetWindow;
                }
            }

            // Priority 4: Find any open text editor
            return FindAnyTextEditor();
        }

        private IntPtr FindAnyTextEditor()
        {
            // Search for Word first (more likely for document work)
            var wordWindow = FindWindowByProcess("WINWORD");
            if (wordWindow != IntPtr.Zero)
            {
                _logger?.LogDebug("[SMART] Found open Word window");
                return wordWindow;
            }

            // Then try Notepad
            var notepadWindow = FindWindowByProcess("notepad");
            if (notepadWindow != IntPtr.Zero)
            {
                _logger?.LogDebug("[SMART] Found open Notepad window");
                return notepadWindow;
            }

            // Try WordPad
            var wordpadWindow = FindWindowByProcess("wordpad");
            if (wordpadWindow != IntPtr.Zero)
            {
                _logger?.LogDebug("[SMART] Found open WordPad window");
                return wordpadWindow;
            }

            return IntPtr.Zero;
        }

        private bool ForceFocusWindow(IntPtr hWnd, int maxRetries = 3)
        {
            // Forcefully bring window to foreground with thread attachment
            for (int i = 0; i < maxRetries; i++)
            {
                try
                {
                    // Get thread IDs
                    int foregroundThreadId = GetWindowThreadProcessId(GetForegroundWindow(), out _);
                    int currentThreadId = GetCurrentThreadId();

                    // Attach threads to allow SetForegroundWindow to work
                    if (foregroundThreadId != currentThreadId)
                    {
                        AttachThreadInput(currentThreadId, foregroundThreadId, true);
                        SetForegroundWindow(hWnd);
                        AttachThreadInput(currentThreadId, foregroundThreadId, false);
                    }
                    else
                    {
                        SetForegroundWindow(hWnd);
                    }

                    System.Threading.Thread.Sleep(50);

                    // Verify focus was acquired
                    if (GetForegroundWindow() == hWnd)
                    {
                        _logger?.LogDebug("[FOCUS] Successfully focused window on attempt {Attempt}", i + 1);
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    _logger?.LogWarning(ex, "[FOCUS] Attempt {Attempt} failed", i + 1);
                }

                System.Threading.Thread.Sleep(100);
            }

            return false;
        }

        private string GetActiveWindowTitle(IntPtr hWnd)
        {
            if (hWnd == IntPtr.Zero) return "(none)";
            
            var sb = new StringBuilder(256);
            GetWindowText(hWnd, sb, 256);
            return sb.ToString();
        }

        public string GetCurrentActiveWindowInfo()
        {
            var hWnd = GetForegroundWindow();
            var title = GetActiveWindowTitle(hWnd);
            return $"Window: '{title}' (Handle: {hWnd})";
        }

        private static string EscapeForSendKeys(string text)
        {
            // SendKeys uses special characters that need escaping
            // + = Shift, ^ = Ctrl, % = Alt, {} = special keys
            var escaped = text
                .Replace("+", "{+}")
                .Replace("^", "{^}")
                .Replace("%", "{%}")
                .Replace("(", "{(}")
                .Replace(")", "{)}")
                .Replace("{", "{{}")
                .Replace("}", "{}}");
            return escaped;
        }

        private CommandResult ExecuteKeyboardShortcut(string keys, string description, string speechResponse)
        {
            // Execute a keyboard shortcut using SendKeys on TARGET window
            // keys format: ^ = Ctrl, + = Shift, % = Alt
            try
            {
                // Use captured target window instead of foreground
                var targetWindow = GetTargetWindow();
                var windowTitle = GetActiveWindowTitle(targetWindow);
                _logger?.LogInformation("[SHORTCUT] Executing '{Keys}' on target '{Title}'", keys, windowTitle);

                if (targetWindow == IntPtr.Zero || windowTitle.Contains("AI Companion", StringComparison.OrdinalIgnoreCase))
                {
                    var error = "No target window. Click on app first, then use voice command.";
                    FocusError?.Invoke(this, error);
                    return new CommandResult(false, "No target window", error);
                }

                // Force focus on TARGET window before shortcut
                if (!ForceFocusWindow(targetWindow))
                {
                    var error = $"Could not focus: {windowTitle}";
                    FocusError?.Invoke(this, error);
                    return new CommandResult(false, "Focus failed", error);
                }
                
                System.Threading.Thread.Sleep(150);

                SendKeys.SendWait(keys);
                
                _logger?.LogInformation("[SHORTCUT] Successfully executed: {Description} on {Title}", description, windowTitle);
                ActionExecuted?.Invoke(this, $"{description} → {windowTitle}");
                
                return new CommandResult(true, description, speechResponse);
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Failed to execute shortcut: {Keys}", keys);
                FocusError?.Invoke(this, $"Shortcut failed: {ex.Message}");
                return new CommandResult(false, $"Could not execute {description}",
                    "I couldn't perform that action.");
            }
        }

        private CommandResult ExecuteHowDoI(string task)
        {
            // Teaching mode: explain how to perform common tasks
            var taskLower = task.ToLowerInvariant();
            
            var explanations = new Dictionary<string, string>
            {
                { "open", "To open an application, say 'Open' followed by the app name. For example: 'Open Notepad', 'Open Calculator', or 'Open Word'." },
                { "type", "To type text, say 'Type:' followed by what you want to write. For example: 'Type: Hello world'. The text will appear in the active window." },
                { "copy", "To copy text, first select it by saying 'Select all', then say 'Copy'. The text is now in your clipboard." },
                { "paste", "To paste text, position your cursor where you want it and say 'Paste'. This inserts text from your clipboard." },
                { "format", "To format text, first select it, then say 'Bold', 'Italic', or 'Underline'. This works in Word and other rich text editors." },
                { "bold", "To make text bold, select the text first, then say 'Bold' or 'Make this bold'." },
                { "save", "To save a document, say 'Save'. This is the same as pressing Ctrl+S." },
                { "undo", "To undo your last action, say 'Undo'. You can undo multiple times." },
                { "search", "To search the web, say 'Search' followed by your query. For example: 'Search weather today'." },
                { "close", "To close the current window, say 'Close' or 'Close this window'." },
            };

            foreach (var kvp in explanations)
            {
                if (taskLower.Contains(kvp.Key))
                {
                    return new CommandResult(true, $"Help: {task}", kvp.Value);
                }
            }

            return new CommandResult(true, $"Help: {task}",
                $"I can help you open applications, type text, format documents, search the web, and more!");
        }

        private CommandResult ExecuteOpenFile(string fileName)
        {
            // Open file dialog and optionally type a filename
            try
            {
                var activeWindow = GetForegroundWindow();
                var windowTitle = GetActiveWindowTitle(activeWindow);
                _logger?.LogInformation("[OPEN FILE] Active window: '{Title}'", windowTitle);

                if (activeWindow == IntPtr.Zero)
                {
                    var error = "No active window. Open an application first.";
                    FocusError?.Invoke(this, error);
                    return new CommandResult(false, "No active window", error);
                }

                ForceFocusWindow(activeWindow);
                System.Threading.Thread.Sleep(100);

                // Send Ctrl+O to open file dialog
                SendKeys.SendWait("^o");
                System.Threading.Thread.Sleep(300);

                // If filename provided, type it into the dialog
                if (!string.IsNullOrWhiteSpace(fileName))
                {
                    System.Threading.Thread.Sleep(200);
                    SendKeys.SendWait(EscapeForSendKeys(fileName));
                    _logger?.LogInformation("[OPEN FILE] Typed filename: {FileName}", fileName);
                    ActionExecuted?.Invoke(this, $"Opening: {fileName}");
                    return new CommandResult(true, $"Opening file: {fileName}",
                        $"I've opened the file dialog and typed {fileName}. Press Enter to open.");
                }

                ActionExecuted?.Invoke(this, "Opened file dialog");
                return new CommandResult(true, "Opened file dialog",
                    "I've opened the file dialog. Say the filename or navigate to your file.");
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Failed to open file dialog");
                FocusError?.Invoke(this, $"Error: {ex.Message}");
                return new CommandResult(false, "Could not open file dialog",
                    "I couldn't open the file dialog.");
            }
        }

        private CommandResult ExecuteSaveAs(string fileName)
        {
            // Open Save As dialog and optionally type a filename
            try
            {
                var activeWindow = GetForegroundWindow();
                var windowTitle = GetActiveWindowTitle(activeWindow);
                _logger?.LogInformation("[SAVE AS] Active window: '{Title}'", windowTitle);

                if (activeWindow == IntPtr.Zero)
                {
                    var error = "No active window for Save As.";
                    FocusError?.Invoke(this, error);
                    return new CommandResult(false, "No active window", error);
                }

                ForceFocusWindow(activeWindow);
                System.Threading.Thread.Sleep(100);

                // F12 is the standard Save As shortcut in Office apps
                // Ctrl+Shift+S works in many other applications
                SendKeys.SendWait("{F12}");
                System.Threading.Thread.Sleep(300);

                // If filename provided, type it
                if (!string.IsNullOrWhiteSpace(fileName))
                {
                    System.Threading.Thread.Sleep(200);
                    SendKeys.SendWait(EscapeForSendKeys(fileName));
                    _logger?.LogInformation("[SAVE AS] Typed filename: {FileName}", fileName);
                    ActionExecuted?.Invoke(this, $"Save as: {fileName}");
                    return new CommandResult(true, $"Saving as: {fileName}",
                        $"I've opened Save As and typed {fileName}. Press Enter to save.");
                }

                ActionExecuted?.Invoke(this, "Opened Save As dialog");
                return new CommandResult(true, "Opened Save As dialog",
                    "I've opened the Save As dialog. Say a filename or choose a location.");
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Failed to open Save As dialog");
                FocusError?.Invoke(this, $"Error: {ex.Message}");
                return new CommandResult(false, "Could not open Save As dialog",
                    "I couldn't open the Save As dialog.");
            }
        }

        private CommandResult ExecuteShowCommands()
        {
            // Show a categorized list of all commands
            var commandList = @"Here are all the commands I understand:

APPLICATIONS: open, start, launch, close, minimize
TYPING: type, write (followed by your text)
EDITING: select all, copy, paste, cut, delete, undo, redo
FORMATTING: bold, italic, underline, save
NAVIGATION: go to start, go to end, new line
UTILITIES: search, time, date, help

Try saying 'How do I copy text?' for detailed instructions on any command.";

            return new CommandResult(true, "Command list", commandList);
        }

        private string GetFriendlyName(string processName)
        {
            var friendlyNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "notepad", "Notepad" },
                { "calc", "Calculator" },
                { "msedge", "Microsoft Edge" },
                { "chrome", "Google Chrome" },
                { "explorer", "File Explorer" },
                { "WINWORD", "Microsoft Word" },
                { "EXCEL", "Microsoft Excel" },
                { "mspaint", "Paint" },
                { "cmd", "Command Prompt" },
                { "wt", "Windows Terminal" },
            };

            return friendlyNames.TryGetValue(processName, out var friendly) ? friendly : processName;
        }

        // ====== WINDOW-AWARE COMMAND IMPLEMENTATIONS ======

        /// <summary>
        /// Type text directly to the current foreground window (where cursor is)
        /// </summary>
        private CommandResult ExecuteTypeToCurrentWindow(string textToType)
        {
            try
            {
                // Get current foreground window
                var foregroundWindow = GetForegroundWindow();
                
                if (foregroundWindow == IntPtr.Zero)
                {
                    return new CommandResult(false, "No active window",
                        "I couldn't find an active window to type into.");
                }

                // Get window title for feedback
                var titleBuilder = new System.Text.StringBuilder(256);
                GetWindowText(foregroundWindow, titleBuilder, 256);
                var windowTitle = titleBuilder.ToString();
                
                // Detect if it's a text editor
                var lowerTitle = windowTitle.ToLowerInvariant();
                var isTextEditor = lowerTitle.Contains("word") || 
                                   lowerTitle.Contains("notepad") || 
                                   lowerTitle.Contains("блокнот") ||
                                   lowerTitle.Contains(".txt") ||
                                   lowerTitle.Contains(".doc") ||
                                   lowerTitle.Contains("code") ||
                                   lowerTitle.Contains("sublime") ||
                                   lowerTitle.Contains("notepad++");

                if (!isTextEditor)
                {
                    _logger?.LogWarning("[TYPE] Window '{Title}' may not be a text editor", windowTitle);
                }

                // Focus the window and type
                SetForegroundWindow(foregroundWindow);
                System.Threading.Thread.Sleep(100);
                
                SendKeys.SendWait(EscapeForSendKeys(textToType));

                var shortTitle = windowTitle.Length > 30 ? windowTitle[..30] + "..." : windowTitle;
                _logger?.LogInformation("[TYPE] Typed to '{Window}': {Text}", shortTitle, textToType);
                ActionExecuted?.Invoke(this, $"Typed: {textToType}");

                return new CommandResult(true, $"Typed to {shortTitle}",
                    $"I've typed '{textToType}' in {shortTitle}.");
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "[TYPE] Error typing to current window");
                return new CommandResult(false, "Typing failed",
                    "Something went wrong while typing. Please try again.");
            }
        }

        /// <summary>
        /// Find a window by process name and type text into it
        /// </summary>
        private CommandResult ExecuteInTargetWindow(string processName, string friendlyName, string textToType)
        {
            try
            {
                // Resolve process name from aliases
                var resolvedProcess = ResolveProcessName(processName);
                _logger?.LogInformation("[WINDOW-CMD] Looking for {App} (process: {Process}) to type: {Text}",
                    friendlyName, resolvedProcess, textToType);

                // Find the window for this process
                var targetWindow = FindWindowByProcess(resolvedProcess);

                if (targetWindow == IntPtr.Zero)
                {
                    // Try to start the application if not found
                    _logger?.LogInformation("[WINDOW-CMD] {App} not found, starting it...", friendlyName);
                    try
                    {
                        Process.Start(new ProcessStartInfo
                        {
                            FileName = resolvedProcess,
                            UseShellExecute = true
                        });

                        // Wait for app to open with retries
                        for (int i = 0; i < 10; i++)
                        {
                            System.Threading.Thread.Sleep(500);
                            targetWindow = FindWindowByProcess(resolvedProcess);
                            if (targetWindow != IntPtr.Zero)
                            {
                                _logger?.LogInformation("[WINDOW-CMD] {App} opened after {Ms}ms", friendlyName, (i + 1) * 500);
                                break;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger?.LogError(ex, "[WINDOW-CMD] Failed to start {App}", friendlyName);
                        return new CommandResult(false, $"{friendlyName} not found",
                            $"I couldn't open {friendlyName}. Please open it manually.");
                    }
                }

                if (targetWindow == IntPtr.Zero)
                {
                    return new CommandResult(false, $"{friendlyName} not found",
                        $"I couldn't find or open {friendlyName}. Please try opening it manually.");
                }

                // SMART MEMORY: Remember this as our active text editor
                var windowTitle = GetActiveWindowTitle(targetWindow);
                if (IsTextEditorWindow(windowTitle) || resolvedProcess.Equals("WINWORD", StringComparison.OrdinalIgnoreCase)
                    || resolvedProcess.Equals("notepad", StringComparison.OrdinalIgnoreCase))
                {
                    _lastTextEditorWindow = targetWindow;
                    _lastTextEditorTitle = windowTitle;
                    _lastTextEditorProcess = resolvedProcess;
                    _logger?.LogInformation("[SMART] Updated active editor to: {Title}", windowTitle);
                }

                // Focus the window and type
                if (!ForceFocusWindow(targetWindow))
                {
                    return new CommandResult(false, $"Could not focus {friendlyName}",
                        $"I found {friendlyName} but couldn't switch to it.");
                }

                System.Threading.Thread.Sleep(300);
                SendKeys.SendWait(EscapeForSendKeys(textToType));

                _logger?.LogInformation("[WINDOW-CMD] Typed to {App}: {Text}", friendlyName, textToType);
                ActionExecuted?.Invoke(this, $"✍️ Typed in {friendlyName}: {textToType}");

                return new CommandResult(true, $"✍️ Typed in {friendlyName}: {textToType}",
                    $"I've typed '{textToType}' in {friendlyName}.");
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "[WINDOW-CMD] Error executing in {App}", friendlyName);
                return new CommandResult(false, $"Error in {friendlyName}",
                    $"Something went wrong while typing in {friendlyName}.");
            }
        }

        /// <summary>
        /// Delete a file in Explorer (requires security confirmation for dangerous ops)
        /// </summary>
        private CommandResult ExecuteDeleteInExplorer(string fileName)
        {
            try
            {
                // Find Explorer window
                var explorerWindow = FindWindowByProcess("explorer");
                
                if (explorerWindow == IntPtr.Zero)
                {
                    return new CommandResult(false, "Explorer not found",
                        "Please open File Explorer first, then try again.");
                }

                // This is a DANGEROUS operation - return special marker
                _logger?.LogWarning("[WINDOW-CMD] Delete requested for: {File} (requires confirmation)", fileName);
                
                return new CommandResult(false, $"Delete requires confirmation: {fileName}",
                    $"Deleting '{fileName}' is a dangerous operation. Please confirm by saying the security code, or say 'cancel' to abort.");
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "[WINDOW-CMD] Error in Explorer delete");
                return new CommandResult(false, "Delete failed",
                    "I couldn't perform the delete operation.");
            }
        }

        /// <summary>
        /// Open URL in browser
        /// </summary>
        private CommandResult ExecuteOpenInBrowser(string urlOrSearch)
        {
            try
            {
                string url;
                
                // Check if it's already a URL
                if (urlOrSearch.Contains(".") && !urlOrSearch.Contains(" "))
                {
                    // Looks like a URL
                    url = urlOrSearch.StartsWith("http") ? urlOrSearch : $"https://{urlOrSearch}";
                }
                else
                {
                    // Treat as search query
                    url = $"https://www.google.com/search?q={Uri.EscapeDataString(urlOrSearch)}";
                }

                Process.Start(new ProcessStartInfo
                {
                    FileName = url,
                    UseShellExecute = true
                });

                _logger?.LogInformation("[WINDOW-CMD] Opened in browser: {Url}", url);
                ActionExecuted?.Invoke(this, $"Opened: {url}");

                return new CommandResult(true, $"Opened in browser: {urlOrSearch}",
                    urlOrSearch.Contains(".") 
                        ? $"Opening {urlOrSearch} in your browser."
                        : $"Searching for '{urlOrSearch}' in your browser.");
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "[WINDOW-CMD] Error opening in browser");
                return new CommandResult(false, "Browser open failed",
                    "I couldn't open that in your browser.");
            }
        }

        /// <summary>
        /// Find window handle by process name
        /// </summary>
        private IntPtr FindWindowByProcess(string processName)
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
        }

        /// <summary>
        /// Clean text from speech recognition artifacts (quotes, punctuation, filler words)
        /// ElevenLabs often adds quotes, commas, or transcribes "type hello world" as "type "hello world""
        /// </summary>
        private string CleanTextForTyping(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return text;

            var cleaned = text.Trim();

            // Remove surrounding quotes of all types
            cleaned = cleaned.Trim('"', '\'', '"', '"', '«', '»');

            // Remove trailing punctuation that might be added by STT
            cleaned = cleaned.TrimEnd('.', ',', '!', '?', ';', ':');

            // Remove common filler words that STT might add
            var fillers = new[] { "please", "the text", "следующее", "текст" };
            foreach (var filler in fillers)
            {
                if (cleaned.StartsWith(filler + " ", StringComparison.OrdinalIgnoreCase))
                {
                    cleaned = cleaned.Substring(filler.Length + 1).TrimStart();
                }
            }

            _logger?.LogDebug("[CLEAN] '{Original}' -> '{Cleaned}'", text, cleaned);
            return cleaned;
        }

        /// <summary>
        /// Resolve process name from common aliases (handles speech recognition variations)
        /// </summary>
        private string ResolveProcessName(string input)
        {
            var aliases = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                // Word variations
                { "word", "WINWORD" },
                { "microsoft word", "WINWORD" },
                { "ms word", "WINWORD" },
                { "the word", "WINWORD" },
                { "ворд", "WINWORD" },
                { "ворде", "WINWORD" },
                { "ворду", "WINWORD" },
                { "документ", "WINWORD" },

                // Notepad variations
                { "notepad", "notepad" },
                { "note pad", "notepad" },
                { "блокнот", "notepad" },
                { "блокноте", "notepad" },
                { "блокноту", "notepad" },

                // Excel
                { "excel", "EXCEL" },
                { "эксель", "EXCEL" },

                // Explorer
                { "explorer", "explorer" },
                { "проводник", "explorer" },

                // Browsers
                { "chrome", "chrome" },
                { "хром", "chrome" },
                { "браузер", "msedge" },
                { "browser", "msedge" },
                { "edge", "msedge" },
            };

            return aliases.TryGetValue(input.Trim(), out var resolved) ? resolved : input;
        }
    }

    public class CommandResult
    {
        public bool Success { get; }
        public string Description { get; }
        public string SpeechResponse { get; }

        public CommandResult(bool success, string description, string speechResponse)
        {
            Success = success;
            Description = description;
            SpeechResponse = speechResponse;
        }
    }
}
